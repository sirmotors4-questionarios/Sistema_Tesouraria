/* Base mãe de Facturação/Recebimentos — motor (lógica pura, sem DOM nem rede).
 * Converte listagens Excel exportadas do PRIMAVERA (documentos de venda e recibos) em registos canónicos, compara cada
 * importação com a base mãe (Excel SharePoint) e faz a reconciliação interna: facturas ↔ recibos, saldos por cliente
 * e antiguidade. A base mãe é a fonte validada; o PRIMAVERA é a origem. Nada aqui grava: prepara o que o fluxo grava.
 * Navegador: window.TES_BASE; Node: module.exports (testes).
 */
(function (root, factory) {
  const api = factory(root.TES_CONCILIACAO || (typeof require === "function" ? require("./conciliacao-core.js") : null));
  if (typeof module === "object" && module.exports) module.exports = api;
  root.TES_BASE = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (C) {
  "use strict";
  if (!C) throw Error("faturacao-core.js precisa de conciliacao-core.js");
  const texto = (v) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
  const chave = (v) => texto(v).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const cent = (n) => Math.round(Number(n) * 100);
  const arred = (n) => Math.round(Number(n) * 100) / 100;

  // ---------- Tipos de documento ----------
  // efeito: +1 aumenta a dívida do cliente; -1 reduz. pago: liquidado no acto (factura-recibo, venda a dinheiro).
  const TIPOS = {
    FA: {nome: "Factura", efeito: 1}, FT: {nome: "Factura", efeito: 1}, FS: {nome: "Factura simplificada", efeito: 1, pago: true},
    FR: {nome: "Factura-recibo", efeito: 1, pago: true}, VD: {nome: "Venda a dinheiro", efeito: 1, pago: true},
    ND: {nome: "Nota de débito", efeito: 1}, NC: {nome: "Nota de crédito", efeito: -1},
    RE: {nome: "Recibo", recibo: true}, RC: {nome: "Recibo", recibo: true}, REC: {nome: "Recibo", recibo: true},
  };
  const infoTipo = (t) => TIPOS[texto(t).toUpperCase()] || null;

  // "FA 2026/123", "FA.2026.123", "FA 2026 / 123", "FA A/123", "NC 1/45" → {tipo, serie, numero}
  function partirDocumento(v) {
    const t = texto(v).toUpperCase();
    let m = t.match(/^([A-Z]{1,5})\s*[ .\-]\s*([A-Z0-9][A-Z0-9\-_]*)\s*[\/.]\s*0*(\d+)$/);
    if (m) return {tipo: m[1], serie: m[2], numero: m[3]};
    m = t.match(/^([A-Z]{1,5})\s+0*(\d+)$/);
    if (m) return {tipo: m[1], serie: "", numero: m[2]};
    return null;
  }
  const idDocumento = (empresa, tipo, serie, numero) =>
    [chave(empresa).toUpperCase(), texto(tipo).toUpperCase(), texto(serie).toUpperCase(), String(Number(numero) || texto(numero))].join("|");
  const nomeDocumento = (tipo, serie, numero) => texto(tipo).toUpperCase() + " " + (serie ? texto(serie).toUpperCase() + "/" : "") + numero;

  // ---------- Colunas das listagens PRIMAVERA ----------
  const SIN = {
    empresa: ["empresa", "cod empresa", "codigo empresa"],
    documento: ["documento", "doc", "doc.", "documento venda", "n documento", "no documento", "n.o documento", "documento completo", "recibo", "n recibo"],
    tipo: ["tipo doc", "tipodoc", "tipo documento", "tipo de documento", "tipo doc.", "tipo"],
    serie: ["serie", "serie doc"],
    numero: ["numero", "numdoc", "num doc", "n doc", "no doc", "nr documento", "numero documento", "num. doc.", "n.o"],
    data: ["data", "data doc", "data doc.", "data documento", "data emissao", "data de emissao", "data recibo"],
    vencimento: ["data venc", "data venc.", "data vencimento", "data de vencimento", "vencimento", "datavenc"],
    clienteCodigo: ["entidade", "cliente", "cod cliente", "cod. cliente", "codigo cliente", "cod. entidade", "codigo entidade", "cod entidade"],
    clienteNome: ["nome", "nome cliente", "nome entidade", "nome do cliente", "designacao", "razao social", "nome/designacao"],
    nuit: ["nuit", "nif", "contribuinte", "num contribuinte", "numcontrib", "n contribuinte", "no contribuinte", "nuit cliente"],
    moeda: ["moeda"],
    total: ["total", "total documento", "totaldoc", "total doc", "total doc.", "valor total", "total liquido", "total c/ iva", "total recibo", "valor recibo", "montante"],
    estado: ["estado", "anulado", "situacao", "anulada"],
    docLiquidado: ["documento liquidado", "doc liquidado", "doc. liquidado", "documento origem", "doc origem", "doc. origem", "factura", "fatura", "documento pago", "doc. pendente", "documento pendente", "doc liq."],
    valorLinha: ["valor liquidado", "valor pago", "valor recebido", "valor linha", "valor liq.", "valor regularizado", "valor"],
    modo: ["modo pagamento", "modo pag", "modo pag.", "forma pagamento", "forma de pagamento", "meio pagamento", "modo", "forma pag."],
    referencia: ["referencia", "ref", "ref.", "cheque", "n cheque", "referencia pagamento", "ref. pagamento"],
    conta: ["conta", "conta bancaria", "banco", "conta banco"],
  };
  const chaveCab = (h) => chave(h).replace(/[º°ª]/g, "o").replace(/[:*]/g, "").replace(/\s+/g, " ").trim();
  function campo(h) {
    const k = chaveCab(h);
    if (!k) return "";
    for (const [c, l] of Object.entries(SIN)) if (l.includes(k) || l.includes(k.replace(/\.$/, ""))) return c;
    return "";
  }
  // Cabeçalho: linha com documento (ou número), data e um valor.
  function cabecalho(matriz, tipo) {
    let melhor = null;
    for (let i = 0; i < Math.min(matriz.length, 40); i++) {
      const mapa = {};
      (matriz[i] || []).forEach((h, j) => { const c = campo(h); if (c && mapa[c] == null) mapa[c] = j; });
      // "Valor" sozinho num ficheiro de facturas é o total
      if (tipo === "faturas" && mapa.total == null && mapa.valorLinha != null) { mapa.total = mapa.valorLinha; delete mapa.valorLinha; }
      // Num recibo sem documento liquidado, "Valor" é o total do recibo
      if (tipo === "recibos" && mapa.docLiquidado == null && mapa.total == null && mapa.valorLinha != null) { mapa.total = mapa.valorLinha; delete mapa.valorLinha; }
      const ok = (mapa.documento != null || mapa.numero != null) && mapa.data != null && (mapa.total != null || mapa.valorLinha != null);
      const pontos = Object.keys(mapa).length;
      if (ok && (!melhor || pontos > melhor.pontos)) melhor = {linha: i, mapa, pontos};
    }
    return melhor;
  }

  const estadoAnulado = (v) => /anulad|cancelad|^sim$|^s$|^true$|^1$/.test(chave(v));
  function hash(obj, campos) {
    let h = 2166136261;
    const s = campos.map((c) => (typeof obj[c] === "number" ? cent(obj[c]) : texto(obj[c]))).join("\u001f");
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return (h >>> 0).toString(16).padStart(8, "0");
  }
  const CAMPOS_FATURA = ["Empresa", "Tipo_Documento", "Serie", "Numero", "Data", "Data_Vencimento", "Codigo_Cliente", "Cliente", "NUIT", "Moeda", "Total", "Estado"];
  const CAMPOS_RECIBO = ["Empresa", "Tipo_Documento", "Serie", "Numero", "Data", "Codigo_Cliente", "Cliente", "Valor", "Modo_Pagamento", "Referencia", "Estado", "Liquidacoes"];

  // ---------- Leitura das listagens ----------
  // Devolve {registos, rejeitados:[{linha, motivo}], mapa, linhaCabecalho, avisos}. empresa: a da importação quando o
  // ficheiro não tem coluna Empresa (as listagens do PRIMAVERA são por empresa).
  function lerListagem(matriz, {tipo, empresa = "", mapa: mapaManual, linhaCabecalho} = {}) {
    const cab = mapaManual ? {linha: linhaCabecalho ?? 0, mapa: mapaManual} : cabecalho(matriz, tipo);
    if (!cab) return {registos: [], rejeitados: [], mapa: null, linhaCabecalho: -1, avisos: ["Não foi encontrado o cabeçalho (Documento ou Número, Data e Total/Valor). Indique as colunas."]};
    const {mapa} = cab, cel = (l, c) => (mapa[c] == null ? "" : l[mapa[c]]);
    const linhas = [], rejeitados = [];
    for (let i = cab.linha + 1; i < matriz.length; i++) {
      const l = matriz[i] || [];
      if (!l.some((x) => texto(x))) continue;
      const nLinha = i + 1;
      let doc = partirDocumento(cel(l, "documento"));
      const tipoCol = texto(cel(l, "tipo")).toUpperCase(), serieCol = texto(cel(l, "serie")).toUpperCase(), numCol = texto(cel(l, "numero"));
      if (tipoCol || numCol) doc = {tipo: tipoCol || doc?.tipo || "", serie: serieCol || doc?.serie || "", numero: (numCol.replace(/^0+(?=\d)/, "") || doc?.numero || "")};
      const emp = texto(cel(l, "empresa")) || empresa;
      // Linhas de totais/subtotais das listagens
      if (/^(total|totais|subtotal|sub-total)\b/.test(chave(l.find((x) => texto(x)) || ""))) continue;
      if (!doc || !doc.tipo || !/^\d+$/.test(String(doc.numero))) { rejeitados.push({linha: nLinha, motivo: "documento não reconhecido: «" + texto(cel(l, "documento") || numCol) + "»"}); continue; }
      const info = infoTipo(doc.tipo);
      if (tipo === "faturas" && info?.recibo) { rejeitados.push({linha: nLinha, motivo: doc.tipo + " é um recibo: importe-o em Recebimentos"}); continue; }
      if (tipo === "recibos" && info && !info.recibo) { rejeitados.push({linha: nLinha, motivo: doc.tipo + " não é um recibo: importe-o em Facturação"}); continue; }
      const dataDoc = C.data(cel(l, "data"));
      if (!dataDoc) { rejeitados.push({linha: nLinha, motivo: "data inválida: «" + texto(cel(l, "data")) + "»"}); continue; }
      if (!emp) { rejeitados.push({linha: nLinha, motivo: "empresa em falta (indique a empresa da importação)"}); continue; }
      linhas.push({nLinha, l, doc, emp, dataDoc});
    }
    const registos = tipo === "recibos" ? montarRecibos(linhas, cel, rejeitados) : montarFaturas(linhas, cel, rejeitados);
    rejeitados.sort((a, b) => a.linha - b.linha);
    return {registos, rejeitados, mapa, linhaCabecalho: cab.linha, avisos: []};
  }
  function montarFaturas(linhas, cel, rejeitados) {
    const out = [], vistos = new Map();
    for (const {nLinha, l, doc, emp, dataDoc} of linhas) {
      const total = C.numero(cel(l, "total"));
      if (!Number.isFinite(total)) { rejeitados.push({linha: nLinha, motivo: "total inválido"}); continue; }
      const id = idDocumento(emp, doc.tipo, doc.serie, doc.numero);
      if (vistos.has(id)) { rejeitados.push({linha: nLinha, motivo: "documento repetido no ficheiro (linha " + vistos.get(id) + ")"}); continue; }
      vistos.set(id, nLinha);
      const r = {ID_Documento: id, Empresa: emp, Tipo_Documento: doc.tipo, Serie: doc.serie, Numero: Number(doc.numero),
        Documento: nomeDocumento(doc.tipo, doc.serie, doc.numero), Data: dataDoc, Data_Vencimento: C.data(cel(l, "vencimento")) || dataDoc,
        Codigo_Cliente: texto(cel(l, "clienteCodigo")), Cliente: texto(cel(l, "clienteNome")) || texto(cel(l, "clienteCodigo")),
        NUIT: texto(cel(l, "nuit")).replace(/\s/g, ""), Moeda: texto(cel(l, "moeda")).toUpperCase() || "MZN",
        Total: arred(Math.abs(total)), Estado: estadoAnulado(cel(l, "estado")) ? "Anulado" : "Normal", _linha: nLinha};
      if (!r.Codigo_Cliente && !r.Cliente) { rejeitados.push({linha: nLinha, motivo: "cliente em falta"}); vistos.delete(id); continue; }
      r.Hash = hash(r, CAMPOS_FATURA);
      out.push(r);
    }
    return out;
  }
  // Recibos: uma linha por documento liquidado (ou só o cabeçalho); agrupados pelo número do recibo.
  function montarRecibos(linhas, cel, rejeitados) {
    const porId = new Map();
    for (const {nLinha, l, doc, emp, dataDoc} of linhas) {
      const id = idDocumento(emp, doc.tipo, doc.serie, doc.numero);
      let r = porId.get(id);
      const total = C.numero(cel(l, "total"));
      if (!r) {
        r = {ID_Recibo: id, Empresa: emp, Tipo_Documento: doc.tipo, Serie: doc.serie, Numero: Number(doc.numero), Documento: nomeDocumento(doc.tipo, doc.serie, doc.numero),
          Data: dataDoc, Codigo_Cliente: texto(cel(l, "clienteCodigo")), Cliente: texto(cel(l, "clienteNome")) || texto(cel(l, "clienteCodigo")),
          Valor: Number.isFinite(total) ? arred(Math.abs(total)) : null, Modo_Pagamento: texto(cel(l, "modo")), Referencia: texto(cel(l, "referencia")),
          Conta_Bancaria: texto(cel(l, "conta")), Estado: estadoAnulado(cel(l, "estado")) ? "Anulado" : "Normal", linhas: [], _linha: nLinha};
        porId.set(id, r);
      }
      const liq = texto(cel(l, "docLiquidado"));
      if (liq) {
        const d = partirDocumento(liq);
        const v = C.numero(cel(l, "valorLinha"));
        if (!d) { rejeitados.push({linha: nLinha, motivo: "documento liquidado não reconhecido: «" + liq + "»"}); continue; }
        if (!Number.isFinite(v)) { rejeitados.push({linha: nLinha, motivo: "valor liquidado inválido"}); continue; }
        const idDoc = idDocumento(emp, d.tipo, d.serie, d.numero);
        const existente = r.linhas.find((x) => x.ID_Documento === idDoc);
        if (existente) existente.Valor = arred(existente.Valor + Math.abs(v));
        else r.linhas.push({ID_Liquidacao: id + ">" + idDoc, ID_Recibo: id, ID_Documento: idDoc, Documento_Liquidado: nomeDocumento(d.tipo, d.serie, d.numero), Valor: arred(Math.abs(v))});
      }
    }
    const out = [];
    for (const r of porId.values()) {
      const soma = arred(r.linhas.reduce((s, x) => s + x.Valor, 0));
      if (r.Valor == null) r.Valor = soma;
      if (!(r.Valor > 0)) { rejeitados.push({linha: r._linha, motivo: "recibo " + r.Documento + " sem valor"}); continue; }
      r.Liquidacoes = r.linhas.map((x) => x.ID_Documento + "=" + cent(x.Valor)).sort().join(";");
      r.Hash = hash(r, CAMPOS_RECIBO);
      out.push(r);
    }
    return out;
  }

  // ---------- Importação ↔ base mãe ----------
  // base: linhas da tabela (tbl_Faturas/tbl_Recibos) como vêm do Excel. Classifica cada registo importado e lista os
  // documentos da base que deviam estar no ficheiro (mesma empresa, tipos e período) e não estão.
  const COMPARAR = {
    faturas: [["Data", "data"], ["Data_Vencimento", "data"], ["Codigo_Cliente"], ["Cliente"], ["NUIT"], ["Moeda"], ["Total", "num"], ["Estado"]],
    recibos: [["Data", "data"], ["Codigo_Cliente"], ["Cliente"], ["Valor", "num"], ["Modo_Pagamento"], ["Referencia"], ["Estado"], ["Liquidacoes"]],
  };
  function normalizarBase(linha, tipo) {
    const r = {...linha};
    for (const [c, t] of COMPARAR[tipo]) {
      if (t === "data") r[c] = C.data(r[c]) || texto(r[c]);
      else if (t === "num") r[c] = arred(C.numero(r[c]));
      else r[c] = texto(r[c]);
    }
    return r;
  }
  function comparar(importados, base, tipo) {
    const idCampo = tipo === "recibos" ? "ID_Recibo" : "ID_Documento";
    const mapa = new Map(base.map((b) => [texto(b[idCampo]), normalizarBase(b, tipo)]));
    const novos = [], alterados = [], iguais = [];
    for (const r of importados) {
      const b = mapa.get(r[idCampo]);
      if (!b) { novos.push(r); continue; }
      if (b.Hash && b.Hash === r.Hash) { iguais.push(r); continue; }
      const mudancas = COMPARAR[tipo].filter(([c, t]) => (t === "num" ? cent(b[c]) !== cent(r[c]) : texto(b[c]) !== texto(r[c])))
        .map(([c]) => ({campo: c, antes: b[c], depois: r[c]}));
      if (mudancas.length) alterados.push({...r, _mudancas: mudancas}); else iguais.push(r);
    }
    // Ausentes: na base, mesma empresa e tipo, dentro do período do ficheiro, mas não no ficheiro
    const ids = new Set(importados.map((r) => r[idCampo]));
    const empresas = new Set(importados.map((r) => chave(r.Empresa))), tipos = new Set(importados.map((r) => r.Tipo_Documento));
    const datas = importados.map((r) => r.Data).sort();
    const ausentes = datas.length ? base.map((b) => normalizarBase(b, tipo)).filter((b) => !ids.has(texto(b[idCampo])) && empresas.has(chave(b.Empresa)) &&
      tipos.has(texto(b.Tipo_Documento).toUpperCase()) && b.Data >= datas[0] && b.Data <= datas[datas.length - 1] && b.Estado !== "Anulado") : [];
    return {novos, alterados, iguais, ausentes};
  }

  // Linhas a gravar (sem campos internos) para o fluxo. Recibos levam as liquidações à parte.
  function paraGravar(registos, tipo, idImportacao) {
    const agora = new Date().toISOString();
    if (tipo === "recibos") {
      return {
        linhas: registos.map(({linhas, _linha, _mudancas, ...r}) => ({...r, ID_Importacao: idImportacao, Data_Importacao: agora})),
        liquidacoes: registos.flatMap((r) => r.linhas.map((x) => ({...x, ID_Importacao: idImportacao}))),
      };
    }
    return {linhas: registos.map(({_linha, _mudancas, ...r}) => ({...r, ID_Importacao: idImportacao, Data_Importacao: agora}))};
  }

  // ---------- Reconciliação interna: facturas ↔ recibos ----------
  // faturas/recibos/liquidacoes: linhas da base mãe (ou importadas). hoje: "aaaa-mm-dd".
  function reconciliar({faturas = [], recibos = [], liquidacoes = [], hoje}) {
    const dia = hoje || new Date().toISOString().slice(0, 10);
    const F = faturas.map((f) => normalizarBase(f, "faturas")).filter((f) => f.Estado !== "Anulado");
    const R = recibos.map((r) => normalizarBase(r, "recibos")).filter((r) => r.Estado !== "Anulado");
    const recibosValidos = new Set(R.map((r) => texto(r.ID_Recibo)));
    const L = liquidacoes.filter((l) => recibosValidos.has(texto(l.ID_Recibo))).map((l) => ({...l, Valor: arred(C.numero(l.Valor))}));
    const porDoc = new Map(F.map((f) => [texto(f.ID_Documento), f]));
    const docRecibo = new Map(R.map((r) => [texto(r.ID_Recibo), texto(r.Documento)]));
    const liquidadoPorDoc = new Map(), excepcoes = [];
    for (const l of L) {
      const id = texto(l.ID_Documento);
      if (!porDoc.has(id)) excepcoes.push({tipo: "Liquidação de documento inexistente na base", documento: l.Documento_Liquidado || id, recibo: l.ID_Recibo, reciboDoc: docRecibo.get(texto(l.ID_Recibo)), valor: l.Valor, cliente: R.find((r) => texto(r.ID_Recibo) === texto(l.ID_Recibo))?.Cliente});
      liquidadoPorDoc.set(id, arred((liquidadoPorDoc.get(id) || 0) + l.Valor));
    }
    const comLinhas = new Set(L.map((l) => texto(l.ID_Recibo)));
    for (const r of R) {
      const soma = arred(L.filter((l) => texto(l.ID_Recibo) === texto(r.ID_Recibo)).reduce((s, l) => s + l.Valor, 0));
      if (!comLinhas.has(texto(r.ID_Recibo))) excepcoes.push({tipo: "Recibo sem documento liquidado (adiantamento ou por alocar)", documento: "—", recibo: r.ID_Recibo, reciboDoc: r.Documento, valor: r.Valor, cliente: r.Cliente});
      else if (cent(soma) !== cent(r.Valor)) excepcoes.push({tipo: "Soma das liquidações difere do valor do recibo", documento: "—", recibo: r.ID_Recibo, reciboDoc: r.Documento, valor: arred(r.Valor - soma), cliente: r.Cliente});
    }
    const documentos = F.map((f) => {
      const info = infoTipo(f.Tipo_Documento) || {efeito: 1};
      const liquidado = liquidadoPorDoc.get(texto(f.ID_Documento)) || 0;
      const total = f.Total;
      let saldo = info.efeito < 0 ? 0 : info.pago ? 0 : arred(total - liquidado);
      if (saldo < 0) { excepcoes.push({tipo: "Liquidado acima do total do documento", documento: f.Documento, valor: -saldo, cliente: f.Cliente}); saldo = 0; }
      const vencido = saldo > 0 && f.Data_Vencimento && f.Data_Vencimento < dia;
      const diasAtraso = vencido ? C.data(f.Data_Vencimento) && Math.round((Date.parse(dia) - Date.parse(f.Data_Vencimento)) / 86400000) : 0;
      const estado = info.efeito < 0 ? "Crédito" : saldo === 0 ? "Liquidado" : liquidado > 0 ? (vencido ? "Parcial vencido" : "Parcial") : vencido ? "Vencido" : "Em aberto";
      return {...f, efeito: info.efeito, liquidado: info.pago ? total : liquidado, saldo, estado, diasAtraso};
    });
    return {documentos, excepcoes, clientes: porCliente(documentos, R, L, dia)};
  }
  // Saldo por cliente e antiguidade (não vencido, 1–30, 31–60, 61–90, > 90 dias). Créditos: NC e recibos por alocar.
  const ESCALOES = [["naoVencido", "Não vencido"], ["d30", "1–30"], ["d60", "31–60"], ["d90", "61–90"], ["mais90", "> 90"]];
  function porCliente(documentos, recibos, liquidacoes, dia) {
    const mapa = new Map();
    const cli = (codigo, nome, empresa) => {
      const k = chave(empresa) + "|" + (chave(codigo) || chave(nome));
      if (!mapa.has(k)) mapa.set(k, {chave: k, Empresa: empresa, Codigo_Cliente: codigo, Cliente: nome, faturado: 0, emAberto: 0, creditos: 0,
        naoVencido: 0, d30: 0, d60: 0, d90: 0, mais90: 0, documentos: 0, ultimoRecibo: "", maiorAtraso: 0});
      const c = mapa.get(k); if (!c.Cliente && nome) c.Cliente = nome; return c;
    };
    for (const d of documentos) {
      const c = cli(d.Codigo_Cliente, d.Cliente, d.Empresa);
      if (d.efeito < 0) { c.creditos = arred(c.creditos + d.Total); continue; }
      c.faturado = arred(c.faturado + d.Total);
      if (d.saldo <= 0) continue;
      c.emAberto = arred(c.emAberto + d.saldo); c.documentos++;
      const k = !d.diasAtraso ? "naoVencido" : d.diasAtraso <= 30 ? "d30" : d.diasAtraso <= 60 ? "d60" : d.diasAtraso <= 90 ? "d90" : "mais90";
      c[k] = arred(c[k] + d.saldo); c.maiorAtraso = Math.max(c.maiorAtraso, d.diasAtraso || 0);
    }
    const alocado = new Map();
    for (const l of liquidacoes) alocado.set(texto(l.ID_Recibo), arred((alocado.get(texto(l.ID_Recibo)) || 0) + l.Valor));
    for (const r of recibos) {
      const c = cli(r.Codigo_Cliente, r.Cliente, r.Empresa);
      const livre = arred(r.Valor - (alocado.get(texto(r.ID_Recibo)) || 0));
      if (livre > 0) c.creditos = arred(c.creditos + livre);
      if (r.Data > c.ultimoRecibo) c.ultimoRecibo = r.Data;
    }
    return [...mapa.values()].map((c) => ({...c, saldoLiquido: arred(c.emAberto - c.creditos)})).sort((a, b) => b.emAberto - a.emAberto);
  }

  // Recibos como contrapartidas para os créditos do extracto (motor de conciliação bancária).
  function recibosParaBanco(recibos, {de = "", ate = "", toleranciaDias = 5, ligados = []} = {}) {
    const usados = new Set(ligados.map(texto));
    const margem = (iso, d) => (iso ? new Date(Date.parse(iso) + d * 86400000).toISOString().slice(0, 10) : "");
    const ini = margem(de, -toleranciaDias), fim = margem(ate, toleranciaDias);
    return recibos.map((r) => normalizarBase(r, "recibos")).filter((r) => r.Estado !== "Anulado" && r.Valor > 0 && !usados.has(texto(r.ID_Recibo)) &&
      (!ini || r.Data >= ini) && (!fim || r.Data <= fim) && !/numerario|dinheiro|cash/.test(chave(r.Modo_Pagamento)))
      .map((r) => ({id: "R:" + texto(r.ID_Recibo), idRecibo: texto(r.ID_Recibo), valor: r.Valor, data: r.Data, referencia: texto(r.Referencia),
        despacho: "", requisicao: texto(r.Documento), fornecedor: texto(r.Cliente), estado: "Recibo", tipo: "recibo"}));
  }

  return Object.freeze({TIPOS, ESCALOES, partirDocumento, idDocumento, campo, cabecalho, lerListagem, comparar, paraGravar,
    reconciliar, recibosParaBanco, hash, CAMPOS_FATURA, CAMPOS_RECIBO});
});
