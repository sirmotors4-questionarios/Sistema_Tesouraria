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
    // Vistos nas listagens reais (SIR / Metrobus, 2026)
    DV: {nome: "Devolução a dinheiro", efeito: -1}, VDS: {nome: "Venda a dinheiro (SchoolBus)", efeito: 1, pago: true},
    VDR: {nome: "Recibo de venda a dinheiro", recibo: true, liquida: ["VD", "VDS"]},
    NPC: {nome: "Pagamento a cliente (devolução)", recibo: true, saida: true, liquida: ["DV", "NC"]},
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
    totalMerc: ["total merc.", "total merc", "total mercadoria", "valor mercadoria", "mercadoria", "total iliquido"],
    totalIva: ["total iva", "iva", "valor iva"],
    desconto: ["total desc.", "total desc", "total desconto", "desconto", "valor desconto", "descontos"],
    totalOutros: ["total outros", "outros"],
    impostos: ["impostos", "outros impostos", "total impostos"],
    filial: ["filial"],
    motivo: ["motivo", "motivo estorno"],
    estorno: ["estorno", "estornado"],
    tipoOrig: ["tipo doc. original", "tipo doc original", "tipo documento original", "tipodoc original"],
    serieOrig: ["serie doc. original", "serie doc original", "serie documento original"],
    numeroOrig: ["numero doc. orig.", "numero doc. orig", "numero doc original", "numero doc. original", "numero documento original"],
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
  const CAMPOS_FATURA = ["Empresa", "Tipo_Documento", "Serie", "Numero", "Data", "Data_Vencimento", "Codigo_Cliente", "Cliente", "NUIT", "Moeda", "Total", "Estado",
    "Total_Mercadoria", "Total_IVA", "Total_Desconto", "Total_Outros", "Impostos", "Sinal", "Tipo_Nome"];
  const CAMPOS_RECIBO = ["Empresa", "Tipo_Documento", "Serie", "Numero", "Data", "Codigo_Cliente", "Cliente", "Valor", "Modo_Pagamento", "Referencia", "Estado", "Liquidacoes",
    "Valor_Desconto", "Filial", "Moeda", "Motivo", "Estorno", "Tipo_Nome"];
  // Códigos de cliente com a letra O no lugar de zeros (ex.: «CLOOO101» → «CL000101»)
  function corrigirCodigo(c) {
    const t = texto(c), m = t.match(/^(CL|FO|FR)([0-9Oo]+)$/i); // prefixos de entidades do PRIMAVERA
    return m && /[Oo]/.test(m[2]) && /\d/.test(m[2]) ? m[1].toUpperCase() + m[2].replace(/[Oo]/g, "0") : t;
  }
  // Empresa do cadastro a partir de um texto (cabeçalho do relatório ou nome do ficheiro): «METRO-BUS, SA» → «Metrobus»
  function empresaPorTexto(t, empresas) {
    const k = (x) => chave(x).replace(/\b(lda|limitada|sa|s\.a\.|sarl|su|e\.?i)\b/g, "").replace(/[^a-z0-9]/g, "");
    const alvo = k(t);
    if (!alvo) return "";
    const nomes = (empresas || []).map((e) => (typeof e === "string" ? e : e.nome || e.Nome_Empresa || "")).filter(Boolean);
    const exacto = nomes.find((n) => k(n) === alvo);
    if (exacto) return exacto;
    // Nome do cadastro no início do texto («SIR» em «SIR Comércio Internacional») ou contido («Metrobus» em «METRO-BUS, SA»)
    const parcial = nomes.filter((n) => k(n).length >= 3 && (alvo.startsWith(k(n)) || (k(n).length >= 4 && (alvo.includes(k(n)) || k(n).includes(alvo)))));
    parcial.sort((a, b) => k(b).length - k(a).length);
    return parcial.length && (parcial.length === 1 || k(parcial[0]).length > k(parcial[1]).length) ? parcial[0] : "";
  }
  // Empresa pelos clientes: códigos (e nomes) do ficheiro comparados com os clientes já facturados de cada empresa.
  function empresaPelosClientes(clientes, faturas) {
    const porEmp = new Map();
    for (const f of faturas || []) {
      const e = texto(f.Empresa); if (!e) continue;
      if (!porEmp.has(e)) porEmp.set(e, new Set());
      porEmp.get(e).add(chave(corrigirCodigo(f.Codigo_Cliente)) + "|" + chave(f.Cliente).replace(/[^a-z0-9]/g, "").slice(0, 12));
    }
    const unicos = [...new Set(clientes.map((c) => chave(corrigirCodigo(c.codigo)) + "|" + chave(c.nome).replace(/[^a-z0-9]/g, "").slice(0, 12)))];
    if (!unicos.length || !porEmp.size) return null;
    const pont = [...porEmp.entries()].map(([e, set]) => ({empresa: e, n: unicos.filter((u) => set.has(u)).length})).sort((a, b) => b.n - a.n);
    const [a, b] = pont;
    return a.n >= Math.max(3, unicos.length * 0.4) && (!b || a.n >= 2 * b.n) ? {empresa: a.empresa, comuns: a.n, total: unicos.length} : null;
  }

  // ---------- Leitura das listagens ----------
  // Devolve {registos, rejeitados:[{linha, motivo}], mapa, linhaCabecalho, avisos}. empresa: a da importação quando o
  // ficheiro não tem coluna Empresa (as listagens do PRIMAVERA são por empresa).
  // Metadados de relatórios PRIMAVERA (Crystal): empresa na 1.ª linha («SIR Comércio Internacional … | Pág. 1/12»),
  // título («Documentos Emitidos de Vendas») e moeda («Valores em MT»).
  function metadadosRelatorio(matriz, linhaCab) {
    const meta = {empresa: "", titulo: "", moeda: ""};
    for (const l of (matriz || []).slice(0, Math.max(0, linhaCab))) {
      const cels = (l || []).map(texto).filter(Boolean);
      if (!cels.length) continue;
      if (!meta.empresa && cels.some((c) => /p[aá]g\.?\s*\d+\s*\/\s*\d+/i.test(c))) meta.empresa = cels[0];
      const moeda = cels.map((c) => c.match(/valores em\s+([A-Z]{2,3})/i)).find(Boolean);
      if (moeda) meta.moeda = moeda[1].toUpperCase();
      if (!meta.titulo && cels[0] !== meta.empresa && !moeda) meta.titulo = cels[0];
      else if (!meta.titulo && moeda && cels[0] && !/valores em/i.test(cels[0])) meta.titulo = cels[0];
    }
    return meta;
  }
  function lerListagem(matriz, {tipo, empresa = "", mapa: mapaManual, linhaCabecalho, empresas = [], nomeFicheiro = "", faturasBase = []} = {}) {
    const cab = mapaManual ? {linha: linhaCabecalho ?? 0, mapa: mapaManual} : cabecalho(matriz, tipo);
    if (!cab) return {registos: [], rejeitados: [], mapa: null, linhaCabecalho: -1, avisos: ["Não foi encontrado o cabeçalho (Documento ou Número, Data e Total/Valor). Indique as colunas."], controlo: null, meta: {}};
    const {mapa} = cab, cel = (l, c) => (mapa[c] == null ? "" : l[mapa[c]]);
    const linhas = [], rejeitados = [], avisos = [], totaisRelatorio = [];
    // Empresa: coluna do ficheiro > escolhida no ecrã > cabeçalho do relatório > nome do ficheiro
    const meta = metadadosRelatorio(matriz, cab.linha);
    if (!empresa) {
      const doRelatorio = meta.empresa ? empresaPorTexto(meta.empresa, empresas) || meta.empresa : "";
      const doNome = nomeFicheiro ? empresaPorTexto(nomeFicheiro.replace(/\.[a-z0-9]+$/i, "").split(/[-–_()]/).map(texto).filter(Boolean).pop() || "", empresas) : "";
      empresa = doRelatorio || doNome;
      if (empresa) { meta.empresaDetectada = empresa; avisos.push("Empresa detectada " + (doRelatorio ? "no cabeçalho do relatório" : "no nome do ficheiro") + ": «" + empresa + "». Confirme."); }
      else if (mapa.empresa == null) {
        // Sem empresa no ficheiro: pelos clientes já facturados de cada empresa
        const clientes = matriz.slice(cab.linha + 1).filter((l) => texto(cel(l || [], "clienteCodigo")) || texto(cel(l || [], "clienteNome")))
          .map((l) => ({codigo: cel(l, "clienteCodigo"), nome: cel(l, "clienteNome")}));
        const sug = empresaPelosClientes(clientes, faturasBase);
        if (sug) {
          empresa = sug.empresa; meta.empresaDetectada = empresa;
          avisos.push("O ficheiro não indica a empresa. Pelos clientes (" + sug.comuns + " de " + sug.total + " já facturados por «" + sug.empresa + "»), é da «" + sug.empresa + "». Confirme.");
        } else {
          meta.precisaEmpresa = true;
          return {registos: [], rejeitados: [], mapa, linhaCabecalho: cab.linha, controlo: null, meta,
            avisos: ["Este ficheiro não indica a empresa (nem no conteúdo nem no nome). Escolha a empresa da listagem acima e os documentos são lidos."]};
        }
      }
    }
    const periodo = texto(nomeFicheiro).match(/(\d{4}-\d{2}-\d{2})\s*(?:at[eé]|a|-)\s*(\d{4}-\d{2}-\d{2})/i);
    if (periodo) meta.periodo = {de: periodo[1], ate: periodo[2]};
    for (let i = cab.linha + 1; i < matriz.length; i++) {
      const l = matriz[i] || [];
      if (!l.some((x) => texto(x))) continue;
      const nLinha = i + 1;
      let doc = partirDocumento(cel(l, "documento"));
      // Coluna «Documento» só com o tipo («RE», «VDR») e Série/Número noutras colunas
      const docSoTipo = /^[A-Z]{1,5}$/.test(texto(cel(l, "documento")).toUpperCase()) ? texto(cel(l, "documento")).toUpperCase() : "";
      const tipoCol = texto(cel(l, "tipo")).toUpperCase() || docSoTipo, serieCol = texto(cel(l, "serie")).toUpperCase(), numCol = texto(cel(l, "numero"));
      if (tipoCol || numCol) doc = {tipo: tipoCol || doc?.tipo || "", serie: serieCol || doc?.serie || "", numero: (numCol.replace(/^0+(?=\d)/, "") || doc?.numero || "")};
      const emp = texto(cel(l, "empresa")) || empresa;
      // Linhas de totais/subtotais das listagens: guardadas como totais de controlo («Total  Factura», «Total Geral»)
      const primeira = texto(l.find((x) => texto(x)) || "");
      if (/^(total|totais|subtotal|sub-total)\b/.test(chave(primeira))) {
        const v = C.numero(cel(l, "total"));
        if (Number.isFinite(v)) totaisRelatorio.push({rotulo: primeira.replace(/\s+/g, " "), valor: arred(v), linha: nLinha});
        continue;
      }
      if (/^©|primavera bss|p[aá]g\.\s*\d/i.test(primeira) || (chave(primeira) === chave(meta.empresa) && meta.empresa)) continue;
      if (!doc || !doc.tipo || !/^\d+$/.test(String(doc.numero))) { rejeitados.push({linha: nLinha, motivo: "documento não reconhecido: «" + texto(cel(l, "documento") || numCol) + "»"}); continue; }
      const info = infoTipo(doc.tipo);
      if (tipo === "faturas" && info?.recibo) { rejeitados.push({linha: nLinha, motivo: doc.tipo + " é um recibo: importe-o em Recebimentos"}); continue; }
      if (tipo === "recibos" && info && !info.recibo) { rejeitados.push({linha: nLinha, motivo: doc.tipo + " não é um recibo: importe-o em Facturação"}); continue; }
      const dataDoc = C.data(cel(l, "data"));
      if (!dataDoc) { rejeitados.push({linha: nLinha, motivo: "data inválida: «" + texto(cel(l, "data")) + "»"}); continue; }
      if (!emp) { rejeitados.push({linha: nLinha, motivo: "empresa em falta (indique a empresa da importação)"}); continue; }
      linhas.push({nLinha, l, doc, emp, dataDoc});
    }
    const registos = tipo === "recibos" ? montarRecibos(linhas, cel, rejeitados, avisos) : montarFaturas(linhas, cel, rejeitados, avisos);
    rejeitados.sort((a, b) => a.linha - b.linha);
    // Tipos desconhecidos: importados (efeito pelo sinal do valor), com aviso
    const desconhecidos = [...new Set(registos.map((r) => r.Tipo_Documento).filter((t) => !infoTipo(t)))];
    if (desconhecidos.length) avisos.push("Tipo(s) de documento sem regra própria: " + desconhecidos.join(", ") + " — importados; o efeito na conta do cliente segue o sinal do valor.");
    const controlo = conferirTotais(totaisRelatorio, registos, tipo);
    if (controlo && !controlo.ok) avisos.push("Os totais do relatório não batem com o importado: " + controlo.linhas.filter((x) => !x.confere).map((x) => x.rotulo).join("; ") + ". Veja os rejeitados.");
    return {registos, rejeitados, mapa, linhaCabecalho: cab.linha, avisos, controlo, meta};
  }
  // Totais de controlo do relatório («Total  Factura», «Total Nota Crédito», «Total Geral») contra a soma importada
  // (com sinal). Prova que nenhum documento ficou de fora.
  function conferirTotais(totais, registos, tipo) {
    if (!totais.length) return null;
    const campoValor = tipo === "recibos" ? "Valor" : "Total";
    const porTipo = new Map();
    for (const r of registos) {
      const v = cent(r[campoValor]) * (r.Sinal === -1 ? -1 : 1);
      const t = porTipo.get(r.Tipo_Documento) || {soma: 0, n: 0};
      t.soma += v; t.n++; porTipo.set(r.Tipo_Documento, t);
    }
    const geral = [...porTipo.values()].reduce((s, t) => s + t.soma, 0), usados = new Set();
    const linhas = totais.map((t) => {
      const alvo = cent(t.valor);
      if (/geral|total$/i.test(chave(t.rotulo).replace(/\s+/g, " ").trim()) && /geral/i.test(t.rotulo)) return {...t, tipo: "(todos)", soma: geral / 100, confere: alvo === geral};
      // Pelo nome do tipo («Factura» → FA) ou pelo valor
      const nome = chave(t.rotulo).replace(/^total\s+/, "");
      let cod = [...porTipo.keys()].find((k) => !usados.has(k) && infoTipo(k) && chave(infoTipo(k).nome).replace(/-/g, " ") === nome.replace(/-/g, " "));
      if (!cod) cod = [...porTipo.entries()].find(([k, v]) => !usados.has(k) && v.soma === alvo)?.[0];
      if (cod) usados.add(cod);
      const soma = cod ? porTipo.get(cod).soma : 0;
      return {...t, tipo: cod || "", n: cod ? porTipo.get(cod).n : 0, soma: soma / 100, confere: !!cod && soma === alvo};
    });
    return {linhas, ok: linhas.every((x) => x.confere), importado: geral / 100};
  }
  function montarFaturas(linhas, cel, rejeitados, avisos = []) {
    const out = [], vistos = new Map(), corrigidos = new Set();
    const num = (l, c) => { const v = C.numero(cel(l, c)); return Number.isFinite(v) ? arred(v) : ""; };
    for (const {nLinha, l, doc, emp, dataDoc} of linhas) {
      const total = C.numero(cel(l, "total"));
      if (!Number.isFinite(total)) { rejeitados.push({linha: nLinha, motivo: "total inválido"}); continue; }
      const id = idDocumento(emp, doc.tipo, doc.serie, doc.numero);
      if (vistos.has(id)) { rejeitados.push({linha: nLinha, motivo: "documento repetido no ficheiro (linha " + vistos.get(id) + ")"}); continue; }
      vistos.set(id, nLinha);
      const codigo = corrigirCodigo(cel(l, "clienteCodigo"));
      if (codigo !== texto(cel(l, "clienteCodigo"))) corrigidos.add(texto(cel(l, "clienteCodigo")) + " → " + codigo);
      const info = infoTipo(doc.tipo);
      const r = {ID_Documento: id, Empresa: emp, Tipo_Documento: doc.tipo, Serie: doc.serie, Numero: Number(doc.numero),
        Documento: nomeDocumento(doc.tipo, doc.serie, doc.numero), Data: dataDoc, Data_Vencimento: C.data(cel(l, "vencimento")) || dataDoc,
        Codigo_Cliente: codigo, Cliente: texto(cel(l, "clienteNome")) || codigo,
        NUIT: texto(cel(l, "nuit")).replace(/\s/g, ""), Moeda: texto(cel(l, "moeda")).toUpperCase() || "MZN",
        Total: arred(Math.abs(total)), Estado: estadoAnulado(cel(l, "estado")) ? "Anulado" : "Normal",
        Total_Mercadoria: num(l, "totalMerc"), Total_IVA: num(l, "totalIva"), Total_Desconto: num(l, "desconto"), Total_Outros: num(l, "totalOutros"), Impostos: num(l, "impostos"),
        // Sinal: +1 aumenta a dívida, -1 reduz. Do tipo conhecido; senão do sinal do valor no ficheiro.
        Sinal: info?.efeito ?? (total < 0 ? -1 : 1), Tipo_Nome: info?.nome || "", _linha: nLinha};
      if (!r.Codigo_Cliente && !r.Cliente) { rejeitados.push({linha: nLinha, motivo: "cliente em falta"}); vistos.delete(id); continue; }
      r.Hash = hash(r, CAMPOS_FATURA);
      out.push(r);
    }
    if (corrigidos.size) avisos.push("Código(s) de cliente com a letra O em vez de zero, corrigido(s): " + [...corrigidos].join(", ") + ".");
    return out;
  }
  // Recibos: uma linha por documento liquidado (ou só o cabeçalho); agrupados pelo número do recibo.
  function montarRecibos(linhas, cel, rejeitados, avisos = []) {
    const porId = new Map(), corrigidos = new Set();
    const num = (l, c) => { const v = C.numero(cel(l, c)); return Number.isFinite(v) ? arred(v) : ""; };
    for (const {nLinha, l, doc, emp, dataDoc} of linhas) {
      const id = idDocumento(emp, doc.tipo, doc.serie, doc.numero);
      let r = porId.get(id);
      const total = C.numero(cel(l, "total"));
      if (!r) {
        const codigo = corrigirCodigo(cel(l, "clienteCodigo"));
        if (codigo !== texto(cel(l, "clienteCodigo"))) corrigidos.add(texto(cel(l, "clienteCodigo")) + " → " + codigo);
        const estorno = texto(cel(l, "estorno"));
        r = {ID_Recibo: id, Empresa: emp, Tipo_Documento: doc.tipo, Serie: doc.serie, Numero: Number(doc.numero), Documento: nomeDocumento(doc.tipo, doc.serie, doc.numero),
          Data: dataDoc, Codigo_Cliente: codigo, Cliente: texto(cel(l, "clienteNome")) || codigo,
          Valor: Number.isFinite(total) ? arred(Math.abs(total)) : null, Modo_Pagamento: texto(cel(l, "modo")), Referencia: texto(cel(l, "referencia")),
          Conta_Bancaria: texto(cel(l, "conta")), Estado: estadoAnulado(cel(l, "estado")) ? "Anulado" : estorno ? "Estornado" : "Normal",
          Valor_Desconto: num(l, "desconto"), Filial: texto(cel(l, "filial")), Moeda: texto(cel(l, "moeda")).toUpperCase(), Motivo: texto(cel(l, "motivo")), Estorno: estorno,
          Tipo_Nome: infoTipo(doc.tipo)?.nome || "", linhas: [], _linha: nLinha};
        porId.set(id, r);
      }
      // Documento liquidado: coluna própria ou «Tipo/Série/Número Doc. Original» (o valor é o da linha ou o do recibo)
      let liq = texto(cel(l, "docLiquidado"));
      const tOrig = texto(cel(l, "tipoOrig")).toUpperCase(), nOrig = texto(cel(l, "numeroOrig"));
      if (!liq && tOrig && nOrig) liq = tOrig + " " + (texto(cel(l, "serieOrig")) ? texto(cel(l, "serieOrig")) + "/" : "") + nOrig;
      if (liq) {
        const d = partirDocumento(liq);
        const v = Number.isFinite(C.numero(cel(l, "valorLinha"))) ? C.numero(cel(l, "valorLinha")) : total;
        if (!d) { rejeitados.push({linha: nLinha, motivo: "documento liquidado não reconhecido: «" + liq + "»"}); continue; }
        if (!Number.isFinite(v)) { rejeitados.push({linha: nLinha, motivo: "valor liquidado inválido"}); continue; }
        const idDoc = idDocumento(emp, d.tipo, d.serie, d.numero);
        const existente = r.linhas.find((x) => x.ID_Documento === idDoc);
        if (existente) existente.Valor = arred(existente.Valor + Math.abs(v));
        else r.linhas.push({ID_Liquidacao: id + ">" + idDoc, ID_Recibo: id, ID_Documento: idDoc, Documento_Liquidado: nomeDocumento(d.tipo, d.serie, d.numero), Valor: arred(Math.abs(v))});
      }
    }
    const out = [], zero = [];
    for (const r of porId.values()) {
      const soma = arred(r.linhas.reduce((s, x) => s + x.Valor, 0));
      if (r.Valor == null) r.Valor = soma;
      if (!Number.isFinite(r.Valor)) { rejeitados.push({linha: r._linha, motivo: "recibo " + r.Documento + " sem valor"}); continue; }
      if (r.Valor === 0) zero.push(r.Documento);
      r.Liquidacoes = r.linhas.map((x) => x.ID_Documento + "=" + cent(x.Valor)).sort().join(";");
      r.Hash = hash(r, CAMPOS_RECIBO);
      out.push(r);
    }
    if (zero.length) avisos.push("Recibo(s) com valor 0, importados para não faltar nenhum nº: " + zero.join(", ") + ".");
    if (corrigidos.size) avisos.push("Código(s) de cliente com a letra O em vez de zero, corrigido(s): " + [...corrigidos].join(", ") + ".");
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
    const F = faturas.map((f) => normalizarBase(f, "faturas")).filter((f) => f.Estado !== "Anulado" && f.Estado !== "Estornado");
    const R = recibos.map((r) => normalizarBase(r, "recibos")).filter((r) => r.Estado !== "Anulado" && r.Estado !== "Estornado");
    const recibosValidos = new Set(R.map((r) => texto(r.ID_Recibo)));
    const L0 = liquidacoes.filter((l) => recibosValidos.has(texto(l.ID_Recibo))).map((l) => ({...l, Valor: arred(C.numero(l.Valor))}));
    // Exportações sem «Doc. Original»: ligações inferidas, só quando são únicas (marcadas como inferidas)
    const inferidas = inferirLiquidacoes(F, R, L0);
    const L = [...L0, ...inferidas];
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
      if (!comLinhas.has(texto(r.ID_Recibo))) excepcoes.push(infoTipo(r.Tipo_Documento)?.saida
        ? {tipo: "Pagamento a cliente sem devolução/nota de crédito correspondente", documento: "—", recibo: r.ID_Recibo, reciboDoc: r.Documento, valor: r.Valor, cliente: r.Cliente}
        : excepcaoPorAlocar(r, F, L));
      else if (cent(soma) !== cent(r.Valor)) excepcoes.push({tipo: "Soma das liquidações difere do valor do recibo", documento: "—", recibo: r.ID_Recibo, reciboDoc: r.Documento, valor: arred(r.Valor - soma), cliente: r.Cliente});
    }
    const documentos = F.map((f) => {
      const info = infoTipo(f.Tipo_Documento) || {efeito: Number(f.Sinal) === -1 ? -1 : 1};
      const liquidado = liquidadoPorDoc.get(texto(f.ID_Documento)) || 0;
      const total = f.Total;
      let saldo = info.efeito < 0 ? 0 : info.pago ? 0 : arred(total - liquidado);
      if (info.efeito < 0 && liquidado > total + 0.004) excepcoes.push({tipo: "Devolvido acima do valor da nota de crédito/devolução", documento: f.Documento, valor: arred(liquidado - total), cliente: f.Cliente});
      if (saldo < 0) { excepcoes.push({tipo: "Liquidado acima do total do documento", documento: f.Documento, valor: -saldo, cliente: f.Cliente}); saldo = 0; }
      const vencido = saldo > 0 && f.Data_Vencimento && f.Data_Vencimento < dia;
      const diasAtraso = vencido ? C.data(f.Data_Vencimento) && Math.round((Date.parse(dia) - Date.parse(f.Data_Vencimento)) / 86400000) : 0;
      const estado = info.efeito < 0 ? "Crédito" : saldo === 0 ? "Liquidado" : liquidado > 0 ? (vencido ? "Parcial vencido" : "Parcial") : vencido ? "Vencido" : "Em aberto";
      return {...f, efeito: info.efeito, liquidado: info.pago ? total : liquidado, devolvido: info.efeito < 0 ? liquidado : 0, saldo, estado, diasAtraso};
    });
    return {documentos, excepcoes, inferidas, clientes: porCliente(documentos, R, L, dia)};
  }
  // Recibo sem ligação: diz porquê e quais os documentos candidatos
  function excepcaoPorAlocar(r, F, L) {
    const k = (x) => chave(x.Empresa) + "|" + chave(corrigirCodigo(x.Codigo_Cliente) || x.Cliente);
    const liq = new Map();
    for (const l of L) liq.set(texto(l.ID_Documento), (liq.get(texto(l.ID_Documento)) || 0) + cent(l.Valor));
    const docs = F.filter((d) => k(d) === k(r) && efeitoDe(d) > 0 && !infoTipo(d.Tipo_Documento)?.pago);
    const abertos = docs.filter((d) => d.Data <= r.Data && cent(d.Total) - (liq.get(texto(d.ID_Documento)) || 0) > 0);
    const base = {documento: "—", recibo: r.ID_Recibo, reciboDoc: r.Documento, valor: r.Valor, cliente: r.Cliente};
    const cand = abertos.slice(0, 3).map((d) => d.Documento + " (" + ((cent(d.Total) - (liq.get(texto(d.ID_Documento)) || 0)) / 100).toFixed(2) + ")").join(", ");
    if (!docs.some((d) => d.Data <= r.Data)) return {...base, tipo: "Recibo de documentos anteriores ao período importado (ou adiantamento)"};
    if (abertos.some((d) => cent(d.Total) - (liq.get(texto(d.ID_Documento)) || 0) > cent(r.Valor))) return {...base, tipo: "Pagamento parcial por alocar", candidatos: cand};
    return {...base, tipo: "Recibo sem documento liquidado (por alocar)", candidatos: cand};
  }
  const efeitoDe = (f) => (infoTipo(f.Tipo_Documento) || {efeito: Number(f.Sinal) === -1 ? -1 : 1}).efeito;
  // Liga recibos sem documento liquidado:
  //  1) VDR → VD/VDS e NPC → DV/NC: mesmo cliente e valor, data mais próxima (no máximo 31 dias antes do recibo);
  //  2) RE (e outros): documento do cliente com saldo igual ao recibo (vários iguais: o mais antigo, para rever), pago
  //     sem o IVA, ou uma única combinação de até 4 documentos em aberto cuja soma é o recibo. Recibos por ordem de
  //     data; cada ligação consome o saldo.
  function inferirLiquidacoes(F, R, L0) {
    const comLinhas = new Set(L0.map((l) => texto(l.ID_Recibo)));
    const saldo = new Map(F.map((f) => [texto(f.ID_Documento), cent(f.Total)]));
    for (const l of L0) saldo.set(texto(l.ID_Documento), (saldo.get(texto(l.ID_Documento)) || 0) - cent(l.Valor));
    const cli = (x) => chave(x.Empresa) + "|" + chave(corrigirCodigo(x.Codigo_Cliente) || x.Cliente);
    const docsPorCliente = new Map();
    for (const f of [...F].sort((a, b) => a.Data.localeCompare(b.Data) || a.Numero - b.Numero)) {
      const k = cli(f); if (!docsPorCliente.has(k)) docsPorCliente.set(k, []); docsPorCliente.get(k).push(f);
    }
    const out = [];
    const ligar = (r, docs, regra) => {
      for (const d of docs) {
        const v = docs.length === 1 ? cent(r.Valor) : saldo.get(texto(d.ID_Documento));
        saldo.set(texto(d.ID_Documento), saldo.get(texto(d.ID_Documento)) - v);
        out.push({ID_Liquidacao: texto(r.ID_Recibo) + ">" + texto(d.ID_Documento), ID_Recibo: texto(r.ID_Recibo), ID_Documento: texto(d.ID_Documento),
          Documento_Liquidado: d.Documento, Valor: v / 100, inferida: true, regra});
      }
    };
    for (const r of [...R].filter((x) => !comLinhas.has(texto(x.ID_Recibo)) && cent(x.Valor) > 0).sort((a, b) => a.Data.localeCompare(b.Data))) {
      const info = infoTipo(r.Tipo_Documento) || {};
      const docs = (docsPorCliente.get(cli(r)) || []).filter((d) => d.Data <= r.Data);
      if (info.liquida) {
        const cand = docs.filter((d) => info.liquida.includes(d.Tipo_Documento) && cent(d.Total) === cent(r.Valor) && saldo.get(texto(d.ID_Documento)) > 0 &&
          (Date.parse(r.Data) - Date.parse(d.Data)) / 86400000 <= 31).sort((a, b) => b.Data.localeCompare(a.Data));
        if (cand.length) ligar(r, [cand[0]], r.Tipo_Documento + " → " + cand[0].Tipo_Documento + " (mesmo cliente e valor)");
        continue;
      }
      const abertos = docs.filter((d) => efeitoDe(d) > 0 && !infoTipo(d.Tipo_Documento)?.pago && saldo.get(texto(d.ID_Documento)) > 0);
      const alvo = cent(r.Valor);
      const iguais = abertos.filter((d) => saldo.get(texto(d.ID_Documento)) === alvo);
      if (iguais.length === 1) { ligar(r, iguais, "único documento do cliente com saldo igual ao recibo"); continue; }
      // Vários documentos com o mesmo saldo (ex.: rendas mensais iguais): o mais antigo primeiro (FIFO), marcado para revisão
      if (iguais.length > 1) { ligar(r, [iguais[0]], "o mais antigo de " + iguais.length + " documentos do cliente com o mesmo valor (rever se não for este)"); continue; }
      // Pago sem o IVA (IVA retido pelo cliente): um único documento com total − IVA = recibo; o IVA fica em aberto
      const semIva = abertos.filter((d) => C.numero(d.Total_IVA) > 0 && saldo.get(texto(d.ID_Documento)) === cent(d.Total) && cent(d.Total) - cent(C.numero(d.Total_IVA)) === alvo);
      if (semIva.length === 1) { ligar(r, semIva, "pago sem o IVA (" + arred(C.numero(semIva[0].Total_IVA)).toFixed(2) + " por regularizar: IVA retido pelo cliente?)"); continue; }
      // Combinação única de até 4 documentos (dos 20 mais antigos em aberto)
      const lista = abertos.slice(0, 20), sols = [];
      const procura = (i, soma, esc) => {
        if (sols.length > 1) return;
        if (soma === alvo && esc.length >= 2) { sols.push(esc.slice()); return; }
        if (soma >= alvo || esc.length >= 4 || i >= lista.length) return;
        esc.push(lista[i]); procura(i + 1, soma + saldo.get(texto(lista[i].ID_Documento)), esc); esc.pop();
        procura(i + 1, soma, esc);
      };
      procura(0, 0, []);
      if (sols.length === 1) ligar(r, sols[0], "única combinação de " + sols[0].length + " documentos em aberto que soma o recibo");
    }
    return out;
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
      // Crédito (NC/DV) ainda por devolver: o que um NPC já pagou ao cliente deixa de ser crédito
      if (d.efeito < 0) { c.creditos = arred(c.creditos + Math.max(0, d.Total - (d.devolvido || 0))); continue; }
      c.faturado = arred(c.faturado + d.Total);
      if (d.saldo <= 0) continue;
      c.emAberto = arred(c.emAberto + d.saldo); c.documentos++;
      const k = !d.diasAtraso ? "naoVencido" : d.diasAtraso <= 30 ? "d30" : d.diasAtraso <= 60 ? "d60" : d.diasAtraso <= 90 ? "d90" : "mais90";
      c[k] = arred(c[k] + d.saldo); c.maiorAtraso = Math.max(c.maiorAtraso, d.diasAtraso || 0);
    }
    const alocado = new Map();
    for (const l of liquidacoes) alocado.set(texto(l.ID_Recibo), arred((alocado.get(texto(l.ID_Recibo)) || 0) + l.Valor));
    for (const r of recibos) {
      if (infoTipo(r.Tipo_Documento)?.saida) continue; // pagamento ao cliente: não é crédito dele
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
    return recibos.map((r) => normalizarBase(r, "recibos")).filter((r) => r.Estado !== "Anulado" && r.Estado !== "Estornado" && r.Valor > 0 && !usados.has(texto(r.ID_Recibo)) &&
      !infoTipo(r.Tipo_Documento)?.saida &&
      (!ini || r.Data >= ini) && (!fim || r.Data <= fim) && !/numerario|dinheiro|cash/.test(chave(r.Modo_Pagamento)))
      .map((r) => ({id: "R:" + texto(r.ID_Recibo), idRecibo: texto(r.ID_Recibo), valor: r.Valor, data: r.Data, referencia: texto(r.Referencia),
        despacho: "", requisicao: texto(r.Documento), fornecedor: texto(r.Cliente), estado: "Recibo", tipo: "recibo"}));
  }

  return Object.freeze({TIPOS, ESCALOES, partirDocumento, idDocumento, campo, cabecalho, lerListagem, comparar, paraGravar,
    reconciliar, recibosParaBanco, hash, CAMPOS_FATURA, CAMPOS_RECIBO, empresaPorTexto, empresaPelosClientes, corrigirCodigo, conferirTotais});
});
