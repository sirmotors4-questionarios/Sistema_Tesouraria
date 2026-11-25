/* Conciliação bancária — motor (lógica pura, sem DOM nem rede).
 * Lê extractos já convertidos em linhas (Excel/CSV, PDF com texto ou OCR), valida-os e emparelha os movimentos
 * do banco com as saídas registadas na Tesouraria. Nada aqui grava no Excel: o resultado é revisto e exportado.
 * Funciona no navegador (window.TES_CONCILIACAO) e em Node (module.exports) para os testes.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.TES_CONCILIACAO = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  // ---------- Normalização ----------
  const texto = (v) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
  const chaveTexto = (v) =>
    texto(v).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const centimos = (n) => Math.round(Number(n) * 100);

  // Valores no formato de Moçambique/Portugal e dos bancos internacionais:
  // "1.234,56", "1 234,56", "1,234.56", "-1234.56", "(1 234,56)", "1.234,56-", "1.234,56 D", "1.234,56 CR", "MZN 1.234,56".
  function numero(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
    let s = texto(v).replace(/ /g, " ");
    if (!s || s === "-" || s === "—") return NaN;
    let sinal = 1;
    if (/^\(.*\)$/.test(s)) { sinal = -1; s = s.slice(1, -1); }
    const sufixo = s.match(/\s*(DR|CR|D|C)\.?$/i);
    if (sufixo) { if (/^D/i.test(sufixo[1])) sinal = -sinal; s = s.slice(0, sufixo.index); }
    s = s.replace(/^(MZN|MT|USD|EUR|ZAR|\$|€)\s*/i, "").replace(/\s*(MZN|MT|USD|EUR|ZAR)$/i, "");
    if (/-$/.test(s)) { sinal = -sinal; s = s.slice(0, -1); }
    if (/^-/.test(s)) { sinal = -sinal; s = s.slice(1); }
    if (/^\+/.test(s)) s = s.slice(1);
    s = s.replace(/[\s']/g, "");
    if (!/^[\d.,]+$/.test(s) || !/\d/.test(s)) return NaN;
    const ultPonto = s.lastIndexOf("."), ultVirgula = s.lastIndexOf(",");
    if (ultPonto >= 0 && ultVirgula >= 0) {
      s = ultVirgula > ultPonto ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    } else if (ultVirgula >= 0) {
      // "1,234" com 3 casas e um só separador é ambíguo; nos extractos em MZN a vírgula é decimal.
      const partes = s.split(",");
      s = partes.length > 2 ? partes.join("") : s.replace(",", ".");
    } else if (ultPonto >= 0) {
      const partes = s.split(".");
      if (partes.length > 2) s = partes.join("");
      else if (partes[1].length === 3 && partes[0].length <= 3 && partes[0] !== "0") s = partes.join(""); // "1.234" milhares
    }
    const n = Number(s);
    return Number.isFinite(n) ? sinal * n : NaN;
  }

  const MESES = {jan: 1, fev: 2, feb: 2, mar: 3, abr: 4, apr: 4, mai: 5, may: 5, jun: 6, jul: 7, ago: 8, aug: 8,
    set: 9, sep: 9, out: 10, oct: 10, nov: 11, dez: 12, dec: 12};
  // Datas: série Excel, ISO, dd/mm/aaaa, dd-mm-aa, dd.mm.aaaa, "05 Nov 2026", "05-NOV-26". Devolve "aaaa-mm-dd" ou "".
  function data(v, anoPorOmissao) {
    if (v == null || v === "") return "";
    if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
    const t = texto(v);
    const n = typeof v === "number" ? v : /^\d{5}(?:\.\d+)?$/.test(t) ? Number(t) : NaN;
    if (Number.isFinite(n) && n > 20000 && n < 80000)
      return new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000).toISOString().slice(0, 10);
    const valida = (a, m, d) => {
      if (a < 100) a += a < 70 ? 2000 : 1900;
      const dt = new Date(Date.UTC(a, m - 1, d));
      return dt.getUTCFullYear() === a && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d
        ? a + "-" + String(m).padStart(2, "0") + "-" + String(d).padStart(2, "0") : "";
    };
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ]|$)/);
    if (m) return valida(+m[1], +m[2], +m[3]);
    m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(?:\s|$)/);
    if (m) return valida(+m[3], +m[2], +m[1]);
    m = t.match(/^(\d{1,2})[\s.\-\/]*([a-zç]{3})[a-zç]*\.?[\s.\-\/]*(\d{2,4})?(?:\s|$)/i);
    if (m && MESES[chaveTexto(m[2]).slice(0, 3)]) {
      const ano = m[3] ? +m[3] : anoPorOmissao;
      return ano ? valida(ano, MESES[chaveTexto(m[2]).slice(0, 3)], +m[1]) : "";
    }
    m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})$/); // "05/11" sem ano (alguns PDF): usa o ano do extracto
    if (m && anoPorOmissao) return valida(anoPorOmissao, +m[2], +m[1]);
    return "";
  }
  const diasEntre = (a, b) => Math.round((Date.parse(a) - Date.parse(b)) / 86400000);

  // ---------- Colunas do extracto ----------
  const SINONIMOS = {
    data: ["data", "data mov", "data movimento", "data do movimento", "data operacao", "data lancamento", "data contabilistica",
      "dt mov", "date", "transaction date", "posting date", "booking date", "data transaccao", "data transacao",
      "tran date", "trans date", "txn date", "trn date", "trans. date"],
    dataValor: ["data valor", "data-valor", "dt valor", "value date", "valor data"],
    descricao: ["descricao", "descritivo", "descricao movimento", "descricao do movimento", "detalhes", "detalhe", "narrativa", "historico",
      "movimento", "description", "details", "narrative", "particulars", "transaction details", "operacao", "designacao"],
    referencia: ["referencia", "ref", "ref.", "n doc", "no doc", "n. doc", "documento", "n documento", "num doc", "cheque",
      "n cheque", "reference", "doc no", "document", "ref. banco", "referencia banco"],
    debito: ["debito", "debitos", "saidas", "saida", "levantamentos", "debit", "debits", "withdrawal", "withdrawals",
      "money out", "valor debito", "montante debito", "dr"],
    credito: ["credito", "creditos", "entradas", "entrada", "depositos", "credit", "credits", "deposit", "deposits",
      "money in", "valor credito", "montante credito", "cr"],
    montante: ["montante", "valor", "importancia", "amount", "quantia", "valor movimento", "montante movimento"],
    saldo: ["saldo", "saldo contabilistico", "saldo disponivel", "balance", "running balance", "saldo apos movimento", "saldo final"],
  };
  const CAMPOS = Object.keys(SINONIMOS);
  const chaveCabecalho = (h) => chaveTexto(String(h ?? "").replace(/_/g, " ")).replace(/[º°ª]/g, "").replace(/[:*]/g, "").replace(/\s+/g, " ").trim();
  function campoExacto(h) {
    const k = chaveCabecalho(h).replace(/\.$/, "");
    for (const c of CAMPOS) if (SINONIMOS[c].includes(k) || SINONIMOS[c].includes(k + ".")) return c;
    return "";
  }
  function campoDoCabecalho(h) {
    const k = chaveCabecalho(h);
    if (!k) return "";
    const exacto = campoExacto(h);
    if (exacto) return exacto;
    // Contém (ex.: "Débito (MZN)", "Saldo Contabilístico MZN")
    for (const c of ["dataValor", "debito", "credito", "saldo", "descricao", "referencia", "data", "montante"])
      if (SINONIMOS[c].some((s) => s.length > 3 && (k.startsWith(s + " ") || k.endsWith(" " + s) || k.includes(" " + s + " ") || k.startsWith(s + "(")))) return c;
    return "";
  }
  // Procura a linha de cabeçalho nas primeiras 40 linhas: a que reconhece mais campos (pelo menos data e um valor).
  function detectarCabecalho(matriz) {
    let melhor = null;
    const limite = Math.min(matriz.length, 40);
    for (let i = 0; i < limite; i++) {
      const linha = matriz[i] || [];
      const mapa = {};
      // Títulos exactos primeiro ("Descricao"), só depois os parciais ("Num_Operacao" não pode tirar o lugar ao descritivo)
      linha.forEach((h, j) => { const c = campoExacto(h); if (c && mapa[c] == null) mapa[c] = j; });
      const usadas = new Set(Object.values(mapa));
      linha.forEach((h, j) => { if (usadas.has(j)) return; const c = campoDoCabecalho(h); if (c && mapa[c] == null) mapa[c] = j; });
      const temValor = mapa.debito != null || mapa.credito != null || mapa.montante != null;
      const pontos = Object.keys(mapa).length;
      if ((mapa.data != null || mapa.dataValor != null) && temValor && (!melhor || pontos > melhor.pontos))
        melhor = {linha: i, mapa, pontos};
    }
    if (melhor && melhor.mapa.data == null) { melhor.mapa.data = melhor.mapa.dataValor; }
    if (melhor && melhor.mapa.saldo == null) {
      const cab = matriz[melhor.linha] || [], ult = Math.max(melhor.mapa.debito ?? -1, melhor.mapa.credito ?? -1, melhor.mapa.montante ?? -1);
      const j = ult + 1, amostra = matriz.slice(melhor.linha + 1, melhor.linha + 21).filter((l) => texto(l[j]));
      if (ult >= 0 && !texto(cab[j]) && amostra.length >= 2 && amostra.every((l) => Number.isFinite(numero(l[j])))) melhor.mapa.saldo = j;
    }
    return melhor;
  }

  // ---------- Matriz (Excel/CSV/PDF reconstruído) → movimentos ----------
  // mapa: {data, descricao, referencia, debito, credito, montante, saldo} → índice da coluna.
  // Débito sai como valor negativo; crédito positivo. Linhas de continuação (sem data, só texto) juntam-se à descrição anterior.
  function matrizParaMovimentos(matriz, opcoes = {}) {
    const cab = opcoes.mapa ? {linha: opcoes.linhaCabecalho ?? -1, mapa: opcoes.mapa} : detectarCabecalho(matriz);
    const avisos = [];
    if (!cab) return {movimentos: [], avisos: ["Não foi encontrado o cabeçalho (Data e Débito/Crédito ou Montante). Indique as colunas manualmente."], mapa: null, linhaCabecalho: -1, saldoInicial: null, saldoFinal: null};
    const {mapa} = cab;
    const anoRef = opcoes.ano;
    const movimentos = [];
    let saldoInicial = null, saldoFinal = null;
    const celula = (l, c) => (mapa[c] == null ? "" : l[mapa[c]]);
    for (let i = cab.linha + 1; i < matriz.length; i++) {
      const l = matriz[i] || [];
      if (!l.some((x) => texto(x))) continue;
      const linhaTexto = chaveTexto(l.join(" "));
      const d = data(celula(l, "data"), anoRef) || data(celula(l, "dataValor"), anoRef);
      let valor = NaN;
      const deb = numero(celula(l, "debito")), cred = numero(celula(l, "credito"));
      // Débito negativo é um estorno (entra dinheiro); crédito negativo, o inverso.
      if (Number.isFinite(deb) && deb !== 0) valor = -deb;
      if (Number.isFinite(cred) && cred !== 0) valor = Number.isFinite(valor) ? valor + cred : cred;
      if (!Number.isFinite(valor) && mapa.montante != null) valor = numero(celula(l, "montante"));
      const saldo = numero(celula(l, "saldo"));
      if (!d) {
        // Saldo inicial/final escrito numa linha própria
        if (/saldo (inicial|anterior|de abertura)|opening balance|balance brought forward|saldo transportado/.test(linhaTexto) && Number.isFinite(saldo)) saldoInicial = saldo;
        else if (/saldo (final|actual|atual|de fecho)|closing balance/.test(linhaTexto) && Number.isFinite(saldo)) saldoFinal = saldo;
        else if (movimentos.length && !Number.isFinite(valor) && texto(celula(l, "descricao")))
          movimentos[movimentos.length - 1].descricao = texto(movimentos[movimentos.length - 1].descricao + " " + celula(l, "descricao"));
        continue;
      }
      if (/saldo (inicial|anterior|de abertura)|opening balance|brought forward/.test(chaveTexto(celula(l, "descricao")) || linhaTexto) && (!Number.isFinite(valor) || /saldo (inicial|anterior|de abertura)|opening balance/.test(chaveTexto(celula(l, "descricao"))))) {
        saldoInicial = Number.isFinite(saldo) ? saldo : Number.isFinite(valor) ? valor : saldoInicial;
        continue;
      }
      if (!Number.isFinite(valor) || valor === 0) {
        if (Number.isFinite(saldo) && movimentos.length === 0) saldoInicial = saldo;
        continue;
      }
      movimentos.push({data: d, dataValor: data(celula(l, "dataValor"), anoRef) || "", descricao: texto(celula(l, "descricao")),
        referencia: texto(celula(l, "referencia")), valor, saldo: Number.isFinite(saldo) ? saldo : null, origem: "linha " + (i + 1),
        ...(Number.isFinite(saldo) && /\d[.,]\s*$/.test(texto(celula(l, "saldo"))) ? {saldoTruncado: true} : {})});
    }
    if (!movimentos.length) avisos.push("O cabeçalho foi encontrado, mas nenhuma linha tem data e valor válidos.");
    const res = {movimentos: [], avisos, mapa, linhaCabecalho: cab.linha, saldoInicial, saldoFinal};
    if (movimentos.some((m) => m.saldoTruncado)) {
      const aviso = reconstruirSaldosCortados(movimentos, res);
      if (aviso) avisos.push(aviso);
    }
    res.movimentos = finalizar(movimentos);
    return res;
  }

  // ---------- PDF (texto com posições, ou OCR) → matriz ----------
  // linhas: [{y, itens:[{x, w, str}]}] já agrupadas por y (de cima para baixo). Usa a linha de cabeçalho para definir
  // as colunas pelo centro de cada título; cada palavra vai para a coluna mais próxima.
  function linhasPdfParaMatriz(linhas) {
    let cabecalho = null;
    for (const l of linhas.slice(0, 80)) {
      const cols = [];
      // Títulos com várias palavras ("Data Valor", "Saldo Contabilístico"): juntar itens próximos
      const itens = [...l.itens].sort((a, b) => a.x - b.x);
      const grupos = [];
      for (const it of itens) {
        const u = grupos[grupos.length - 1];
        // Junta palavras próximas só quando formam um título conhecido ("Data Valor", "Saldo Contabilístico"),
        // ou quando o grupo ainda não é título; "Data Mov." seguido de "Data Valor" ficam separados.
        // Espaço entre palavras relativo à largura média de um carácter (PDF em pontos, OCR em píxeis)
        const largCar = u ? ((u.w / Math.max(1, u.str.length)) + (it.w / Math.max(1, it.str.length))) / 2 : 0;
        const junto = u && it.x - (u.x + u.w) < Math.max(4, largCar * 1.6) ? u.str + " " + it.str : null;
        if (junto && (campoExacto(junto) || !campoDoCabecalho(u.str))) { u.str = junto; u.w = it.x + it.w - u.x; } else grupos.push({...it});
      }
      for (const g of grupos) { const c = campoDoCabecalho(g.str); if (c) cols.push({campo: c, x: g.x, fim: g.x + g.w, titulo: g.str}); }
      const campos = new Set(cols.map((c) => c.campo));
      if ((campos.has("data") || campos.has("dataValor")) && (campos.has("debito") || campos.has("credito") || campos.has("montante")) && cols.length >= 3) {
        cabecalho = {linha: l, cols}; break;
      }
    }
    if (!cabecalho) return null;
    const cols = cabecalho.cols.sort((a, b) => a.x - b.x);
    const inferidas = inferirColunasNumericas(cols, linhas.slice(linhas.indexOf(cabecalho.linha) + 1));
    // Os números alinham à direita do título; o texto à esquerda. Usar o intervalo de cada título com folga.
    const NUMERICAS = ["debito", "credito", "montante", "saldo"];
    const haTexto = cols.some((c) => !NUMERICAS.includes(c.campo));
    const colunaDe = (it) => {
      const centro = it.x + it.w / 2, fim = it.x + it.w;
      let melhor = 0, dist = Infinity;
      const numerico = /^[\d\s.,()\-+]+(D|C|DR|CR)?$/i.test(it.str.trim());
      cols.forEach((c, j) => {
        // Texto nunca vai para uma coluna de valores (ex.: OCR junta «AUTHORISATION CODE DEBIT» num só título e a
        // última palavra de um descritivo longo ficava no débito: «INVEST 80,585.23»).
        if (!numerico && haTexto && NUMERICAS.includes(c.campo)) return;
        const alvo = numerico && ["debito", "credito", "montante", "saldo"].includes(c.campo) ? Math.abs(fim - c.fim) : Math.abs(centro - (c.x + c.fim) / 2);
        const dentro = it.x >= c.x - 4 && fim <= c.fim + 4 ? -1000 : 0;
        if (alvo + dentro < dist) { dist = alvo + dentro; melhor = j; }
      });
      return melhor;
    };
    const matriz = [cols.map((c) => c.titulo)];
    const inicio = linhas.indexOf(cabecalho.linha);
    for (const l of linhas.slice(inicio + 1)) {
      const linha = cols.map(() => "");
      for (const it of [...l.itens].sort((a, b) => a.x - b.x)) {
        const j = colunaDe(it);
        linha[j] = texto(linha[j] + " " + it.str);
      }
      matriz.push(linha);
    }
    const mapa = {};
    cols.forEach((c, j) => { if (mapa[c.campo] == null) mapa[c.campo] = j; });
    if (mapa.data == null) mapa.data = mapa.dataValor;
    return {matriz, mapa, inferidas};
  }
  // Valores monetários alinham à direita. Quando o PDF cortou títulos de colunas (impresso do browser com a página
  // estreita: «CREDIT» e «BALANCE» fora da margem), os valores continuam lá: agrupam-se os fins (x da direita) dos
  // números com casas decimais e cada grupo com apoio suficiente é uma coluna. As colunas numéricas com título ficam
  // com o grupo mais próximo; os grupos a mais, à direita, são as colunas em falta (crédito e/ou saldo, por esta ordem).
  const RE_DINHEIRO = /^\(?-?\d[\d.,\s']*[.,](\d{1,2})?\)?-?(\s?(D|C|DR|CR))?$/i;
  function inferirColunasNumericas(cols, dados) {
    const NUM = ["debito", "credito", "montante", "saldo"];
    const fins = [];
    let linhasComNumero = 0;
    for (const l of dados.slice(0, 400)) {
      const ns = l.itens.filter((it) => RE_DINHEIRO.test(String(it.str).trim()));
      if (ns.length) linhasComNumero++;
      for (const it of ns) fins.push(it.x + it.w);
    }
    if (linhasComNumero < 3) return [];
    fins.sort((a, b) => a - b);
    const grupos = [];
    for (const f of fins) {
      const u = grupos[grupos.length - 1];
      if (u && f - u.max <= 3) { u.max = f; u.n++; u.soma += f; } else grupos.push({min: f, max: f, n: 1, soma: f});
    }
    const fortes = grupos.filter((g) => g.n >= Math.max(3, linhasComNumero * 0.15)).map((g) => ({fim: g.soma / g.n}));
    const numericas = cols.filter((c) => NUM.includes(c.campo));
    if (!numericas.length || fortes.length <= numericas.length) return [];
    // As colunas com título ficam com os grupos mais à esquerda (pela ordem); os restantes são colunas em falta.
    numericas.sort((a, b) => a.x - b.x).forEach((c, k) => { c.fim = fortes[k].fim; c.x = Math.min(c.x, fortes[k].fim - 40); });
    const tem = new Set(cols.map((c) => c.campo));
    const faltam = ["credito", "saldo"].filter((c) => !tem.has(c) && !(c === "credito" && tem.has("montante")));
    const novas = [];
    fortes.slice(numericas.length).forEach((g, k) => {
      const campo = faltam[k];
      if (!campo) return;
      const c = {campo, x: g.fim - 40, fim: g.fim, titulo: (campo === "credito" ? "Crédito" : "Saldo") + " (deduzido)"};
      cols.push(c); novas.push(c.titulo);
    });
    cols.sort((a, b) => a.x - b.x);
    return novas;
  }
  // Saldos cortados no PDF («2,306,722.»: sem cêntimos). Os débitos e créditos estão completos, por isso o saldo exacto
  // é saldo inicial + soma dos movimentos. Cada saldo cortado limita os cêntimos do saldo inicial a um intervalo; a
  // intersecção de todos dá o valor exacto (com centenas de linhas fica um só valor). Devolve um aviso ou "".
  function reconstruirSaldosCortados(movimentos, resultado) {
    const cortados = movimentos.filter((m) => m.saldoTruncado).length;
    if (!cortados) return "";
    // Intervalo (em cêntimos) do saldo inicial compatível com cada saldo lido
    const intervalos = [];
    let soma = 0;
    for (const m of movimentos) {
      soma += centimos(m.valor);
      if (m.saldo == null) continue;
      const T = Math.round(m.saldo * 100);
      if (!m.saldoTruncado) intervalos.push([T - soma, T - soma, m]);
      else if (T >= 0) intervalos.push([T - soma, T + 99 - soma, m]);
      else intervalos.push([T - 99 - soma, T - soma, m]);
    }
    for (const m of movimentos) delete m.saldoTruncado;
    // Valor que satisfaz mais saldos (varrimento). Um saldo mal lido (ex.: OCR «544,52].») fica de fora e é assinalado,
    // em vez de invalidar todos.
    const ev = [];
    for (const [a, b] of intervalos) { ev.push([a, 1]); ev.push([b + 1, -1]); }
    ev.sort((x, y) => x[0] - y[0] || x[1] - y[1]);
    let n = 0, melhor = -1, lo = 0, hi = 0;
    for (let k = 0; k < ev.length; k++) {
      n += ev[k][1];
      if (ev[k][1] === 1 && n > melhor) { melhor = n; lo = ev[k][0]; hi = (ev[k + 1] ? ev[k + 1][0] : ev[k][0] + 1) - 1; }
    }
    const fora = intervalos.filter(([a, b]) => lo < a || lo > b).map(([, , m]) => m);
    if (melhor < Math.max(2, intervalos.length * 0.8)) {
      for (const m of movimentos) m.saldo = null;
      return cortados + " saldos vinham cortados no PDF (sem cêntimos) e não batem com os movimentos: os saldos foram ignorados. Confirme se faltam linhas.";
    }
    const inicial = lo; // com intervalo largo usa-se o mínimo (a diferença fica nos cêntimos)
    let s = inicial;
    for (const m of movimentos) { s += centimos(m.valor); m.saldo = s / 100; }
    resultado.saldoInicial = inicial / 100;
    resultado.saldoFinal = s / 100;
    const base = "O PDF cortou a coluna do saldo (sem cêntimos). Os saldos foram reconstruídos a partir dos débitos e créditos";
    const dif = fora.length ? " " + fora.length + " saldo(s) lido(s) não batem — provável erro de leitura, confirme: " + fora.slice(0, 5).map((m) => m.data + " " + Math.abs(m.valor).toFixed(2)).join("; ") + (fora.length > 5 ? "…" : "") + "." : "";
    return (lo === hi || hi - lo < 1
      ? base + " e batem com " + (intervalos.length - fora.length) + " de " + intervalos.length + " saldos do extracto."
      : base + "; os cêntimos do saldo inicial podem diferir até " + ((hi - lo) / 100).toFixed(2).replace(".", ",") + " MT.") + dif;
  }

  // Alternativa sem cabeçalho reconhecível (ex.: OCR imperfeito): cada linha que começa por uma data e termina em
  // valores. Com saldo em cada linha, o sinal vem da variação do saldo; sem saldo, de palavras (débito/crédito).
  const RE_VALOR = /\(?-?\d{1,3}(?:[ .,']\d{3})*(?:[.,]\d{2})\)?-?(?:\s?(?:D|C|DR|CR)\b)?|\(?-?\d+[.,]\d{2}\)?-?/g;
  function textoParaMovimentos(linhasTexto, opcoes = {}) {
    const movimentos = [];
    let saldoAnterior = null, saldoInicial = null;
    for (const bruto of linhasTexto) {
      const l = texto(bruto);
      const md = l.match(/^(\d{1,2}[\/.\-]\d{1,2}(?:[\/.\-]\d{2,4})?|\d{4}-\d{2}-\d{2}|\d{1,2}[\s\-][A-Za-zç]{3}[\s\-]\d{2,4})\s+/);
      if (!md) {
        const lt = chaveTexto(l);
        if (/saldo (inicial|anterior|de abertura)|opening balance|brought forward/.test(lt)) {
          const v = (l.match(RE_VALOR) || []).map(numero).filter(Number.isFinite);
          if (v.length) { saldoInicial = v[v.length - 1]; saldoAnterior = saldoInicial; }
        } else if (movimentos.length && l && !(l.match(RE_VALOR) || []).length) movimentos[movimentos.length - 1].descricao += " " + l;
        continue;
      }
      const d = data(md[1], opcoes.ano);
      if (!d) continue;
      let resto = l.slice(md[0].length);
      const md2 = resto.match(/^(\d{1,2}[\/.\-]\d{1,2}(?:[\/.\-]\d{2,4})?)\s+/); // data-valor
      if (md2) resto = resto.slice(md2[0].length);
      const valores = [];
      let m; RE_VALOR.lastIndex = 0;
      const encontrados = [];
      while ((m = RE_VALOR.exec(resto))) encontrados.push({str: m[0], idx: m.index});
      // Só os valores no fim da linha (os números na descrição, ex. referências, ficam no texto)
      let corte = resto.length;
      for (let k = encontrados.length - 1; k >= 0; k--) {
        const e = encontrados[k];
        if (texto(resto.slice(e.idx + e.str.length, corte)) !== "") break;
        valores.unshift(numero(e.str)); corte = e.idx;
      }
      const descricao = texto(resto.slice(0, corte));
      if (!valores.length || !valores.every(Number.isFinite)) continue;
      let valor, saldo = null;
      if (valores.length >= 2) { saldo = valores[valores.length - 1]; valor = valores[valores.length - 2]; }
      else valor = valores[0];
      if (saldo != null && saldoAnterior != null && Math.abs(Math.abs(saldo - saldoAnterior) - Math.abs(valor)) < 0.005)
        valor = Math.sign(saldo - saldoAnterior) * Math.abs(valor);
      else if (valor > 0 && /\b(debito|levantamento|pagamento|transferencia emitida|comissao|imposto|taxa)\b/.test(chaveTexto(descricao)) && !/\bcredito\b/.test(chaveTexto(descricao))) valor = -valor;
      if (saldo != null) saldoAnterior = saldo;
      if (saldoInicial == null && saldo != null && movimentos.length === 0) saldoInicial = Math.round((saldo - valor) * 100) / 100;
      movimentos.push({data: d, dataValor: md2 ? data(md2[1], opcoes.ano) : "", descricao, referencia: "", valor, saldo, origem: "texto"});
    }
    return {movimentos: finalizar(movimentos), saldoInicial, avisos: movimentos.length ? ["Sem cabeçalho reconhecido: os débitos/créditos foram deduzidos pela variação do saldo. Confirme os sinais."] : []};
  }

  function finalizar(movimentos) {
    return movimentos.map((m, i) => ({...m, id: "m" + (i + 1), chave: m.data + "|" + centimos(m.valor) + "|" + chaveTexto(m.descricao).slice(0, 40)}));
  }

  // ---------- Base mãe: identidade dos movimentos e organização dos extractos ----------
  // O mesmo movimento tem de ter a mesma chave venha de CSV, Excel ou PDF, e em extractos de períodos sobrepostos.
  // Por isso a chave NÃO usa descritivo, referência nem saldo (mudam entre formatos ou vêm cortados no export):
  //   BANCO|CONTA|AAAA-MM-DD|cêntimos|ocorrência
  // «ocorrência» numera, pela ordem do ficheiro, os movimentos iguais no mesmo dia (ex.: duas comissões de 167,99).
  // Extractos cortam o período por dias inteiros, por isso a numeração coincide entre ficheiros.
  const bancoCurto = (b) => chaveTexto(b).toUpperCase().replace(/[^A-Z0-9 ]/g, " ").trim().split(/\s+/)[0] || "BANCO";
  const contaCurta = (c) => texto(c).replace(/\D/g, "").replace(/^0+/, "");
  function idsMovimentos(movimentos, {banco = "", conta = ""} = {}) {
    const b = bancoCurto(banco), c = contaCurta(conta);
    if (!c) throw Error("Indique a conta bancária do extracto antes de gravar (o nº da conta faz parte da identidade dos movimentos).");
    const contagem = new Map();
    return movimentos.map((m) => {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(m.data || "")) throw Error("Movimento sem data válida: «" + texto(m.descricao).slice(0, 40) + "».");
      if (!Number.isFinite(Number(m.valor)) || centimos(m.valor) === 0) throw Error("Movimento sem valor em " + m.data + ".");
      const base = b + "|" + c + "|" + m.data + "|" + centimos(m.valor);
      const n = (contagem.get(base) || 0) + 1;
      contagem.set(base, n);
      return base + "|" + n;
    });
  }
  // BANCO_CONTA_AAAA-MM-DD_AAAA-MM-DD.ext — nome único e ordenável por conta e período.
  function nomePadronizado({banco = "", conta = "", de = "", ate = "", extensao = ""} = {}) {
    const ext = texto(extensao).replace(/^\./, "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const partes = [bancoCurto(banco), texto(conta).replace(/\D/g, "") || "SEM-CONTA", de || "inicio", ate || "fim"];
    return partes.join("_") + (ext ? "." + ext : "");
  }
  // Nº de conta (e banco/NIB, se houver) indicado no próprio ficheiro: cabeçalho do BCI («Conta;2308710007»),
  // coluna «Conta» do formato normalizado ou texto do PDF («Conta nº 0009102002651»).
  function contaDoFicheiro(matriz, textos) {
    const r = {conta: "", banco: "", nib: ""};
    const linhas = Array.isArray(matriz) ? matriz.slice(0, 40) : [];
    for (let i = 0; i < linhas.length && !(r.conta && r.nib); i++) {
      const l = linhas[i] || [];
      const k = l.map((x) => chaveTexto(x).replace(/[^a-z ]/g, "").trim());
      const j = r.conta ? -1 : k.indexOf("conta");
      if (j >= 0 && l.slice(j + 1).some((x) => /\d{6,}/.test(texto(x).replace(/[\s.\-]/g, "")))) {
        r.conta = texto(l.slice(j + 1).find((x) => /\d{6,}/.test(texto(x).replace(/[\s.\-]/g, "")))).replace(/[\s.\-]/g, "");
      } else if (j >= 0 && linhas[i + 1]) {
        const v = texto(linhas[i + 1][j]).replace(/[\s.\-]/g, "");
        if (/^\d{6,}$/.test(v)) {
          r.conta = v;
          const jb = k.indexOf("banco"), jn = k.indexOf("nib");
          if (jb >= 0) r.banco = texto(linhas[i + 1][jb]);
          if (jn >= 0) r.nib = texto(linhas[i + 1][jn]).replace(/\D/g, "");
        }
      }
      const jn = k.indexOf("nib");
      if (!r.nib && jn >= 0) { const v = l.slice(jn + 1).map((x) => texto(x).replace(/\D/g, "")).find((x) => x.length >= 15); if (v) r.nib = v; }
    }
    if (Array.isArray(textos)) {
      for (const t of textos.slice(0, 80)) {
        const l = texto(t);
        // Rótulo explícito («Account Number : 3368», «Nº de conta: 123»): aceita números curtos
        const e = !r.conta && l.match(/\b(?:account\s*(?:number|no\.?|n[ºo°])|n[.º°o]*\s*(?:de\s*)?conta|numero\s*(?:de\s*)?conta)\s*[:\-]?\s*(\d[\d .\-]{1,}\d)/i);
        if (e) r.conta = e[1].replace(/[\s.\-]/g, "");
        const m = !r.conta && l.match(/\b(?:conta|account)\b\s*(?:n[.º°o]*|number|no\.?)?\s*[:\-]?\s*(\d[\d .\-]{5,}\d)/i);
        if (m) r.conta = m[1].replace(/[\s.\-]/g, "");
        // IBAN moçambicano: MZ + 2 dígitos de controlo + NIB de 21 dígitos
        const ib = !r.nib && l.replace(/\s/g, "").match(/MZ\d{2}(\d{21})/i);
        if (ib) r.nib = ib[1];
        const nib = !r.nib && l.match(/\bNIB\b\s*[:\-]?\s*([\d ]{21,27})/i);
        if (nib && nib[1].replace(/\D/g, "").length === 21) r.nib = nib[1].replace(/\D/g, "");
        if (r.conta && r.nib) break;
      }
    }
    return r;
  }
  const mesmaConta = (a, b) => { const x = contaCurta(a), y = contaCurta(b); return !!x && !!y && (x === y || x.endsWith(y) || y.endsWith(x)); };
  // Linhas para tbl_MovimentosBancarios (colunas da base mãe).
  function movimentosParaBase(movimentos, ids, {idConta = "", banco = "", conta = "", idExtrato = "", dataImportacao = ""} = {}) {
    return movimentos.map((m, i) => ({
      ID_Movimento: ids[i], ID_Conta: idConta, Banco: banco, Conta: texto(conta).replace(/\D/g, ""),
      Data_Movimento: m.data, Data_Valor: m.dataValor || "", Referencia: texto(m.referencia).slice(0, 80), Descricao: texto(m.descricao).slice(0, 250),
      Debito: m.valor < 0 ? Math.round(-m.valor * 100) / 100 : 0, Credito: m.valor > 0 ? Math.round(m.valor * 100) / 100 : 0,
      Valor: Math.round(m.valor * 100) / 100, Saldo: m.saldo == null || !Number.isFinite(m.saldo) ? "" : Math.round(m.saldo * 100) / 100,
      ID_Extrato: idExtrato, Data_Importacao: dataImportacao,
    }));
  }
  // Movimentos da base mãe (linhas do Excel) → formato do motor, para conciliar sem voltar a ler o ficheiro.
  function movimentosDaBase(linhas, {conta = "", de = "", ate = "", excluir = new Set()} = {}) {
    const iso = (v) => {
      if (typeof v === "number" && v > 1000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000).toISOString().slice(0, 10);
      const t = texto(v), m = t.match(/^(\d{4})-(\d{2})-(\d{2})/) || t.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      return !m ? "" : m[1].length === 4 ? m[1] + "-" + m[2] + "-" + m[3] : m[3] + "-" + m[2] + "-" + m[1];
    };
    return (linhas || []).map((l) => ({
      id: texto(l.ID_Movimento), idBase: texto(l.ID_Movimento), data: iso(l.Data_Movimento), dataValor: iso(l.Data_Valor),
      descricao: texto(l.Descricao), referencia: texto(l.Referencia), valor: numero(l.Valor), saldo: l.Saldo === "" || l.Saldo == null ? null : numero(l.Saldo),
      conta: texto(l.Conta), origem: "base",
    })).filter((m) => m.id && m.data && Number.isFinite(m.valor) && !excluir.has(m.id) && (!conta || mesmaConta(m.conta, conta)) && (!de || m.data >= de) && (!ate || m.data <= ate))
      .sort((a, b) => a.data.localeCompare(b.data));
  }

  // ---------- Validação ----------
  // Confere: saldo anterior + movimento = saldo de cada linha; saldo inicial + soma = saldo final; duplicados.
  function validar(movimentos, {saldoInicial = null, saldoFinal = null} = {}) {
    const problemas = [];
    let anterior = saldoInicial;
    const quebras = [];
    movimentos.forEach((m, i) => {
      if (m.saldo != null && anterior != null && centimos(anterior + m.valor) !== centimos(m.saldo)) quebras.push(i);
      if (m.saldo != null) anterior = m.saldo;
      else if (anterior != null) anterior += m.valor;
    });
    // Extractos por ordem decrescente: testar a ordem inversa antes de acusar quebras
    let ordemInversa = false;
    if (quebras.length > movimentos.length / 2 && movimentos.length > 1) {
      let ant = null, q = 0;
      [...movimentos].reverse().forEach((m) => { if (m.saldo != null && ant != null && centimos(ant + m.valor) !== centimos(m.saldo)) q++; if (m.saldo != null) ant = m.saldo; });
      if (q < quebras.length) ordemInversa = true;
    }
    if (!ordemInversa && quebras.length) problemas.push(quebras.length + " linha(s) em que o saldo não bate com o movimento (" + quebras.slice(0, 5).map((i) => movimentos[i].data + " " + formatar(movimentos[i].valor)).join("; ") + "). Pode haver um sinal ou valor mal lido.");
    const debitos = movimentos.filter((m) => m.valor < 0).reduce((s, m) => s + m.valor, 0);
    const creditos = movimentos.filter((m) => m.valor > 0).reduce((s, m) => s + m.valor, 0);
    const ultimoSaldo = [...(ordemInversa ? [...movimentos].reverse() : movimentos)].reverse().find((m) => m.saldo != null)?.saldo ?? null;
    const final = saldoFinal ?? ultimoSaldo;
    const primeiro = (ordemInversa ? [...movimentos].reverse() : movimentos)[0];
    const inicial = saldoInicial ?? (primeiro && primeiro.saldo != null ? Math.round((primeiro.saldo - primeiro.valor) * 100) / 100 : null);
    if (inicial != null && final != null && centimos(inicial + debitos + creditos) !== centimos(final))
      problemas.push("Saldo inicial + movimentos (" + formatar(inicial + debitos + creditos) + ") difere do saldo final (" + formatar(final) + ").");
    const vistos = new Map(), duplicados = [];
    for (const m of movimentos) {
      const k = m.chave + "|" + texto(m.referencia).replace(/^0+/, "") + "|" + (m.saldo ?? "");
      if (vistos.has(k)) duplicados.push(m.id); else vistos.set(k, m.id);
    }
    if (duplicados.length) problemas.push(duplicados.length + " movimento(s) repetido(s) no ficheiro (mesma data, valor, descrição e saldo).");
    return {ok: !problemas.length, problemas, debitos, creditos, saldoInicial: inicial, saldoFinal: final, duplicados, quebras: ordemInversa ? [] : quebras.map((i) => movimentos[i].id), ordemInversa};
  }

  // ---------- Emparelhamento com a Tesouraria ----------
  // pagamentos: [{id, valor (positivo), data (Data de Saída), dataPrevista (Data de Despacho, só informativa), referencia, …}]
  // A data que conta é SEMPRE a Data de Saída (quando o dinheiro sai do banco), nunca a do despacho.
  // Regras: valor igual ao cêntimo é obrigatório; a data e as referências dão pontos.
  const PALAVRAS_VAZIAS = new Set(["de", "da", "do", "das", "dos", "e", "a", "o", "para", "lda", "sa", "limitada", "sarl", "mz", "trf", "transf", "pagamento", "pag"]);
  const fichas = (v) => chaveTexto(v).split(/[^a-z0-9]+/).filter((t) => t.length >= 3 && !PALAVRAS_VAZIAS.has(t));
  // ---------- Cheques ----------
  // Nos despachos o meio fica em «F. de Pagamento» e o nº do cheque em «Ref. Pagamento». No extracto o nº aparece no
  // descritivo («Pagamento de Cheque 1294827») ou na referência. O cheque é debitado quando é apresentado, por isso
  // pode cair semanas depois da saída: com o nº igual aceita-se até TOLERANCIA_CHEQUE dias depois.
  const TOLERANCIA_CHEQUE = 90;
  const soDigitos = (v) => texto(v).replace(/\D/g, "").replace(/^0+/, "");
  const eCheque = (meio) => /cheq|chq/.test(chaveTexto(meio));
  function chequesDoMovimento(m) {
    if (m._cheques) return m._cheques;
    const desc = chaveTexto(m.descricao + " " + m.referencia), out = new Set();
    if (/cheq|chq/.test(desc)) for (const t of (texto(m.descricao) + " " + texto(m.referencia)).match(/\d{3,12}/g) || []) out.add(t.replace(/^0+/, ""));
    const ref = texto(m.referencia);
    if (/^\d{3,12}$/.test(ref)) out.add(ref.replace(/^0+/, ""));
    Object.defineProperty(m, "_cheques", {value: out, enumerable: false});
    return out;
  }
  // Nº quase igual (um algarismo trocado, a mais ou a menos, ou dois vizinhos trocados): provável erro de digitação.
  function quaseIgual(a, b) {
    if (a === b || Math.abs(a.length - b.length) > 1 || Math.min(a.length, b.length) < 4) return false;
    if (a.length === b.length) {
      const d = [...a].map((c, i) => (c !== b[i] ? i : -1)).filter((i) => i >= 0);
      return d.length === 1 || (d.length === 2 && d[1] === d[0] + 1 && a[d[0]] === b[d[1]] && a[d[1]] === b[d[0]]);
    }
    const [c, l] = a.length < b.length ? [a, b] : [b, a];
    for (let i = 0; i < l.length; i++) if (l.slice(0, i) + l.slice(i + 1) === c) return true;
    return false;
  }
  // Nº abreviado no despacho: nos despachos é comum registar só o início e o fim do nº («4517» para 45564417).
  // Aceita quando o nº curto (3–5 algarismos) é um prefixo + sufixo do nº do banco (sufixo com 2 ou mais algarismos).
  function abreviado(curto, longo) {
    if (curto.length < 3 || curto.length > 5 || longo.length < curto.length + 2) return false;
    if (longo.endsWith(curto)) return true;
    for (let k = 1; k <= curto.length - 2; k++) if (longo.startsWith(curto.slice(0, k)) && longo.endsWith(curto.slice(k))) return true;
    return false;
  }
  // Relação entre o cheque do pagamento e o movimento: "igual" | "abreviado" | "parecido" | "diferente" | ""
  function relacaoCheque(mov, pag) {
    if (!pag.cheque) return {tipo: ""};
    const ch = chequesDoMovimento(mov);
    if (ch.has(pag.cheque)) return {tipo: "igual"};
    for (const c of ch) if (c.length >= 6 && pag.cheque.length >= 6 && (c.endsWith(pag.cheque) || pag.cheque.endsWith(c))) return {tipo: "igual"};
    for (const c of ch) if (abreviado(pag.cheque, c)) return {tipo: "abreviado", banco: c};
    for (const c of ch) if (quaseIgual(c, pag.cheque)) return {tipo: "parecido", banco: c};
    return {tipo: /cheq|chq/.test(chaveTexto(mov.descricao + " " + mov.referencia)) && ch.size ? "diferente" : "", banco: [...ch][0]};
  }

  function pontuar(mov, pag, tolerancia) {
    if (centimos(Math.abs(mov.valor)) !== centimos(Math.abs(pag.valor))) return null;
    const ref = pag.data; // Data de Saída
    const dias = ref ? Math.abs(diasEntre(mov.data, ref)) : null;
    const ch = relacaoCheque(mov, pag);
    // Sem Data de Saída só se aceita com o nº do cheque igual (e nunca automático)
    if (!ref && ch.tipo !== "igual" && ch.tipo !== "abreviado") return null;
    // Cheque com o mesmo nº: pode ser apresentado semanas depois da saída (nunca muito antes)
    const depois = ref ? diasEntre(mov.data, ref) : null; // dias do débito depois da saída
    const janelaCheque = (ch.tipo === "igual" || ch.tipo === "abreviado") && depois != null && depois >= -3 && depois <= TOLERANCIA_CHEQUE;
    if (dias != null && dias > tolerancia && !janelaCheque) return null;
    let pontos = 50, motivos = ["valor igual"];
    if (!ref) { pontos -= 30; motivos.push("despacho sem Data de Saída: confirmar e registar a saída"); }
    if (ch.tipo === "igual") { pontos += 40; motivos.push("nº do cheque " + pag.cheque + " igual"); }
    else if (ch.tipo === "abreviado") { pontos += 30; motivos.push("nº do cheque abreviado no despacho (" + pag.cheque + " = " + ch.banco + "; registar o nº completo)"); }
    else if (ch.tipo === "parecido") { pontos -= 10; motivos.push("nº do cheque parecido: despacho " + pag.cheque + ", banco " + ch.banco + " (verificar digitação)"); }
    else if (ch.tipo === "diferente") { pontos -= 40; motivos.push("nº do cheque diferente: despacho " + pag.cheque + ", banco " + ch.banco); }
    if (dias === 0) { pontos += 25; motivos.push("mesma data"); }
    else if (dias != null && !janelaCheque) { pontos += Math.max(0, 20 - dias * 3); motivos.push(dias + " dia(s) de diferença"); }
    else if (dias != null) motivos.push(depois >= 0 ? "cheque apresentado " + depois + " dia(s) depois da saída" : "débito " + (-depois) + " dia(s) antes da data de saída");
    const desc = chaveTexto(mov.descricao + " " + mov.referencia);
    const descSemEspacos = desc.replace(/[^a-z0-9]/g, "");
    for (const [campo, nome, peso] of [["referencia", "referência do pagamento", 30], ["despacho", "nº de despacho", 25], ["requisicao", "nº de requisição", 20]]) {
      if (campo === "referencia" && pag.cheque) continue; // o nº do cheque já foi avaliado acima
      const r = chaveTexto(pag[campo]).replace(/[^a-z0-9]/g, "");
      if (r.length >= 4 && descSemEspacos.includes(r)) { pontos += peso; motivos.push(nome + " no descritivo"); }
    }
    const fForn = fichas(pag.fornecedor);
    const comuns = fForn.filter((t) => desc.includes(t));
    if (comuns.length) { pontos += Math.min(15, comuns.length * 8); motivos.push("fornecedor no descritivo"); }
    return {pontos: Math.max(0, Math.min(100, pontos)), motivos, dias, cheque: ch.tipo};
  }

  // Emparelha débitos do extracto com pagamentos. 1:1 primeiro (melhor pontuação global); depois N:1 (vários pagamentos
  // num só débito, ex. transferência em lote) só com o que sobrou, até 5 pagamentos com datas próximas.
  // Automático: pontuação ≥ limiar e sem outro candidato igualmente bom; o resto fica como sugestão para confirmar.
  function conciliar(movimentos, pagamentos, opcoes = {}) {
    const tolerancia = opcoes.toleranciaDias ?? 5, limiar = opcoes.limiarAutomatico ?? 75, maxGrupo = opcoes.maxGrupo ?? 5;
    // sentido "credito": entradas do extracto contra recibos (Recebimentos); por omissão débitos contra pagamentos
    const debitos = movimentos.filter((m) => (opcoes.sentido === "credito" ? m.valor > 0 : m.valor < 0));
    const usadosM = new Set(), usadosP = new Set(), resultado = [];
    // 1.º Um cheque que paga várias despesas: as despesas com o mesmo nº de cheque somam o valor do débito (sem limite
    // de quantidade). É a prova mais forte, por isso vem antes do 1:1 (que podia levar uma das despesas sozinha).
    const porCheque = new Map();
    for (const p of pagamentos) if (p.cheque) porCheque.set(p.cheque, (porCheque.get(p.cheque) || []).concat([p]));
    for (const m of debitos) {
      for (const c of chequesDoMovimento(m)) {
        const ps = [...porCheque.entries()].filter(([k]) => k === c || abreviado(k, c)).flatMap(([, v]) => v).filter((p) => !usadosP.has(p.id));
        if (ps.length < 2 || usadosM.has(m.id)) continue;
        if (ps.reduce((s, p) => s + centimos(p.valor), 0) !== centimos(Math.abs(m.valor))) continue;
        const ref = ps.map((p) => p.data).filter(Boolean).sort()[0]; // Data de Saída
        const depois = ref ? diasEntre(m.data, ref) : 0;
        if (depois < -3 || depois > TOLERANCIA_CHEQUE) continue;
        const semSaida = ps.filter((p) => !p.data).length;
        ps.forEach((p) => usadosP.add(p.id)); usadosM.add(m.id);
        resultado.push({mov: m.id, pags: ps.map((p) => p.id), pontos: semSaida ? 70 : 95,
          motivos: ["cheque nº " + c + " paga " + ps.length + " despesas", "soma igual ao débito", ...(semSaida ? [semSaida + " despacho(s) sem Data de Saída: confirmar e registar a saída"] : [])],
          dias: ref ? Math.abs(depois) : null, tipo: "N:1", estado: semSaida ? "sugerido" : "automatico", ambiguo: false});
      }
    }
    const pares = [];
    for (const m of debitos) for (const p of pagamentos) {
      if (usadosM.has(m.id) || usadosP.has(p.id)) continue;
      const r = pontuar(m, p, tolerancia);
      if (r) pares.push({mov: m.id, pags: [p.id], ...r});
    }
    pares.sort((a, b) => b.pontos - a.pontos || (a.dias ?? 99) - (b.dias ?? 99));
    const candidatosM = new Map(), candidatosP = new Map();
    for (const p of pares) {
      candidatosM.set(p.mov, (candidatosM.get(p.mov) || []).concat(p));
      candidatosP.set(p.pags[0], (candidatosP.get(p.pags[0]) || []).concat(p));
    }
    for (const p of pares) {
      if (usadosM.has(p.mov) || usadosP.has(p.pags[0])) continue;
      // Ambíguo: outro candidato com a mesma pontuação no mesmo movimento ou pagamento (a troca seria igualmente válida)
      const rivais = [...candidatosM.get(p.mov), ...candidatosP.get(p.pags[0])].filter((q) => q !== p && q.pontos === p.pontos);
      usadosM.add(p.mov); usadosP.add(p.pags[0]);
      // Dois despachos com o mesmo cheque e o mesmo valor para um só débito: provável despacho duplicado
      const pagPorId = (id) => pagamentos.find((x) => x.id === id);
      const este = pagPorId(p.pags[0]);
      const gemeos = este?.cheque ? rivais.filter((q) => q.mov === p.mov && pagPorId(q.pags[0])?.cheque === este.cheque) : [];
      const motivos = gemeos.length ? [...p.motivos, "possível despacho duplicado: o cheque " + este.cheque + " aparece em " + (gemeos.length + 1) + " despachos com o mesmo valor"] : p.motivos;
      resultado.push({...p, motivos, tipo: "1:1", estado: p.pontos >= limiar && !rivais.length ? "automatico" : "sugerido", ambiguo: rivais.length > 0});
    }
    // N:1 com o que sobrou
    const pagLivres = pagamentos.filter((p) => !usadosP.has(p.id));
    for (const m of debitos) {
      if (usadosM.has(m.id)) continue;
      const alvo = centimos(Math.abs(m.valor));
      const cand = pagLivres.filter((p) => !usadosP.has(p.id) && centimos(p.valor) < alvo && p.data && Math.abs(diasEntre(m.data, p.data)) <= tolerancia
        && relacaoCheque(m, p).tipo !== "diferente")
        .sort((a, b) => b.valor - a.valor).slice(0, 14);
      const grupo = somaExacta(cand, alvo, maxGrupo);
      if (grupo) {
        grupo.forEach((p) => usadosP.add(p.id)); usadosM.add(m.id);
        resultado.push({mov: m.id, pags: grupo.map((p) => p.id), pontos: 60, motivos: ["soma de " + grupo.length + " pagamentos igual ao débito"], tipo: "N:1", estado: "sugerido"});
      }
    }
    return {
      pares: resultado,
      bancoSemPar: movimentos.filter((m) => !usadosM.has(m.id)).map((m) => m.id),
      tesourariaSemPar: pagamentos.filter((p) => !usadosP.has(p.id)).map((p) => p.id),
    };
  }
  function somaExacta(cand, alvo, max) {
    let achado = null;
    const passo = (i, soma, escolha) => {
      if (achado) return;
      if (soma === alvo && escolha.length >= 2) { achado = escolha.slice(); return; }
      if (soma > alvo || escolha.length >= max || i >= cand.length) return;
      escolha.push(cand[i]); passo(i + 1, soma + centimos(cand[i].valor), escolha); escolha.pop();
      passo(i + 1, soma, escolha);
    };
    passo(0, 0, []);
    return achado;
  }

  // Pagamentos da Tesouraria que entram na conciliação de uma conta e período.
  function pagamentosParaConciliar(registos, {conta = "", idConta = "", banco = "", de = "", ate = "", toleranciaDias = 5} = {}) {
    const kConta = chaveTexto(conta).replace(/[^a-z0-9]/g, ""), kBanco = chaveTexto(banco);
    const margem = (iso, d) => (iso ? new Date(Date.parse(iso) + d * 86400000).toISOString().slice(0, 10) : "");
    const ini = margem(de, -toleranciaDias), fim = margem(ate, toleranciaDias), iniCheque = margem(de, -TOLERANCIA_CHEQUE);
    return registos.filter((r) => {
      const estado = chaveTexto(r.estado);
      if (estado !== "pago" && estado !== "aguardando saida") return false;
      const valor = Number(r.valor);
      if (!Number.isFinite(valor) || valor <= 0) return false;
      // Período pela Data de Saída. Sem saída registada só entra se tiver nº de cheque (pode casar pelo cheque).
      const dt = r.dataSaida || "", comCheque = eCheque(r.meio) && soDigitos(r.referencia).length >= 3;
      if (!dt && !comCheque) return false;
      if (ini && dt && dt < (comCheque ? iniCheque : ini)) return false;
      if (fim && dt && dt > fim) return false;
      // Conta: pelo ID_Conta quando ambos o têm (despachos V9); senão pelo número (aceita número parcial no fim).
      if (idConta && r.idConta) return String(r.idConta).trim() === String(idConta).trim();
      const rc = chaveTexto(r.conta).replace(/[^a-z0-9]/g, ""), rb = chaveTexto(r.banco);
      if (kConta && rc && rc !== kConta && !rc.endsWith(kConta) && !kConta.endsWith(rc)) return false;
      if (!kConta && kBanco && rb && rb !== kBanco) return false;
      return true;
    }).map((r) => ({id: String(r.id), idDocumento: String(r.idDocumento || ""), tipo: "despesa", valor: Number(r.valor), data: r.dataSaida || "", dataPrevista: r.dataDespacho || "",
      referencia: r.referencia || "", meio: r.meio || "", cheque: eCheque(r.meio) ? soDigitos(r.referencia) : "", despacho: r.despacho || "", requisicao: r.requisicao || "", fornecedor: r.fornecedor || "",
      estado: r.estado, conta: r.conta || "", banco: r.banco || "", empresa: r.empresa || ""}));
  }

  // Pagamentos a partir da tblPagamentos (registo de cada saída de dinheiro: Data_Pagamento, ID_Conta, Meio_Pagamento,
  // Referencia_Transaccao = nº do cheque, Valor_Pagamento), com os dados da despesa (fornecedor, requisição, despacho)
  // pelo ID_Despesa. Cada linha é uma unidade de conciliação (uma despesa paga em vários cheques tem várias linhas).
  function pagamentosDaTabela(tblPagamentos, registos, opcoes = {}) {
    const iso = (v) => {
      if (typeof v === "number" && v > 1000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000).toISOString().slice(0, 10);
      const t = texto(v), m = t.match(/^(\d{4})-(\d{2})-(\d{2})/) || t.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      return !m ? "" : m[1].length === 4 ? m[1] + "-" + m[2] + "-" + m[3] : m[3] + "-" + m[2] + "-" + m[1];
    };
    const porDespesa = new Map();
    for (const r of registos || []) { const k = texto(r.idDocumento); if (k) porDespesa.set(k, r); }
    const linhas = (tblPagamentos || []).filter((p) => texto(p.ID_Pagamento) && !/anulad|cancelad|estornad/.test(chaveTexto(p.Estado_Pagamento))).map((p) => {
      const d = porDespesa.get(texto(p.ID_Despesa)) || {};
      return {id: "pag:" + texto(p.ID_Pagamento), idDocumento: texto(p.ID_Pagamento), idDespesa: texto(p.ID_Despesa), estado: "Pago",
        valor: numero(p.Valor_Pagamento), dataSaida: iso(p.Data_Pagamento) || d.dataSaida || "", dataDespacho: d.dataDespacho || "",
        referencia: texto(p.Referencia_Transaccao) || d.referencia || "", meio: texto(p.Meio_Pagamento) || d.meio || "",
        idConta: texto(p.ID_Conta) || d.idConta || "", conta: d.conta || "", banco: d.banco || "",
        requisicao: d.requisicao || "", fornecedor: d.fornecedor || "", despacho: d.despacho || "", empresa: d.empresa || ""};
    });
    const meta = new Map(linhas.map((l) => [l.id, l]));
    return pagamentosParaConciliar(linhas, opcoes).map((p) => ({...p, tipo: "pagamento", idDespesa: meta.get(p.id)?.idDespesa || ""}));
  }

  function resumo(movimentos, pagamentos, resultado) {
    const porId = new Map(movimentos.map((m) => [m.id, m]));
    const confirmados = resultado.pares.filter((p) => p.estado === "automatico" || p.estado === "confirmado");
    const somaMov = (ids) => ids.reduce((s, id) => s + Math.abs(porId.get(id)?.valor || 0), 0);
    const debitos = movimentos.filter((m) => m.valor < 0);
    const pagPorId = new Map(pagamentos.map((p) => [p.id, p]));
    return {
      movimentos: movimentos.length, debitos: debitos.length, creditos: movimentos.length - debitos.length,
      conciliados: confirmados.length, sugeridos: resultado.pares.filter((p) => p.estado === "sugerido").length,
      valorConciliado: somaMov(confirmados.map((p) => p.mov)),
      valorDebitos: debitos.reduce((s, m) => s - m.valor, 0),
      bancoSemPar: resultado.bancoSemPar.length, valorBancoSemPar: somaMov(resultado.bancoSemPar),
      tesourariaSemPar: resultado.tesourariaSemPar.length,
      valorTesourariaSemPar: resultado.tesourariaSemPar.reduce((s, id) => s + (pagPorId.get(id)?.valor || 0), 0),
      percentagem: movimentos.length ? Math.round(100 * confirmados.length / movimentos.length) : 0,
    };
  }

  // Relatório em linhas (para CSV/Excel): um registo por movimento do banco e por pagamento sem par.
  function relatorio(movimentos, pagamentos, resultado) {
    const porMov = new Map(), porPag = new Map(pagamentos.map((p) => [p.id, p]));
    for (const p of resultado.pares) porMov.set(p.mov, p);
    const linhas = movimentos.map((m) => {
      const par = porMov.get(m.id);
      const pags = par ? par.pags.map((id) => porPag.get(id)).filter(Boolean) : [];
      return {
        "Data banco": m.data, "Descritivo banco": m.descricao, "Referência banco": m.referencia,
        "Débito": m.valor < 0 ? -m.valor : "", "Crédito": m.valor > 0 ? m.valor : "", "Saldo": m.saldo ?? "",
        "Estado": !par ? (m.valor > 0 ? "Crédito sem contrapartida" : "Sem correspondência") : par.estado === "sugerido" ? "Sugerido (por confirmar)" : "Conciliado",
        "Tipo": par?.tipo || "", "Pontuação": par?.pontos ?? "", "Motivos": par ? par.motivos.join("; ") : "",
        "Requisição": pags.map((p) => p.requisicao).join(" + "), "Nº Despacho": pags.map((p) => p.despacho).join(" + "),
        "Ref. Pagamento": pags.map((p) => p.referencia).join(" + "), "Fornecedor": pags.map((p) => p.fornecedor).join(" + "),
        "Data de Saída": pags.map((p) => p.data).join(" + "), "Valor Tesouraria": pags.length ? pags.reduce((s, p) => s + p.valor, 0) : "",
      };
    });
    for (const id of resultado.tesourariaSemPar) {
      const p = porPag.get(id);
      linhas.push({"Data banco": "", "Descritivo banco": "", "Referência banco": "", "Débito": "", "Crédito": "", "Saldo": "",
        "Estado": "Pagamento sem movimento no extracto", "Tipo": "", "Pontuação": "", "Motivos": "",
        "Requisição": p.requisicao, "Nº Despacho": p.despacho, "Ref. Pagamento": p.referencia, "Fornecedor": p.fornecedor,
        "Data de Saída": p.data || "", "Valor Tesouraria": p.valor});
    }
    return linhas;
  }
  function csv(linhas) {
    if (!linhas.length) return "";
    const cols = Object.keys(linhas[0]);
    const cel = (v) => {
      if (typeof v === "number") return String(v).replace(".", ",");
      const s = String(v ?? "");
      const seguro = /^[=+\-@\t\r]/.test(s) ? "'" + s : s; // evita fórmulas ao abrir no Excel
      return /[";\n]/.test(seguro) ? '"' + seguro.replace(/"/g, '""') + '"' : seguro;
    };
    return "﻿" + [cols.join(";"), ...linhas.map((l) => cols.map((c) => cel(l[c])).join(";"))].join("\r\n");
  }
  function formatar(n) {
    return new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(Number(n) || 0);
  }

  return Object.freeze({numero, data, campoDoCabecalho, detectarCabecalho, matrizParaMovimentos, linhasPdfParaMatriz,
    textoParaMovimentos, validar, pontuar, conciliar, pagamentosParaConciliar, resumo, relatorio, csv, formatar, CAMPOS,
    idsMovimentos, nomePadronizado, contaDoFicheiro, mesmaConta, movimentosParaBase, movimentosDaBase, pagamentosDaTabela});
});
