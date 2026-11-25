// Cheques na conciliação e leitura de PDFs impressos do browser (colunas cortadas). Dados fictícios.
const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../conciliacao-core.js");
globalThis.window = globalThis;
require("../conciliacao-import.js");
const I = globalThis.TES_IMPORTAR;

const mov = (id, data, valor, descricao, referencia = "") => ({id, data, valor, descricao, referencia, saldo: null});
const despacho = (id, valor, dataSaida, cheque, meio = "Cheque") => ({id, idDocumento: id, estado: "Pago", valor, dataSaida, referencia: cheque, meio, fornecedor: "Fornecedor " + id});
const conciliar = (movs, regs, op = {}) => C.conciliar(movs, C.pagamentosParaConciliar(regs, {}), {toleranciaDias: 5, ...op});

test("cheque com o mesmo nº: automático mesmo apresentado 30 dias depois da saída", () => {
  const r = conciliar([mov("m1", "2026-03-31", -40000, "Chq.Caixa 45564506", "45564506")], [despacho("d1", 40000, "2026-03-01", "45564506")]);
  assert.equal(r.pares.length, 1);
  assert.equal(r.pares[0].estado, "automatico");
  assert.ok(r.pares[0].motivos.some((t) => /nº do cheque 45564506 igual/.test(t)));
  // Sem o meio «Cheque», 30 dias está fora da tolerância
  assert.equal(conciliar([mov("m1", "2026-03-31", -40000, "Chq.Caixa 45564506", "45564506")], [despacho("d1", 40000, "2026-03-01", "45564506", "")]).pares.length, 0);
});

test("um cheque que paga várias despesas (mais de 5): todas ligadas ao débito", () => {
  const regs = [1, 2, 3, 4, 5, 6, 7].map((k) => despacho("d" + k, 1000 * k, "2026-05-10", "1294827"));
  const r = conciliar([mov("m1", "2026-05-12", -28000, "Pagamento de Cheque 1294827")], regs);
  assert.equal(r.pares.length, 1);
  assert.equal(r.pares[0].pags.length, 7);
  assert.equal(r.pares[0].estado, "automatico");
  assert.match(r.pares[0].motivos[0], /cheque nº 1294827 paga 7 despesas/);
});

test("nº abreviado no despacho («4517» = 45564417): automático com nota para registar o nº completo", () => {
  const r = conciliar([mov("m1", "2026-06-19", -313501.67, "Chq.Caixa 45564417", "45564417")], [despacho("d1", 313501.67, "2026-06-19", "4517")]);
  assert.equal(r.pares[0].estado, "automatico");
  assert.ok(r.pares[0].motivos.some((t) => /abreviado.*4517 = 45564417/.test(t)));
});

test("nº parecido (erro de digitação) fica como sugestão; nº diferente nunca é automático", () => {
  const parecido = conciliar([mov("m1", "2026-08-06", -50000, "Chq.Caixa 45564539", "45564539")], [despacho("d1", 50000, "2026-08-06", "45564439")]);
  assert.equal(parecido.pares[0].estado, "sugerido");
  assert.ok(parecido.pares[0].motivos.some((t) => /parecido: despacho 45564439, banco 45564539/.test(t)));
  const diferente = conciliar([mov("m1", "2026-07-08", -25000, "Chq.Caixa 31080287", "31080287")], [despacho("d1", 25000, "2026-07-08", "29990011")]);
  assert.notEqual(diferente.pares[0]?.estado, "automatico");
  // O cheque certo ganha ao errado com o mesmo valor
  const dois = conciliar([mov("m1", "2026-07-08", -25000, "Chq.Caixa 31080287", "31080287")], [despacho("d1", 25000, "2026-07-08", "29990011"), despacho("d2", 25000, "2026-07-01", "31080287")]);
  assert.deepEqual(dois.pares[0].pags, ["d2"]);
  assert.equal(dois.pares[0].estado, "automatico");
});

// Linhas de um PDF impresso do browser com a página estreita: títulos CREDIT e BALANCE cortados, saldo sem cêntimos.
function linhasPdf() {
  const it = (x, w, str) => ({x, w, str});
  const cab = {y: 0, itens: [it(40, 36, "TRAN DATE"), it(79, 42, "VALUE DATE"), it(126, 39, "REFERENCE"), it(189, 27, "DETAILS"), it(447, 20, "DEBIT")]};
  const dados = [
    ["02/06/2025", "FT1", "Despesa de transf. Local", "510.00", "0.00", "2,306,722."],
    ["02/06/2025", "FT1", "Transferencia MTR", "695,590.00", "0.00", "1,611,132."],
    ["03/06/2025", "FT2", "Trf recebida", "0.00", "33,680.00", "1,644,812."],
    ["04/06/2025", "FT3", "Pagamento de Cheque", "577,000.00", "0.00", "1,067,812."],
    ["05/06/2025", "FT4", "Deposito numerario", "0.00", "2,444.00", "1,070,256."],
  ];
  const linhas = [cab];
  dados.forEach((d, k) => linhas.push({y: -10 * (k + 1), itens: [it(40, 29, d[0]), it(79, 29, d[0]), it(119, 40, d[1]), it(187, 90, d[2]),
    it(500 - d[3].length * 4.4, d[3].length * 4.4, d[3]), it(545.5 - d[4].length * 4.4, d[4].length * 4.4, d[4]), it(584.6 - d[5].length * 4.4, d[5].length * 4.4, d[5])]}));
  return linhas;
}

test("PDF com colunas cortadas: crédito e saldo deduzidos pela posição; saldo sem cêntimos reconstruído", () => {
  const r = I._testes.interpretarLinhas(linhasPdf(), 2025);
  assert.deepEqual(r.inferidas, ["Crédito (deduzido)", "Saldo (deduzido)"]);
  const m = r.resultado.movimentos;
  assert.deepEqual(m.map((x) => x.valor), [-510, -695590, 33680, -577000, 2444]);
  // Saldo inicial 2 307 232,xx: os cêntimos ficam determinados pelos saldos cortados (todos batem com ,14 a ,99… → mínimo)
  assert.ok(r.resultado.avisos.some((a) => /cortou a coluna do saldo/.test(a)));
  const v = C.validar(m, {saldoInicial: r.resultado.saldoInicial, saldoFinal: r.resultado.saldoFinal});
  assert.equal(v.ok, true);
  assert.equal(Math.floor(m[0].saldo), 2306722);
  assert.equal(Math.floor(m[4].saldo), 1070256);
});

test("OCR: confusões em números corrigidas, texto intacto", () => {
  const L = I._testes.limparOcr;
  assert.equal(L("544,52]."), "544,521.");
  assert.equal(L("l0,000.00"), "10,000.00");
  assert.equal(L("1,O00.00"), "1,000.00");
  assert.equal(L("Local"), "Local");
  assert.equal(L("FT25160D70Q4"), "FT25160D70Q4");
  assert.equal(L("09/06/2025 —"), "09/06/2025");
});

test("OCR: palavra do descritivo não vai para a coluna do débito", () => {
  const linhas = linhasPdf();
  linhas[0].itens.splice(4, 1, {x: 369, w: 98, str: "AUTHORISATION CODE DEBIT"}); // títulos juntos pelo OCR
  linhas[2].itens.splice(3, 1, {x: 187, w: 60, str: "Transferencia MTR"}, {x: 380, w: 40, str: "INVEST"});
  const m = I._testes.interpretarLinhas(linhas, 2025).resultado.movimentos;
  assert.equal(m.length, 5);
  assert.equal(m[1].valor, -695590);
  assert.match(m[1].descricao, /INVEST/);
});

test("conta indicada no PDF: «Account Number : 3368» e NIB a partir do IBAN", () => {
  const r = C.contaDoFicheiro(null, ["SIR COMERCIO Account Number : 3368", "MAPUTO IBAN : MZ59001100020000000336822"]);
  assert.equal(r.conta, "3368");
  assert.equal(r.nib, "001100020000000336822");
});

test("mesmo cheque e valor em dois despachos: sugestão com aviso de possível duplicado", () => {
  const r = conciliar([mov("m1", "2026-05-22", -5641.39, "Ch.Comp.45564407", "45564407")],
    [despacho("d1", 5641.39, "2026-05-20", "45564407"), despacho("d2", 5641.39, "2026-05-18", "45564407")]);
  assert.equal(r.pares.length, 1);
  assert.equal(r.pares[0].estado, "sugerido");
  assert.ok(r.pares[0].motivos.some((t) => /possível despacho duplicado: o cheque 45564407 aparece em 2 despachos/.test(t)));
});

test("a data que conta é a Data de Saída, não a do despacho", () => {
  // Despacho a 01/04, saída a 20/04; débito a 20/04 → par pela saída (mesma data)
  const reg = {id: "d1", idDocumento: "d1", estado: "Pago", valor: 12000, dataSaida: "2026-04-20", dataDespacho: "2026-04-01", referencia: "", meio: "Transferência", fornecedor: "X"};
  const r = conciliar([mov("m1", "2026-04-20", -12000, "Transferencia MTR")], [reg]);
  assert.equal(r.pares.length, 1);
  assert.ok(r.pares[0].motivos.includes("mesma data"));
  // Débito na data do despacho (01/04), longe da saída: não é par
  assert.equal(conciliar([mov("m1", "2026-04-01", -12000, "Transferencia MTR")], [reg]).pares.length, 0);
});

test("despacho sem Data de Saída: só pelo nº do cheque, e como sugestão", () => {
  const semSaida = (cheque, meio = "Cheque") => ({id: "d1", idDocumento: "d1", estado: "Aguardando Saída", valor: 9000, dataSaida: "", dataDespacho: "2026-05-02", referencia: cheque, meio, fornecedor: "Y"});
  // Sem cheque: fica de fora (não se usa a data do despacho)
  assert.equal(C.pagamentosParaConciliar([semSaida("", "Transferência")], {}).length, 0);
  // Com o mesmo nº de cheque: sugestão com aviso para registar a saída
  const r = conciliar([mov("m1", "2026-05-05", -9000, "Chq.Caixa 45564600", "45564600")], [semSaida("45564600")]);
  assert.equal(r.pares.length, 1);
  assert.equal(r.pares[0].estado, "sugerido");
  assert.ok(r.pares[0].motivos.some((t) => /sem Data de Saída/.test(t)));
  // Com outro cheque: não é par
  assert.equal(conciliar([mov("m1", "2026-05-05", -9000, "Chq.Caixa 45564601", "45564601")], [semSaida("45564600")]).pares.length, 0);
});

test("tblPagamentos: cada pagamento é uma linha (cheque, data, conta), com os dados da despesa", () => {
  const despesas = [{id: "x1", idDocumento: "D1", estado: "Pago", valor: 30000, dataSaida: "2026-03-01", requisicao: "R-77", fornecedor: "Fornecedor A", referencia: "", meio: ""}];
  const tbl = [
    {ID_Pagamento: "P1", ID_Despesa: "D1", Data_Pagamento: 46087, ID_Conta: "C9", Meio_Pagamento: "Cheque", Referencia_Transaccao: "45564600", Valor_Pagamento: 20000, Estado_Pagamento: "Confirmado"},
    {ID_Pagamento: "P2", ID_Despesa: "D1", Data_Pagamento: "2026-03-20", ID_Conta: "C9", Meio_Pagamento: "Cheque", Referencia_Transaccao: "45564601", Valor_Pagamento: 10000, Estado_Pagamento: "Confirmado"},
    {ID_Pagamento: "P3", ID_Despesa: "D1", Data_Pagamento: "2026-03-20", ID_Conta: "C9", Meio_Pagamento: "Cheque", Referencia_Transaccao: "45564602", Valor_Pagamento: 999, Estado_Pagamento: "Anulado"},
  ];
  const pags = C.pagamentosDaTabela(tbl, despesas, {idConta: "C9"});
  assert.equal(pags.length, 2, "o anulado fica de fora");
  assert.deepEqual(pags.map((p) => [p.idDocumento, p.cheque, p.data, p.tipo, p.idDespesa, p.fornecedor]),
    [["P1", "45564600", "2026-03-06", "pagamento", "D1", "Fornecedor A"], ["P2", "45564601", "2026-03-20", "pagamento", "D1", "Fornecedor A"]]);
  // Despesa paga em 2 cheques: cada débito liga ao seu pagamento
  const r = C.conciliar([mov("m1", "2026-03-10", -20000, "Chq.Caixa 45564600", "45564600"), mov("m2", "2026-03-25", -10000, "Chq.Caixa 45564601", "45564601")], pags, {toleranciaDias: 5});
  assert.deepEqual(r.pares.map((p) => [p.mov, p.pags[0], p.estado]), [["m1", "pag:P1", "automatico"], ["m2", "pag:P2", "automatico"]]);
});
