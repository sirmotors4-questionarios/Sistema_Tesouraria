const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../conciliacao-core.js");

test("números nos formatos dos extractos", () => {
  const casos = [["1.234,56", 1234.56], ["1 234,56", 1234.56], ["1,234.56", 1234.56], ["-1234.56", -1234.56],
    ["(1 234,56)", -1234.56], ["1.234,56-", -1234.56], ["1.234,56 D", -1234.56], ["1.234,56 CR", 1234.56],
    ["MZN 25.000,00", 25000], ["12.500", 12500], ["0,50", 0.5], [1500, 1500], ["2.345.678,90", 2345678.9], ["1'250.00", 1250]];
  for (const [v, esperado] of casos) assert.equal(C.numero(v), esperado, String(v));
  for (const v of ["", "-", "abc", "12/11/2026"]) assert.ok(Number.isNaN(C.numero(v)), v);
});

test("datas nos formatos dos extractos", () => {
  assert.equal(C.data("05/11/2026"), "2026-11-05");
  assert.equal(C.data("5-11-26"), "2026-11-05");
  assert.equal(C.data("05.11.2026"), "2026-11-05");
  assert.equal(C.data("2026-11-05"), "2026-11-05");
  assert.equal(C.data(46331), "2026-11-05");
  assert.equal(C.data("05 Nov 2026"), "2026-11-05");
  assert.equal(C.data("05-DEZ-26"), "2026-12-05");
  assert.equal(C.data("05/11", 2026), "2026-11-05");
  assert.equal(C.data("31/02/2026"), "");
  assert.equal(C.data("Saldo"), "");
});

test("cabeçalhos com acentos, unidades e sinónimos em inglês", () => {
  assert.equal(C.campoDoCabecalho("Data Mov."), "data");
  assert.equal(C.campoDoCabecalho("Data Valor"), "dataValor");
  assert.equal(C.campoDoCabecalho("Débito (MZN)"), "debito");
  assert.equal(C.campoDoCabecalho("Crédito"), "credito");
  assert.equal(C.campoDoCabecalho("Saldo Contabilístico"), "saldo");
  assert.equal(C.campoDoCabecalho("Descritivo"), "descricao");
  assert.equal(C.campoDoCabecalho("Nº Doc"), "referencia");
  assert.equal(C.campoDoCabecalho("Withdrawals"), "debito");
});

// Extracto típico exportado do internet banking: título, dados da conta, cabeçalho na linha 6, saldo inicial.
const EXCEL = [
  ["BANCO EXEMPLO, SA"], ["Extracto de Conta"], ["Conta:", "0001234567"], ["Período:", "01/11/2026 a 30/11/2026"], [],
  ["Data Mov.", "Data Valor", "Descritivo", "Nº Doc", "Débito", "Crédito", "Saldo"],
  ["", "", "Saldo Anterior", "", "", "", "100.000,00"],
  ["03/11/2026", "03/11/2026", "TRF FORNECEDOR ALFA LDA REF PAG-0045", "778812", "25.000,00", "", "75.000,00"],
  ["", "", "despacho D-2026-118", "", "", "", ""],
  ["05/11/2026", "05/11/2026", "DEPOSITO CLIENTE", "", "", "40.000,00", "115.000,00"],
  ["06/11/2026", "06/11/2026", "COMISSAO MANUTENCAO", "", "350,00", "", "114.650,00"],
  ["07/11/2026", "07/11/2026", "TRF LOTE SALARIOS", "", "30.000,00", "", "84.650,00"],
  [], ["", "", "Saldo Final", "", "", "", "84.650,00"],
];

test("Excel: encontra o cabeçalho, junta linhas de continuação e lê saldos", () => {
  const r = C.matrizParaMovimentos(EXCEL);
  assert.equal(r.linhaCabecalho, 5);
  assert.equal(r.movimentos.length, 4);
  assert.deepEqual(r.movimentos.map((m) => m.valor), [-25000, 40000, -350, -30000]);
  assert.match(r.movimentos[0].descricao, /despacho D-2026-118$/);
  assert.equal(r.movimentos[0].referencia, "778812");
  assert.equal(r.saldoInicial, 100000);
  assert.equal(r.saldoFinal, 84650);
  const v = C.validar(r.movimentos, r);
  assert.equal(v.ok, true, v.problemas.join(" "));
  assert.equal(v.debitos, -55350); assert.equal(v.creditos, 40000);
});

test("Excel: coluna única de montante com sinal e mapeamento manual", () => {
  const m = [["Quando", "O quê", "Quanto"], ["03/11/2026", "Pagamento", "-1.500,00"], ["04/11/2026", "Recebido", "2.000,00"]];
  assert.equal(C.matrizParaMovimentos(m).movimentos.length, 0);
  const r = C.matrizParaMovimentos(m, {mapa: {data: 0, descricao: 1, montante: 2}, linhaCabecalho: 0});
  assert.deepEqual(r.movimentos.map((x) => x.valor), [-1500, 2000]);
});

test("validação: detecta saldo que não bate e movimentos repetidos", () => {
  const r = C.matrizParaMovimentos(EXCEL);
  const errado = r.movimentos.map((m, i) => (i === 2 ? {...m, valor: 350} : m)); // comissão lida como crédito
  const v = C.validar(errado, r);
  assert.equal(v.ok, false);
  assert.ok(v.problemas.some((p) => /saldo não bate/.test(p)));
  const dup = C.validar([...r.movimentos, r.movimentos[1]], {});
  assert.ok(dup.problemas.some((p) => /repetido/.test(p)));
});

test("PDF com texto: reconstrói as colunas pelas posições do cabeçalho", () => {
  const it = (x, str, w = str.length * 5) => ({x, w, str});
  const linhas = [
    {y: 800, itens: [it(40, "BANCO EXEMPLO - EXTRACTO")]},
    {y: 760, itens: [it(40, "Data", 20), it(100, "Descritivo", 50), it(330, "Débito", 30), it(410, "Crédito", 35), it(490, "Saldo", 25)]},
    {y: 740, itens: [it(40, "03/11/2026", 50), it(100, "TRF FORNECEDOR ALFA LDA", 120), it(310, "25.000,00", 50)]},
    {y: 730, itens: [it(100, "REF PAG-0045", 60)]},
    {y: 720, itens: [it(40, "05/11/2026", 50), it(100, "DEPOSITO CLIENTE", 80), it(395, "40.000,00", 50), it(470, "115.000,00", 45)]},
  ];
  // linha de 03/11 sem saldo (algumas páginas omitem) — só testa colunas
  const p = C.linhasPdfParaMatriz(linhas);
  assert.ok(p);
  const r = C.matrizParaMovimentos(p.matriz, {mapa: p.mapa, linhaCabecalho: 0});
  assert.deepEqual(r.movimentos.map((m) => m.valor), [-25000, 40000]);
  assert.match(r.movimentos[0].descricao, /REF PAG-0045/);
  assert.equal(r.movimentos[1].saldo, 115000);
});

test("texto de OCR sem cabeçalho: sinal pela variação do saldo e números da descrição preservados", () => {
  const linhas = [
    "BANCO EXEMPLO EXTRACTO", "SALDO ANTERIOR 100.000,00",
    "03/11/2026 TRF FORNECEDOR ALFA REF 0045 25.000,00 75.000,00",
    "05/11/2026 DEPOSITO CLIENTE 40.000,00 115.000,00",
    "06/11/2026 06/11/2026 COMISSAO 350,00 114.650,00",
  ];
  const r = C.textoParaMovimentos(linhas);
  assert.deepEqual(r.movimentos.map((m) => m.valor), [-25000, 40000, -350]);
  assert.match(r.movimentos[0].descricao, /REF 0045$/);
  assert.equal(r.saldoInicial, 100000);
  assert.equal(C.validar(r.movimentos, r).ok, true);
});

const REGISTOS = [
  {id: "D1", estado: "Pago", valor: 25000, dataSaida: "2026-11-03", referencia: "PAG-0045", despacho: "D-2026-118", requisicao: "REQ-501", fornecedor: "Fornecedor Alfa, Lda", conta: "0001234567", banco: "Banco Exemplo"},
  {id: "D2", estado: "Pago", valor: 12000, dataSaida: "2026-11-07", requisicao: "REQ-510", fornecedor: "Salários Lote A", conta: "0001234567"},
  {id: "D3", estado: "Pago", valor: 18000, dataSaida: "2026-11-07", requisicao: "REQ-511", fornecedor: "Salários Lote B", conta: "0001234567"},
  {id: "D4", estado: "Pago", valor: 9999, dataSaida: "2026-11-10", requisicao: "REQ-520", fornecedor: "Gama", conta: "0001234567"},
  {id: "D5", estado: "Pendente", valor: 350, dataSaida: "", requisicao: "REQ-530", conta: "0001234567"},
  {id: "D6", estado: "Pago", valor: 5000, dataSaida: "2026-11-04", requisicao: "REQ-540", conta: "999888777"},
];

test("selecção de pagamentos: estado, período com margem e conta", () => {
  const p = C.pagamentosParaConciliar(REGISTOS, {conta: "0001234567", de: "2026-11-01", ate: "2026-11-30"});
  assert.deepEqual(p.map((x) => x.id), ["D1", "D2", "D3", "D4"]);
  assert.deepEqual(C.pagamentosParaConciliar(REGISTOS, {de: "2026-12-10", ate: "2026-12-31"}).map((x) => x.id), []);
});

test("conciliação: 1:1 automático por valor+data+referência, N:1 sugerido, sobras de cada lado", () => {
  const {movimentos} = C.matrizParaMovimentos(EXCEL);
  const pags = C.pagamentosParaConciliar(REGISTOS, {conta: "0001234567", de: "2026-11-01", ate: "2026-11-30"});
  const r = C.conciliar(movimentos, pags);
  const par1 = r.pares.find((p) => p.pags[0] === "D1");
  assert.equal(par1.estado, "automatico");
  assert.equal(par1.mov, movimentos[0].id);
  assert.ok(par1.motivos.includes("referência do pagamento no descritivo"));
  assert.ok(par1.motivos.includes("nº de despacho no descritivo"));
  const lote = r.pares.find((p) => p.tipo === "N:1");
  assert.equal(lote.mov, movimentos[3].id);
  assert.deepEqual(lote.pags.sort(), ["D2", "D3"]);
  assert.equal(lote.estado, "sugerido");
  assert.deepEqual(r.bancoSemPar, [movimentos[1].id, movimentos[2].id]); // depósito e comissão
  assert.deepEqual(r.tesourariaSemPar, ["D4"]);
  const s = C.resumo(movimentos, pags, r);
  assert.equal(s.conciliados, 1); assert.equal(s.sugeridos, 1); assert.equal(s.valorTesourariaSemPar, 9999);
});

test("conciliação: valor diferente ou data fora da tolerância nunca emparelha; empate fica como sugestão", () => {
  const mov = [{id: "m1", data: "2026-11-03", valor: -1000, descricao: "TRF", referencia: ""}, {id: "m2", data: "2026-11-03", valor: -1000, descricao: "TRF", referencia: ""}];
  assert.equal(C.conciliar(mov, [{id: "p", valor: 1000.01, data: "2026-11-03"}]).pares.length, 0);
  assert.equal(C.conciliar(mov, [{id: "p", valor: 1000, data: "2026-11-20"}]).pares.length, 0);
  const r = C.conciliar(mov, [{id: "p1", valor: 1000, data: "2026-11-03"}, {id: "p2", valor: 1000, data: "2026-11-03"}]);
  assert.equal(r.pares.length, 2);
  assert.ok(r.pares.every((p) => p.estado === "sugerido" && p.ambiguo));
});

test("relatório CSV: separador ;, decimais com vírgula e protecção contra fórmulas", () => {
  const {movimentos} = C.matrizParaMovimentos(EXCEL);
  const pags = C.pagamentosParaConciliar(REGISTOS, {conta: "0001234567"});
  const r = C.conciliar(movimentos, pags);
  const linhas = C.relatorio(movimentos, pags, r);
  assert.equal(linhas.length, movimentos.length + r.tesourariaSemPar.length);
  const texto = C.csv([...linhas, {...linhas[0], "Descritivo banco": "=HYPERLINK(\"x\")"}]);
  assert.ok(texto.startsWith("﻿Data banco;"));
  assert.match(texto, /;25000;/);
  assert.match(texto, /"'=HYPERLINK\(""x""\)"/);
});

test("PDF: títulos vizinhos (Data Mov. | Data Valor) e títulos de duas palavras", () => {
  const it = (x, str, w) => ({x, w, str});
  const linhas = [
    {y: 760, itens: [it(40, "Data", 18), it(60, "Mov.", 18), it(82, "Data", 18), it(102, "Valor", 22), it(140, "Descritivo", 50), it(330, "Débito", 30), it(410, "Crédito", 35), it(470, "Saldo", 25), it(497, "Contabilístico", 50)]},
    {y: 740, itens: [it(40, "03/11/2026", 40), it(82, "04/11/2026", 40), it(140, "TRF ALFA", 40), it(310, "25.000,00", 50), it(490, "75.000,00", 50)]},
  ];
  const p = C.linhasPdfParaMatriz(linhas);
  assert.deepEqual(p.matriz[0], ["Data Mov.", "Data Valor", "Descritivo", "Débito", "Crédito", "Saldo Contabilístico"]);
  const r = C.matrizParaMovimentos(p.matriz, {mapa: p.mapa, linhaCabecalho: 0});
  assert.equal(r.movimentos[0].data, "2026-11-03"); assert.equal(r.movimentos[0].dataValor, "2026-11-04");
  assert.equal(r.movimentos[0].valor, -25000); assert.equal(r.movimentos[0].saldo, 75000);
});

test("selecção de pagamentos pelo ID_Conta quando existe", () => {
  const regs = [{id: "a", estado: "Pago", valor: 10, dataSaida: "2026-11-02", idConta: "C7", conta: ""}, {id: "b", estado: "Pago", valor: 10, dataSaida: "2026-11-02", idConta: "C8", conta: "0001234567"}];
  assert.deepEqual(C.pagamentosParaConciliar(regs, {idConta: "C7", conta: "0001234567"}).map((x) => x.id), ["a"]);
});
