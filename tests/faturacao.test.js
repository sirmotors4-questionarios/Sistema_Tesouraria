const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../conciliacao-core.js");
const B = require("../faturacao-core.js");

// Listagem de documentos de venda (formato típico exportado do PRIMAVERA para Excel)
const FATURAS = [
  ["SIR MOTORS, LDA"], ["Listagem de Documentos de Venda"], [],
  ["Documento", "Data", "Data Venc.", "Entidade", "Nome", "NUIT", "Moeda", "Total Documento", "Anulado"],
  ["FA 2026/101", "02/11/2026", "02/12/2026", "C001", "Cliente Beta, Lda", "400123456", "MZN", "300.000,00", "Não"],
  ["FA 2026/102", "05/11/2026", "20/11/2026", "C002", "Cliente Omega", "400999888", "MZN", "120.000,00", "Não"],
  ["FA 2026/103", "06/11/2026", "06/11/2026", "C001", "Cliente Beta, Lda", "400123456", "MZN", "50.000,00", "Não"],
  ["NC 2026/7", "08/11/2026", "", "C002", "Cliente Omega", "400999888", "MZN", "20.000,00", "Não"],
  ["FR 2026/55", "09/11/2026", "", "C003", "Consumidor Final", "", "MZN", "8.500,00", "Não"],
  ["FA 2026/104", "10/11/2026", "", "C003", "Consumidor Final", "", "MZN", "1.000,00", "Sim"],
  ["FA 2026/101", "02/11/2026", "02/12/2026", "C001", "Cliente Beta, Lda", "400123456", "MZN", "300.000,00", "Não"],
  ["RE 2026/9", "10/11/2026", "", "C001", "Cliente Beta, Lda", "", "MZN", "1,00", "Não"],
  ["XPTO", "10/11/2026", "", "C001", "x", "", "MZN", "1,00", "Não"],
  ["Total", "", "", "", "", "", "", "499.500,00", ""],
];
// Listagem de recibos com uma linha por documento liquidado
const RECIBOS = [
  ["Listagem de Recibos"],
  ["Recibo", "Data", "Cliente", "Nome", "Modo Pagamento", "Referência", "Documento Liquidado", "Valor Liquidado", "Total Recibo"],
  ["RE 2026/10", "04/11/2026", "C001", "Cliente Beta, Lda", "Transferência", "TRF889", "FA 2026/101", "300.000,00", "300.000,00"],
  ["RE 2026/11", "12/11/2026", "C002", "Cliente Omega", "Transferência", "", "FA 2026/102", "60.000,00", "100.000,00"],
  ["RE 2026/11", "12/11/2026", "C002", "Cliente Omega", "Transferência", "", "FA 2026/999", "40.000,00", "100.000,00"],
  ["RE 2026/12", "15/11/2026", "C004", "Cliente Novo", "Numerário", "", "", "", "25.000,00"],
];

test("documentos PRIMAVERA: tipo, série e número", () => {
  assert.deepEqual(B.partirDocumento("FA 2026/101"), {tipo: "FA", serie: "2026", numero: "101"});
  assert.deepEqual(B.partirDocumento("fa.2026.0101"), {tipo: "FA", serie: "2026", numero: "101"});
  assert.deepEqual(B.partirDocumento("NC A/7"), {tipo: "NC", serie: "A", numero: "7"});
  assert.deepEqual(B.partirDocumento("VD 12"), {tipo: "VD", serie: "", numero: "12"});
  assert.equal(B.partirDocumento("XPTO"), null);
  assert.equal(B.idDocumento("Sir Motors", "fa", "2026", "0101"), "SIR MOTORS|FA|2026|101");
});

test("facturas: cabeçalho, valores, anulados, rejeitados com motivo", () => {
  const r = B.lerListagem(FATURAS, {tipo: "faturas", empresa: "SIR Motors"});
  assert.equal(r.linhaCabecalho, 3);
  assert.deepEqual(r.registos.map((x) => x.Documento), ["FA 2026/101", "FA 2026/102", "FA 2026/103", "NC 2026/7", "FR 2026/55", "FA 2026/104"]);
  const fa = r.registos[0];
  assert.equal(fa.Total, 300000); assert.equal(fa.Data, "2026-11-02"); assert.equal(fa.Data_Vencimento, "2026-12-02"); assert.equal(fa.NUIT, "400123456");
  assert.equal(r.registos[3].Total, 20000);
  assert.equal(r.registos[5].Estado, "Anulado");
  assert.equal(r.registos[4].Data_Vencimento, "2026-11-09"); // sem vencimento: a data do documento
  assert.deepEqual(r.rejeitados.map((x) => x.motivo.split(" ")[0]), ["documento", "RE", "documento"]);
  assert.match(r.rejeitados[0].motivo, /repetido no ficheiro \(linha 5\)/);
  assert.match(r.rejeitados[1].motivo, /importe-o em Recebimentos/);
  assert.match(r.registos[0].Hash, /^[0-9a-f]{8}$/);
});

test("facturas em colunas separadas (Tipo Doc, Série, Núm. Doc.) com Empresa no ficheiro", () => {
  const m = [["Empresa", "Tipo Doc", "Série", "NumDoc", "Data Doc.", "Cód. Cliente", "Nome Cliente", "Total"],
    ["SIR", "FA", "2026", "00012", 46330, "C9", "Gama", 1500.5]];
  const r = B.lerListagem(m, {tipo: "faturas"});
  assert.equal(r.registos[0].ID_Documento, "SIR|FA|2026|12");
  assert.equal(r.registos[0].Data, "2026-11-04"); assert.equal(r.registos[0].Total, 1500.5);
});

test("recibos: agrupa linhas de liquidação e aceita recibo sem documento", () => {
  const r = B.lerListagem(RECIBOS, {tipo: "recibos", empresa: "SIR Motors"});
  assert.equal(r.registos.length, 3);
  const re11 = r.registos.find((x) => x.Documento === "RE 2026/11");
  assert.equal(re11.Valor, 100000);
  assert.deepEqual(re11.linhas.map((l) => [l.Documento_Liquidado, l.Valor]), [["FA 2026/102", 60000], ["FA 2026/999", 40000]]);
  assert.equal(r.registos.find((x) => x.Documento === "RE 2026/12").linhas.length, 0);
  assert.equal(r.rejeitados.length, 0);
  const g = B.paraGravar(r.registos, "recibos", "IMP-1");
  assert.equal(g.linhas.length, 3); assert.equal(g.liquidacoes.length, 3);
  assert.ok(!("linhas" in g.linhas[0]) && !("_linha" in g.linhas[0]));
  assert.equal(g.liquidacoes[0].ID_Liquidacao, "SIR MOTORS|RE|2026|10>SIR MOTORS|FA|2026|101");
});

test("importação ↔ base mãe: novos, iguais, alterados (com campos) e ausentes", () => {
  const imp = B.lerListagem(FATURAS, {tipo: "faturas", empresa: "SIR Motors"}).registos;
  const base = [
    {...imp[0]},                                                         // igual (mesmo hash)
    {...imp[1], Total: "110000", Data_Vencimento: 46336, Hash: "antigo"}, // alterado: total e vencimento (data Excel)
    {ID_Documento: "SIR MOTORS|FA|2026|100", Empresa: "SIR Motors", Tipo_Documento: "FA", Documento: "FA 2026/100", Data: "2026-11-03", Total: 5000, Estado: "Normal"}, // ausente
    {ID_Documento: "SIR MOTORS|FA|2025|9", Empresa: "SIR Motors", Tipo_Documento: "FA", Data: "2025-01-03", Total: 5000, Estado: "Normal"}, // fora do período
  ];
  const d = B.comparar(imp, base, "faturas");
  assert.equal(d.iguais.length, 1);
  assert.equal(d.alterados.length, 1);
  assert.deepEqual(d.alterados[0]._mudancas.map((m) => m.campo).sort(), ["Data_Vencimento", "Total"]);
  assert.equal(d.alterados[0]._mudancas.find((m) => m.campo === "Total").antes, 110000);
  assert.equal(d.novos.length, imp.length - 2);
  assert.deepEqual(d.ausentes.map((a) => a.Documento), ["FA 2026/100"]);
});

test("reconciliação interna: saldos, estados, créditos, excepções e antiguidade por cliente", () => {
  const faturas = B.lerListagem(FATURAS, {tipo: "faturas", empresa: "SIR Motors"}).registos;
  const rec = B.lerListagem(RECIBOS, {tipo: "recibos", empresa: "SIR Motors"}).registos;
  const g = B.paraGravar(rec, "recibos", "IMP");
  const r = B.reconciliar({faturas, recibos: g.linhas, liquidacoes: g.liquidacoes, hoje: "2026-12-25"});
  const doc = (n) => r.documentos.find((x) => x.Documento === n);
  assert.equal(doc("FA 2026/101").estado, "Liquidado");
  assert.equal(doc("FA 2026/102").saldo, 60000); assert.equal(doc("FA 2026/102").estado, "Parcial vencido"); assert.equal(doc("FA 2026/102").diasAtraso, 35);
  assert.equal(doc("FA 2026/103").estado, "Vencido"); assert.equal(doc("FA 2026/103").diasAtraso, 49);
  assert.equal(doc("NC 2026/7").estado, "Crédito");
  assert.equal(doc("FR 2026/55").estado, "Liquidado");
  assert.equal(doc("FA 2026/104"), undefined); // anulada fora
  const tipos = r.excepcoes.map((e) => e.tipo);
  assert.ok(tipos.includes("Liquidação de documento inexistente na base"));
  assert.ok(tipos.some((t) => /Recibo sem documento liquidado/.test(t)));
  const omega = r.clientes.find((c) => c.Codigo_Cliente === "C002");
  assert.equal(omega.emAberto, 60000); assert.equal(omega.creditos, 20000); assert.equal(omega.saldoLiquido, 40000); assert.equal(omega.d60, 60000);
  const beta = r.clientes.find((c) => c.Codigo_Cliente === "C001");
  assert.equal(beta.emAberto, 50000); assert.equal(beta.d60, 50000); assert.equal(beta.ultimoRecibo, "2026-11-04");
  assert.equal(r.clientes.find((c) => c.Codigo_Cliente === "C004").creditos, 25000);
  assert.equal(r.clientes[0].emAberto >= r.clientes[1].emAberto, true);
});

test("recibos ↔ créditos do extracto: transferências emparelhadas, numerário excluído", () => {
  const rec = B.lerListagem(RECIBOS, {tipo: "recibos", empresa: "SIR Motors"}).registos;
  const contra = B.recibosParaBanco(B.paraGravar(rec, "recibos", "I").linhas, {de: "2026-11-01", ate: "2026-11-30"});
  assert.deepEqual(contra.map((x) => x.requisicao), ["RE 2026/10", "RE 2026/11"]);
  const movs = [{id: "m1", data: "2026-11-04", valor: 300000, descricao: "TRF CLIENTE BETA LDA TRF889", referencia: ""},
    {id: "m2", data: "2026-11-13", valor: 100000, descricao: "DEP OMEGA", referencia: ""}, {id: "m3", data: "2026-11-05", valor: -5, descricao: "x", referencia: ""}];
  const r = C.conciliar(movs, contra, {sentido: "credito"});
  assert.equal(r.pares.find((p) => p.mov === "m1").estado, "automatico");
  assert.equal(r.pares.find((p) => p.mov === "m2").pags[0], "R:SIR MOTORS|RE|2026|11");
  assert.ok(!r.pares.some((p) => p.mov === "m3"));
});
