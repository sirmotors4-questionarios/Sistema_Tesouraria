// Listagens PRIMAVERA no formato real (relatório Crystal «Documentos Emitidos de Vendas» e «Recebimentos»). Dados fictícios.
const test = require("node:test");
const assert = require("node:assert/strict");
globalThis.window = globalThis;
require("../conciliacao-core.js");
require("../faturacao-core.js");
const B = globalThis.TES_BASE;
const EMP = ["SIR", "Indigena", "Metrobus", "Cometal", "Hotel Inhambane"];
const v = Array(13).fill("");
const linha = (o) => Object.assign([...v], o);
const RELATORIO = [
  linha({0: "METRO-BUS, SA", 10: "25/11/2026 | Pág. 1/5"}), linha({}), linha({0: "Documentos Emitidos de Vendas (2026-01-01 até 2026-11-25)", 10: "Valores em MT"}), linha({}), linha({}),
  ["Entidade", "Nome", "Documento", "", "Data ", "Total Merc.", "Total IVA", "Total Desc.", "Total Outros", "", "", "Impostos", "Total Doc."], linha({}), linha({}),
  ["CL000100", "Cliente A", "DV 2026/1", "", 46173, -5250, 0, 0, 0, "", "", 0, -5250], linha({}), linha({}),
  ["CLOOO101", "Cliente B", "DV 2026/2", "", 46173, -3500, 0, 0, 0, "", "", 0, -3500], linha({}),
  linha({0: "Total  Devolução a Dinheiro", 5: -8750, 6: 0, 7: 0, 8: 0, 11: 0, 12: -8750}), linha({}),
  ["CL00034", "Cliente C", "FA 2026/1", "", 46030, 100000, 16000, 0, 0, "", "", 0, 116000], linha({}),
  ["CL00034", "Cliente C", "FA 2026/2", "", 46061, 100000, 16000, 0, 0, "", "", 0, 116000], linha({}),
  linha({0: "Total  Factura", 5: 200000, 6: 32000, 12: 232000}), linha({}),
  ["CL00137", "Varios Clientes", "VD 2026/1", "", 46100, 3861, 0, 0, 0, "", "", 0, 3861], linha({}),
  ["CL00137", "Varios Clientes", "XYZ 2026/1", "", 46101, -100, 0, 0, 0, "", "", 0, -100], linha({}),
  linha({0: "Total  Venda-a-Dinheiro", 12: 3861}), linha({0: "Total  Xyz", 12: -100}), linha({0: "Total Geral", 12: 227011}), linha({}),
  linha({0: "© PRIMAVERA BSS / Licença"}),
];
const cabRec = [" ", "Documento", "Série", "Número", "Filial", "Data", "Valor ", "Valor Desconto", "Tipo Entidade", "Entidade", "Nome", "Moeda", "Imp", "Estorno", "Motivo", "Tipo Doc. Original", "Série Doc. Original", "Número Doc. Orig."];
const RECIBOS = [cabRec,
  ["", "VDR", "2026", "1", "000", 46100, 3861, 0, "C", "CL00137", "Varios Clientes", "MT", "", "", "", "", "", ""],
  ["", "NPC", "2026", "1", "000", 46173, 5250, 0, "C", "CL000100", "Cliente A", "MT", "", "", "", "", "", ""],
  ["", "RE", "2026", "5", "000", 46070, 116000, 0, "C", "CL00034", "Cliente C", "MT", "", "", "C04", "", "", ""],
  ["", "RE", "2026", "6", "000", 46080, 50000, 0, "C", "CL00034", "Cliente C", "MT", "", "", "", "FA", "2026", "2"],
  ["", "VDR", "2026", "2", "000", 46101, 0, 0, "C", "CL00137", "Varios Clientes", "MT", "", "", "", "", "", ""],
];

test("relatório «Documentos Emitidos de Vendas»: empresa do cabeçalho, todos os documentos, totais de controlo", () => {
  const r = B.lerListagem(RELATORIO, {tipo: "faturas", empresas: EMP, nomeFicheiro: "Documentos Emitidos de Vendas (2026-01-01 até 2026-11-25) -Metrobus.xls"});
  assert.equal(r.rejeitados.length, 0, JSON.stringify(r.rejeitados));
  assert.equal(r.registos.length, 6);
  assert.equal(r.meta.empresaDetectada, "Metrobus");
  assert.deepEqual(r.meta.periodo, {de: "2026-01-01", ate: "2026-11-25"});
  assert.ok(r.registos.every((x) => x.Empresa === "Metrobus"));
  const fa = r.registos.find((x) => x.Documento === "FA 2026/1");
  assert.deepEqual([fa.Total, fa.Total_Mercadoria, fa.Total_IVA, fa.Sinal], [116000, 100000, 16000, 1]);
  assert.equal(r.registos.find((x) => x.Documento === "DV 2026/2").Codigo_Cliente, "CL000101", "CLOOO101 corrigido");
  assert.equal(r.registos.find((x) => x.Tipo_Documento === "XYZ").Sinal, -1, "tipo desconhecido: efeito pelo sinal");
  assert.equal(r.controlo.ok, true, JSON.stringify(r.controlo));
  assert.equal(r.controlo.linhas.find((x) => /Geral/.test(x.rotulo)).soma, 227011);
  assert.ok(r.avisos.some((a) => /XYZ/.test(a)) && r.avisos.some((a) => /CLOOO101 → CL000101/.test(a)));
});

test("totais de controlo apanham documentos em falta", () => {
  const sem = RELATORIO.filter((l) => l[2] !== "FA 2026/2");
  const r = B.lerListagem(sem, {tipo: "faturas", empresas: EMP});
  assert.equal(r.controlo.ok, false);
  assert.ok(r.avisos.some((a) => /não batem/.test(a)));
});

test("«Recebimentos»: tipo na coluna Documento + Série + Número; VDR, NPC, valor 0 e Doc. Original", () => {
  const r = B.lerListagem(RECIBOS, {tipo: "recibos", empresas: EMP, nomeFicheiro: "Recebimentos- Metrobus.xlsx"});
  assert.equal(r.rejeitados.length, 0, JSON.stringify(r.rejeitados));
  assert.deepEqual(r.registos.map((x) => x.Documento), ["VDR 2026/1", "NPC 2026/1", "RE 2026/5", "RE 2026/6", "VDR 2026/2"]);
  assert.equal(r.meta.empresaDetectada, "Metrobus");
  assert.equal(r.registos.find((x) => x.Documento === "RE 2026/5").Motivo, "C04");
  assert.deepEqual(r.registos.find((x) => x.Documento === "RE 2026/6").linhas.map((l) => [l.Documento_Liquidado, l.Valor]), [["FA 2026/2", 50000]]);
  assert.ok(r.avisos.some((a) => /valor 0/.test(a)));
});

test("ficheiro sem empresa: deduzida pelos clientes já facturados; sem base pede a empresa", () => {
  const fat = B.lerListagem(RELATORIO, {tipo: "faturas", empresas: EMP}).registos;
  const rec = RECIBOS.map((l) => l.slice());
  const r = B.lerListagem(rec, {tipo: "recibos", empresas: EMP, nomeFicheiro: "Recebimentos.xlsx", faturasBase: fat});
  assert.equal(r.meta.empresaDetectada, "Metrobus");
  assert.equal(r.registos.length, 5);
  const s = B.lerListagem(rec, {tipo: "recibos", empresas: EMP, nomeFicheiro: "Recebimentos.xlsx"});
  assert.equal(s.meta.precisaEmpresa, true);
  assert.equal(s.rejeitados.length, 0);
  assert.match(s.avisos[0], /não indica a empresa/);
});

test("reconciliação: VDR→VD, NPC→DV, RE ao único/mais antigo documento; NPC não é crédito nem vai ao banco", () => {
  const fat = B.lerListagem(RELATORIO, {tipo: "faturas", empresas: EMP}).registos;
  const rec = B.lerListagem(RECIBOS, {tipo: "recibos", empresas: EMP, nomeFicheiro: "Recebimentos- Metrobus.xlsx"}).registos;
  const g = B.paraGravar(rec, "recibos", "I");
  const res = B.reconciliar({faturas: fat, recibos: g.linhas, liquidacoes: g.liquidacoes, hoje: "2026-11-25"});
  const reg = (doc) => res.inferidas.find((x) => x.ID_Recibo.endsWith(doc))?.regra || "";
  assert.match(reg("VDR|2026|1"), /VDR → VD/);
  assert.match(reg("NPC|2026|1"), /NPC → DV/);
  assert.match(reg("RE|2026|5"), /único documento/, "a FA 2026/2 já tem 50 000 pagos pelo RE 6: só a FA 2026/1 tem saldo 116 000");
  assert.equal(res.documentos.find((d) => d.Documento === "FA 2026/1").estado, "Liquidado");
  const a = res.clientes.find((c) => c.Codigo_Cliente === "CL000100");
  assert.equal(a.creditos, 0, "DV devolvida pelo NPC deixa de ser crédito");
  assert.ok(!B.recibosParaBanco(g.linhas, {}).some((x) => /NPC/.test(x.requisicao)), "NPC é saída, não crédito do banco");
});

test("vários documentos com o mesmo valor (rendas): o recibo paga o mais antigo, marcado para rever", () => {
  const fat = B.lerListagem(RELATORIO, {tipo: "faturas", empresas: EMP}).registos;
  const rec = B.lerListagem([cabRec, ["", "RE", "2026", "9", "000", 46090, 116000, 0, "C", "CL00034", "Cliente C", "MT", "", "", "", "", "", ""]], {tipo: "recibos", empresas: EMP, nomeFicheiro: "Recebimentos- Metrobus.xlsx"}).registos;
  const res = B.reconciliar({faturas: fat, recibos: B.paraGravar(rec, "recibos", "I").linhas, liquidacoes: [], hoje: "2026-11-25"});
  assert.equal(res.inferidas.length, 1);
  assert.equal(res.inferidas[0].Documento_Liquidado, "FA 2026/1");
  assert.match(res.inferidas[0].regra, /mais antigo de 2 documentos.*rever/);
});
