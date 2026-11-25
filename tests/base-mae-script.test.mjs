import test from "node:test";
import assert from "node:assert/strict";
import {createRequire} from "node:module";
const require = createRequire(import.meta.url);
const B = require("../faturacao-core.js");

import {livro, chamar} from "./livro-excel-mock.mjs";

const FATURAS = [["Documento", "Data", "Data Venc.", "Entidade", "Nome", "NUIT", "Total Documento"],
  ["FA 2026/101", "02/11/2026", "02/12/2026", "C001", "Cliente Beta", "0400123456", "300.000,00"],
  ["FA 2026/102", "05/11/2026", "20/11/2026", "C002", "Cliente Omega", "400999888", "120.000,00"]];
const RECIBOS = [["Recibo", "Data", "Cliente", "Nome", "Documento Liquidado", "Valor Liquidado"],
  ["RE 2026/10", "04/11/2026", "C001", "Cliente Beta", "FA 2026/101", "300.000,00"],
  ["RE 2026/11", "12/11/2026", "C002", "Cliente Omega", "FA 2026/102", "60.000,00"]];
const pedido = (tipo, matriz, id = "IMP-1") => {
  const r = B.lerListagem(matriz, {tipo, empresa: "SIR"});
  const g = B.paraGravar(r.registos, tipo, id);
  return {operacao: "base.importar", id_pedido: "p-" + id, tipo, linhas: g.linhas, liquidacoes: g.liquidacoes || [],
    importacao: {ID_Importacao: id, Empresa: "SIR", Ficheiro: "x.xlsx", Utilizador: "rh@sir", Linhas_Ficheiro: matriz.length - 1, Novos: r.registos.length}};
};

test("primeira importação cria as tabelas e grava datas como data e NUIT como texto", () => {
  const w = livro();
  const r = chamar(w, pedido("faturas", FATURAS));
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.novos, 2); assert.equal(r.alterados, 0);
  const f = w.ler("tbl_Faturas");
  assert.equal(f.length, 2); // a linha vazia da tabela nova foi aproveitada
  assert.equal(f[0].ID_Documento, "SIR|FA|2026|101");
  assert.equal(f[0].Data, 46328); // 02/11/2026 em série Excel
  assert.equal(f[0].NUIT, "0400123456");
  assert.equal(f[0].Total, 300000);
  const folha = w.folhas.get("BM_Faturas");
  assert.equal(folha.formatos["1:6"], "dd/mm/yyyy"); // coluna Data
  assert.equal(folha.formatos["1:10"], "@");
  assert.equal(w.ler("tbl_Importacoes")[0].ID_Importacao, "IMP-1");
});

test("reimportar: alterados reescritos por blocos, novos num só addRows, importação sem duplicar", () => {
  const w = livro();
  chamar(w, pedido("faturas", FATURAS));
  const m2 = FATURAS.map((l) => l.slice());
  m2[2][6] = "110.000,00";
  m2.push(["FA 2026/103", "06/11/2026", "", "C001", "Cliente Beta", "", "50.000,00"]);
  const antes = w.cont.escritas;
  const r = chamar(w, pedido("faturas", m2));
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.novos, 1); assert.equal(r.alterados, 2);
  const f = w.ler("tbl_Faturas");
  assert.equal(f.length, 3); assert.equal(f[1].Total, 110000);
  assert.ok(w.cont.escritas - antes < 20, "escritas: " + (w.cont.escritas - antes));
  chamar(w, pedido("faturas", m2));
  assert.equal(w.ler("tbl_Importacoes").length, 1); // mesmo ID_Importacao: actualiza, não duplica
});

test("recibos: grava liquidações e substitui-as quando o recibo é reimportado", () => {
  const w = livro();
  const r = chamar(w, pedido("recibos", RECIBOS));
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.liquidacoes, 2);
  const corr = RECIBOS.map((l) => l.slice());
  corr[2][5] = "70.000,00";
  chamar(w, pedido("recibos", corr, "IMP-2"));
  const liq = w.ler("tbl_Liquidacoes");
  assert.equal(liq.length, 2);
  assert.equal(liq.find((x) => x.ID_Recibo === "SIR|RE|2026|11").Valor, 70000);
  assert.equal(w.ler("tbl_Recibos").find((x) => x.ID_Recibo === "SIR|RE|2026|11").Valor, 70000);
});

test("validação antes de escrever: chave repetida, coluna desconhecida, liquidação de outro recibo", () => {
  const w = livro();
  const p = pedido("faturas", FATURAS);
  p.linhas.push({...p.linhas[0]});
  let r = chamar(w, p);
  assert.equal(r.sucesso, false); assert.equal(r.codigo, "VALIDACAO"); assert.match(r.mensagem, /chave repetida/);
  assert.equal(w.tabelas.size, 0);
  const q = pedido("faturas", FATURAS); q.linhas[0].Inventada = 1;
  assert.match(chamar(w, q).mensagem, /coluna desconhecida: Inventada/);
  const s = pedido("recibos", RECIBOS); s.liquidacoes.push({ID_Recibo: "SIR|RE|2026|99", ID_Documento: "X", Valor: 1});
  assert.match(chamar(w, s).mensagem, /recibo fora do pedido/);
  assert.equal(chamar(w, {operacao: "base.apagar"}).codigo, "OPERACAO_DESCONHECIDA");
});

test("ligações recibo ↔ banco: upsert por ID_Ligacao", () => {
  const w = livro();
  const l = {ID_Ligacao: "SIR|RE|2026|10@2026-11-04|30000000", ID_Recibo: "SIR|RE|2026|10", Data_Banco: "2026-11-04", Valor: 300000, Descritivo_Banco: "TRF BETA", Conta: "0001234567", Utilizador: "t", Data: "2026-11-25T09:00:00Z"};
  assert.equal(chamar(w, {operacao: "base.ligar_banco", ligacoes: [l]}).novos, 1);
  assert.equal(chamar(w, {operacao: "base.ligar_banco", ligacoes: [{...l, Conta: "X"}]}).alterados, 1);
  assert.equal(w.ler("tbl_RecebimentosBanco").length, 1);
});

test("estrutura: tabela existente de uma versão anterior recebe as colunas novas, sem perder dados", () => {
  const w = livro();
  const f = w.addWorksheet("Manual"); f.getRangeByIndexes(0, 0, 1, 2).setValues([["ID_Documento", "Total"]]);
  f.addTable().setName("tbl_Faturas");
  w.getTable("tbl_Faturas").rows = [["ANTIGO|FA|1|1", 123]];
  const r = chamar(w, pedido("faturas", FATURAS));
  assert.equal(r.sucesso, true, r.mensagem);
  const t = w.getTable("tbl_Faturas");
  assert.ok(["Empresa", "Total_IVA", "Sinal", "Tipo_Nome"].every((c) => t.headers.includes(c)));
  assert.deepEqual(t.rows[0].slice(0, 2), ["ANTIGO|FA|1|1", 123], "linha antiga intacta");
});
