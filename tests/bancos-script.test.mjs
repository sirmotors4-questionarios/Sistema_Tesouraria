import test from "node:test";
import assert from "node:assert/strict";
import {createRequire} from "node:module";
import {livro, chamar} from "./livro-excel-mock.mjs";
const require = createRequire(import.meta.url);
const C = require("../conciliacao-core.js");

// Extracto de teste (fictício) com duas comissões iguais no mesmo dia.
const HASH_A = "a".repeat(64), HASH_B = "b".repeat(64);
const MOV = [
  {data: "2026-03-02", descricao: "Fecho TPA", referencia: "1", valor: 22398, saldo: 22398},
  {data: "2026-03-02", descricao: "Comissão TPA", referencia: "2", valor: -167.99, saldo: 22230.01},
  {data: "2026-03-02", descricao: "Comissão TPA", referencia: "3", valor: -167.99, saldo: 22062.02},
  {data: "2026-03-05", descricao: "Pagamento fornecedor", referencia: "4", valor: -15000, saldo: 7062.02},
];
const CONTA = {banco: "BCI", conta: "17724771310001", idConta: "f0fbd887"};
function pedidoExtrato(movs, hash, {final = true, lote = movs} = {}) {
  const ids = C.idsMovimentos(movs, CONTA);
  const linhas = C.movimentosParaBase(movs, ids, {...CONTA, idExtrato: hash});
  const sel = lote === movs ? linhas : linhas.filter((l) => lote.some((m) => m === movs[linhas.indexOf(l)]));
  return {operacao: "banco.importar_extrato", id_pedido: "P-" + hash.slice(0, 4), final,
    extrato: {ID_Extrato: hash, Nome_Original: "trial1.csv", Nome_Padronizado: C.nomePadronizado({...CONTA, de: movs[0].data, ate: movs.at(-1).data, extensao: "csv"}),
      Banco: "BCI", ID_Conta: CONTA.idConta, Conta: CONTA.conta, Formato: "CSV", De: movs[0].data, Ate: movs.at(-1).data, Movimentos_Ficheiro: movs.length},
    movimentos: sel};
}

test("importar extracto: grava movimentos e regista o ficheiro com nome padronizado", () => {
  const w = livro();
  const r = chamar(w, pedidoExtrato(MOV, HASH_A));
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.novos, 4);
  assert.equal(r.estado, "Completo");
  const m = w.ler("tbl_MovimentosBancarios");
  assert.equal(m.length, 4);
  assert.deepEqual(m.map((x) => x.ID_Movimento).slice(1, 3), ["BCI|17724771310001|2026-03-02|-16799|1", "BCI|17724771310001|2026-03-02|-16799|2"]);
  assert.equal(m[0].Conta, "17724771310001", "conta fica como texto");
  const e = w.ler("tbl_Extratos");
  assert.equal(e.length, 1);
  assert.equal(e[0].Nome_Padronizado, "BCI_17724771310001_2026-03-02_2026-03-05.csv");
  assert.equal(e[0].Novos, 4);
  assert.equal(e[0].Ja_Existentes, 0);
});

test("o mesmo ficheiro não é importado duas vezes", () => {
  const w = livro();
  chamar(w, pedidoExtrato(MOV, HASH_A));
  const r = chamar(w, pedidoExtrato(MOV, HASH_A));
  assert.equal(r.sucesso, false);
  assert.equal(r.codigo, "EXTRATO_JA_IMPORTADO");
  assert.equal(w.ler("tbl_MovimentosBancarios").length, 4);
});

test("outro ficheiro com período sobreposto só acrescenta os movimentos novos", () => {
  const w = livro();
  chamar(w, pedidoExtrato(MOV.slice(0, 3), HASH_A));
  const r = chamar(w, pedidoExtrato(MOV, HASH_B));
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.novos, 1);
  assert.equal(r.duplicados, 3);
  assert.equal(w.ler("tbl_MovimentosBancarios").length, 4);
  const e = w.ler("tbl_Extratos").find((x) => x.ID_Extrato === HASH_B);
  assert.equal(e.Ja_Existentes, 3);
  // Um terceiro ficheiro sem nada novo fica registado, sem gravar movimentos
  const r3 = chamar(w, pedidoExtrato(MOV.slice(1, 2), "c".repeat(64)));
  assert.equal(r3.estado, "Sem movimentos novos");
  // Mesmo período e conta, conteúdo diferente: nome único
  assert.equal(r.nome_padronizado, "BCI_17724771310001_2026-03-02_2026-03-05.csv");
  const r4 = chamar(w, pedidoExtrato(MOV, "d".repeat(64)));
  assert.equal(r4.nome_padronizado, "BCI_17724771310001_2026-03-02_2026-03-05_v2.csv");
  assert.equal(w.ler("tbl_MovimentosBancarios").length, 4);
});

test("extracto em lotes: estado Parcial até ao último lote; retomar não duplica", () => {
  const w = livro();
  const p1 = pedidoExtrato(MOV, HASH_A, {final: false}); p1.movimentos = p1.movimentos.slice(0, 2);
  assert.equal(chamar(w, p1).estado, "Parcial");
  assert.equal(chamar(w, p1).novos, 0, "repetir o lote parcial não duplica");
  const p2 = pedidoExtrato(MOV, HASH_A, {final: true}); p2.movimentos = p2.movimentos.slice(2);
  const r = chamar(w, p2);
  assert.equal(r.estado, "Completo");
  assert.equal(w.ler("tbl_MovimentosBancarios").length, 4);
  assert.equal(w.ler("tbl_Extratos")[0].Novos, 4);
});

test("validação: conta diferente, movimento repetido e hash inválido são recusados sem gravar", () => {
  const w = livro();
  const p = pedidoExtrato(MOV, HASH_A);
  p.movimentos[1] = {...p.movimentos[1], Conta: "2308710007"};
  assert.equal(chamar(w, p).codigo, "VALIDACAO");
  const q = pedidoExtrato(MOV, HASH_A);
  q.movimentos.push({...q.movimentos[0]});
  assert.match(chamar(w, q).mensagem, /repetido/);
  const h = pedidoExtrato(MOV, "123");
  assert.equal(chamar(w, h).codigo, "PEDIDO_INVALIDO");
  assert.equal(w.getTable("tbl_MovimentosBancarios"), undefined);
});

const lig = (mov, tipo, doc, valor, banco) => ({ID_Movimento: mov, Tipo_Documento: tipo, ID_Documento: doc, Documento: doc, Valor: valor, Valor_Banco: banco, Data_Banco: "2026-03-05", Conta: "17724771310001", Metodo: "confirmado"});

test("conciliar: grava, é idempotente e não deixa um documento em dois movimentos", () => {
  const w = livro();
  chamar(w, pedidoExtrato(MOV, HASH_A));
  const ids = C.idsMovimentos(MOV, CONTA);
  const p = {operacao: "banco.conciliar", ligacoes: [lig(ids[3], "despesa", "D-1", 10000, -15000), lig(ids[3], "despesa", "D-2", 5000, -15000)]};
  const r = chamar(w, p);
  assert.equal(r.sucesso, true, r.mensagem);
  assert.equal(r.ligacoes, 2);
  assert.equal(chamar(w, p).ligacoes, 0, "repetir não duplica");
  assert.equal(w.ler("tbl_ConciliacaoBancaria").length, 2);
  const outro = chamar(w, {operacao: "banco.conciliar", ligacoes: [lig(ids[1], "despesa", "D-1", 167.99, -167.99)]});
  assert.equal(outro.codigo, "CONFLITO");
  assert.match(outro.mensagem, /DOCUMENTO_JA_CONCILIADO/);
  const mudar = chamar(w, {operacao: "banco.conciliar", ligacoes: [lig(ids[3], "despesa", "D-9", 15000, -15000)]});
  assert.match(mudar.mensagem, /MOVIMENTO_JA_CONCILIADO/);
  assert.equal(w.ler("tbl_ConciliacaoBancaria").length, 2);
});

test("conciliar: soma diferente do banco, movimento fora da base; desfazer liberta", () => {
  const w = livro();
  chamar(w, pedidoExtrato(MOV, HASH_A));
  const ids = C.idsMovimentos(MOV, CONTA);
  assert.equal(chamar(w, {operacao: "banco.conciliar", ligacoes: [lig(ids[3], "despesa", "D-1", 14000, -15000)]}).codigo, "VALIDACAO");
  assert.equal(chamar(w, {operacao: "banco.conciliar", ligacoes: [lig("BCI|1|2026-01-01|-100|1", "despesa", "D-1", 1, -1)]}).codigo, "MOVIMENTO_INEXISTENTE");
  assert.equal(chamar(w, {operacao: "banco.conciliar", ligacoes: [lig(ids[0], "recibo", "SIR|RE|2026|10", 22398, 22398)]}).sucesso, true);
  assert.equal(w.ler("tbl_RecebimentosBanco").length, 1, "recibos também ficam em tbl_RecebimentosBanco");
  const d = chamar(w, {operacao: "banco.desfazer", movimentos: [ids[0]]});
  assert.equal(d.ligacoes, 1);
  assert.equal(w.ler("tbl_RecebimentosBanco").length, 0);
  assert.equal(chamar(w, {operacao: "banco.conciliar", ligacoes: [lig(ids[0], "recibo", "SIR|RE|2026|11", 22398, 22398)]}).sucesso, true);
});

test("movimentosDaBase: converte as linhas do Excel e exclui os já conciliados", () => {
  const ids = C.idsMovimentos(MOV, CONTA);
  const linhas = C.movimentosParaBase(MOV, ids, CONTA).map((l) => ({...l, Data_Movimento: 46083 + (l.Data_Movimento === "2026-03-05" ? 3 : 0)}));
  const m = C.movimentosDaBase(linhas, {conta: "0017724771310001", excluir: new Set([ids[0]])});
  assert.equal(m.length, 3);
  assert.equal(m[0].data, "2026-03-02");
  assert.equal(m[2].data, "2026-03-05");
  assert.equal(m[2].valor, -15000);
});

test("base.consultar: devolve as tabelas (vazias antes da primeira gravação) sem criar nada", () => {
  const w = livro();
  const vazio = chamar(w, {operacao: "base.consultar", tabelas: ["tbl_Extratos", "tbl_MovimentosBancarios"]});
  assert.equal(vazio.sucesso, true);
  assert.deepEqual(JSON.parse(vazio.tabelas_json).map((t) => t.linhas.length), [0, 0]);
  assert.equal(w.getTable("tbl_Extratos"), undefined, "consultar não cria tabelas");
  chamar(w, pedidoExtrato(MOV, HASH_A));
  const r = chamar(w, {operacao: "base.consultar"});
  const m = JSON.parse(r.tabelas_json).find((t) => t.tabela === "tbl_MovimentosBancarios");
  assert.equal(m.linhas.length, 4);
  assert.equal(m.colunas[0], "ID_Movimento");
  assert.equal(chamar(w, {operacao: "base.consultar", tabelas: ["xpto"]}).codigo, "PEDIDO_INVALIDO");
});

test("rapidez: vários lotes sem linhas vazias nem duplicados, poucas leituras, contas em formato texto", () => {
  const w = livro();
  const muitos = Array.from({length: 2400}, (_, i) => ({data: "2026-0" + (1 + (i % 9)) + "-" + String(10 + (i % 18)), descricao: "m" + i, referencia: String(i), valor: -(i + 1), saldo: null}));
  const p = pedidoExtrato(muitos, HASH_A, {final: false});
  const lotes = [p.movimentos.slice(0, 1000), p.movimentos.slice(1000, 2000), p.movimentos.slice(2000)];
  lotes.forEach((l, i) => {
    const antes = w.cont.leituras;
    const r = chamar(w, {...p, final: i === 2, movimentos: l});
    assert.equal(r.sucesso, true, r.mensagem);
    assert.ok(w.cont.leituras - antes <= 9, "leituras por lote (inclui a reserva do livro): " + (w.cont.leituras - antes));
  });
  const m = w.ler("tbl_MovimentosBancarios");
  assert.equal(m.length, 2400);
  assert.equal(new Set(m.map((x) => x.ID_Movimento)).size, 2400);
  assert.equal(w.getTable("tbl_MovimentosBancarios").rows.length, 2400, "sem linhas vazias no corpo");
  const f = w.getWorksheet("BM_Movimentos"), jConta = w.getTable("tbl_MovimentosBancarios").headers.indexOf("Conta");
  assert.ok([1, 1200, 2400].every((r) => f.formatos[r + ":" + jConta] === "@"), "Conta em formato texto");
  const e = w.ler("tbl_Extratos")[0];
  assert.equal(e.Estado, "Completo"); assert.equal(e.Novos, 2400);
  // Repetir o 2.º lote (como depois de um tempo esgotado) não duplica
  assert.equal(chamar(w, {...p, final: false, movimentos: lotes[1], extrato: {...p.extrato, ID_Extrato: HASH_B}}).novos, 0);
  assert.equal(w.ler("tbl_MovimentosBancarios").length, 2400);
});

test("reserva do livro: outra gravação em curso → EM_USO sem gravar; reserva libertada no fim e expirada após 150 s", () => {
  const w = livro();
  const p = {...pedidoExtrato(MOV, HASH_A), id_pedido: "A"};
  assert.equal(chamar(w, p).sucesso, true);
  const controlo = w.getWorksheet("BM_Controlo");
  assert.equal(controlo.celulas["A2:C2"][0], "", "reserva libertada no fim");
  // Simula outra execução a meio (reserva recente de outro pedido)
  controlo.celulas["A2:C2"] = ["B", Date.now(), "banco.importar_extrato"];
  const r = chamar(w, {...pedidoExtrato(MOV, HASH_B), id_pedido: "C"});
  assert.equal(r.codigo, "EM_USO");
  assert.equal(w.ler("tbl_Extratos").length, 1, "nada gravado");
  assert.equal(controlo.celulas["A2:C2"][0], "B", "não mexe na reserva alheia");
  // Reserva antiga (script interrompido): ignorada
  controlo.celulas["A2:C2"] = ["B", Date.now() - 200000, "banco.importar_extrato"];
  assert.equal(chamar(w, {...pedidoExtrato(MOV, HASH_B), id_pedido: "C"}).sucesso, true);
  // Consultar não reserva nem espera
  controlo.celulas["A2:C2"] = ["B", Date.now(), "banco.conciliar"];
  assert.equal(chamar(w, {operacao: "base.consultar"}).sucesso, true);
});

test("diagnóstico do parâmetro «pedido»: vazio, texto, sem operação, embrulhado e codificado duas vezes", async () => {
  const {livro: novoLivro} = await import("./livro-excel-mock.mjs");
  const w = novoLivro();
  const bruto = (x) => JSON.parse(JSON.stringify(globalThis.__mainBaseMae(w, x)));
  assert.equal(bruto("").codigo, "PEDIDO_VAZIO");
  assert.match(bruto("").mensagem, /string\(outputs\('Pedido'\)\)/);
  assert.equal(bruto("string(outputs('Pedido'))").codigo, "PEDIDO_INVALIDO");
  assert.equal(bruto("{}").codigo, "PEDIDO_SEM_OPERACAO");
  assert.equal(bruto(JSON.stringify({body: {operacao: "base.consultar"}})).sucesso, true);
  assert.equal(bruto(JSON.stringify(JSON.stringify({operacao: "base.consultar"}))).sucesso, true);
});

test("base.consultar devolve também a tblPagamentos do livro (só leitura), sem a criar quando não existe", () => {
  const w = livro();
  let r = chamar(w, {operacao: "base.consultar"});
  assert.equal(JSON.parse(r.tabelas_json).some((t) => t.tabela === "tblPagamentos"), false);
  assert.equal(w.getTable("tblPagamentos"), undefined);
  const f = w.addWorksheet("Pagamentos"), cab = ["ID_Pagamento", "ID_Despesa", "Data_Pagamento", "Meio_Pagamento", "Referencia_Transaccao", "Valor_Pagamento"];
  f.getRangeByIndexes(0, 0, 1, cab.length).setValues([cab]);
  const t = f.addTable(f.getRangeByIndexes(0, 0, 1, cab.length), true); t.setName("tblPagamentos");
  w.getTable("tblPagamentos").rows = [["P1", "D1", 46087, "Cheque", "45564600", 20000]];
  r = chamar(w, {operacao: "base.consultar"});
  const tp = JSON.parse(r.tabelas_json).find((x) => x.tabela === "tblPagamentos");
  assert.deepEqual(tp.colunas, cab);
  assert.deepEqual(tp.linhas, [["P1", "D1", 46087, "Cheque", "45564600", 20000]]);
});
