/**
 * Office Script "TesourariaBaseMae": grava a base mãe de Facturação/Recebimentos no livro da Tesouraria.
 * Chamado pelo fluxo de escrita V9 quando `operacao` começa por "base." ou "banco." (parâmetro único: pedido = corpo JSON).
 *
 *  base.importar      — upsert validado de facturas ou recibos (+ liquidações) importados do PRIMAVERA, por chave
 *                       (ID_Documento / ID_Recibo). Linhas novas entram num só addRows; alteradas são reescritas por
 *                       blocos de linhas seguidas; as liquidações dos recibos importados são substituídas. Regista a
 *                       importação em tbl_Importacoes (upsert por ID_Importacao: repetir o mesmo lote não duplica).
 *  base.ligar_banco   — regista ligações recibo ↔ crédito do extracto em tbl_RecebimentosBanco (upsert por ID_Ligacao).
 *
 *  Bancos (a base mãe nunca tem movimentos nem conciliações duplicados):
 *  banco.importar_extrato — grava os movimentos de um extracto em tbl_MovimentosBancarios. Só ACRESCENTA: um movimento
 *                       cuja chave (ID_Movimento = BANCO|CONTA|data|cêntimos|ocorrência) já existe é ignorado, venha do
 *                       mesmo ficheiro, de outro formato (CSV/Excel/PDF) ou de um extracto de período sobreposto.
 *                       O ficheiro é registado em tbl_Extratos pelo SHA-256 (ID_Extrato): um ficheiro já importado por
 *                       completo é recusado. Extractos grandes chegam em lotes; o último traz final=true.
 *  banco.conciliar    — grava ligações movimento ↔ despesa/recibo em tbl_ConciliacaoBancaria. Um documento só pode estar
 *                       ligado a um movimento e um movimento só a um conjunto de documentos (cuja soma = valor do banco).
 *                       Repetir o mesmo pedido não duplica. Os recibos também ficam em tbl_RecebimentosBanco.
 *  banco.desfazer     — remove a conciliação dos movimentos indicados (para corrigir uma ligação errada).
 *
 * Cria as tabelas na primeira utilização (folhas BM_*). Valida tudo antes de escrever; não apaga facturas nem recibos.
 * Desempenho (limite de 120 s do Power Automate): leituras em bloco, sem leituras por linha.
 */
interface Resultado {
  sucesso: boolean; codigo: string; mensagem: string; operacao: string;
  novos: number; alterados: number; liquidacoes: number; ligacoes: number; erros: string[]; ms: number;
  duplicados: number; estado: string;
  nome_padronizado: string; // banco.importar_extrato: nome final do ficheiro (o fluxo usa-o para o guardar no SharePoint)
}
type Valor = string | number | boolean;
type Registo = Record<string, unknown>;

const TABELAS: Record<string, { folha: string; chave: string; colunas: string[] }> = {
  tbl_Faturas: { folha: "BM_Faturas", chave: "ID_Documento", colunas: ["ID_Documento", "Empresa", "Tipo_Documento", "Serie", "Numero", "Documento", "Data", "Data_Vencimento", "Codigo_Cliente", "Cliente", "NUIT", "Moeda", "Total", "Estado", "Hash", "ID_Importacao", "Data_Importacao"] },
  tbl_Recibos: { folha: "BM_Recibos", chave: "ID_Recibo", colunas: ["ID_Recibo", "Empresa", "Tipo_Documento", "Serie", "Numero", "Documento", "Data", "Codigo_Cliente", "Cliente", "Valor", "Modo_Pagamento", "Referencia", "Conta_Bancaria", "Estado", "Liquidacoes", "Hash", "ID_Importacao", "Data_Importacao"] },
  tbl_Liquidacoes: { folha: "BM_Liquidacoes", chave: "ID_Liquidacao", colunas: ["ID_Liquidacao", "ID_Recibo", "ID_Documento", "Documento_Liquidado", "Valor", "ID_Importacao"] },
  tbl_Importacoes: { folha: "BM_Importacoes", chave: "ID_Importacao", colunas: ["ID_Importacao", "Tipo", "Empresa", "Ficheiro", "Hash_Ficheiro", "Utilizador", "Data", "Linhas_Ficheiro", "Novos", "Alterados", "Iguais", "Rejeitados", "Id_Pedido"] },
  tbl_RecebimentosBanco: { folha: "BM_RecebimentosBanco", chave: "ID_Ligacao", colunas: ["ID_Ligacao", "ID_Recibo", "Data_Banco", "Valor", "Descritivo_Banco", "Conta", "Utilizador", "Data"] },
  tbl_Extratos: { folha: "BM_Extratos", chave: "ID_Extrato", colunas: ["ID_Extrato", "Nome_Original", "Nome_Padronizado", "Banco", "ID_Conta", "Conta", "Formato", "De", "Ate", "Movimentos_Ficheiro", "Novos", "Ja_Existentes", "Saldo_Inicial", "Saldo_Final", "Estado", "Utilizador", "Data_Importacao", "Id_Pedido"] },
  tbl_MovimentosBancarios: { folha: "BM_Movimentos", chave: "ID_Movimento", colunas: ["ID_Movimento", "ID_Conta", "Banco", "Conta", "Data_Movimento", "Data_Valor", "Referencia", "Descricao", "Debito", "Credito", "Valor", "Saldo", "ID_Extrato", "Data_Importacao"] },
  tbl_ConciliacaoBancaria: { folha: "BM_Conciliacao", chave: "ID_Conciliacao", colunas: ["ID_Conciliacao", "ID_Movimento", "Tipo_Documento", "ID_Documento", "Documento", "Valor", "Data_Banco", "Valor_Banco", "Descritivo_Banco", "Conta", "Pontos", "Motivos", "Metodo", "Utilizador", "Data_Importacao"] },
};
const TEXTO = ["ID_Documento", "ID_Recibo", "ID_Liquidacao", "ID_Importacao", "ID_Ligacao", "Empresa", "Tipo_Documento", "Serie", "Documento",
  "Codigo_Cliente", "Cliente", "NUIT", "Moeda", "Estado", "Hash", "Modo_Pagamento", "Referencia", "Conta_Bancaria", "Liquidacoes",
  "Documento_Liquidado", "Tipo", "Ficheiro", "Hash_Ficheiro", "Utilizador", "Descritivo_Banco", "Conta", "Id_Pedido",
  "ID_Extrato", "Nome_Original", "Nome_Padronizado", "Banco", "ID_Conta", "Formato", "ID_Movimento", "Descricao",
  "ID_Conciliacao", "Tipo_Documento", "Motivos", "Metodo"];
const DATAS = ["Data", "Data_Vencimento", "Data_Banco", "De", "Ate", "Data_Movimento", "Data_Valor"];
const DATA_HORA = ["Data_Importacao"];
const NUMEROS = ["Numero", "Total", "Valor", "Linhas_Ficheiro", "Novos", "Alterados", "Iguais", "Rejeitados",
  "Movimentos_Ficheiro", "Ja_Existentes", "Saldo_Inicial", "Saldo_Final", "Debito", "Credito", "Saldo", "Valor_Banco", "Pontos"];
const ISO = /^\d{4}-\d{2}-\d{2}$/;
const digitos = (v: unknown) => texto(v).replace(/\D/g, "").replace(/^0+/, "");
const cent = (v: unknown) => Math.round(Number(v) * 100);
const LIMITE = 1500;

function texto(v: unknown): string { return v === null || v === undefined ? "" : String(v).trim(); }
// "aaaa-mm-dd" (ou data-hora ISO) → número de série do Excel; outros valores ficam como vêm.
function serie(v: unknown, comHora: boolean): Valor {
  const t = texto(v);
  const m = t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/);
  if (!m) return typeof v === "number" ? v : t;
  const dias = Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
  return comHora && m[4] ? dias + (+m[4] * 60 + +m[5]) / 1440 : dias;
}
function celula(coluna: string, v: unknown): Valor {
  if (v === null || v === undefined || v === "") return "";
  if (DATAS.indexOf(coluna) >= 0) return serie(v, false);
  if (DATA_HORA.indexOf(coluna) >= 0) return serie(v, true);
  if (NUMEROS.indexOf(coluna) >= 0) { const n = Number(v); return isFinite(n) ? n : texto(v); }
  return texto(v);
}

function main(workbook: ExcelScript.Workbook, pedido: string): Resultado {
  const t0 = Date.now();
  const r: Resultado = { sucesso: false, codigo: "", mensagem: "", operacao: "", novos: 0, alterados: 0, liquidacoes: 0, ligacoes: 0, erros: [], ms: 0, duplicados: 0, estado: "", nome_padronizado: "" };
  const fim = (codigo: string, mensagem: string) => { r.codigo = codigo; r.mensagem = mensagem; r.ms = Date.now() - t0; return r; };

  let p: Registo;
  try { p = JSON.parse(pedido || "{}") as Registo; } catch { return fim("PEDIDO_INVALIDO", "O pedido não é JSON válido."); }
  r.operacao = texto(p.operacao);

  if (r.operacao === "base.importar") {
    const tipo = texto(p.tipo);
    if (tipo !== "faturas" && tipo !== "recibos") return fim("TIPO_INVALIDO", "tipo deve ser faturas ou recibos.");
    const nome = tipo === "faturas" ? "tbl_Faturas" : "tbl_Recibos";
    const linhas = p.linhas as Registo[], liquidacoes = (p.liquidacoes ?? []) as Registo[], imp = (p.importacao ?? {}) as Registo;
    if (!Array.isArray(linhas) || !Array.isArray(liquidacoes)) return fim("PEDIDO_INVALIDO", "linhas/liquidacoes têm de ser listas.");
    if (linhas.length > LIMITE || liquidacoes.length > LIMITE * 3) return fim("LOTE_GRANDE", "Máximo " + LIMITE + " registos por chamada.");
    if (!texto(imp.ID_Importacao)) return fim("PEDIDO_INVALIDO", "importacao.ID_Importacao em falta.");
    // Validação antes de qualquer escrita
    const def = TABELAS[nome], vistos = new Set<string>();
    linhas.forEach((l, i) => {
      const k = texto(l?.[def.chave]);
      if (!k) r.erros.push("linha " + (i + 1) + ": " + def.chave + " em falta");
      else if (vistos.has(k)) r.erros.push("chave repetida no pedido: " + k);
      vistos.add(k);
      for (const c of Object.keys(l ?? {})) if (def.colunas.indexOf(c) < 0) r.erros.push("coluna desconhecida: " + c);
      const valor = Number(l?.[tipo === "faturas" ? "Total" : "Valor"]);
      if (!isFinite(valor) || valor < 0) r.erros.push("linha " + (i + 1) + ": valor inválido");
    });
    for (const q of liquidacoes) {
      if (!texto(q?.ID_Recibo) || !texto(q?.ID_Documento) || !isFinite(Number(q?.Valor))) r.erros.push("liquidação inválida: " + JSON.stringify(q).slice(0, 80));
      else if (!vistos.has(texto(q.ID_Recibo))) r.erros.push("liquidação de recibo fora do pedido: " + texto(q.ID_Recibo));
    }
    if (r.erros.length) return fim("VALIDACAO", r.erros.slice(0, 5).join("; ") + (r.erros.length > 5 ? " (e mais " + (r.erros.length - 5) + ")" : ""));

    const res = upsert(workbook, nome, linhas);
    if (res.erro) return fim("ESTRUTURA", res.erro);
    r.novos = res.novos; r.alterados = res.alterados;
    if (tipo === "recibos") {
      const t = obterTabela(workbook, "tbl_Liquidacoes");
      if (typeof t === "string") return fim("ESTRUTURA", t);
      // Substitui as liquidações dos recibos deste pedido (o recibo no PRIMAVERA pode ter sido corrigido)
      const cab = cabecalho(t), j = cab.indexOf("ID_Recibo");
      if (j < 0) return fim("ESTRUTURA", "tbl_Liquidacoes sem coluna ID_Recibo.");
      const n = t.getRowCount();
      if (n > 0) {
        const col = t.getRangeBetweenHeaderAndTotal().getColumn(j).getValues().map(v => texto(v[0]));
        apagar(t, col.map((k, i) => (vistos.has(k) ? i : -1)).filter(i => i >= 0));
      }
      const l = acrescentar(t, liquidacoes);
      if (l) return fim("ESTRUTURA", l);
      r.liquidacoes = liquidacoes.length;
    }
    const regImp: Registo = { ...imp, Tipo: tipo, Data: imp.Data ?? new Date().toISOString(), Id_Pedido: texto(p.id_pedido) };
    for (const c of Object.keys(regImp)) if (TABELAS.tbl_Importacoes.colunas.indexOf(c) < 0) delete regImp[c];
    const ri = upsert(workbook, "tbl_Importacoes", [regImp]);
    if (ri.erro) return fim("ESTRUTURA", ri.erro);
    r.sucesso = true;
    return fim("OK", r.novos + " novo(s), " + r.alterados + " alterado(s)" + (tipo === "recibos" ? ", " + r.liquidacoes + " liquidação(ões)" : "") + ".");
  }

  if (r.operacao === "base.ligar_banco") {
    const ligacoes = p.ligacoes as Registo[];
    if (!Array.isArray(ligacoes) || !ligacoes.length || ligacoes.length > LIMITE) return fim("PEDIDO_INVALIDO", "ligacoes deve ser uma lista (1 a " + LIMITE + ").");
    for (const l of ligacoes) {
      if (!texto(l?.ID_Ligacao) || !texto(l?.ID_Recibo) || !isFinite(Number(l?.Valor))) r.erros.push("ligação inválida: " + JSON.stringify(l).slice(0, 80));
      for (const c of Object.keys(l ?? {})) if (TABELAS.tbl_RecebimentosBanco.colunas.indexOf(c) < 0) r.erros.push("coluna desconhecida: " + c);
    }
    if (r.erros.length) return fim("VALIDACAO", r.erros.slice(0, 5).join("; "));
    const res = upsert(workbook, "tbl_RecebimentosBanco", ligacoes);
    if (res.erro) return fim("ESTRUTURA", res.erro);
    r.ligacoes = res.novos + res.alterados; r.novos = res.novos; r.alterados = res.alterados;
    r.sucesso = true;
    return fim("OK", r.ligacoes + " ligação(ões) gravada(s).");
  }
  if (r.operacao === "banco.importar_extrato") {
    const ext = (p.extrato ?? {}) as Registo, movs = p.movimentos as Registo[], final = p.final === true;
    const agora = new Date().toISOString();
    const idExt = texto(ext.ID_Extrato).toLowerCase(), conta = digitos(ext.Conta);
    if (!/^[0-9a-f]{64}$/.test(idExt)) return fim("PEDIDO_INVALIDO", "extrato.ID_Extrato tem de ser o SHA-256 do ficheiro.");
    if (!conta) return fim("PEDIDO_INVALIDO", "extrato.Conta em falta.");
    if (!texto(ext.Nome_Padronizado)) return fim("PEDIDO_INVALIDO", "extrato.Nome_Padronizado em falta.");
    if (!Array.isArray(movs) || movs.length > LIMITE) return fim("PEDIDO_INVALIDO", "movimentos deve ser uma lista (máximo " + LIMITE + " por lote).");
    for (const c of Object.keys(ext)) if (TABELAS.tbl_Extratos.colunas.indexOf(c) < 0) r.erros.push("coluna desconhecida em extrato: " + c);
    const vistos = new Set<string>();
    movs.forEach((m, i) => {
      const k = texto(m?.ID_Movimento), onde = "movimento " + (i + 1) + ": ";
      if (!k) r.erros.push(onde + "ID_Movimento em falta");
      else if (vistos.has(k)) r.erros.push("movimento repetido no pedido: " + k);
      vistos.add(k);
      if (digitos(m?.Conta) !== conta) r.erros.push(onde + "conta " + texto(m?.Conta) + " diferente da conta do extracto");
      if (!ISO.test(texto(m?.Data_Movimento))) r.erros.push(onde + "data inválida");
      if (!isFinite(Number(m?.Valor)) || cent(m?.Valor) === 0) r.erros.push(onde + "valor inválido");
      for (const c of Object.keys(m ?? {})) if (TABELAS.tbl_MovimentosBancarios.colunas.indexOf(c) < 0) r.erros.push("coluna desconhecida: " + c);
    });
    if (r.erros.length) return fim("VALIDACAO", r.erros.slice(0, 5).join("; ") + (r.erros.length > 5 ? " (e mais " + (r.erros.length - 5) + ")" : ""));

    const te = obterTabela(workbook, "tbl_Extratos");
    if (typeof te === "string") return fim("ESTRUTURA", te);
    const anterior = linhaPorChave(te, "ID_Extrato", idExt);
    if (anterior && ["Completo", "Sem movimentos novos"].indexOf(texto(anterior.Estado)) >= 0) {
      r.duplicados = movs.length; r.estado = texto(anterior.Estado);
      return fim("EXTRATO_JA_IMPORTADO", "Este ficheiro já foi importado (" + texto(anterior.Nome_Padronizado) + ", estado " + texto(anterior.Estado) + "). Nada foi gravado.");
    }
    // Nome único: outro ficheiro (conteúdo diferente) da mesma conta e período recebe _v2, _v3…
    let nome = anterior ? texto(anterior.Nome_Padronizado) : texto(ext.Nome_Padronizado);
    if (!anterior) {
      const usados = new Set(coluna(te, "Nome_Padronizado").map((x) => x.toLowerCase()));
      const ponto = nome.lastIndexOf("."), raiz = ponto > 0 ? nome.slice(0, ponto) : nome, extensao = ponto > 0 ? nome.slice(ponto) : "";
      for (let v = 2; usados.has(nome.toLowerCase()); v++) nome = raiz + "_v" + v + extensao;
    }
    r.nome_padronizado = nome;
    const tm = obterTabela(workbook, "tbl_MovimentosBancarios");
    if (typeof tm === "string") return fim("ESTRUTURA", tm);
    const existentes = new Set(coluna(tm, "ID_Movimento"));
    const doExtrato = coluna(tm, "ID_Extrato").filter((x) => x.toLowerCase() === idExt).length;
    const novos = movs.filter((m) => !existentes.has(texto(m.ID_Movimento)))
      .map((m) => ({ ...m, ID_Extrato: idExt, Data_Importacao: texto(m.Data_Importacao) || agora }));
    const e = acrescentar(tm, novos);
    if (e) return fim("ESTRUTURA", e);
    r.novos = novos.length; r.duplicados = movs.length - novos.length;
    const totalNovos = doExtrato + r.novos, noFicheiro = Number(ext.Movimentos_Ficheiro);
    r.estado = !final ? "Parcial" : totalNovos > 0 ? "Completo" : "Sem movimentos novos";
    const reg: Registo = { ...ext, ID_Extrato: idExt, Nome_Padronizado: nome, Novos: totalNovos, Estado: r.estado, Data_Importacao: texto(ext.Data_Importacao) || agora, Id_Pedido: texto(p.id_pedido),
      Ja_Existentes: final && isFinite(noFicheiro) ? Math.max(0, noFicheiro - totalNovos) : "" };
    const ri = upsert(workbook, "tbl_Extratos", [reg]);
    if (ri.erro) return fim("ESTRUTURA", ri.erro);
    r.sucesso = true;
    return fim("OK", r.novos + " movimento(s) novo(s), " + r.duplicados + " já existiam na base (ignorados)." + (final ? " Extracto registado: " + nome + "." : ""));
  }

  if (r.operacao === "banco.conciliar") {
    const ligs = p.ligacoes as Registo[], agora = new Date().toISOString(), utilizador = texto(p.utilizador);
    if (!Array.isArray(ligs) || !ligs.length || ligs.length > LIMITE) return fim("PEDIDO_INVALIDO", "ligacoes deve ser uma lista (1 a " + LIMITE + ").");
    const porMov = new Map<string, Registo[]>(), docs = new Set<string>();
    ligs.forEach((l, i) => {
      const onde = "ligação " + (i + 1) + ": ", tipo = texto(l?.Tipo_Documento);
      if (!texto(l?.ID_Movimento)) r.erros.push(onde + "ID_Movimento em falta");
      if (tipo !== "despesa" && tipo !== "recibo") r.erros.push(onde + "Tipo_Documento deve ser despesa ou recibo");
      if (!texto(l?.ID_Documento)) r.erros.push(onde + "ID_Documento em falta");
      if (!isFinite(Number(l?.Valor)) || Number(l?.Valor) <= 0 || !isFinite(Number(l?.Valor_Banco))) r.erros.push(onde + "valores inválidos");
      for (const c of Object.keys(l ?? {})) if (TABELAS.tbl_ConciliacaoBancaria.colunas.indexOf(c) < 0) r.erros.push("coluna desconhecida: " + c);
      const d = tipo + "|" + texto(l?.ID_Documento);
      if (docs.has(d)) r.erros.push("documento repetido no pedido: " + texto(l?.ID_Documento));
      docs.add(d);
      const k = texto(l?.ID_Movimento);
      porMov.set(k, (porMov.get(k) ?? []).concat([l]));
    });
    porMov.forEach((ls, k) => {
      const soma = ls.reduce((s, l) => s + cent(l.Valor), 0), banco = Math.abs(cent(ls[0].Valor_Banco));
      if (soma !== banco) r.erros.push("movimento " + k + ": a soma dos documentos (" + soma / 100 + ") difere do valor do banco (" + banco / 100 + ")");
    });
    if (r.erros.length) return fim("VALIDACAO", r.erros.slice(0, 5).join("; ") + (r.erros.length > 5 ? " (e mais " + (r.erros.length - 5) + ")" : ""));
    const tm = workbook.getTable("tbl_MovimentosBancarios");
    const naBase = new Set(tm ? coluna(tm, "ID_Movimento") : []);
    const faltam = Array.from(porMov.keys()).filter((k) => !naBase.has(k));
    if (faltam.length) return fim("MOVIMENTO_INEXISTENTE", faltam.length + " movimento(s) não estão na base mãe (ex.: " + faltam[0] + "). Grave primeiro o extracto.");
    const tc = obterTabela(workbook, "tbl_ConciliacaoBancaria");
    if (typeof tc === "string") return fim("ESTRUTURA", tc);
    const ids = coluna(tc, "ID_Conciliacao"), movCol = coluna(tc, "ID_Movimento");
    const docParaMov = new Map<string, string>(), movParaDocs = new Map<string, string[]>();
    ids.forEach((id, i) => { if (!id) return; docParaMov.set(id, movCol[i]); movParaDocs.set(movCol[i], (movParaDocs.get(movCol[i]) ?? []).concat([id])); });
    const novos: Registo[] = [];
    porMov.forEach((ls, mov) => {
      const pedidos = ls.map((l) => texto(l.Tipo_Documento) + "|" + texto(l.ID_Documento));
      for (const d of pedidos) {
        const outro = docParaMov.get(d);
        if (outro && outro !== mov) r.erros.push("DOCUMENTO_JA_CONCILIADO: " + d.replace("|", " ") + " já está ligado ao movimento " + outro);
      }
      const atuais = movParaDocs.get(mov) ?? [];
      if (atuais.length && (atuais.length !== pedidos.length || atuais.some((d) => pedidos.indexOf(d) < 0)))
        r.erros.push("MOVIMENTO_JA_CONCILIADO: " + mov + " já está conciliado com " + atuais.join(", ") + " (desfaça antes de alterar)");
      ls.forEach((l, j) => { if (!docParaMov.has(pedidos[j])) novos.push({ ...l, ID_Conciliacao: pedidos[j], Utilizador: texto(l.Utilizador) || utilizador, Data_Importacao: agora }); });
    });
    if (r.erros.length) return fim("CONFLITO", r.erros.slice(0, 5).join("; ") + (r.erros.length > 5 ? " (e mais " + (r.erros.length - 5) + ")" : ""));
    const e = acrescentar(tc, novos);
    if (e) return fim("ESTRUTURA", e);
    const recibos = novos.filter((l) => texto(l.Tipo_Documento) === "recibo").map((l) => ({ ID_Ligacao: texto(l.ID_Documento) + "@" + texto(l.ID_Movimento), ID_Recibo: texto(l.ID_Documento),
      Data_Banco: l.Data_Banco, Valor: l.Valor, Descritivo_Banco: l.Descritivo_Banco, Conta: l.Conta, Utilizador: texto(l.Utilizador) || utilizador, Data: agora.slice(0, 10) }));
    if (recibos.length) { const rr = upsert(workbook, "tbl_RecebimentosBanco", recibos); if (rr.erro) return fim("ESTRUTURA", rr.erro); }
    r.ligacoes = novos.length; r.novos = novos.length; r.duplicados = ligs.length - novos.length;
    r.sucesso = true;
    return fim("OK", novos.length + " ligação(ões) gravada(s)" + (r.duplicados ? ", " + r.duplicados + " já existiam" : "") + ".");
  }

  if (r.operacao === "banco.desfazer") {
    const movs = (p.movimentos ?? []) as unknown[];
    if (!Array.isArray(movs) || !movs.length || movs.length > LIMITE) return fim("PEDIDO_INVALIDO", "movimentos deve ser uma lista de ID_Movimento (1 a " + LIMITE + ").");
    const alvo = new Set(movs.map(texto));
    const tc = workbook.getTable("tbl_ConciliacaoBancaria");
    if (tc && !soLinhaVazia(tc)) {
      const col = coluna(tc, "ID_Movimento"), ind = col.map((k, i) => (alvo.has(k) ? i : -1)).filter((i) => i >= 0);
      apagar(tc, ind); r.ligacoes = ind.length;
    }
    const tr = workbook.getTable("tbl_RecebimentosBanco");
    if (tr && !soLinhaVazia(tr)) {
      const col = coluna(tr, "ID_Ligacao");
      apagar(tr, col.map((k, i) => (alvo.has(k.slice(k.indexOf("@") + 1)) && k.indexOf("@") > 0 ? i : -1)).filter((i) => i >= 0));
    }
    r.sucesso = true;
    return fim("OK", r.ligacoes + " ligação(ões) removida(s).");
  }
  return fim("OPERACAO_DESCONHECIDA", "Operação não suportada: " + r.operacao);
}

// ---------- Tabelas ----------
// Valores de uma coluna (texto), [] se a tabela estiver vazia. Uma leitura.
function coluna(t: ExcelScript.Table, nome: string): string[] {
  if (soLinhaVazia(t) || t.getRowCount() === 0) return [];
  const j = cabecalho(t).indexOf(nome);
  if (j < 0) return [];
  return t.getRangeBetweenHeaderAndTotal().getColumn(j).getValues().map((v) => texto(v[0]));
}
// Linha (objecto) cuja coluna chave = valor, ou null. Duas leituras (coluna chave e a linha encontrada).
function linhaPorChave(t: ExcelScript.Table, chave: string, valor: string): Registo | null {
  const i = coluna(t, chave).findIndex((k) => k.toLowerCase() === valor.toLowerCase());
  if (i < 0) return null;
  const cab = cabecalho(t), v = t.getRangeBetweenHeaderAndTotal().getRow(i).getValues()[0];
  const o: Registo = {};
  cab.forEach((c, j) => { o[c] = v[j]; });
  return o;
}

function cabecalho(t: ExcelScript.Table): string[] { return t.getHeaderRowRange().getValues()[0].map(v => String(v)); }
function obterTabela(workbook: ExcelScript.Workbook, nome: string): ExcelScript.Table | string {
  const def = TABELAS[nome];
  let t = workbook.getTable(nome);
  if (!t) {
    const folha = workbook.getWorksheet(def.folha) ?? workbook.addWorksheet(def.folha);
    const n = def.colunas.length;
    const cab = folha.getRangeByIndexes(0, 0, 1, n);
    cab.setValues([def.colunas]);
    t = folha.addTable(cab, true);
    t.setName(nome);
  }
  const cab = cabecalho(t);
  const faltam = def.colunas.filter(c => cab.indexOf(c) < 0);
  if (faltam.length) return nome + ": faltam as colunas " + faltam.join(", ") + ".";
  return t;
}
// Tabela acabada de criar tem uma linha vazia: devolve true se for o caso (para ser aproveitada).
function soLinhaVazia(t: ExcelScript.Table): boolean {
  if (t.getRowCount() !== 1) return false;
  return t.getRangeBetweenHeaderAndTotal().getValues()[0].every(v => texto(v) === "");
}
function acrescentar(t: ExcelScript.Table, registos: Registo[]): string {
  if (!registos.length) return "";
  const cab = cabecalho(t);
  const matriz = registos.map(o => cab.map(c => celula(c, o[c])));
  const vazia = soLinhaVazia(t);
  const inicio = vazia ? 0 : t.getRowCount();
  if (vazia) {
    t.getRangeBetweenHeaderAndTotal().getRow(0).setValues([matriz[0]]);
    if (matriz.length > 1) t.addRows(-1, matriz.slice(1));
  } else t.addRows(-1, matriz);
  formatar(t, cab, inicio, matriz);
  return "";
}
// Texto fica texto (NUIT, códigos, IDs com zeros à esquerda); datas com formato de data.
function formatar(t: ExcelScript.Table, cab: string[], inicio: number, matriz: Valor[][]) {
  const corpo = t.getRangeBetweenHeaderAndTotal(), folha = corpo.getWorksheet();
  const l0 = corpo.getRowIndex() + inicio, c0 = corpo.getColumnIndex();
  cab.forEach((c, j) => {
    const rg = folha.getRangeByIndexes(l0, c0 + j, matriz.length, 1);
    if (TEXTO.indexOf(c) >= 0) { rg.setNumberFormat("@"); rg.setValues(matriz.map(l => [String(l[j])])); }
    else if (DATAS.indexOf(c) >= 0) rg.setNumberFormat("dd/mm/yyyy");
    else if (DATA_HORA.indexOf(c) >= 0) rg.setNumberFormat("dd/mm/yyyy hh:mm");
  });
}
function apagar(t: ExcelScript.Table, indices: number[]) {
  const ord = indices.slice().sort((a, b) => b - a);
  for (let p = 0; p < ord.length;) {
    let q = p;
    while (q + 1 < ord.length && ord[q + 1] === ord[q] - 1) q++;
    t.deleteRowsAt(ord[q], q - p + 1);
    p = q + 1;
  }
}
// Upsert por chave: uma leitura da coluna chave; alterados reescritos por blocos (uma leitura do bloco para manter
// as colunas que não vêm no pedido); novos num só addRows.
function upsert(workbook: ExcelScript.Workbook, nome: string, registos: Registo[]): { novos: number; alterados: number; erro: string } {
  const t = obterTabela(workbook, nome);
  if (typeof t === "string") return { novos: 0, alterados: 0, erro: t };
  const def = TABELAS[nome], cab = cabecalho(t), jChave = cab.indexOf(def.chave);
  const n = soLinhaVazia(t) ? 0 : t.getRowCount();
  const existentes = new Map<string, number>();
  if (n > 0) t.getRangeBetweenHeaderAndTotal().getColumn(jChave).getValues().forEach((v, i) => { const k = texto(v[0]); if (k && !existentes.has(k)) existentes.set(k, i); });
  const novos: Registo[] = [], alterar = new Map<number, Registo>();
  for (const o of registos) {
    const i = existentes.get(texto(o[def.chave]));
    if (i === undefined) novos.push(o); else alterar.set(i, o);
  }
  if (alterar.size) {
    const ind = Array.from(alterar.keys()).sort((a, b) => a - b);
    const corpo = t.getRangeBetweenHeaderAndTotal(), folha = corpo.getWorksheet(), l0 = corpo.getRowIndex(), c0 = corpo.getColumnIndex();
    const ini = ind[0], bloco = folha.getRangeByIndexes(l0 + ini, c0, ind[ind.length - 1] - ini + 1, cab.length).getValues();
    for (let p = 0; p < ind.length;) {
      let q = p;
      while (q + 1 < ind.length && ind[q + 1] === ind[q] + 1) q++;
      const linhas: Valor[][] = [];
      for (let k = p; k <= q; k++) {
        const o = alterar.get(ind[k]) as Registo, atual = bloco[ind[k] - ini];
        linhas.push(cab.map((c, j) => (Object.prototype.hasOwnProperty.call(o, c) ? celula(c, o[c]) : atual[j])));
      }
      folha.getRangeByIndexes(l0 + ind[p], c0, q - p + 1, cab.length).setValues(linhas);
      p = q + 1;
    }
  }
  const erro = acrescentar(t, novos);
  return { novos: novos.length, alterados: alterar.size, erro };
}
