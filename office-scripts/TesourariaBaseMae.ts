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
 *  base.consultar     — devolve as tabelas da base mãe (colunas + linhas em matriz, compacto) para o site. Tabelas
 *                       ainda não criadas vêm vazias. Só leitura: não cria nem altera nada.
 *
 * Cria as tabelas na primeira utilização (folhas BM_*). Valida tudo antes de escrever; não apaga facturas nem recibos.
 * Desempenho (limite de 120 s do Power Automate; cada leitura é uma ida ao servidor do Excel):
 *  - cada tabela é aberta uma vez por execução com 2 leituras (cabeçalho + endereço do corpo); só se lêem as colunas
 *    necessárias (a chave), nunca a tabela inteira de movimentos;
 *  - cada bloco novo é escrito de uma vez (formatos + valores + redimensionar a tabela), sem leituras;
 *  - base.consultar lê cada tabela numa só leitura.
 */
interface Resultado {
  sucesso: boolean; codigo: string; mensagem: string; operacao: string;
  novos: number; alterados: number; liquidacoes: number; ligacoes: number; erros: string[]; ms: number;
  duplicados: number; estado: string;
  nome_padronizado: string; // banco.importar_extrato: nome final do ficheiro (o fluxo usa-o para o guardar no SharePoint)
  tabelas: TabelaLida[];     // base.consultar
}
interface TabelaLida { tabela: string; colunas: string[]; linhas: Valor[][]; }
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
const LIMITE = 1500, LIMITE_MOV = 2500;

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
  ABERTAS.clear();
  const r: Resultado = { sucesso: false, codigo: "", mensagem: "", operacao: "", novos: 0, alterados: 0, liquidacoes: 0, ligacoes: 0, erros: [], ms: 0, duplicados: 0, estado: "", nome_padronizado: "", tabelas: [] };
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
      const t = abrir(workbook, "tbl_Liquidacoes", true);
      if (typeof t === "string" || !t) return fim("ESTRUTURA", String(t));
      // Substitui as liquidações dos recibos deste pedido (o recibo no PRIMAVERA pode ter sido corrigido)
      apagar(t, coluna(t, "ID_Recibo").map((k, i) => (vistos.has(k) ? i : -1)).filter(i => i >= 0));
      acrescentar(t, liquidacoes);
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
  if (r.operacao === "base.consultar") {
    const pedidas = Array.isArray(p.tabelas) && (p.tabelas as unknown[]).length ? (p.tabelas as unknown[]).map(texto) : Object.keys(TABELAS);
    for (const nome of pedidas) {
      if (!TABELAS[nome]) return fim("PEDIDO_INVALIDO", "Tabela desconhecida: " + nome);
      const t = workbook.getTable(nome);
      if (!t) { r.tabelas.push({ tabela: nome, colunas: TABELAS[nome].colunas.slice(), linhas: [] }); continue; }
      const tudo = t.getRange().getValues() as Valor[][]; // cabeçalho + corpo numa só leitura
      const colunas = tudo[0].map((v) => String(v));
      const linhas = tudo.slice(1).filter((l) => l.some((v) => texto(v) !== ""));
      r.tabelas.push({ tabela: nome, colunas, linhas });
      r.novos += linhas.length;
    }
    r.sucesso = true;
    return fim("OK", r.tabelas.length + " tabela(s), " + r.novos + " linha(s).");
  }

  if (r.operacao === "banco.importar_extrato") {
    const ext = (p.extrato ?? {}) as Registo, movs = p.movimentos as Registo[], final = p.final === true;
    const agora = new Date().toISOString();
    const idExt = texto(ext.ID_Extrato).toLowerCase(), conta = digitos(ext.Conta);
    if (!/^[0-9a-f]{64}$/.test(idExt)) return fim("PEDIDO_INVALIDO", "extrato.ID_Extrato tem de ser o SHA-256 do ficheiro.");
    if (!conta) return fim("PEDIDO_INVALIDO", "extrato.Conta em falta.");
    if (!texto(ext.Nome_Padronizado)) return fim("PEDIDO_INVALIDO", "extrato.Nome_Padronizado em falta.");
    if (!Array.isArray(movs) || movs.length > LIMITE_MOV) return fim("PEDIDO_INVALIDO", "movimentos deve ser uma lista (máximo " + LIMITE_MOV + " por lote).");
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

    const te = abrir(workbook, "tbl_Extratos", true);
    if (typeof te === "string" || !te) return fim("ESTRUTURA", String(te));
    // tbl_Extratos é pequena: lê-se inteira numa leitura (registo anterior + nomes já usados)
    const regs = corpo(te), jId = te.cab.indexOf("ID_Extrato"), jNome = te.cab.indexOf("Nome_Padronizado");
    const iAnt = regs.findIndex((l) => texto(l[jId]).toLowerCase() === idExt);
    const anterior: Registo | null = iAnt < 0 ? null : objecto(te.cab, regs[iAnt]);
    if (anterior && ["Completo", "Sem movimentos novos"].indexOf(texto(anterior.Estado)) >= 0) {
      r.duplicados = movs.length; r.estado = texto(anterior.Estado);
      return fim("EXTRATO_JA_IMPORTADO", "Este ficheiro já foi importado (" + texto(anterior.Nome_Padronizado) + ", estado " + texto(anterior.Estado) + "). Nada foi gravado.");
    }
    // Nome único: outro ficheiro (conteúdo diferente) da mesma conta e período recebe _v2, _v3…
    let nome = anterior ? texto(anterior.Nome_Padronizado) : texto(ext.Nome_Padronizado);
    if (!anterior) {
      const usados = new Set(regs.map((l) => texto(l[jNome]).toLowerCase()));
      const ponto = nome.lastIndexOf("."), raiz = ponto > 0 ? nome.slice(0, ponto) : nome, extensao = ponto > 0 ? nome.slice(ponto) : "";
      for (let v = 2; usados.has(nome.toLowerCase()); v++) nome = raiz + "_v" + v + extensao;
    }
    r.nome_padronizado = nome;
    const tm = abrir(workbook, "tbl_MovimentosBancarios", true);
    if (typeof tm === "string" || !tm) return fim("ESTRUTURA", String(tm));
    const existentes = new Set(coluna(tm, "ID_Movimento")); // a única leitura grande: a coluna chave
    const novos = movs.filter((m) => !existentes.has(texto(m.ID_Movimento)))
      .map((m) => ({ ...m, ID_Extrato: idExt, Data_Importacao: texto(m.Data_Importacao) || agora }));
    acrescentar(tm, novos);
    r.novos = novos.length; r.duplicados = movs.length - novos.length;
    // Total de novos do extracto: acumulado dos lotes; no último lote confirma-se pela coluna ID_Extrato (1 leitura)
    const totalNovos = final ? coluna(tm, "ID_Extrato").filter((x) => x.toLowerCase() === idExt).length
      : (anterior ? Number(anterior.Novos) || 0 : 0) + r.novos;
    const noFicheiro = Number(ext.Movimentos_Ficheiro);
    r.estado = !final ? "Parcial" : totalNovos > 0 ? "Completo" : "Sem movimentos novos";
    const reg: Registo = { ...ext, ID_Extrato: idExt, Nome_Padronizado: nome, Novos: totalNovos, Estado: r.estado, Data_Importacao: texto(ext.Data_Importacao) || agora, Id_Pedido: texto(p.id_pedido),
      Ja_Existentes: final && isFinite(noFicheiro) ? Math.max(0, noFicheiro - totalNovos) : "" };
    for (const c of Object.keys(reg)) if (te.cab.indexOf(c) < 0) delete reg[c];
    if (iAnt >= 0) escreverLinhas(te, iAnt, [te.cab.map((c, j) => (Object.prototype.hasOwnProperty.call(reg, c) ? celula(c, reg[c]) : regs[iAnt][j]))]);
    else acrescentar(te, [reg]);
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
    const tm = abrir(workbook, "tbl_MovimentosBancarios", false);
    if (typeof tm === "string") return fim("ESTRUTURA", tm);
    const naBase = new Set(tm ? coluna(tm, "ID_Movimento") : []);
    const faltam = Array.from(porMov.keys()).filter((k) => !naBase.has(k));
    if (faltam.length) return fim("MOVIMENTO_INEXISTENTE", faltam.length + " movimento(s) não estão na base mãe (ex.: " + faltam[0] + "). Grave primeiro o extracto.");
    const tc = abrir(workbook, "tbl_ConciliacaoBancaria", true);
    if (typeof tc === "string" || !tc) return fim("ESTRUTURA", String(tc));
    const [ids, movCol] = colunas(tc, ["ID_Conciliacao", "ID_Movimento"]); // colunas vizinhas: uma leitura
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
    acrescentar(tc, novos);
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
    const tc = abrir(workbook, "tbl_ConciliacaoBancaria", false);
    if (typeof tc === "string") return fim("ESTRUTURA", tc);
    if (tc) {
      const ind = coluna(tc, "ID_Movimento").map((k, i) => (alvo.has(k) ? i : -1)).filter((i) => i >= 0);
      apagar(tc, ind); r.ligacoes = ind.length;
    }
    const tr = abrir(workbook, "tbl_RecebimentosBanco", false);
    if (typeof tr === "string") return fim("ESTRUTURA", tr);
    if (tr) {
      const col = coluna(tr, "ID_Ligacao");
      apagar(tr, col.map((k, i) => (alvo.has(k.slice(k.indexOf("@") + 1)) && k.indexOf("@") > 0 ? i : -1)).filter((i) => i >= 0));
    }
    r.sucesso = true;
    return fim("OK", r.ligacoes + " ligação(ões) removida(s).");
  }
  return fim("OPERACAO_DESCONHECIDA", "Operação não suportada: " + r.operacao);
}

// ---------- Tabelas (camada rápida) ----------
// Uma tabela aberta: posição do corpo (l0, c0), nº de linhas com dados (n) e cabeçalho. Aberta uma vez por execução.
interface Tab { nome: string; t: ExcelScript.Table; folha: ExcelScript.Worksheet; cab: string[]; l0: number; c0: number; n: number; }
const ABERTAS = new Map<string, Tab>();
function indiceColuna(letras: string): number { let n = 0; for (const ch of letras) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
// Abre (ou cria) a tabela. Existente: 2 leituras (cabeçalho e endereço do corpo) e mais 1 só se o corpo tiver uma linha.
function abrir(workbook: ExcelScript.Workbook, nome: string, criar: boolean): Tab | string | null {
  const ja = ABERTAS.get(nome);
  if (ja) return ja;
  const def = TABELAS[nome];
  const t = workbook.getTable(nome);
  let tab: Tab;
  if (!t) {
    if (!criar) return null;
    const folha = workbook.getWorksheet(def.folha) ?? workbook.addWorksheet(def.folha);
    const cab = folha.getRangeByIndexes(0, 0, 1, def.colunas.length);
    cab.setValues([def.colunas]);
    const nova = folha.addTable(cab, true);
    nova.setName(nome);
    tab = { nome, t: nova, folha, cab: def.colunas.slice(), l0: 1, c0: 0, n: 0 }; // tabela nova: 1 linha vazia
  } else {
    const cab = t.getHeaderRowRange().getValues()[0].map((v) => String(v));
    const end = t.getRangeBetweenHeaderAndTotal().getAddress();
    const m = end.slice(end.lastIndexOf("!") + 1).replace(/\$/g, "").match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) return nome + ": endereço inesperado (" + end + ").";
    const l0 = Number(m[2]) - 1, n = (m[4] ? Number(m[4]) : Number(m[2])) - Number(m[2]) + 1;
    tab = { nome, t, folha: t.getWorksheet(), cab, l0, c0: indiceColuna(m[1]), n };
    if (n === 1 && tab.folha.getRangeByIndexes(l0, tab.c0, 1, cab.length).getValues()[0].every((v) => texto(v) === "")) tab.n = 0;
  }
  const faltam = def.colunas.filter((c) => tab.cab.indexOf(c) < 0);
  if (faltam.length) return nome + ": faltam as colunas " + faltam.join(", ") + ".";
  ABERTAS.set(nome, tab);
  return tab;
}
// Uma coluna (texto): 1 leitura; [] se vazia.
function coluna(tab: Tab, nome: string): string[] {
  const j = tab.cab.indexOf(nome);
  if (j < 0 || !tab.n) return [];
  return tab.folha.getRangeByIndexes(tab.l0, tab.c0 + j, tab.n, 1).getValues().map((v) => texto(v[0]));
}
// Várias colunas vizinhas numa só leitura.
function colunas(tab: Tab, nomes: string[]): string[][] {
  const js = nomes.map((c) => tab.cab.indexOf(c));
  if (!tab.n || js.some((j) => j < 0)) return nomes.map(() => []);
  const j0 = Math.min(...js), j1 = Math.max(...js);
  if (j1 - j0 > 4) return nomes.map((c) => coluna(tab, c));
  const v = tab.folha.getRangeByIndexes(tab.l0, tab.c0 + j0, tab.n, j1 - j0 + 1).getValues();
  return js.map((j) => v.map((l) => texto(l[j - j0])));
}
function corpo(tab: Tab): Valor[][] {
  return tab.n ? (tab.folha.getRangeByIndexes(tab.l0, tab.c0, tab.n, tab.cab.length).getValues() as Valor[][]) : [];
}
function objecto(cab: string[], l: Valor[]): Registo { const o: Registo = {}; cab.forEach((c, j) => { o[c] = l[j]; }); return o; }
const formato = (c: string) => (TEXTO.indexOf(c) >= 0 ? "@" : DATAS.indexOf(c) >= 0 ? "dd/mm/yyyy" : DATA_HORA.indexOf(c) >= 0 ? "dd/mm/yyyy hh:mm" : "General");
// Escreve linhas a partir da linha «inicio» do corpo: formatos primeiro (texto fica texto: NUIT, contas, IDs), depois valores.
function escreverLinhas(tab: Tab, inicio: number, matriz: Valor[][]) {
  if (!matriz.length) return;
  const fmts = tab.cab.map(formato), rg = tab.folha.getRangeByIndexes(tab.l0 + inicio, tab.c0, matriz.length, tab.cab.length);
  rg.setNumberFormats(matriz.map(() => fmts));
  rg.setValues(matriz.map((l) => l.map((v, j) => (fmts[j] === "@" ? String(v) : v))));
}
// Acrescenta no fim: um bloco escrito + redimensionar a tabela. Sem leituras.
function acrescentar(tab: Tab, registos: Registo[]) {
  if (!registos.length) return;
  escreverLinhas(tab, tab.n, registos.map((o) => tab.cab.map((c) => celula(c, o[c]))));
  const total = tab.n + registos.length;
  if (total > Math.max(tab.n, 1)) tab.t.resize(tab.folha.getRangeByIndexes(tab.l0 - 1, tab.c0, total + 1, tab.cab.length));
  tab.n = total;
}
// Apaga linhas (índices do corpo) por blocos seguidos, de baixo para cima.
function apagar(tab: Tab, indices: number[]) {
  const ord = indices.slice().sort((a, b) => b - a);
  for (let p = 0; p < ord.length;) {
    let q = p;
    while (q + 1 < ord.length && ord[q + 1] === ord[q] - 1) q++;
    tab.t.deleteRowsAt(ord[q], q - p + 1);
    tab.n -= q - p + 1;
    p = q + 1;
  }
}
// Upsert por chave: 1 leitura da coluna chave; alterados reescritos por blocos (lê-se o bloco só se o pedido não trouxer
// todas as colunas); novos num só bloco.
function upsert(workbook: ExcelScript.Workbook, nome: string, registos: Registo[]): { novos: number; alterados: number; erro: string } {
  const tab = abrir(workbook, nome, true);
  if (typeof tab === "string" || !tab) return { novos: 0, alterados: 0, erro: String(tab) };
  const chave = TABELAS[nome].chave, existentes = new Map<string, number>();
  coluna(tab, chave).forEach((k, i) => { if (k && !existentes.has(k)) existentes.set(k, i); });
  const novos: Registo[] = [], alterar = new Map<number, Registo>();
  for (const o of registos) {
    const i = existentes.get(texto(o[chave]));
    if (i === undefined) novos.push(o); else alterar.set(i, o);
  }
  if (alterar.size) {
    const ind = Array.from(alterar.keys()).sort((a, b) => a - b), ini = ind[0];
    const completos = Array.from(alterar.values()).every((o) => tab.cab.every((c) => Object.prototype.hasOwnProperty.call(o, c)));
    const bloco = completos ? [] : (tab.folha.getRangeByIndexes(tab.l0 + ini, tab.c0, ind[ind.length - 1] - ini + 1, tab.cab.length).getValues() as Valor[][]);
    for (let p = 0; p < ind.length;) {
      let q = p;
      while (q + 1 < ind.length && ind[q + 1] === ind[q] + 1) q++;
      const linhas: Valor[][] = [];
      for (let k = p; k <= q; k++) {
        const o = alterar.get(ind[k]) as Registo;
        linhas.push(tab.cab.map((c, j) => (Object.prototype.hasOwnProperty.call(o, c) ? celula(c, o[c]) : bloco[ind[k] - ini][j])));
      }
      escreverLinhas(tab, ind[p], linhas);
      p = q + 1;
    }
  }
  acrescentar(tab, novos);
  return { novos: novos.length, alterados: alterar.size, erro: "" };
}
