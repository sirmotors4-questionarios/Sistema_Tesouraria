/* Facturação, Recebimentos e Cobranças — interface da base mãe.
 * Importa listagens Excel do PRIMAVERA, mostra o que muda face à base mãe (novos, alterados, rejeitados, ausentes) e
 * grava pelo endpoint V9 (operação base.importar → Office Script TesourariaBaseMae). As vistas calculam a reconciliação
 * interna (facturas ↔ recibos), as excepções e a antiguidade de saldos a partir da base mãe.
 */
(function () {
  "use strict";
  const B = window.TES_BASE, C = window.TES_CONCILIACAO, IMP = window.TES_IMPORTAR;
  const $ = (id) => document.getElementById(id);
  if (!B || !C || !IMP || !$("facturacao")) return;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"})[c]);
  const mt = (n) => (n === "" || n == null || !Number.isFinite(Number(n)) ? "" : C.formatar(n));
  const dataPt = (v) => { const iso = C.data(v); return iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—"; };
  const ctx = () => window.TES_CONTEXTO || {ligado: () => false, tabela: () => [], empresas: () => [], escritaDisponivel: () => false};
  const LOTE = 1000; // menos execuções do Office Script (cada uma tem um custo fixo de arranque)
  const TIPO = {
    faturas: {tabela: "tbl_Faturas", chave: "ID_Documento", nome: "facturas e notas de crédito", ficheiro: "Listagem de documentos de venda (FA, FR, VD, NC, ND)"},
    recibos: {tabela: "tbl_Recibos", chave: "ID_Recibo", nome: "recibos", ficheiro: "Listagem de recibos, com o documento liquidado em cada linha"},
  };

  // ---------- Base mãe: Excel + o que foi gravado nesta sessão (até à próxima consulta) ----------
  const local = {faturas: new Map(), recibos: new Map(), liquidacoes: new Map(), ligacoes: new Map(), soSessao: 0};
  function juntar(base, extra, chave) {
    if (!extra.size) return base;
    const fora = new Set(extra.keys());
    return [...base.filter((b) => !fora.has(String(b[chave] ?? "").trim())), ...extra.values()];
  }
  function dados() {
    const t = ctx().tabela;
    const recibosLocais = new Set(local.recibos.keys());
    return {
      faturas: juntar(t("tbl_Faturas"), local.faturas, "ID_Documento"),
      recibos: juntar(t("tbl_Recibos"), local.recibos, "ID_Recibo"),
      liquidacoes: [...t("tbl_Liquidacoes").filter((l) => !recibosLocais.has(String(l.ID_Recibo ?? "").trim())), ...local.liquidacoes.values()],
      ligacoes: juntar(t("tbl_RecebimentosBanco"), local.ligacoes, "ID_Ligacao"),
      importacoes: t("tbl_Importacoes"),
    };
  }
  function registarLocal(tipo, registos, soSessao) {
    for (const r of registos) {
      const {linhas, _linha, _mudancas, ...limpo} = r;
      local[tipo].set(r[TIPO[tipo].chave], limpo);
      if (tipo === "recibos") {
        for (const k of [...local.liquidacoes.keys()]) if (k.startsWith(r.ID_Recibo + ">")) local.liquidacoes.delete(k);
        for (const l of linhas || []) local.liquidacoes.set(l.ID_Liquidacao, l);
      }
    }
    if (soSessao) local.soSessao += registos.length;
    desenharTudo();
  }
  window.TES_BASE_DADOS = Object.freeze({
    dados,
    registarLigacoes: (ligs) => { for (const l of ligs) local.ligacoes.set(l.ID_Ligacao, l); desenharTudo(); },
  });

  // ---------- Importador (um por secção) ----------
  const importadores = {};
  function montarImportador(caixa) {
    const tipo = caixa.dataset.tipo, info = TIPO[tipo], p = "bm-" + tipo + "-";
    caixa.innerHTML = '<div class="panel"><div class="panel-head"><h3>Importar do PRIMAVERA</h3><span class="note">' + esc(info.ficheiro) + ', exportada para Excel.</span></div>' +
      '<div class="form-grid bm-grid"><label>Empresa da listagem<input id="' + p + 'empresa" list="' + p + 'empresas" placeholder="Usada se o ficheiro não tiver a coluna Empresa" autocomplete="off"><datalist id="' + p + 'empresas"></datalist></label>' +
      '<label class="conc-drop bm-drop"><input type="file" id="' + p + 'ficheiro" accept=".xlsx,.xls,.xlsm,.ods,.csv"><strong>Seleccione ou arraste o ficheiro Excel</strong><span>A base mãe compara-o com o que já tem antes de gravar</span></label></div>' +
      '<p class="note" id="' + p + 'progresso" role="status" aria-live="polite"></p>' +
      '<div id="' + p + 'previa" hidden><div class="cards bm-cards" id="' + p + 'cards"></div>' +
      '<div class="conc-tabs" role="tablist">' + [["novos", "Novos"], ["alterados", "Alterados"], ["rejeitados", "Rejeitados"], ["ausentes", "Na base mas não no ficheiro"], ["iguais", "Iguais"]]
        .map(([k, t], i) => '<button type="button" class="conc-tab' + (i ? "" : " active") + '" data-bm-tab="' + k + '">' + t + "</button>").join("") + "</div>" +
      '<div class="table-wrap bm-tabela"><table><thead id="' + p + 'head"></thead><tbody id="' + p + 'body"></tbody></table></div>' +
      '<div class="bm-accoes"><button type="button" class="btn btn-primary" id="' + p + 'gravar">Gravar na base mãe</button><button type="button" class="btn" id="' + p + 'sessao" hidden>Usar só nesta sessão</button>' +
      '<button type="button" class="btn" id="' + p + 'colunas">Indicar colunas</button><button type="button" class="btn" id="' + p + 'csv">Exportar diferenças</button><button type="button" class="text-btn" id="' + p + 'cancelar">Cancelar</button><span class="note" id="' + p + 'nota"></span></div>' +
      '<div id="' + p + 'mapa" class="conc-mapa" hidden></div></div></div>';
    const st = {tipo, p, leitura: null, matriz: null, folha: "", hash: "", ficheiro: "", diff: null, tab: "novos", mapa: null};
    importadores[tipo] = st;
    const inp = $(p + "ficheiro");
    inp.addEventListener("change", (e) => ler(st, e.target.files[0]));
    const drop = inp.closest(".bm-drop");
    drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", () => drop.classList.remove("over"));
    drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); ler(st, e.dataTransfer.files[0]); });
    caixa.querySelectorAll("[data-bm-tab]").forEach((b) => b.addEventListener("click", () => { st.tab = b.dataset.bmTab; desenharPrevia(st); }));
    $(p + "gravar").addEventListener("click", () => gravar(st));
    $(p + "sessao").addEventListener("click", () => { const g = aGravar(st); registarLocal(tipo, g, true); fecharPrevia(st, g.length + " registo(s) aplicados só nesta sessão (não gravados no Excel).", "conc-ok"); });
    $(p + "cancelar").addEventListener("click", () => fecharPrevia(st, "", ""));
    $(p + "csv").addEventListener("click", () => exportarDiferencas(st));
    $(p + "colunas").addEventListener("click", () => abrirColunas(st));
    $(p + "empresa").addEventListener("change", () => { if (st.matriz) interpretar(st); });
  }
  const prog = (st, t, cls = "") => { const el = $(st.p + "progresso"); el.textContent = t; el.className = "note " + cls; };

  async function ler(st, ficheiro) {
    if (!ficheiro) return;
    prog(st, "A ler «" + ficheiro.name + "»…", "conc-a-ler");
    try {
      const {folhas, hash} = await IMP.lerFolhas(ficheiro);
      // A folha que produz mais registos válidos
      let melhor = null;
      for (const f of folhas) {
        const r = B.lerListagem(f.matriz, {...opcoesLeitura(st, ficheiro.name)});
        if (!melhor || r.registos.length > melhor.r.registos.length) melhor = {f, r};
      }
      Object.assign(st, {matriz: melhor?.f.matriz || [], folha: melhor?.f.folha || "", hash, ficheiro: ficheiro.name, mapa: null});
      // Empresa detectada (cabeçalho do relatório, nome do ficheiro ou clientes): preenchida no campo, editável
      const det = melhor?.r?.meta?.empresaDetectada;
      if (det && !$(st.p + "empresa").value.trim()) $(st.p + "empresa").value = det;
      interpretar(st);
      const repetida = dados().importacoes.find((i) => hash && String(i.Hash_Ficheiro || "") === hash);
      if (repetida) $(st.p + "nota").textContent = "Atenção: este mesmo ficheiro já foi importado (" + String(repetida.ID_Importacao) + ").";
    } catch (e) { prog(st, e.message || String(e), "conc-erro"); }
    finally { $(st.p + "ficheiro").value = ""; }
  }
  // Opções de leitura: empresa escolhida, nomes do cadastro, nome do ficheiro (empresa/período) e facturas da base
  // (para deduzir a empresa de ficheiros que não a indicam)
  function opcoesLeitura(st, nomeFicheiro) {
    return {tipo: st.tipo, empresa: $(st.p + "empresa").value.trim(), empresas: ctx().empresas().map((e) => e.nome), nomeFicheiro: nomeFicheiro || st.ficheiro || "",
      faturasBase: st.tipo === "recibos" ? dados().faturas : []};
  }
  function interpretar(st) {
    const opcoes = opcoesLeitura(st);
    if (st.mapa) Object.assign(opcoes, st.mapa);
    const r = B.lerListagem(st.matriz, opcoes);
    st.leitura = r;
    const base = dados()[st.tipo];
    st.diff = B.comparar(r.registos, base, st.tipo);
    st.tab = st.diff.novos.length ? "novos" : st.diff.alterados.length ? "alterados" : r.rejeitados.length ? "rejeitados" : "iguais";
    $(st.p + "previa").hidden = false;
    $(st.p + "nota").textContent = "";
    if (r.meta?.precisaEmpresa) {
      prog(st, r.avisos[0], "conc-erro"); $(st.p + "empresa").focus();
    } else if (!r.registos.length && !r.rejeitados.length) {
      prog(st, (r.avisos[0] || "Nenhum registo reconhecido.") + " Use «Indicar colunas».", "conc-erro");
    } else {
      const semEmpresa = r.rejeitados.some((x) => /empresa em falta/.test(x.motivo));
      prog(st, "«" + st.ficheiro + "» (folha " + st.folha + "): " + r.registos.length + " " + TIPO[st.tipo].nome + " lidos" + (r.rejeitados.length ? ", " + r.rejeitados.length + " linha(s) rejeitada(s)" : "") + "." + (semEmpresa ? " Indique a empresa da listagem." : ""), r.registos.length ? "conc-ok" : "conc-erro");
    }
    desenharPrevia(st);
  }
  function desenharPrevia(st) {
    const d = st.diff, r = st.leitura; if (!d) return;
    const n = {novos: d.novos.length, alterados: d.alterados.length, rejeitados: r.rejeitados.length, ausentes: d.ausentes.length, iguais: d.iguais.length};
    $(st.p + "cards").innerHTML = [["Lidos no ficheiro", r.registos.length, r.controlo ? (r.controlo.ok ? "totais do relatório conferem" : "totais NÃO conferem") : ""], ["Novos", n.novos, "entram na base"], ["Alterados", n.alterados, "diferem da base"], ["Rejeitados", n.rejeitados, "não serão gravados"]]
      .map(([t, v, s]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(s) + "</span></div>").join("");
    // Totais de controlo do relatório e avisos da leitura
    let ctl = document.getElementById(st.p + "controlo");
    if (!ctl) { ctl = document.createElement("div"); ctl.id = st.p + "controlo"; ctl.className = "bm-controlo"; $(st.p + "cards").after(ctl); }
    ctl.innerHTML = (r.controlo ? '<div class="bm-controlo-tit">Totais de controlo do relatório ' + (r.controlo.ok ? '<span class="tag conc-tag-ok">todos conferem</span>' : '<span class="tag waiting">há diferenças</span>') + "</div><table><tbody>" +
      r.controlo.linhas.map((x) => "<tr><td>" + esc(x.rotulo) + "</td><td>" + esc(x.tipo || "?") + (x.n ? " · " + x.n + " doc." : "") + '</td><td class="num">' + esc(mt(x.valor)) + '</td><td class="num">' + esc(mt(x.soma)) + "</td><td>" + (x.confere ? "✓" : "✗ " + esc(mt(x.valor - x.soma))) + "</td></tr>").join("") + "</tbody></table>" : "") +
      (r.avisos?.length ? "<ul>" + r.avisos.map((a) => "<li>" + esc(a) + "</li>").join("") + "</ul>" : "");
    document.querySelectorAll("#" + (st.tipo === "faturas" ? "facturacao" : "recebimentos") + " [data-bm-tab]").forEach((b) => { b.classList.toggle("active", b.dataset.bmTab === st.tab); b.dataset.n = n[b.dataset.bmTab]; });
    const chave = TIPO[st.tipo].chave;
    let head, linhas;
    const valor = (x) => mt(st.tipo === "faturas" ? x.Total : x.Valor);
    if (st.tab === "rejeitados") {
      head = "<tr><th>Linha do ficheiro</th><th>Motivo</th></tr>";
      linhas = r.rejeitados.map((x) => "<tr><td>" + esc(x.linha) + '</td><td class="conc-wrap">' + esc(x.motivo) + "</td></tr>");
    } else if (st.tab === "alterados") {
      head = '<tr><th>Documento</th><th>Cliente</th><th>O que muda</th></tr>';
      linhas = d.alterados.map((x) => "<tr><td>" + esc(x.Documento) + '</td><td class="conc-wrap">' + esc(x.Cliente) + '</td><td class="conc-wrap">' +
        x._mudancas.map((m) => "<b>" + esc(m.campo) + "</b>: " + esc(fmtCampo(m.campo, m.antes)) + " → " + esc(fmtCampo(m.campo, m.depois))).join("<br>") + "</td></tr>");
    } else {
      const lista = st.tab === "novos" ? d.novos : st.tab === "ausentes" ? d.ausentes : d.iguais;
      head = '<tr><th>Documento</th><th>Data</th><th>Cliente</th><th class="num">' + (st.tipo === "faturas" ? "Total" : "Valor") + "</th><th>" + (st.tipo === "faturas" ? "Vencimento" : "Documentos liquidados") + "</th><th>Estado</th></tr>";
      linhas = lista.map((x) => "<tr><td>" + esc(x.Documento || x[chave]) + "</td><td>" + esc(dataPt(x.Data)) + '</td><td class="conc-wrap">' + esc(x.Cliente) + '</td><td class="num">' + esc(valor(x)) + "</td><td>" +
        (st.tipo === "faturas" ? esc(dataPt(x.Data_Vencimento)) : esc((x.linhas || []).map((l) => l.Documento_Liquidado + " (" + mt(l.Valor) + ")").join(", ") || String(x.Liquidacoes || "—"))) + "</td><td>" + esc(x.Estado) + "</td></tr>");
    }
    $(st.p + "head").innerHTML = head;
    $(st.p + "body").innerHTML = linhas.slice(0, 500).join("") || '<tr><td colspan="6" class="empty">Nada nesta lista.</td></tr>';
    if (st.tab === "ausentes" && d.ausentes.length) $(st.p + "nota").textContent = "Documentos que a base mãe tem no mesmo período e não vêm no ficheiro: confirme no PRIMAVERA se foram anulados. Não são apagados.";
    const escrita = ctx().escritaDisponivel();
    const g = aGravar(st).length;
    $(st.p + "gravar").disabled = !escrita || !g;
    $(st.p + "gravar").textContent = g ? "Gravar na base mãe (" + g + ")" : "Nada a gravar";
    $(st.p + "gravar").title = escrita ? "" : "Ligue a base de dados e indique o endpoint de escrita V9 em «Conectar base de dados».";
    $(st.p + "sessao").hidden = escrita || !g;
  }
  const fmtCampo = (c, v) => (/^Data/.test(c) ? dataPt(v) : /Total|Valor/.test(c) ? mt(v) : c === "Liquidacoes" ? String(v || "").split(";").length + " linha(s)" : String(v ?? ""));
  const aGravar = (st) => (st.diff ? [...st.diff.novos, ...st.diff.alterados] : []);
  function fecharPrevia(st, mensagem, cls) {
    $(st.p + "previa").hidden = true; st.diff = null; st.leitura = null;
    prog(st, mensagem, cls);
  }

  async function gravar(st) {
    const registos = aGravar(st);
    if (!registos.length) return;
    const botao = $(st.p + "gravar"); botao.disabled = true;
    const idImp = "IMP-" + new Date().toISOString().replace(/\D/g, "").slice(0, 14) + "-" + Math.random().toString(36).slice(2, 6).toUpperCase();
    const importacao = {ID_Importacao: idImp, Empresa: $(st.p + "empresa").value.trim() || [...new Set(registos.map((r) => r.Empresa))].join(", "),
      Ficheiro: st.ficheiro, Hash_Ficheiro: st.hash, Data: new Date().toISOString(), Linhas_Ficheiro: st.leitura.registos.length + st.leitura.rejeitados.length,
      Novos: st.diff.novos.length, Alterados: st.diff.alterados.length, Iguais: st.diff.iguais.length, Rejeitados: st.leitura.rejeitados.length};
    let feitos = 0, novos = 0, alterados = 0;
    try {
      for (let k = 0; k < registos.length; k += LOTE) {
        const lote = registos.slice(k, k + LOTE);
        prog(st, "A gravar na base mãe: " + feitos + " de " + registos.length + "…", "conc-a-ler");
        const g = B.paraGravar(lote, st.tipo, idImp);
        const r = await ctx().escrever({operacao: "base.importar", id_pedido: crypto.randomUUID(), tipo: st.tipo, linhas: g.linhas, liquidacoes: g.liquidacoes || [], importacao});
        novos += Number(r.novos) || 0; alterados += Number(r.alterados) || 0; feitos += lote.length;
        registarLocal(st.tipo, lote, false);
      }
      fecharPrevia(st, "Gravado na base mãe (" + idImp + "): " + novos + " novo(s), " + alterados + " alterado(s).", "conc-ok");
    } catch (e) {
      botao.disabled = false;
      prog(st, (feitos ? feitos + " de " + registos.length + " já gravados. " : "") + e.message + " Pode repetir: os registos já gravados são actualizados, não duplicados.", "conc-erro");
    }
  }
  function exportarDiferencas(st) {
    if (!st.diff) return;
    const linhas = [
      ...st.diff.novos.map((x) => ({Situação: "Novo", Documento: x.Documento, Data: x.Data, Cliente: x.Cliente, Valor: st.tipo === "faturas" ? x.Total : x.Valor, Detalhe: ""})),
      ...st.diff.alterados.map((x) => ({Situação: "Alterado", Documento: x.Documento, Data: x.Data, Cliente: x.Cliente, Valor: st.tipo === "faturas" ? x.Total : x.Valor, Detalhe: x._mudancas.map((m) => m.campo + ": " + fmtCampo(m.campo, m.antes) + " → " + fmtCampo(m.campo, m.depois)).join("; ")})),
      ...st.leitura.rejeitados.map((x) => ({Situação: "Rejeitado", Documento: "", Data: "", Cliente: "", Valor: "", Detalhe: "linha " + x.linha + ": " + x.motivo})),
      ...st.diff.ausentes.map((x) => ({Situação: "Na base, não no ficheiro", Documento: x.Documento, Data: x.Data, Cliente: x.Cliente, Valor: st.tipo === "faturas" ? x.Total : x.Valor, Detalhe: ""})),
    ];
    descarregar(C.csv(linhas), "diferencas-" + st.tipo + "-" + st.ficheiro.replace(/\.[^.]+$/, "") + ".csv");
  }
  function descarregar(conteudo, nome) {
    const url = URL.createObjectURL(new Blob([conteudo], {type: "text/csv;charset=utf-8"}));
    const a = document.createElement("a"); a.href = url; a.download = nome.replace(/[^\w.\-]+/g, "_"); a.click(); URL.revokeObjectURL(url);
  }
  // Colunas indicadas à mão (quando o cabeçalho tem nomes diferentes)
  const CAMPOS_MAPA = {faturas: [["documento", "Documento (ex.: FA 2026/1)"], ["tipo", "Tipo de documento"], ["serie", "Série"], ["numero", "Número"], ["data", "Data"], ["vencimento", "Vencimento"], ["clienteCodigo", "Código do cliente"], ["clienteNome", "Nome do cliente"], ["nuit", "NUIT"], ["moeda", "Moeda"], ["total", "Total"], ["estado", "Estado/Anulado"], ["empresa", "Empresa"]],
    recibos: [["documento", "Recibo (ex.: RE 2026/1)"], ["tipo", "Tipo de documento"], ["serie", "Série"], ["numero", "Número"], ["data", "Data"], ["clienteCodigo", "Código do cliente"], ["clienteNome", "Nome do cliente"], ["total", "Total do recibo"], ["docLiquidado", "Documento liquidado"], ["valorLinha", "Valor liquidado"], ["modo", "Modo de pagamento"], ["referencia", "Referência"], ["estado", "Estado/Anulado"], ["empresa", "Empresa"]]};
  function abrirColunas(st) {
    const m = st.matriz || [], cab = st.leitura?.linhaCabecalho >= 0 ? st.leitura.linhaCabecalho : Math.max(0, m.findIndex((l) => l.filter((x) => String(x).trim()).length >= 3));
    const largura = Math.max(0, ...m.slice(0, 60).map((l) => l.length));
    const ops = '<option value="">—</option>' + Array.from({length: largura}, (_, j) => '<option value="' + j + '">' + esc("Coluna " + (j + 1) + (String(m[cab]?.[j] ?? "").trim() ? " — " + String(m[cab][j]).trim().slice(0, 28) : "")) + "</option>").join("");
    const caixa = $(st.p + "mapa"), actual = st.leitura?.mapa || {};
    caixa.innerHTML = "<p>Indique as colunas do ficheiro. É obrigatório o Documento (ou Tipo + Número), a Data e o Total/Valor.</p><div class=\"conc-mapa-campos\"><label>Linha do cabeçalho<input type=\"number\" min=\"1\" data-bm-linha value=\"" + (cab + 1) + "\"></label>" +
      CAMPOS_MAPA[st.tipo].map(([c, t]) => "<label>" + esc(t) + '<select data-bm-campo="' + c + '">' + ops + "</select></label>").join("") + '</div><div class="actions"><button type="button" class="btn btn-primary" data-bm-aplicar>Aplicar colunas</button><button type="button" class="btn" data-bm-fechar>Fechar</button></div>';
    for (const [c, j] of Object.entries(actual)) { const s = caixa.querySelector('[data-bm-campo="' + c + '"]'); if (s) s.value = String(j); }
    caixa.hidden = false;
    caixa.querySelector("[data-bm-fechar]").onclick = () => { caixa.hidden = true; };
    caixa.querySelector("[data-bm-aplicar]").onclick = () => {
      const mapa = {};
      caixa.querySelectorAll("[data-bm-campo]").forEach((s) => { if (s.value !== "") mapa[s.dataset.bmCampo] = Number(s.value); });
      st.mapa = {mapa, linhaCabecalho: Math.max(0, Number(caixa.querySelector("[data-bm-linha]").value || 1) - 1)};
      caixa.hidden = true; interpretar(st);
    };
  }

  // ---------- Vistas ----------
  function hoje() { return $("bm-cob-data").value || new Date().toISOString().slice(0, 10); }
  let calculo = null;
  function recalcular() { const d = dados(); calculo = {d, r: B.reconciliar({faturas: d.faturas, recibos: d.recibos, liquidacoes: d.liquidacoes, hoje: hoje()})}; return calculo; }
  const corEstado = (e) => ({"Liquidado": "pago", "Crédito": "pago", "Parcial": "saida", "Em aberto": "saida", "Vencido": "bm-vencido", "Parcial vencido": "bm-vencido"})[e] || "";
  function desenharEstado() {
    const d = calculo.d, ligado = ctx().ligado();
    const temBase = d.faturas.length + d.recibos.length > 0;
    const txt = (temBase ? "Base mãe: " + d.faturas.length + " documentos · " + d.recibos.length + " recibos" : ligado ? "Base mãe vazia (a consulta não devolveu tbl_Faturas/tbl_Recibos)" : "Modo demonstração · sem base mãe") +
      (local.soSessao ? " · " + local.soSessao + " só nesta sessão, não gravados" : "");
    document.querySelectorAll(".bm-estado").forEach((e) => { e.textContent = txt; });
  }
  function desenharFaturas() {
    const f = chave($("bm-fat-pesquisa").value), est = $("bm-fat-estado").value;
    const docs = calculo.r.documentos.filter((x) => (!est || x.estado === est) && (!f || chave([x.Documento, x.Cliente, x.Codigo_Cliente, x.NUIT].join(" ")).includes(f)))
      .sort((a, b) => String(b.Data).localeCompare(String(a.Data)));
    $("bm-fat-body").innerHTML = docs.slice(0, 500).map((x) => "<tr><td>" + esc(x.Documento) + "</td><td>" + esc(dataPt(x.Data)) + "</td><td>" + esc(dataPt(x.Data_Vencimento)) + (x.diasAtraso ? ' <small class="bm-atraso">' + x.diasAtraso + " d</small>" : "") + '</td><td class="conc-wrap">' + esc(x.Cliente) + "</td><td>" + esc(x.NUIT) + '</td><td class="num">' + esc(mt(x.efeito < 0 ? -x.Total : x.Total)) + '</td><td class="num">' + esc(mt(x.liquidado)) + '</td><td class="num">' + esc(mt(x.saldo)) + '</td><td><span class="badge ' + corEstado(x.estado) + '">' + esc(x.estado) + "</span></td></tr>").join("") ||
      '<tr><td colspan="9" class="empty">' + (calculo.d.faturas.length ? "Nenhum documento com estes filtros." : "A base mãe ainda não tem documentos. Importe a listagem de documentos de venda do PRIMAVERA.") + "</td></tr>";
    $("bm-fat-info").textContent = docs.length > 500 ? "A mostrar 500 de " + docs.length + " documentos. Use a pesquisa ou exporte para CSV." : docs.length ? docs.length + " documento(s)." : "";
  }
  const chave = (v) => String(v ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
  function desenharRecebimentos() {
    const {d, r} = calculo;
    const liqPorRecibo = new Map();
    for (const l of [...d.liquidacoes, ...(r.inferidas || [])]) { const k = String(l.ID_Recibo ?? "").trim(); liqPorRecibo.set(k, [...(liqPorRecibo.get(k) || []), l]); }
    const ligados = new Map(d.ligacoes.map((l) => [String(l.ID_Recibo ?? "").trim(), l]));
    const f = chave($("bm-rec-pesquisa").value);
    const recs = d.recibos.filter((x) => !f || chave([x.Documento, x.Cliente, x.Codigo_Cliente, x.Referencia].join(" ")).includes(f)).sort((a, b) => String(C.data(b.Data)).localeCompare(String(C.data(a.Data))));
    $("bm-rec-body").innerHTML = recs.slice(0, 500).map((x) => {
      const id = String(x.ID_Recibo ?? "").trim(), lig = ligados.get(id);
      return "<tr><td>" + esc(x.Documento) + "</td><td>" + esc(dataPt(x.Data)) + '</td><td class="conc-wrap">' + esc(x.Cliente) + '</td><td class="num">' + esc(mt(C.numero(x.Valor))) + "</td><td>" + esc(x.Modo_Pagamento) + "</td><td>" + esc(x.Referencia) + '</td><td class="conc-wrap">' +
        ((liqPorRecibo.get(id) || []).map((l) => esc((l.Documento_Liquidado || l.ID_Documento) + " (" + mt(C.numero(l.Valor)) + ")") + (l.inferida ? ' <span class="note" title="' + esc(l.regra) + '">(inferida)</span>' : "")).join(", ") || "por alocar") + "</td><td>" +
        (lig ? '<span class="badge pago" title="' + esc(lig.Descritivo_Banco) + '">' + esc(dataPt(lig.Data_Banco)) + "</span>" : /numer|dinheiro/.test(chave(x.Modo_Pagamento)) ? '<span class="note">numerário</span>' : '<span class="badge saida">por ligar</span>') + "</td></tr>";
    }).join("") || '<tr><td colspan="8" class="empty">' + (d.recibos.length ? "Nenhum recibo com esta pesquisa." : "A base mãe ainda não tem recibos. Importe a listagem de recibos do PRIMAVERA.") + "</td></tr>";
    const valorRecebido = d.recibos.filter((x) => String(x.Estado) !== "Anulado").reduce((s, x) => s + (C.numero(x.Valor) || 0), 0);
    const activos = d.recibos.filter((x) => String(x.Estado) !== "Anulado");
    const nLigados = activos.filter((x) => ligados.has(String(x.ID_Recibo ?? "").trim())).length;
    const porLigar = activos.filter((x) => !ligados.has(String(x.ID_Recibo ?? "").trim()) && !/numer|dinheiro/.test(chave(x.Modo_Pagamento))).length;
    $("bm-rec-cards").innerHTML = [["Recibos", d.recibos.length, mt(valorRecebido) + " MT"], ["Ligados ao banco", nLigados, porLigar + " por ligar (excepto numerário)"],
      ["Ligações inferidas", (r.inferidas || []).length, (r.inferidas || []).filter((x) => /rever/.test(x.regra)).length + " a rever (vários com o mesmo valor)"], ["Excepções", r.excepcoes.length, "a resolver"],
      ["Facturas em aberto", r.documentos.filter((x) => x.saldo > 0).length, mt(r.documentos.reduce((s, x) => s + x.saldo, 0)) + " MT"]]
      .map(([t, v, s]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(s) + "</span></div>").join("");
    $("bm-exc-body").innerHTML = r.excepcoes.map((e) => '<tr><td class="conc-wrap">' + esc(e.tipo) + '</td><td class="conc-wrap">' + esc(e.candidatos ? "candidatos: " + e.candidatos : e.documento) + "</td><td>" + esc(e.reciboDoc || "") + '</td><td class="conc-wrap">' + esc(e.cliente) + '</td><td class="num">' + esc(mt(e.valor)) + "</td></tr>").join("") ||
      '<tr><td colspan="5" class="empty">' + (d.recibos.length ? "Sem excepções: todos os recibos estão alocados a documentos da base." : "Sem dados.") + "</td></tr>";
    $("bm-rec-resumo").textContent = "Facturas ↔ recibos calculado em " + dataPt(hoje()) + ".";
  }
  let clienteAberto = "";
  function desenharCobrancas() {
    const {r} = calculo, f = chave($("bm-cob-pesquisa").value);
    const cl = r.clientes.filter((c) => (c.emAberto > 0 || c.creditos > 0) && (!f || chave(c.Cliente + " " + c.Codigo_Cliente).includes(f)));
    const soma = (k) => r.clientes.reduce((s, c) => s + c[k], 0);
    $("bm-cob-cards").innerHTML = [["Em aberto", mt(soma("emAberto")) + " MT", r.clientes.filter((c) => c.emAberto > 0).length + " cliente(s)"],
      ["Vencido", mt(soma("d30") + soma("d60") + soma("d90") + soma("mais90")) + " MT", "à data de " + dataPt(hoje())], ["Mais de 90 dias", mt(soma("mais90")) + " MT", "prioridade de cobrança"],
      ["Créditos por regularizar", mt(soma("creditos")) + " MT", "notas de crédito e recibos por alocar"]]
      .map(([t, v, s]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(s) + "</span></div>").join("");
    $("bm-cob-body").innerHTML = cl.map((c) => {
      const linha = '<tr class="bm-cliente" data-cliente="' + esc(c.chave) + '"><td class="conc-wrap"><b>' + esc(c.Cliente || c.Codigo_Cliente) + "</b><br><small>" + esc(c.Codigo_Cliente) + (c.Empresa ? " · " + esc(c.Empresa) : "") + '</small></td><td class="num"><b>' + esc(mt(c.emAberto)) + "</b></td>" +
        ["naoVencido", "d30", "d60", "d90", "mais90"].map((k) => '<td class="num' + (k === "mais90" && c[k] > 0 ? " bm-vencido-txt" : "") + '">' + (c[k] ? esc(mt(c[k])) : "") + "</td>").join("") +
        '<td class="num">' + (c.creditos ? esc(mt(c.creditos)) : "") + '</td><td class="num"><b>' + esc(mt(c.saldoLiquido)) + "</b></td><td>" + (c.maiorAtraso ? esc(c.maiorAtraso + " dias") : "—") + "</td><td>" + esc(c.ultimoRecibo ? dataPt(c.ultimoRecibo) : "—") + "</td></tr>";
      if (c.chave !== clienteAberto) return linha;
      const docs = r.documentos.filter((x) => x.saldo > 0 && chave(x.Empresa) + "|" + (chave(x.Codigo_Cliente) || chave(x.Cliente)) === c.chave).sort((a, b) => b.diasAtraso - a.diasAtraso);
      return linha + '<tr class="bm-detalhe"><td colspan="11"><table><thead><tr><th>Documento</th><th>Data</th><th>Vencimento</th><th class="num">Total</th><th class="num">Liquidado</th><th class="num">Saldo</th><th>Atraso</th></tr></thead><tbody>' +
        docs.map((x) => "<tr><td>" + esc(x.Documento) + "</td><td>" + esc(dataPt(x.Data)) + "</td><td>" + esc(dataPt(x.Data_Vencimento)) + '</td><td class="num">' + esc(mt(x.Total)) + '</td><td class="num">' + esc(mt(x.liquidado)) + '</td><td class="num">' + esc(mt(x.saldo)) + "</td><td>" + (x.diasAtraso ? esc(x.diasAtraso + " dias") : "não vencido") + "</td></tr>").join("") + "</tbody></table></td></tr>";
    }).join("") || '<tr><td colspan="11" class="empty">' + (calculo.d.faturas.length ? "Sem saldos em aberto." : "Sem dados: importe facturas e recibos do PRIMAVERA.") + "</td></tr>";
  }
  $("bm-cob-body").addEventListener("click", (e) => { const tr = e.target.closest("tr.bm-cliente"); if (!tr) return; clienteAberto = clienteAberto === tr.dataset.cliente ? "" : tr.dataset.cliente; desenharCobrancas(); });
  function desenharTudo() {
    recalcular(); desenharEstado(); desenharFaturas(); desenharRecebimentos(); desenharCobrancas();
    for (const st of Object.values(importadores)) {
      const emp = ctx().empresas();
      $(st.p + "empresas").innerHTML = emp.map((e) => '<option value="' + esc(e.nome) + '">').join("");
      if (st.diff) desenharPrevia(st);
    }
  }
  $("bm-fat-pesquisa").addEventListener("input", desenharFaturas);
  $("bm-fat-estado").addEventListener("change", desenharFaturas);
  $("bm-rec-pesquisa").addEventListener("input", desenharRecebimentos);
  $("bm-cob-pesquisa").addEventListener("input", desenharCobrancas);
  $("bm-cob-data").addEventListener("change", desenharTudo);
  $("bm-fat-csv").addEventListener("click", () => {
    const linhas = calculo.r.documentos.map((x) => ({Documento: x.Documento, Empresa: x.Empresa, Data: x.Data, Vencimento: x.Data_Vencimento, "Código cliente": x.Codigo_Cliente, Cliente: x.Cliente, NUIT: x.NUIT,
      Total: x.efeito < 0 ? -x.Total : x.Total, Liquidado: x.liquidado, Saldo: x.saldo, Estado: x.estado, "Dias de atraso": x.diasAtraso || ""}));
    descarregar(C.csv(linhas), "facturacao-base-mae-" + hoje() + ".csv");
  });
  $("bm-cob-csv").addEventListener("click", () => {
    const linhas = calculo.r.clientes.filter((c) => c.emAberto > 0 || c.creditos > 0).map((c) => ({Empresa: c.Empresa, "Código cliente": c.Codigo_Cliente, Cliente: c.Cliente, "Em aberto": c.emAberto,
      "Não vencido": c.naoVencido, "1-30": c.d30, "31-60": c.d60, "61-90": c.d90, "> 90": c.mais90, Créditos: c.creditos, "Saldo líquido": c.saldoLiquido, "Maior atraso (dias)": c.maiorAtraso, "Último recibo": c.ultimoRecibo}));
    descarregar(C.csv(linhas), "cobrancas-antiguidade-" + hoje() + ".csv");
  });

  document.querySelectorAll(".bm-importador").forEach(montarImportador);
  $("bm-cob-data").value = new Date().toISOString().slice(0, 10);
  document.addEventListener("tes:contexto", desenharTudo);
  ["facturacao", "recebimentos", "cobrancas"].forEach((v) => document.querySelector('.nav[data-view="' + v + '"]')?.addEventListener("click", desenharTudo));
  desenharTudo();
})();
