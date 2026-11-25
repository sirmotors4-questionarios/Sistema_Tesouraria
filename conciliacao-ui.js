/* Conciliação bancária — interface. Lê as despesas/contas carregadas pelo app.js (TES_CONTEXTO),
 * importa o extracto (TES_IMPORTAR), emparelha (TES_CONCILIACAO) e exporta o relatório.
 * Gravação na base mãe, sempre pelo fluxo de escrita do Power Automate (Office Script TesourariaBaseMae):
 *   banco.importar_extrato — movimentos do extracto (o script ignora os que já existem; o ficheiro é registado pelo SHA-256);
 *   banco.conciliar / banco.desfazer — correspondências movimento ↔ despesa/recibo (um documento ↔ um movimento).
 * No navegador só se guardam o mapeamento de colunas por conta e a opção «guardar o ficheiro no SharePoint».
 */
(function () {
  "use strict";
  const C = window.TES_CONCILIACAO, IMP = window.TES_IMPORTAR;
  const $ = (id) => document.getElementById(id);
  if (!C || !IMP || !$("conciliacao")) return;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"})[c]);
  const mt = (n) => (n === "" || n == null || !Number.isFinite(Number(n)) ? "" : C.formatar(n));
  const dataPt = (iso) => (/^\d{4}-\d{2}-\d{2}$/.test(iso || "") ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—");
  const ctx = () => window.TES_CONTEXTO || {ligado: () => false, contas: () => [], registos: () => []};
  const MAPAS_KEY = "sir.tesouraria.conciliacao.mapas.v1", ENVIAR_KEY = "sir.tesouraria.bancos.enviarFicheiro.v1";
  const LOTE_MOV = 1000, LOTE_LIG = 1000, MAX_ENVIO = 10 * 1024 * 1024;

  // ---------- Base mãe (tabelas devolvidas pela consulta + o que foi gravado nesta sessão) ----------
  const sessao = {extratos: [], movimentos: [], conciliacoes: [], desfeitos: new Set()};
  const T = (n) => (ctx().tabela ? ctx().tabela(n) : []);
  const extratosBase = () => [...T("tbl_Extratos"), ...sessao.extratos];
  const movimentosBase = () => [...T("tbl_MovimentosBancarios"), ...sessao.movimentos];
  const idsBase = () => new Set(movimentosBase().map((l) => String(l.ID_Movimento ?? "").trim()).filter(Boolean));
  const conciliacoesBase = () => [...T("tbl_ConciliacaoBancaria").filter((l) => !sessao.desfeitos.has(String(l.ID_Movimento))), ...sessao.conciliacoes];
  const consultaTemBase = () => ctx().ligado() && (T("tbl_MovimentosBancarios").length > 0 || T("tbl_Extratos").length > 0);
  const escrita = () => !!(ctx().escritaDisponivel && ctx().escritaDisponivel());
  const avisarBancos = () => document.dispatchEvent(new Event("tes:bancos"));
  window.TES_BANCOS = Object.freeze({extratos: extratosBase, movimentos: movimentosBase, conciliacoes: conciliacoesBase, consultaTemBase});
  const banco1 = (b) => String(b ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9 ]/g, " ").trim().split(/\s+/)[0] || "";

  const estado = {leitura: null, movimentos: [], saldoInicial: null, saldoFinal: null, pagamentos: [], resultado: null,
    tab: "conciliados", confirmados: [], rejeitados: new Set(), desactualizado: false,
    ficheiro: null, origem: "ficheiro", gravados: [], msgBase: null, aGravar: false};

  // ---------- Conta ----------
  const contaTexto = (c) => [c.Banco ?? c.Nome_Banco, c.Designacao_Conta, c.Numero_Conta ?? c["Nº de Conta"] ?? c.Nr_Conta, c.Moeda].map((v) => String(v ?? "").trim()).filter(Boolean).join(" · ");
  function contasActivas() {
    return ctx().contas().filter((c) => String(c.Estado ?? "Activo").trim().toLowerCase() !== "inactivo" && String(c.ID_Conta ?? "").trim());
  }
  function preencherContas() {
    const sel = $("conc-conta"), actual = sel.value;
    const contas = contasActivas();
    sel.replaceChildren(new Option(contas.length ? "Todas as contas (sem filtro)" : "Sem cadastro de contas — filtrar pelo nº abaixo", ""),
      ...contas.map((c) => new Option(contaTexto(c) || String(c.ID_Conta), String(c.ID_Conta))));
    if ([...sel.options].some((o) => o.value === actual)) sel.value = actual;
    const regs = ctx().registos();
    $("conc-fonte").textContent = ctx().ligado()
      ? "Dados da Tesouraria: consulta ao Excel (" + regs.length + " despesas)."
      : "Modo demonstração: são usadas as despesas de demonstração deste navegador (" + regs.length + ").";
  }
  function contaSeleccionada() {
    const id = $("conc-conta").value;
    const c = contasActivas().find((x) => String(x.ID_Conta) === id);
    return {id, numero: $("conc-conta-num").value.trim() || String(c?.Numero_Conta ?? c?.["Nº de Conta"] ?? "").trim(), banco: String(c?.Banco ?? "").trim(), texto: c ? contaTexto(c) : ""};
  }
  $("conc-conta").addEventListener("change", () => {
    const c = contaSeleccionada();
    $("conc-conta-num").value = c.id ? String(contasActivas().find((x) => String(x.ID_Conta) === c.id)?.Numero_Conta ?? "") : "";
    if (estado.origem === "base" && estado.leitura) reiniciar(true);
    desenharBase(); invalidar();
  });
  $("conc-conta-num").addEventListener("input", () => { desenharBase(); invalidar(); });
  // A conta indicada no ficheiro selecciona a conta do cadastro (mesmo nº e, se o ficheiro o disser, mesmo banco).
  function escolherContaDoFicheiro(lf) {
    if (!lf || !lf.conta || $("conc-conta").value) return;
    const cand = contasActivas().filter((c) => C.mesmaConta(c.Numero_Conta ?? c["Nº de Conta"], lf.conta) && (!lf.banco || !c.Banco || banco1(c.Banco) === banco1(lf.banco)));
    if (cand.length === 1) { $("conc-conta").value = String(cand[0].ID_Conta); $("conc-conta-num").value = String(cand[0].Numero_Conta ?? lf.conta); }
    else if (!$("conc-conta-num").value) $("conc-conta-num").value = lf.conta;
  }

  // ---------- Leitura do extracto ----------
  const progresso = (t, tipo = "") => { const p = $("conc-progresso"); p.textContent = t; p.className = "note " + tipo; };
  async function ler(ficheiro) {
    if (!ficheiro) return;
    reiniciar(false);
    progresso("A ler «" + ficheiro.name + "»…", "conc-a-ler");
    $("conc-ficheiro").disabled = true;
    try {
      const leitura = await IMP.lerFicheiro(ficheiro, {aoProgresso: (t) => progresso(t, "conc-a-ler"), forcarOcr: $("conc-ocr").checked});
      // Sem cabeçalho reconhecido num Excel/CSV: tentar o mapeamento guardado para esta conta
      if (!leitura.resultado.movimentos.length && leitura.matriz) {
        const guardado = mapas()[chaveConta()];
        if (guardado) {
          const r = C.matrizParaMovimentos(leitura.matriz, {mapa: guardado.mapa, linhaCabecalho: guardado.linhaCabecalho});
          if (r.movimentos.length) leitura.resultado = r;
        }
      }
      estado.leitura = leitura; estado.ficheiro = ficheiro; estado.origem = "ficheiro";
      escolherContaDoFicheiro(leitura.contaFicheiro);
      aplicarResultado(leitura.resultado);
      const n = estado.movimentos.length;
      progresso(n ? "«" + ficheiro.name + "» lido por " + leitura.metodo + ": " + n + " movimento(s)." : "Não foram encontrados movimentos em «" + ficheiro.name + "» (" + leitura.metodo + ").", n ? "conc-ok" : "conc-erro");
      if (!n && leitura.matriz) abrirMapa();
    } catch (e) {
      progresso(e.message || String(e), "conc-erro");
    } finally {
      $("conc-ficheiro").disabled = false; $("conc-ficheiro").value = "";
    }
  }
  function aplicarResultado(r) {
    estado.movimentos = r.movimentos.map((m) => ({...m}));
    estado.saldoInicial = r.saldoInicial ?? null; estado.saldoFinal = r.saldoFinal ?? null;
    estado.avisos = r.avisos || [];
    $("conc-passo2").hidden = false;
    $("conc-passo3").hidden = !estado.movimentos.length;
    const datas = estado.movimentos.map((m) => m.data).sort();
    if (datas.length) { $("conc-de").value = datas[0]; $("conc-ate").value = datas[datas.length - 1]; }
    desenharMovimentos(); invalidar();
  }
  $("conc-ficheiro").addEventListener("change", (e) => ler(e.target.files[0]));
  const drop = $("conc-drop");
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("over"));
  drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); ler(e.dataTransfer.files[0]); });
  $("conc-limpar").addEventListener("click", () => reiniciar(true));
  function reiniciar(limparProgresso) {
    Object.assign(estado, {leitura: null, movimentos: [], saldoInicial: null, saldoFinal: null, resultado: null, confirmados: [], rejeitados: new Set(), tab: "conciliados",
      ficheiro: null, origem: "ficheiro", gravados: [], msgBase: null});
    $("conc-passo2").hidden = true; $("conc-passo3").hidden = true; $("conc-mapa").hidden = true; $("conc-base").hidden = true;
    if (limparProgresso) progresso("");
  }

  // ---------- Movimentos (revisão) ----------
  function desenharMovimentos() {
    const daBase = estado.origem === "base", ro = daBase ? " disabled" : "";
    const v = daBase ? {quebras: [], duplicados: [], problemas: [], ok: true, saldoInicial: null, saldoFinal: null, debitos: estado.movimentos.filter((m) => m.valor < 0).reduce((s, m) => s + m.valor, 0), creditos: estado.movimentos.filter((m) => m.valor > 0).reduce((s, m) => s + m.valor, 0)}
      : C.validar(estado.movimentos, {saldoInicial: estado.saldoInicial, saldoFinal: estado.saldoFinal});
    const quebra = new Set(v.quebras), dup = new Set(v.duplicados);
    $("conc-mov-body").innerHTML = estado.movimentos.map((m) =>
      '<tr data-id="' + esc(m.id) + '" class="' + (quebra.has(m.id) ? "conc-quebra " : "") + (dup.has(m.id) ? "conc-dup" : "") + '">' +
      '<td><input type="date" value="' + esc(m.data) + '" data-campo="data"' + ro + '></td>' +
      '<td><input value="' + esc(m.descricao) + '" data-campo="descricao" class="conc-desc"' + ro + '></td>' +
      '<td><input value="' + esc(m.referencia) + '" data-campo="referencia" class="conc-ref"' + ro + '></td>' +
      '<td class="num"><input value="' + (m.valor < 0 ? esc(mt(-m.valor)) : "") + '" data-campo="debito" class="conc-valor" inputmode="decimal"' + ro + '></td>' +
      '<td class="num"><input value="' + (m.valor > 0 ? esc(mt(m.valor)) : "") + '" data-campo="credito" class="conc-valor" inputmode="decimal"' + ro + '></td>' +
      '<td class="num">' + esc(mt(m.saldo)) + '</td>' +
      '<td class="conc-acoes">' + (daBase ? "" : '<button type="button" class="icon-btn" data-accao="sinal" title="Trocar débito/crédito">±</button><button type="button" class="icon-btn" data-accao="remover" title="Remover esta linha">✕</button>') + '</td></tr>').join("") ||
      '<tr><td colspan="7" class="empty">Sem movimentos. Use «Indicar colunas» se o ficheiro tiver outro formato.</td></tr>';
    const avisos = [...(estado.avisos || []), ...v.problemas];
    $("conc-validacao").innerHTML =
      '<div class="conc-saldos"><span>Saldo inicial <b>' + (v.saldoInicial == null ? "—" : esc(mt(v.saldoInicial))) + '</b></span><span>Débitos <b>' + esc(mt(-v.debitos)) + '</b></span><span>Créditos <b>' + esc(mt(v.creditos)) + '</b></span><span>Saldo final <b>' + (v.saldoFinal == null ? "—" : esc(mt(v.saldoFinal))) + '</b></span>' +
      '<span class="tag ' + (v.ok ? "conc-tag-ok" : "waiting") + '">' + (daBase ? "Movimentos da base mãe por conciliar (só leitura)" : v.ok ? "Saldos conferidos" : "Rever antes de conciliar") + "</span></div>" +
      (avisos.length ? "<ul>" + avisos.map((a) => "<li>" + esc(a) + "</li>").join("") + "</ul>" : "");
    estado.saldosOk = v.ok;
    desenharBase();
  }

  // ---------- Gravar o extracto na base mãe ----------
  function contaParaIds() {
    const c = contaSeleccionada();
    return {banco: c.banco || $("conc-banco").value.trim() || estado.leitura?.contaFicheiro?.banco || "", conta: c.numero, idConta: c.id};
  }
  // Identidade de cada movimento na base mãe (igual entre CSV/Excel/PDF): m.idBase
  function calcularIds() {
    if (estado.origem === "base") return "";
    const k = contaParaIds();
    if (!k.banco) { estado.movimentos.forEach((m) => { delete m.idBase; }); return "Indique o banco da conta."; }
    try { C.idsMovimentos(estado.movimentos, k).forEach((id, i) => { estado.movimentos[i].idBase = id; }); return ""; }
    catch (e) { estado.movimentos.forEach((m) => { delete m.idBase; }); return e.message; }
  }
  const baseEstado = (t, tipo = "") => { const p = $("conc-base-estado"); p.textContent = t; p.className = "note " + tipo; };
  const estadosFinais = ["Completo", "Sem movimentos novos"];
  function desenharBase() {
    const painel = $("conc-base");
    if (!estado.leitura || estado.origem === "base" || !estado.movimentos.length) { painel.hidden = true; return; }
    painel.hidden = false;
    const lf = estado.leitura.contaFicheiro || {}, c = contaSeleccionada();
    $("conc-banco-wrap").hidden = !!c.banco;
    if (!c.banco && !$("conc-banco").value && lf.banco) $("conc-banco").value = lf.banco;
    const erroIds = calcularIds(), k = contaParaIds(), info = [];
    const datas = estado.movimentos.map((m) => m.data).filter(Boolean).sort();
    estado.nomePadronizado = C.nomePadronizado({banco: k.banco, conta: k.conta, de: datas[0], ate: datas[datas.length - 1], extensao: estado.leitura.extensao});
    let bloqueio = "";
    if (lf.conta) {
      const confere = k.conta && C.mesmaConta(lf.conta, k.conta);
      info.push("Conta indicada no ficheiro: <b>" + esc(lf.conta) + "</b>" + (lf.nib ? " · NIB " + esc(lf.nib) : "") + (confere ? ' <span class="tag conc-tag-ok">confere com a conta seleccionada</span>' : ""));
      if (!k.conta) bloqueio = "Seleccione a conta bancária (o ficheiro indica a conta " + lf.conta + ").";
      else if (!confere) bloqueio = "O ficheiro é da conta " + lf.conta + " mas está seleccionada a conta " + k.conta + ". Corrija a conta antes de gravar.";
    } else if (!k.conta) bloqueio = "Seleccione a conta bancária: o ficheiro não indica o número da conta.";
    if (!bloqueio && erroIds) bloqueio = erroIds;
    info.push("Nome padronizado: <b>" + esc(estado.nomePadronizado) + "</b> <span class=\"note\">(original: " + esc(estado.leitura.nome) + ")</span>");
    const ja = extratosBase().find((e) => String(e.ID_Extrato ?? "").toLowerCase() === estado.leitura.hash);
    if (ja) {
      info.push("Este ficheiro já foi importado: <b>" + esc(ja.Nome_Padronizado) + "</b> · " + esc(ja.Estado) + (ja.Novos !== undefined && ja.Novos !== "" ? " · " + esc(ja.Novos) + " movimento(s) novos" : ""));
      if (estadosFinais.includes(String(ja.Estado)) && !bloqueio) bloqueio = "Este ficheiro já está na base mãe: não é gravado outra vez.";
    }
    if (!erroIds) {
      const s = idsBase(), existentes = estado.movimentos.filter((m) => s.has(m.idBase)).length;
      info.push("<b>" + (estado.movimentos.length - existentes) + "</b> movimento(s) novo(s) · <b>" + existentes + "</b> já na base mãe (não são gravados outra vez)" +
        (consultaTemBase() || sessao.movimentos.length ? "" : ' <span class="note">— a consulta não devolveu a base de movimentos; o Power Automate volta a verificar ao gravar.</span>'));
    }
    $("conc-forcar-wrap").hidden = !!estado.saldosOk;
    if (!bloqueio && !estado.saldosOk && !$("conc-forcar").checked) bloqueio = "Os saldos não conferem: reveja as linhas ou marque «Gravar mesmo com avisos de saldo».";
    if (!bloqueio && !escrita()) bloqueio = "Para gravar, ligue a base de dados e indique o endpoint de escrita do Power Automate (V9) em «Conectar base de dados».";
    $("conc-base-info").innerHTML = info.map((t) => "<div>" + t + "</div>").join("");
    const b = $("conc-gravar-extrato");
    b.disabled = !!bloqueio || estado.aGravar; b.title = bloqueio;
    if (estado.msgBase) baseEstado(estado.msgBase.t, estado.msgBase.tipo);
    else if (!estado.aGravar) baseEstado(bloqueio, bloqueio ? "conc-erro" : "");
  }
  ["input", "change"].forEach((ev) => $("conc-banco").addEventListener(ev, () => { desenharBase(); invalidar(); }));
  $("conc-forcar").addEventListener("change", desenharBase);
  try { $("conc-enviar-ficheiro").checked = localStorage.getItem(ENVIAR_KEY) === "1"; } catch {}
  $("conc-enviar-ficheiro").addEventListener("change", (e) => { try { localStorage.setItem(ENVIAR_KEY, e.target.checked ? "1" : "0"); } catch {} });
  const paraBase64 = (f) => new Promise((ok, erro) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(",")[1] || ""); r.onerror = () => erro(r.error); r.readAsDataURL(f); });

  async function gravarExtrato() {
    desenharBase();
    if ($("conc-gravar-extrato").disabled) return;
    const k = contaParaIds(), agora = new Date().toISOString(), hash = estado.leitura.hash;
    const ids = estado.movimentos.map((m) => m.idBase);
    const todas = C.movimentosParaBase(estado.movimentos, ids, {idConta: k.idConta, banco: k.banco, conta: k.conta, idExtrato: hash, dataImportacao: agora});
    const conhecidos = idsBase(), enviar = todas.filter((l) => !conhecidos.has(l.ID_Movimento));
    const datas = todas.map((l) => l.Data_Movimento).sort();
    const v = C.validar(estado.movimentos, {saldoInicial: estado.saldoInicial, saldoFinal: estado.saldoFinal});
    const extrato = {ID_Extrato: hash, Nome_Original: estado.leitura.nome, Nome_Padronizado: estado.nomePadronizado, Banco: k.banco, ID_Conta: k.idConta,
      Conta: k.conta.replace(/\D/g, ""), Formato: estado.leitura.metodo, De: datas[0], Ate: datas[datas.length - 1], Movimentos_Ficheiro: todas.length,
      Saldo_Inicial: v.saldoInicial ?? "", Saldo_Final: v.saldoFinal ?? "", Utilizador: ""};
    const lotes = [];
    for (let i = 0; i < enviar.length; i += LOTE_MOV) lotes.push(enviar.slice(i, i + LOTE_MOV));
    if (!lotes.length) lotes.push([]);
    let ficheiro = null;
    if ($("conc-enviar-ficheiro").checked && estado.ficheiro) {
      if (estado.ficheiro.size > MAX_ENVIO) { estado.msgBase = {t: "O ficheiro tem mais de 10 MB: não é enviado para o SharePoint. Desmarque a opção ou guarde-o manualmente com o nome padronizado.", tipo: "conc-erro"}; desenharBase(); return; }
      ficheiro = {nome: estado.nomePadronizado, pasta: "02_Importacoes/Bancos/" + estado.nomePadronizado.split("_")[0], conteudo_base64: await paraBase64(estado.ficheiro)};
    }
    estado.aGravar = true; estado.msgBase = null; desenharBase();
    let novos = 0, duplicados = 0, ultimo = null;
    try {
      for (let i = 0; i < lotes.length; i++) {
        const final = i === lotes.length - 1;
        baseEstado("A gravar no Power Automate… lote " + (i + 1) + " de " + lotes.length + " (" + lotes[i].length + " movimentos).", "conc-a-ler");
        ultimo = await ctx().escrever({operacao: "banco.importar_extrato", id_pedido: crypto.randomUUID(), final, extrato, movimentos: lotes[i], ...(final && ficheiro ? {ficheiro} : {})});
        novos += Number(ultimo.novos) || 0; duplicados += Number(ultimo.duplicados) || 0;
        sessao.movimentos.push(...lotes[i]);
      }
      duplicados += todas.length - enviar.length;
      if (ultimo?.nome_padronizado) { extrato.Nome_Padronizado = ultimo.nome_padronizado; estado.nomePadronizado = ultimo.nome_padronizado; }
      sessao.extratos.push({...extrato, Novos: novos, Ja_Existentes: duplicados, Estado: ultimo?.estado || (novos ? "Completo" : "Sem movimentos novos"), Data_Importacao: agora});
      estado.msgBase = {t: "Gravado na base mãe: " + novos + " movimento(s) novo(s), " + duplicados + " já existiam (ignorados). Extracto registado como «" + estado.nomePadronizado + "»" + (ficheiro ? " e enviado para o SharePoint." : "."), tipo: "conc-ok"};
      avisarBancos();
    } catch (e) {
      const ja = /EXTRATO_JA_IMPORTADO/.test(e.message);
      estado.msgBase = {t: (ja ? "Este ficheiro já tinha sido importado: nada foi gravado. " : "") + e.message + (novos ? " (Lotes anteriores gravados: " + novos + " movimento(s); repetir é seguro, não duplica.)" : ""), tipo: "conc-erro"};
    } finally { estado.aGravar = false; desenharBase(); invalidar(); }
  }
  $("conc-gravar-extrato").addEventListener("click", gravarExtrato);
  $("conc-descarregar").addEventListener("click", () => {
    if (!estado.ficheiro) return;
    const url = URL.createObjectURL(estado.ficheiro), a = document.createElement("a");
    a.href = url; a.download = estado.nomePadronizado || estado.ficheiro.name; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  });

  // ---------- Conciliar a partir da base mãe (sem voltar a ler o ficheiro) ----------
  $("conc-da-base").addEventListener("click", () => {
    const c = contaSeleccionada();
    if (!c.numero) { progresso("Seleccione primeiro a conta bancária.", "conc-erro"); return; }
    const conc = new Set(conciliacoesBase().map((l) => String(l.ID_Movimento)));
    const movs = C.movimentosDaBase(movimentosBase(), {conta: c.numero, excluir: conc});
    if (!movs.length) {
      progresso(movimentosBase().length ? "Não há movimentos por conciliar da conta " + c.numero + " na base mãe." : "A consulta do Power Automate não devolveu a tabela tbl_MovimentosBancarios (ver docs/BANCOS_EXTRATOS.md, passo 3.4) ou ainda não foram gravados extractos.", "conc-erro");
      return;
    }
    reiniciar(false);
    estado.origem = "base"; estado.leitura = {metodo: "base mãe", nome: "base mãe", origem: "base"};
    aplicarResultado({movimentos: movs, saldoInicial: null, saldoFinal: null, avisos: []});
    progresso(movs.length + " movimento(s) por conciliar da conta " + c.numero + " carregados da base mãe.", "conc-ok");
  });
  $("conc-mov-body").addEventListener("change", (e) => {
    const tr = e.target.closest("tr[data-id]"); if (!tr) return;
    const m = estado.movimentos.find((x) => x.id === tr.dataset.id); if (!m) return;
    const campo = e.target.dataset.campo, v = e.target.value;
    if (campo === "debito" || campo === "credito") {
      const n = C.numero(v);
      if (v.trim() && !Number.isFinite(n)) { e.target.classList.add("conc-invalido"); return; }
      e.target.classList.remove("conc-invalido");
      if (Number.isFinite(n) && n !== 0) m.valor = campo === "debito" ? -Math.abs(n) : Math.abs(n);
    } else m[campo] = campo === "data" ? v : v.trim();
    desenharMovimentos(); invalidar();
  });
  $("conc-mov-body").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-accao]"); if (!b) return;
    const id = b.closest("tr").dataset.id, i = estado.movimentos.findIndex((x) => x.id === id);
    if (i < 0) return;
    if (b.dataset.accao === "sinal") estado.movimentos[i].valor = -estado.movimentos[i].valor;
    else estado.movimentos.splice(i, 1);
    desenharMovimentos(); invalidar();
  });

  // ---------- Mapeamento manual de colunas ----------
  const NOMES = {data: "Data do movimento", dataValor: "Data-valor", descricao: "Descritivo", referencia: "Referência", debito: "Débito", credito: "Crédito", montante: "Montante (com sinal)", saldo: "Saldo"};
  const mapas = () => { try { return JSON.parse(localStorage.getItem(MAPAS_KEY) || "{}") || {}; } catch { return {}; } };
  const chaveConta = () => { const c = contaSeleccionada(); return c.id || c.numero || "geral"; };
  function abrirMapa() {
    const matriz = estado.leitura?.matriz;
    if (!matriz || !matriz.length) { progresso("Este ficheiro não tem colunas para indicar (texto lido de PDF/OCR). Corrija as linhas directamente na tabela.", "conc-erro"); return; }
    const linhaCab = estado.leitura.resultado.linhaCabecalho >= 0 ? estado.leitura.resultado.linhaCabecalho : Math.max(0, matriz.findIndex((l) => l.filter((x) => String(x).trim()).length >= 3));
    const cab = matriz[linhaCab] || [];
    const largura = Math.max(...matriz.slice(0, 60).map((l) => l.length));
    const opcoes = ['<option value="">—</option>', ...Array.from({length: largura}, (_, j) => '<option value="' + j + '">' + esc("Coluna " + String.fromCharCode(65 + (j % 26)) + (j >= 26 ? Math.floor(j / 26) : "") + (String(cab[j] ?? "").trim() ? " — " + String(cab[j]).trim().slice(0, 30) : "")) + "</option>")].join("");
    const actual = estado.leitura.resultado.mapa || mapas()[chaveConta()]?.mapa || {};
    $("conc-mapa-campos").innerHTML = '<label>Linha do cabeçalho<input type="number" min="1" id="conc-mapa-linha" value="' + (linhaCab + 1) + '"></label>' +
      Object.keys(NOMES).map((c) => '<label>' + NOMES[c] + '<select data-campo-mapa="' + c + '">' + opcoes + "</select></label>").join("");
    for (const [c, j] of Object.entries(actual)) { const s = document.querySelector('[data-campo-mapa="' + c + '"]'); if (s && j != null) s.value = String(j); }
    $("conc-mapa").hidden = false;
  }
  $("conc-mapear").addEventListener("click", abrirMapa);
  $("conc-mapa-fechar").addEventListener("click", () => { $("conc-mapa").hidden = true; });
  $("conc-mapa-aplicar").addEventListener("click", () => {
    const mapa = {};
    document.querySelectorAll("[data-campo-mapa]").forEach((s) => { if (s.value !== "") mapa[s.dataset.campoMapa] = Number(s.value); });
    if (mapa.data == null || (mapa.debito == null && mapa.credito == null && mapa.montante == null)) { progresso("Indique pelo menos a Data e o Débito/Crédito ou o Montante.", "conc-erro"); return; }
    const linhaCabecalho = Math.max(0, Number($("conc-mapa-linha").value || 1) - 1);
    const r = C.matrizParaMovimentos(estado.leitura.matriz, {mapa, linhaCabecalho});
    estado.leitura.resultado = r;
    aplicarResultado(r);
    try { const todos = mapas(); todos[chaveConta()] = {mapa, linhaCabecalho}; localStorage.setItem(MAPAS_KEY, JSON.stringify(todos)); } catch {}
    $("conc-mapa").hidden = true;
    progresso(r.movimentos.length + " movimento(s) com as colunas indicadas. Colunas memorizadas para esta conta.", r.movimentos.length ? "conc-ok" : "conc-erro");
  });

  // ---------- Conciliação ----------
  function invalidar() {
    if (estado.resultado) { estado.desactualizado = true; $("conc-executar").textContent = "Conciliar novamente"; desenharResultado(); }
  }
  function executar() {
    const conta = contaSeleccionada(), tol = Math.max(0, Math.min(30, Number($("conc-tol").value) || 0));
    const despesas = C.pagamentosParaConciliar(ctx().registos(), {idConta: conta.id, conta: conta.numero, banco: conta.banco, de: $("conc-de").value, ate: $("conc-ate").value, toleranciaDias: tol});
    // Entradas do extracto ↔ recibos da base mãe (Recebimentos); os recibos já ligados a um crédito ficam de fora.
    const BD = window.TES_BASE_DADOS, BASE = window.TES_BASE;
    let recibos = [];
    if (BD && BASE) { const d = BD.dados(); recibos = BASE.recibosParaBanco(d.recibos, {de: $("conc-de").value, ate: $("conc-ate").value, toleranciaDias: tol, ligados: d.ligacoes.map((l) => l.ID_Recibo)}); }
    // Já conciliados na base mãe: o movimento e os documentos ficam de fora (cada um só é conciliado uma vez)
    const gravadas = conciliacoesBase(), movConc = new Set(gravadas.map((l) => String(l.ID_Movimento)));
    const docConc = new Set(gravadas.map((l) => String(l.Tipo_Documento) + "|" + String(l.ID_Documento)));
    estado.gravados = estado.movimentos.filter((m) => m.idBase && movConc.has(m.idBase)).map((m) => m.id);
    const despesasLivres = despesas.filter((p) => !(p.idDocumento && docConc.has("despesa|" + p.idDocumento)));
    const recibosLivres = recibos.filter((p) => !docConc.has("recibo|" + p.idRecibo));
    estado.docsJaConciliados = despesas.length - despesasLivres.length + recibos.length - recibosLivres.length;
    estado.pagamentos = [...despesasLivres, ...recibosLivres];
    // Decisões manuais mantêm-se: confirmados saem do cálculo; rejeitados não voltam a ser sugeridos.
    const idsMov = new Set(estado.movimentos.map((m) => m.id)), idsPag = new Set(estado.pagamentos.map((p) => p.id));
    estado.confirmados = estado.confirmados.filter((p) => idsMov.has(p.mov) && p.pags.every((id) => idsPag.has(id)));
    const fixosM = new Set([...estado.confirmados.map((p) => p.mov), ...estado.gravados]), fixosP = new Set(estado.confirmados.flatMap((p) => p.pags));
    const movLivres = estado.movimentos.filter((m) => !fixosM.has(m.id)), pagLivres = estado.pagamentos.filter((p) => !fixosP.has(p.id));
    const r1 = C.conciliar(movLivres, pagLivres.filter((p) => p.tipo !== "recibo"), {toleranciaDias: tol});
    const r2 = C.conciliar(movLivres, pagLivres.filter((p) => p.tipo === "recibo"), {toleranciaDias: tol, sentido: "credito"});
    const pares = [...r1.pares, ...r2.pares], comPar = new Set(pares.map((p) => p.mov)), pagComPar = new Set(pares.flatMap((p) => p.pags));
    const r = {pares, bancoSemPar: movLivres.filter((m) => !comPar.has(m.id)).map((m) => m.id), tesourariaSemPar: pagLivres.filter((p) => !pagComPar.has(p.id)).map((p) => p.id)};
    const aceites = r.pares.filter((p) => !p.pags.some((id) => estado.rejeitados.has(p.mov + "|" + id)));
    const recusados = r.pares.filter((p) => !aceites.includes(p));
    estado.resultado = {
      pares: [...estado.confirmados, ...aceites],
      bancoSemPar: [...r.bancoSemPar, ...recusados.map((p) => p.mov)],
      tesourariaSemPar: [...r.tesourariaSemPar, ...recusados.flatMap((p) => p.pags)],
    };
    estado.desactualizado = false;
    $("conc-executar").textContent = "Conciliar";
    $("conc-fonte-pag").textContent = despesasLivres.length + " saída(s) da Tesouraria e " + recibosLivres.length + " recibo(s) por ligar no período" + (conta.texto ? " · " + conta.texto : conta.numero ? " · conta " + conta.numero : " · todas as contas") + "." +
      (estado.gravados.length || estado.docsJaConciliados ? " Já conciliados na base mãe: " + estado.gravados.length + " movimento(s), " + estado.docsJaConciliados + " documento(s)." : "");
    desenharResultado();
  }
  $("conc-executar").addEventListener("click", executar);

  const porMov = () => new Map(estado.movimentos.map((m) => [m.id, m]));
  const porPag = () => new Map(estado.pagamentos.map((p) => [p.id, p]));
  function desenharResultado() {
    const res = estado.resultado;
    if (!res) { $("conc-resumo").innerHTML = ""; $("conc-res-head").innerHTML = ""; $("conc-res-body").innerHTML = '<tr><td class="empty">Clique em «Conciliar» para comparar o extracto com as saídas da Tesouraria.</td></tr>'; return; }
    const gravados = new Set(estado.gravados), disponiveis = estado.movimentos.filter((m) => !gravados.has(m.id));
    const s = C.resumo(disponiveis, estado.pagamentos, res);
    $("conc-resumo").innerHTML = [
      ["Conciliado", s.conciliados + " de " + s.movimentos + " movimentos", s.percentagem + "% · " + mt(s.valorConciliado) + " MT"],
      ["Por confirmar", String(s.sugeridos), "sugestões a rever"],
      ["Banco sem correspondência", String(s.bancoSemPar), mt(s.valorBancoSemPar) + " MT"],
      ["Tesouraria sem correspondência", String(s.tesourariaSemPar), mt(s.valorTesourariaSemPar) + " MT sem movimento"],
    ].map(([t, v, d]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(d) + "</span></div>").join("") +
      (estado.desactualizado ? '<p class="conc-aviso">Os movimentos ou a conta mudaram. Clique em «Conciliar novamente».</p>' : "");
    document.querySelectorAll("[data-conc-tab]").forEach((b) => {
      b.classList.toggle("active", b.dataset.concTab === estado.tab);
      const n = {conciliados: s.conciliados, sugeridos: s.sugeridos, banco: s.bancoSemPar, tesouraria: s.tesourariaSemPar, gravados: estado.gravados.length}[b.dataset.concTab];
      b.dataset.n = n;
    });
    const M = porMov(), P = porPag();
    const descPags = (ids) => ids.map((id) => P.get(id)).filter(Boolean);
    const colPar = '<tr><th>Data banco</th><th>Descritivo banco</th><th class="num">Valor</th><th>Requisição / recibo</th><th>Fornecedor / cliente</th><th>Despacho / Ref.</th><th>Data</th><th>Motivos</th><th></th></tr>';
    const linhaPar = (p, botoes) => {
      const m = M.get(p.mov), ps = descPags(p.pags);
      return '<tr data-mov="' + esc(p.mov) + '"><td>' + esc(dataPt(m?.data)) + '</td><td class="conc-wrap">' + esc(m?.descricao) + '</td><td class="num">' + esc(mt(Math.abs(m?.valor || 0))) + "</td>" +
        "<td>" + ps.map((x) => esc(x.requisicao)).join("<br>") + '</td><td class="conc-wrap">' + ps.map((x) => esc(x.fornecedor)).join("<br>") + "</td>" +
        "<td>" + ps.map((x) => esc([x.despacho, x.referencia].filter(Boolean).join(" / ") || "—")).join("<br>") + "</td><td>" + ps.map((x) => esc(dataPt(x.data || x.dataPrevista))).join("<br>") + "</td>" +
        '<td><div class="conc-motivos"><span class="conc-pontos" title="Pontuação (0–100)">' + esc(p.pontos) + "</span>" + p.motivos.map((t) => '<span class="conc-chip">' + esc(curto(t)) + "</span>").join("") + (p.ambiguo ? '<span class="conc-chip amb">outro candidato igual</span>' : "") + "</div></td>" +
        '<td class="conc-acoes">' + botoes + "</td></tr>";
    };
    let head = "", body = "";
    if (estado.tab === "conciliados") {
      head = colPar;
      body = res.pares.filter((p) => p.estado !== "sugerido").map((p) => linhaPar(p, '<button type="button" class="text-btn" data-res="desfazer">Desfazer</button>')).join("");
    } else if (estado.tab === "sugeridos") {
      head = colPar;
      body = res.pares.filter((p) => p.estado === "sugerido").map((p) => linhaPar(p, '<button type="button" class="btn btn-primary conc-mini" data-res="confirmar">Confirmar</button><button type="button" class="text-btn" data-res="rejeitar">Rejeitar</button>')).join("");
    } else if (estado.tab === "banco") {
      const temRecibos = estado.pagamentos.some((p) => p.tipo === "recibo");
      head = '<tr><th>Data</th><th>Descritivo</th><th>Referência</th><th class="num">Débito</th><th class="num">Crédito</th><th></th></tr>';
      body = res.bancoSemPar.map((id) => M.get(id)).filter(Boolean).map((m) => '<tr data-mov="' + esc(m.id) + '"><td>' + esc(dataPt(m.data)) + '</td><td class="conc-wrap">' + esc(m.descricao) + "</td><td>" + esc(m.referencia) + '</td><td class="num">' + (m.valor < 0 ? esc(mt(-m.valor)) : "") + '</td><td class="num">' + (m.valor > 0 ? esc(mt(m.valor)) : "") + "</td><td>" +
        (m.valor < 0 || temRecibos ? '<button type="button" class="text-btn" data-res="associar">Associar…</button>' : '<span class="note">Sem recibos na base mãe</span>') + "</td></tr>").join("");
    } else if (estado.tab === "gravados") {
      const porMovBase = new Map();
      for (const l of conciliacoesBase()) porMovBase.set(String(l.ID_Movimento), (porMovBase.get(String(l.ID_Movimento)) || []).concat([l]));
      head = '<tr><th>Data banco</th><th>Descritivo banco</th><th class="num">Valor</th><th>Documento(s)</th><th>Tipo</th><th>Método</th><th></th></tr>';
      body = estado.gravados.map((id) => M.get(id)).filter(Boolean).map((m) => { const ls = porMovBase.get(m.idBase) || [];
        return '<tr data-mov="' + esc(m.id) + '"><td>' + esc(dataPt(m.data)) + '</td><td class="conc-wrap">' + esc(m.descricao) + '</td><td class="num">' + esc(mt(Math.abs(m.valor))) + "</td><td>" + ls.map((l) => esc(l.Documento || l.ID_Documento) + " (" + esc(mt(l.Valor)) + ")").join("<br>") + "</td><td>" + esc([...new Set(ls.map((l) => l.Tipo_Documento))].join(", ")) + "</td><td>" + esc([...new Set(ls.map((l) => l.Metodo))].join(", ")) + '</td><td class="conc-acoes"><button type="button" class="text-btn" data-res="desfazer-base" title="Remove esta conciliação da base mãe">Desfazer</button></td></tr>'; }).join("");
    } else {
      head = '<tr><th>Requisição / recibo</th><th>Fornecedor / cliente</th><th class="num">Valor</th><th>Estado</th><th>Data</th><th>Despacho</th><th>Referência</th><th>Conta</th></tr>';
      body = res.tesourariaSemPar.map((id) => P.get(id)).filter(Boolean).map((p) => "<tr><td>" + esc(p.requisicao) + '</td><td class="conc-wrap">' + esc(p.fornecedor) + '</td><td class="num">' + esc(mt(p.valor)) + "</td><td>" + esc(p.estado) + "</td><td>" + esc(dataPt(p.data || p.dataPrevista)) + "</td><td>" + esc(p.despacho) + "</td><td>" + esc(p.referencia) + "</td><td>" + esc([p.banco, p.conta].filter(Boolean).join(" · ")) + "</td></tr>").join("");
    }
    $("conc-res-head").innerHTML = head;
    $("conc-res-body").innerHTML = body || '<tr><td colspan="9" class="empty">Nada nesta lista.</td></tr>';
  }
  document.querySelectorAll("[data-conc-tab]").forEach((b) => b.addEventListener("click", () => { estado.tab = b.dataset.concTab; desenharResultado(); }));
  $("conc-res-body").addEventListener("click", (e) => {
    const b = e.target.closest("[data-res]"); if (!b || !estado.resultado) return;
    const mov = b.closest("tr").dataset.mov, res = estado.resultado;
    const par = res.pares.find((p) => p.mov === mov);
    if (b.dataset.res === "confirmar" && par) {
      par.estado = "confirmado"; estado.confirmados.push(par);
    } else if ((b.dataset.res === "rejeitar" || b.dataset.res === "desfazer") && par) {
      res.pares = res.pares.filter((p) => p !== par);
      estado.confirmados = estado.confirmados.filter((p) => p.mov !== mov);
      if (b.dataset.res === "rejeitar" || par.estado === "automatico") par.pags.forEach((id) => estado.rejeitados.add(mov + "|" + id));
      res.bancoSemPar.push(mov); res.tesourariaSemPar.push(...par.pags);
    } else if (b.dataset.res === "associar") { abrirAssociar(mov); return; }
    else if (b.dataset.res === "desfazer-base") { desfazerBase(mov, b); return; }
    desenharResultado();
  });

  // ---------- Associação manual (1 ou vários pagamentos com a mesma soma) ----------
  // Motivos em etiquetas curtas (o texto completo vai no relatório exportado)
  function curto(t) {
    return t.replace(" no descritivo", "").replace("referência do pagamento", "ref. pagamento").replace("nº de despacho", "despacho")
      .replace("nº de requisição", "requisição").replace(" dia(s) de diferença", " dia(s)").replace("soma de ", "soma ").replace(" pagamentos igual ao débito", " pagamentos");
  }
  let associarMov = null;
  function abrirAssociar(mov) {
    const m = porMov().get(mov); if (!m) return;
    associarMov = m;
    $("conc-assoc-mov").textContent = dataPt(m.data) + " · " + m.descricao + " · " + mt(Math.abs(m.valor)) + " MT" + (m.valor > 0 ? " (entrada: associar a recibos)" : "");
    $("conc-assoc-filtro").value = "";
    listarAssociar();
    $("conc-associar").showModal();
  }
  function listarAssociar() {
    const P = porPag(), alvo = Math.abs(associarMov.valor), f = $("conc-assoc-filtro").value.trim().toLowerCase();
    const marcados = new Set([...document.querySelectorAll("#conc-assoc-lista input:checked")].map((i) => i.value));
    const credito = associarMov.valor > 0;
    const livres = estado.resultado.tesourariaSemPar.map((id) => P.get(id)).filter((p) => p && (p.tipo === "recibo") === credito)
      .filter((p) => !f || [p.requisicao, p.fornecedor, p.despacho, p.referencia, mt(p.valor), String(p.valor)].join(" ").toLowerCase().includes(f))
      .sort((a, b) => Math.abs(a.valor - alvo) - Math.abs(b.valor - alvo));
    $("conc-assoc-lista").innerHTML = livres.map((p) => '<label class="conc-assoc-item"><input type="checkbox" value="' + esc(p.id) + '"' + (marcados.has(p.id) ? " checked" : "") + "><span><b>" + esc(mt(p.valor)) + " MT</b> · " + esc(p.requisicao) + " · " + esc(p.fornecedor) + "<br><small>" + esc(dataPt(p.data || p.dataPrevista)) + (p.despacho ? " · despacho " + esc(p.despacho) : "") + (p.referencia ? " · ref. " + esc(p.referencia) : "") + "</small></span></label>").join("") ||
      '<p class="note">Não há ' + (credito ? "recibos" : "pagamentos") + ' sem correspondência neste período. Alargue as datas ou a tolerância e concilie novamente.</p>';
    somaAssociar();
  }
  function somaAssociar() {
    const P = porPag(), ids = [...document.querySelectorAll("#conc-assoc-lista input:checked")].map((i) => i.value);
    const soma = ids.reduce((s, id) => s + (P.get(id)?.valor || 0), 0), alvo = Math.abs(associarMov.valor);
    const igual = Math.round(soma * 100) === Math.round(alvo * 100);
    $("conc-assoc-soma").textContent = ids.length ? "Seleccionado: " + mt(soma) + " MT de " + mt(alvo) + " MT" + (igual ? " ✓" : " — a soma tem de ser igual ao movimento") : "Seleccione um ou mais documentos cuja soma seja igual ao movimento.";
    $("conc-assoc-ok").disabled = !igual;
  }
  $("conc-assoc-lista").addEventListener("change", somaAssociar);
  $("conc-assoc-filtro").addEventListener("input", listarAssociar);
  $("conc-assoc-cancelar").addEventListener("click", () => $("conc-associar").close());
  $("conc-assoc-ok").addEventListener("click", () => {
    const ids = [...document.querySelectorAll("#conc-assoc-lista input:checked")].map((i) => i.value), res = estado.resultado;
    const par = {mov: associarMov.id, pags: ids, pontos: 100, motivos: ["associação manual"], tipo: ids.length > 1 ? "N:1" : "1:1", estado: "confirmado"};
    res.pares.push(par); estado.confirmados.push(par);
    res.bancoSemPar = res.bancoSemPar.filter((id) => id !== associarMov.id);
    res.tesourariaSemPar = res.tesourariaSemPar.filter((id) => !ids.includes(id));
    $("conc-associar").close();
    estado.tab = "conciliados"; desenharResultado();
  });

  // ---------- Exportação ----------
  function nomeFicheiro(ext) {
    const c = contaSeleccionada();
    return ("conciliacao-" + (c.numero || c.id || "conta") + "-" + ($("conc-de").value || "") + "_" + ($("conc-ate").value || "")).replace(/[^\w.\-]+/g, "_") + "." + ext;
  }
  function linhasRelatorio() {
    if (!estado.resultado) { progresso("Concilie antes de exportar.", "conc-erro"); return null; }
    if (estado.desactualizado) { progresso("Os movimentos mudaram desde a última conciliação: clique em «Conciliar novamente» antes de exportar.", "conc-erro"); return null; }
    return C.relatorio(estado.movimentos, estado.pagamentos, estado.resultado);
  }
  // Correspondências conciliadas (automáticas, confirmadas e manuais) → base mãe, pelo Power Automate (banco.conciliar).
  async function gravarConciliacao() {
    if (!estado.resultado || estado.desactualizado) { progresso("Concilie antes de gravar.", "conc-erro"); return; }
    if (!escrita()) { progresso("Para gravar, ligue a base de dados e indique o endpoint de escrita do Power Automate (V9).", "conc-erro"); return; }
    const M = porMov(), P = porPag(), naBase = idsBase(), conta = contaSeleccionada();
    const ligacoes = [], semId = [], foraBase = [];
    for (const p of estado.resultado.pares.filter((x) => x.estado !== "sugerido")) {
      const m = M.get(p.mov), ps = p.pags.map((id) => P.get(id));
      if (!m || !m.idBase || (consultaTemBase() || sessao.movimentos.length ? !naBase.has(m.idBase) : false)) { foraBase.push(p); continue; }
      if (ps.some((x) => !x || !(x.tipo === "recibo" ? x.idRecibo : x.idDocumento))) { semId.push(p); continue; }
      const metodo = p.motivos.includes("associação manual") ? "manual" : p.estado === "automatico" ? "automatico" : "confirmado";
      for (const x of ps) ligacoes.push({ID_Movimento: m.idBase, Tipo_Documento: x.tipo === "recibo" ? "recibo" : "despesa", ID_Documento: x.tipo === "recibo" ? x.idRecibo : x.idDocumento,
        Documento: String(x.requisicao || "").slice(0, 80), Valor: Math.round(x.valor * 100) / 100, Data_Banco: m.data, Valor_Banco: Math.round(m.valor * 100) / 100,
        Descritivo_Banco: String(m.descricao || "").slice(0, 250), Conta: conta.numero || "", Pontos: p.pontos, Motivos: p.motivos.join("; ").slice(0, 250), Metodo: metodo});
    }
    if (!ligacoes.length) {
      progresso(foraBase.length ? "Grave primeiro o extracto na base mãe (passo 2): as conciliações ligam-se aos movimentos gravados." : semId.length ? "As despesas conciliadas não têm ID_Despesa nem Nr na consulta: não é possível gravá-las." : "Não há correspondências conciliadas para gravar (confirme as sugestões primeiro).", "conc-erro");
      return;
    }
    const b = $("conc-gravar"); b.disabled = true;
    let gravadas = 0;
    try {
      for (let i = 0; i < ligacoes.length; i += LOTE_LIG) {
        const lote = ligacoes.slice(i, i + LOTE_LIG);
        progresso("A gravar a conciliação no Power Automate… " + (i + lote.length) + " de " + ligacoes.length + ".", "conc-a-ler");
        const r = await ctx().escrever({operacao: "banco.conciliar", id_pedido: crypto.randomUUID(), ligacoes: lote});
        gravadas += Number(r.ligacoes) || 0;
        sessao.conciliacoes.push(...lote.map((l) => ({...l, ID_Conciliacao: l.Tipo_Documento + "|" + l.ID_Documento})));
        window.TES_BASE_DADOS?.registarLigacoes(lote.filter((l) => l.Tipo_Documento === "recibo").map((l) => ({ID_Ligacao: l.ID_Documento + "@" + l.ID_Movimento, ID_Recibo: l.ID_Documento, Data_Banco: l.Data_Banco, Valor: l.Valor, Descritivo_Banco: l.Descritivo_Banco, Conta: l.Conta})));
      }
      estado.confirmados = []; executar(); estado.tab = "gravados"; desenharResultado(); avisarBancos();
      progresso(gravadas + " ligação(ões) gravada(s) na base mãe." + (foraBase.length ? " " + foraBase.length + " correspondência(s) de movimentos ainda não gravados ficaram de fora." : "") + (semId.length ? " " + semId.length + " sem identificação da despesa ficaram de fora." : ""), "conc-ok");
    } catch (e) { progresso(e.message + (gravadas ? " (Já gravadas: " + gravadas + "; repetir é seguro.)" : ""), "conc-erro"); }
    finally { b.disabled = false; }
  }
  $("conc-gravar").addEventListener("click", gravarConciliacao);
  async function desfazerBase(mov, botao) {
    const m = porMov().get(mov);
    if (!m?.idBase || !escrita()) { progresso("Para desfazer na base mãe é preciso o endpoint de escrita do Power Automate.", "conc-erro"); return; }
    if (!confirm("Remover da base mãe a conciliação deste movimento (" + dataPt(m.data) + ", " + mt(Math.abs(m.valor)) + " MT)?")) return;
    botao.disabled = true;
    try {
      const r = await ctx().escrever({operacao: "banco.desfazer", id_pedido: crypto.randomUUID(), movimentos: [m.idBase]});
      sessao.desfeitos.add(m.idBase); sessao.conciliacoes = sessao.conciliacoes.filter((l) => l.ID_Movimento !== m.idBase);
      executar(); avisarBancos();
      progresso(String(r.mensagem || "Conciliação removida.") + " O movimento voltou a estar por conciliar.", "conc-ok");
    } catch (e) { progresso(e.message, "conc-erro"); botao.disabled = false; }
  }
  $("conc-csv").addEventListener("click", () => {
    const linhas = linhasRelatorio(); if (!linhas) return;
    const url = URL.createObjectURL(new Blob([C.csv(linhas)], {type: "text/csv;charset=utf-8"}));
    const a = document.createElement("a"); a.href = url; a.download = nomeFicheiro("csv"); a.click(); URL.revokeObjectURL(url);
  });
  $("conc-xlsx").addEventListener("click", async () => {
    const linhas = linhasRelatorio(); if (!linhas) return;
    try { await IMP.exportarExcel(linhas, nomeFicheiro("xlsx")); } catch (e) { progresso(e.message, "conc-erro"); }
  });

  // ---------- Arranque ----------
  function modo() { $("conc-modo").textContent = escrita() ? "Grava na base mãe pelo Power Automate" : ctx().ligado() ? "Consulta ligada · sem endpoint de escrita" : "Modo demonstração · não grava"; }
  document.addEventListener("tes:contexto", () => { preencherContas(); modo(); desenharBase(); invalidar(); });
  modo();
  document.querySelector('[data-view="conciliacao"]')?.addEventListener("click", preencherContas);
  preencherContas();
  desenharResultado();
})();
