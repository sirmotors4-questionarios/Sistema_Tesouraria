/* Conciliação bancária — interface. Lê as despesas/contas carregadas pelo app.js (TES_CONTEXTO, só leitura),
 * importa o extracto (TES_IMPORTAR), emparelha (TES_CONCILIACAO) e exporta o relatório. Não grava no Excel.
 * O extracto fica apenas na memória desta página; no navegador só se guarda o mapeamento de colunas por conta.
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
  const MAPAS_KEY = "sir.tesouraria.conciliacao.mapas.v1";

  const estado = {leitura: null, movimentos: [], saldoInicial: null, saldoFinal: null, pagamentos: [], resultado: null,
    tab: "conciliados", confirmados: [], rejeitados: new Set(), desactualizado: false};

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
    if (c.id) $("conc-conta-num").value = c.numero;
    invalidar();
  });
  $("conc-conta-num").addEventListener("input", invalidar);

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
      estado.leitura = leitura;
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
    Object.assign(estado, {leitura: null, movimentos: [], saldoInicial: null, saldoFinal: null, resultado: null, confirmados: [], rejeitados: new Set(), tab: "conciliados"});
    $("conc-passo2").hidden = true; $("conc-passo3").hidden = true; $("conc-mapa").hidden = true;
    if (limparProgresso) progresso("");
  }

  // ---------- Movimentos (revisão) ----------
  function desenharMovimentos() {
    const v = C.validar(estado.movimentos, {saldoInicial: estado.saldoInicial, saldoFinal: estado.saldoFinal});
    const quebra = new Set(v.quebras), dup = new Set(v.duplicados);
    $("conc-mov-body").innerHTML = estado.movimentos.map((m) =>
      '<tr data-id="' + esc(m.id) + '" class="' + (quebra.has(m.id) ? "conc-quebra " : "") + (dup.has(m.id) ? "conc-dup" : "") + '">' +
      '<td><input type="date" value="' + esc(m.data) + '" data-campo="data"></td>' +
      '<td><input value="' + esc(m.descricao) + '" data-campo="descricao" class="conc-desc"></td>' +
      '<td><input value="' + esc(m.referencia) + '" data-campo="referencia" class="conc-ref"></td>' +
      '<td class="num"><input value="' + (m.valor < 0 ? esc(mt(-m.valor)) : "") + '" data-campo="debito" class="conc-valor" inputmode="decimal"></td>' +
      '<td class="num"><input value="' + (m.valor > 0 ? esc(mt(m.valor)) : "") + '" data-campo="credito" class="conc-valor" inputmode="decimal"></td>' +
      '<td class="num">' + esc(mt(m.saldo)) + '</td>' +
      '<td class="conc-acoes"><button type="button" class="icon-btn" data-accao="sinal" title="Trocar débito/crédito">±</button><button type="button" class="icon-btn" data-accao="remover" title="Remover esta linha">✕</button></td></tr>').join("") ||
      '<tr><td colspan="7" class="empty">Sem movimentos. Use «Indicar colunas» se o ficheiro tiver outro formato.</td></tr>';
    const avisos = [...(estado.avisos || []), ...v.problemas];
    $("conc-validacao").innerHTML =
      '<div class="conc-saldos"><span>Saldo inicial <b>' + (v.saldoInicial == null ? "—" : esc(mt(v.saldoInicial))) + '</b></span><span>Débitos <b>' + esc(mt(-v.debitos)) + '</b></span><span>Créditos <b>' + esc(mt(v.creditos)) + '</b></span><span>Saldo final <b>' + (v.saldoFinal == null ? "—" : esc(mt(v.saldoFinal))) + '</b></span>' +
      '<span class="tag ' + (v.ok ? "conc-tag-ok" : "waiting") + '">' + (v.ok ? "Saldos conferidos" : "Rever antes de conciliar") + "</span></div>" +
      (avisos.length ? "<ul>" + avisos.map((a) => "<li>" + esc(a) + "</li>").join("") + "</ul>" : "");
  }
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
    estado.pagamentos = C.pagamentosParaConciliar(ctx().registos(), {idConta: conta.id, conta: conta.numero, banco: conta.banco, de: $("conc-de").value, ate: $("conc-ate").value, toleranciaDias: tol});
    // Decisões manuais mantêm-se: confirmados saem do cálculo; rejeitados não voltam a ser sugeridos.
    const idsMov = new Set(estado.movimentos.map((m) => m.id)), idsPag = new Set(estado.pagamentos.map((p) => p.id));
    estado.confirmados = estado.confirmados.filter((p) => idsMov.has(p.mov) && p.pags.every((id) => idsPag.has(id)));
    const fixosM = new Set(estado.confirmados.map((p) => p.mov)), fixosP = new Set(estado.confirmados.flatMap((p) => p.pags));
    const r = C.conciliar(estado.movimentos.filter((m) => !fixosM.has(m.id)), estado.pagamentos.filter((p) => !fixosP.has(p.id)), {toleranciaDias: tol});
    const aceites = r.pares.filter((p) => !p.pags.some((id) => estado.rejeitados.has(p.mov + "|" + id)));
    const recusados = r.pares.filter((p) => !aceites.includes(p));
    estado.resultado = {
      pares: [...estado.confirmados, ...aceites],
      bancoSemPar: [...r.bancoSemPar, ...recusados.map((p) => p.mov)],
      tesourariaSemPar: [...r.tesourariaSemPar, ...recusados.flatMap((p) => p.pags)],
    };
    estado.desactualizado = false;
    $("conc-executar").textContent = "Conciliar";
    $("conc-fonte-pag").textContent = estado.pagamentos.length + " saída(s) da Tesouraria no período" + (conta.texto ? " · " + conta.texto : conta.numero ? " · conta " + conta.numero : " · todas as contas") + ".";
    desenharResultado();
  }
  $("conc-executar").addEventListener("click", executar);

  const porMov = () => new Map(estado.movimentos.map((m) => [m.id, m]));
  const porPag = () => new Map(estado.pagamentos.map((p) => [p.id, p]));
  function desenharResultado() {
    const res = estado.resultado;
    if (!res) { $("conc-resumo").innerHTML = ""; $("conc-res-head").innerHTML = ""; $("conc-res-body").innerHTML = '<tr><td class="empty">Clique em «Conciliar» para comparar o extracto com as saídas da Tesouraria.</td></tr>'; return; }
    const s = C.resumo(estado.movimentos, estado.pagamentos, res);
    $("conc-resumo").innerHTML = [
      ["Conciliado", s.conciliados + " de " + s.debitos + " débitos", s.percentagem + "% · " + mt(s.valorConciliado) + " MT"],
      ["Por confirmar", String(s.sugeridos), "sugestões a rever"],
      ["Banco sem correspondência", String(s.bancoSemPar), mt(s.valorBancoSemPar) + " MT em débitos"],
      ["Tesouraria sem correspondência", String(s.tesourariaSemPar), mt(s.valorTesourariaSemPar) + " MT pagos sem movimento"],
    ].map(([t, v, d]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(d) + "</span></div>").join("") +
      (estado.desactualizado ? '<p class="conc-aviso">Os movimentos ou a conta mudaram. Clique em «Conciliar novamente».</p>' : "");
    document.querySelectorAll("[data-conc-tab]").forEach((b) => {
      b.classList.toggle("active", b.dataset.concTab === estado.tab);
      const n = {conciliados: s.conciliados, sugeridos: s.sugeridos, banco: s.bancoSemPar, tesouraria: s.tesourariaSemPar}[b.dataset.concTab];
      b.dataset.n = n;
    });
    const M = porMov(), P = porPag();
    const descPags = (ids) => ids.map((id) => P.get(id)).filter(Boolean);
    const colPar = '<tr><th>Data banco</th><th>Descritivo banco</th><th class="num">Valor</th><th>Requisição</th><th>Fornecedor</th><th>Despacho / Ref.</th><th>Saída</th><th>Motivos</th><th></th></tr>';
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
      head = '<tr><th>Data</th><th>Descritivo</th><th>Referência</th><th class="num">Débito</th><th class="num">Crédito</th><th></th></tr>';
      body = res.bancoSemPar.map((id) => M.get(id)).filter(Boolean).map((m) => '<tr data-mov="' + esc(m.id) + '"><td>' + esc(dataPt(m.data)) + '</td><td class="conc-wrap">' + esc(m.descricao) + "</td><td>" + esc(m.referencia) + '</td><td class="num">' + (m.valor < 0 ? esc(mt(-m.valor)) : "") + '</td><td class="num">' + (m.valor > 0 ? esc(mt(m.valor)) : "") + "</td><td>" +
        (m.valor < 0 ? '<button type="button" class="text-btn" data-res="associar">Associar…</button>' : '<span class="note">Recebimentos ainda não integrados</span>') + "</td></tr>").join("");
    } else {
      head = '<tr><th>Requisição</th><th>Fornecedor</th><th class="num">Valor</th><th>Estado</th><th>Saída</th><th>Despacho</th><th>Ref. pagamento</th><th>Conta</th></tr>';
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
    $("conc-assoc-mov").textContent = dataPt(m.data) + " · " + m.descricao + " · " + mt(-m.valor) + " MT";
    $("conc-assoc-filtro").value = "";
    listarAssociar();
    $("conc-associar").showModal();
  }
  function listarAssociar() {
    const P = porPag(), alvo = Math.abs(associarMov.valor), f = $("conc-assoc-filtro").value.trim().toLowerCase();
    const marcados = new Set([...document.querySelectorAll("#conc-assoc-lista input:checked")].map((i) => i.value));
    const livres = estado.resultado.tesourariaSemPar.map((id) => P.get(id)).filter(Boolean)
      .filter((p) => !f || [p.requisicao, p.fornecedor, p.despacho, p.referencia, mt(p.valor), String(p.valor)].join(" ").toLowerCase().includes(f))
      .sort((a, b) => Math.abs(a.valor - alvo) - Math.abs(b.valor - alvo));
    $("conc-assoc-lista").innerHTML = livres.map((p) => '<label class="conc-assoc-item"><input type="checkbox" value="' + esc(p.id) + '"' + (marcados.has(p.id) ? " checked" : "") + "><span><b>" + esc(mt(p.valor)) + " MT</b> · " + esc(p.requisicao) + " · " + esc(p.fornecedor) + "<br><small>" + esc(dataPt(p.data || p.dataPrevista)) + (p.despacho ? " · despacho " + esc(p.despacho) : "") + (p.referencia ? " · ref. " + esc(p.referencia) : "") + "</small></span></label>").join("") ||
      '<p class="note">Não há pagamentos sem correspondência neste período. Alargue as datas ou a tolerância e concilie novamente.</p>';
    somaAssociar();
  }
  function somaAssociar() {
    const P = porPag(), ids = [...document.querySelectorAll("#conc-assoc-lista input:checked")].map((i) => i.value);
    const soma = ids.reduce((s, id) => s + (P.get(id)?.valor || 0), 0), alvo = Math.abs(associarMov.valor);
    const igual = Math.round(soma * 100) === Math.round(alvo * 100);
    $("conc-assoc-soma").textContent = ids.length ? "Seleccionado: " + mt(soma) + " MT de " + mt(alvo) + " MT" + (igual ? " ✓" : " — a soma tem de ser igual ao débito") : "Seleccione um ou mais pagamentos cuja soma seja igual ao débito.";
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
  document.addEventListener("tes:contexto", () => { preencherContas(); invalidar(); });
  document.querySelector('[data-view="conciliacao"]')?.addEventListener("click", preencherContas);
  preencherContas();
  desenharResultado();
})();
