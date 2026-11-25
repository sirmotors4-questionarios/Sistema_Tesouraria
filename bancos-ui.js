/* Bancos e Extractos — vista da base mãe bancária (só leitura): contas com movimentos, registo de extractos importados.
 * Lê as tabelas devolvidas pela consulta do Power Automate (tbl_Extratos, tbl_MovimentosBancarios, tbl_ConciliacaoBancaria)
 * mais o que foi gravado nesta sessão (window.TES_BANCOS, definido em conciliacao-ui.js). A importação faz-se em Conciliação.
 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  if (!$("bancos")) return;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"})[c]);
  const C = window.TES_CONCILIACAO;
  const num = (v) => { const n = C ? C.numero(v) : Number(v); return Number.isFinite(n) ? n : 0; };
  const mt = (n) => new Intl.NumberFormat("pt-PT", {minimumFractionDigits: 2, maximumFractionDigits: 2}).format(Number(n) || 0);
  const iso = (v) => {
    if (typeof v === "number" && v > 1000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v * 86400000)).toISOString().slice(0, v % 1 ? 16 : 10).replace("T", " ");
    const t = String(v ?? "").trim(); return t.slice(0, 16).replace("T", " ");
  };
  const dataPt = (v) => { const t = iso(v); return /^\d{4}-\d{2}-\d{2}/.test(t) ? t.slice(8, 10) + "/" + t.slice(5, 7) + "/" + t.slice(0, 4) + t.slice(10) : t || "—"; };
  const digitos = (v) => String(v ?? "").replace(/\D/g, "").replace(/^0+/, "");

  function desenhar() {
    const B = window.TES_BANCOS, ctx = window.TES_CONTEXTO;
    if (!B || !ctx) return;
    const extratos = B.extratos(), movs = B.movimentos(), conc = B.conciliacoes();
    const contas = ctx.contas();
    const conciliados = new Set(conc.map((l) => String(l.ID_Movimento)));
    const grupos = new Map();
    for (const m of movs) {
      const k = String(m.Banco ?? "").trim().toUpperCase().split(/\s+/)[0] + "|" + digitos(m.Conta);
      const g = grupos.get(k) || {banco: m.Banco, conta: m.Conta, idConta: m.ID_Conta, n: 0, conc: 0, primeiro: "", ultimo: "", saldo: null};
      const d = iso(m.Data_Movimento);
      g.n++; if (conciliados.has(String(m.ID_Movimento))) g.conc++;
      if (!g.primeiro || d < g.primeiro) g.primeiro = d;
      if (!g.ultimo || d >= g.ultimo) { g.ultimo = d; if (m.Saldo !== "" && m.Saldo != null) g.saldo = num(m.Saldo); }
      grupos.set(k, g);
    }
    const porConciliar = [...grupos.values()].reduce((s, g) => s + g.n - g.conc, 0);
    $("bk-kpis").innerHTML = [
      ["Extractos importados", extratos.length, "ficheiros registados (cada um uma só vez)"],
      ["Movimentos na base mãe", movs.length, "sem duplicados"],
      ["Por conciliar", porConciliar, "movimentos sem despesa/recibo ligado"],
      ["Contas com extracto", grupos.size + " de " + contas.length, "contas do cadastro"],
    ].map(([t, v, d]) => '<div class="metric"><small>' + esc(t) + "</small><strong>" + esc(v) + "</strong><span>" + esc(d) + "</span></div>").join("");
    const designacao = (g) => {
      const c = contas.find((x) => (g.idConta && String(x.ID_Conta) === String(g.idConta)) || (digitos(x.Numero_Conta) === digitos(g.conta) && String(x.Banco ?? "").toUpperCase().split(/\s+/)[0] === String(g.banco ?? "").toUpperCase().split(/\s+/)[0]));
      return c ? String(c.Designacao_Conta ?? "") : "não cadastrada";
    };
    $("bk-contas").innerHTML = [...grupos.values()].sort((a, b) => String(a.banco).localeCompare(String(b.banco)) || String(a.conta).localeCompare(String(b.conta))).map((g) =>
      "<tr><td>" + esc(g.banco) + "</td><td>" + esc(g.conta) + "</td><td>" + esc(designacao(g)) + "</td><td>" + esc(dataPt(g.primeiro)) + "</td><td>" + esc(dataPt(g.ultimo)) +
      '</td><td class="num">' + g.n + '</td><td class="num">' + (g.saldo == null ? "—" : esc(mt(g.saldo))) + '</td><td class="num">' + g.conc + '</td><td class="num">' + (g.n - g.conc) + "</td></tr>").join("") ||
      '<tr><td colspan="9" class="empty">' + (ctx.ligado() ? "Sem movimentos na base mãe. Use «Importar extracto»." : "Ligue a base de dados para ver os extractos importados.") + "</td></tr>";
    const f = $("bk-pesquisa").value.trim().toLowerCase();
    const lista = extratos.filter((e) => !f || [e.Nome_Padronizado, e.Nome_Original, e.Conta, e.Banco].join(" ").toLowerCase().includes(f))
      .sort((a, b) => String(iso(b.Data_Importacao)).localeCompare(String(iso(a.Data_Importacao))));
    $("bk-extratos").innerHTML = lista.map((e) => {
      const ok = ["Completo", "Sem movimentos novos"].includes(String(e.Estado));
      return '<tr><td><b>' + esc(e.Nome_Padronizado) + "</b></td><td>" + esc(e.Nome_Original) + "</td><td>" + esc([e.Banco, e.Conta].filter(Boolean).join(" · ")) + "</td><td>" + esc(dataPt(e.De)) + " – " + esc(dataPt(e.Ate)) + "</td><td>" + esc(e.Formato) +
        '</td><td class="num">' + esc(e.Movimentos_Ficheiro) + '</td><td class="num">' + esc(e.Novos) + '</td><td class="num">' + esc(e.Ja_Existentes) + '</td><td><span class="tag ' + (ok ? "conc-tag-ok" : "waiting") + '">' + esc(e.Estado) + "</span></td><td>" + esc(dataPt(e.Data_Importacao)) + "</td></tr>";
    }).join("") || '<tr><td colspan="10" class="empty">Nenhum extracto importado.</td></tr>';
    $("bk-nota").textContent = !ctx.ligado() ? "Modo demonstração: sem ligação à base mãe."
      : B.consultaTemBase() ? "Fonte: consulta do Power Automate (base mãe no SharePoint) e gravações desta sessão."
      : "A consulta do Power Automate não devolveu tbl_Extratos / tbl_MovimentosBancarios: só aparecem as gravações desta sessão. Ver docs/BANCOS_EXTRATOS.md (passo 3.4).";
    $("bk-modo").textContent = ctx.escritaDisponivel() ? "Base mãe · gravação activa" : ctx.ligado() ? "Base mãe · só leitura" : "Modo demonstração";
  }
  $("bk-pesquisa").addEventListener("input", desenhar);
  document.addEventListener("tes:contexto", desenhar);
  document.addEventListener("tes:bancos", desenhar);
  document.querySelector('[data-view="bancos"]')?.addEventListener("click", desenhar);
  desenhar();
})();
