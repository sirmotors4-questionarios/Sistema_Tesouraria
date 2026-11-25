import {readFileSync} from "node:fs";
import {stripTypeScriptTypes} from "node:module";
import vm from "node:vm";

// Executa o Office Script real sobre um livro em memória (nenhuma chamada ao Excel).
const fonte = stripTypeScriptTypes(readFileSync(new URL("../office-scripts/TesourariaBaseMae.ts", import.meta.url), "utf8"));
const ctx = vm.createContext({});
vm.runInContext(fonte, ctx);
// Chamada directa com o texto bruto do parâmetro «pedido» (para testar o diagnóstico)
globalThis.__mainBaseMae = (w, texto) => ctx.main(w, texto);

export function livro() {
  const tabelas = new Map(), folhas = new Map();
  const cont = {leituras: 0, escritas: 0};
  function novaFolha(nome) {
    const f = {nome, cab: null, tabela: null, formatos: {}, abaixo: []};
    f.getRangeByIndexes = (r, c, nr, nc) => intervalo(f, r, c, nr, nc);
    // Células soltas por endereço (ex.: reserva em BM_Controlo!A2:C2)
    f.celulas = {};
    f.getRange = (end) => ({
      getValues: () => { cont.leituras++; return [f.celulas[end] ? f.celulas[end].slice() : ["", "", ""]]; },
      setValues: (vs) => { cont.escritas++; f.celulas[end] = vs[0].slice(); },
    });
    f.addTable = () => {
      const t = {headers: f.cab.slice(), rows: [f.cab.map(() => "")], folha: f};
      Object.assign(t, {
        getHeaderRowRange: () => ({getValues: () => { cont.leituras++; return [t.headers]; }}),
        getRangeBetweenHeaderAndTotal: () => intervalo(f, 1, 0, t.rows.length, t.headers.length),
        getRowCount: () => { cont.leituras++; return t.rows.length; },
        addRows: (i, vs) => { cont.escritas++; t.rows.push(...vs.map((l) => [...l])); },
        // Como no Excel: a tabela fica sempre com pelo menos uma linha (vazia) no corpo.
        deleteRowsAt: (i, n = 1) => { cont.escritas++; t.rows.splice(i, n); if (!t.rows.length) t.rows.push(t.headers.map(() => "")); },
        getRange: () => intervalo(f, 0, 0, t.rows.length + 1, t.headers.length),
        getWorksheet: () => f,
        // resize(intervalo cabeçalho+corpo): o corpo passa a ter nr-1 linhas (as células já escritas abaixo entram na tabela)
        resize: (rg) => { cont.escritas++; const nr = rg.nr - 1; while (t.rows.length < nr) t.rows.push(t.headers.map(() => "")); t.rows.length = nr; t.rows.forEach((l, k) => { if (f.abaixo[k]) { f.abaixo[k].forEach((v, j) => { if (v !== undefined) l[j] = v; }); } }); f.abaixo = []; },
        setName: (n) => tabelas.set(n, t),
      });
      f.tabela = t;
      return t;
    };
    folhas.set(nome, f);
    return f;
  }
  // Coordenadas absolutas: linha 0 = cabeçalho da tabela da folha; linhas 1.. = corpo.
  const letra = (j) => { let t = ""; j++; while (j > 0) { const m = (j - 1) % 26; t = String.fromCharCode(65 + m) + t; j = Math.floor((j - 1) / 26); } return t; };
  function intervalo(f, r, c, nr, nc) {
    const corpo = () => f.tabela.rows;
    // getRange() da tabela (r=0) devolve cabeçalho + corpo
    const linhas = () => (r === 0 && f.tabela ? [f.tabela.headers, ...corpo()] : corpo().slice(r - 1)).slice(0, nr);
    return {
      nr,
      getValues: () => { cont.leituras++; return linhas().map((l) => l.slice(c, c + nc)); },
      getAddress: () => { cont.leituras++; return f.nome + "!" + letra(c) + (r + 1) + ":" + letra(c + nc - 1) + (r + nr); },
      setNumberFormats: (fs) => { cont.escritas++; fs.forEach((l, i) => l.forEach((x, j) => { f.formatos[(r + i) + ":" + (c + j)] = x; })); },
      setValues: (vs) => {
        cont.escritas++;
        if (r === 0) { f.cab = vs[0].slice(); return; }
        // Linhas abaixo do corpo ficam "fora da tabela" até ao resize
        vs.forEach((l, i) => l.forEach((v, j) => {
          const k = r - 1 + i;
          if (k < corpo().length) corpo()[k][c + j] = v;
          else { (f.abaixo[k] = f.abaixo[k] || [])[c + j] = v; }
        }));
      },
      setNumberFormat: (fmt) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) f.formatos[(r + i) + ":" + (c + j)] = fmt; },
      getColumn: (j) => intervalo(f, r, c + j, nr, 1),
      getRow: (i) => intervalo(f, r + i, c, 1, nc),
      getWorksheet: () => f, getRowIndex: () => r, getColumnIndex: () => c,
    };
  }
  return {
    tabelas, folhas, cont,
    getTable: (n) => tabelas.get(n),
    getWorksheet: (n) => folhas.get(n),
    addWorksheet: (n) => novaFolha(n),
    ler: (n) => { const t = tabelas.get(n); return t.rows.filter((l) => l.some((v) => v !== "" && v != null)).map((l) => Object.fromEntries(t.headers.map((h, j) => [h, l[j]]))); },
  };
}
export const chamar = (w, pedido) => JSON.parse(JSON.stringify(ctx.main(w, JSON.stringify(pedido))));
