import {readFileSync} from "node:fs";
import {stripTypeScriptTypes} from "node:module";
import vm from "node:vm";

// Executa o Office Script real sobre um livro em memória (nenhuma chamada ao Excel).
const fonte = stripTypeScriptTypes(readFileSync(new URL("../office-scripts/TesourariaBaseMae.ts", import.meta.url), "utf8"));
const ctx = vm.createContext({});
vm.runInContext(fonte, ctx);

export function livro() {
  const tabelas = new Map(), folhas = new Map();
  const cont = {leituras: 0, escritas: 0};
  function novaFolha(nome) {
    const f = {nome, cab: null, tabela: null, formatos: {}};
    f.getRangeByIndexes = (r, c, nr, nc) => intervalo(f, r, c, nr, nc);
    f.addTable = () => {
      const t = {headers: f.cab.slice(), rows: [f.cab.map(() => "")], folha: f};
      Object.assign(t, {
        getHeaderRowRange: () => ({getValues: () => { cont.leituras++; return [t.headers]; }}),
        getRangeBetweenHeaderAndTotal: () => intervalo(f, 1, 0, t.rows.length, t.headers.length),
        getRowCount: () => { cont.leituras++; return t.rows.length; },
        addRows: (i, vs) => { cont.escritas++; t.rows.push(...vs.map((l) => [...l])); },
        deleteRowsAt: (i, n = 1) => { cont.escritas++; t.rows.splice(i, n); },
        setName: (n) => tabelas.set(n, t),
      });
      f.tabela = t;
      return t;
    };
    folhas.set(nome, f);
    return f;
  }
  // Coordenadas absolutas: linha 0 = cabeçalho da tabela da folha; linhas 1.. = corpo.
  function intervalo(f, r, c, nr, nc) {
    const corpo = () => f.tabela.rows;
    return {
      getValues: () => { cont.leituras++; return corpo().slice(r - 1, r - 1 + nr).map((l) => l.slice(c, c + nc)); },
      setValues: (vs) => {
        cont.escritas++;
        if (r === 0) { f.cab = vs[0].slice(); return; }
        vs.forEach((l, i) => l.forEach((v, j) => { corpo()[r - 1 + i][c + j] = v; }));
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
    ler: (n) => { const t = tabelas.get(n); return t.rows.map((l) => Object.fromEntries(t.headers.map((h, j) => [h, l[j]]))); },
  };
}
export const chamar = (w, pedido) => JSON.parse(JSON.stringify(ctx.main(w, JSON.stringify(pedido))));
