/* Conciliação bancária — leitura de extractos no navegador.
 * Excel/CSV → SheetJS (ou leitor CSV próprio); PDF com texto → pdf.js; PDF digitalizado/imagem → OCR (Tesseract.js).
 * As bibliotecas são carregadas só quando necessárias, em versões fixas. O ficheiro do extracto nunca sai do
 * navegador: o OCR corre localmente (apenas o motor e o dicionário de português são descarregados do CDN).
 */
(function (root) {
  "use strict";
  const C = root.TES_CONCILIACAO;
  const LIBS = Object.freeze({
    xlsx: "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs",
    pdf: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs",
    pdfWorker: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs",
    tesseract: "https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js",
  });
  const cache = {};
  const importar = (url) => (cache[url] ||= import(url).catch((e) => { delete cache[url]; throw Error("Não foi possível carregar o leitor (" + url.split("/")[2] + "). Verifique a ligação à Internet."); }));
  const script = (url, global) => (cache[url] ||= new Promise((ok, falha) => {
    if (root[global]) return ok(root[global]);
    const s = document.createElement("script");
    s.src = url; s.async = true; s.crossOrigin = "anonymous";
    s.onload = () => (root[global] ? ok(root[global]) : falha(Error("Leitor OCR indisponível.")));
    s.onerror = () => { delete cache[url]; falha(Error("Não foi possível carregar o leitor OCR. Verifique a ligação à Internet.")); };
    document.head.appendChild(s);
  }));

  const extensao = (nome) => (String(nome).toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1] || "";
  // Ano mais frequente no texto (para datas "05/11" sem ano em alguns PDF)
  function anoProvavel(textos) {
    const conta = {};
    for (const t of textos) for (const m of String(t).matchAll(/\b(20\d{2})\b/g)) conta[m[1]] = (conta[m[1]] || 0) + 1;
    const [ano] = Object.entries(conta).sort((a, b) => b[1] - a[1])[0] || [];
    return ano ? Number(ano) : new Date().getFullYear();
  }

  // ---------- CSV ----------
  function decodificar(buffer) {
    const utf8 = new TextDecoder("utf-8").decode(buffer);
    return utf8.includes("�") ? new TextDecoder("windows-1252").decode(buffer) : utf8;
  }
  function lerCsv(textoCsv) {
    const linhas = textoCsv.replace(/^﻿/, "").split(/\r?\n/).slice(0, 30);
    const contar = (sep) => linhas.reduce((s, l) => s + (l.split(sep).length - 1), 0);
    const sep = [";", "\t", ",", "|"].sort((a, b) => contar(b) - contar(a))[0];
    const out = [];
    let linha = [], celula = "", aspas = false;
    const t = textoCsv.replace(/^﻿/, "");
    for (let i = 0; i < t.length; i++) {
      const c = t[i];
      if (aspas) {
        if (c === '"' && t[i + 1] === '"') { celula += '"'; i++; } else if (c === '"') aspas = false; else celula += c;
      } else if (c === '"') aspas = true;
      else if (c === sep) { linha.push(celula); celula = ""; }
      else if (c === "\n" || c === "\r") { if (c === "\r" && t[i + 1] === "\n") i++; linha.push(celula); out.push(linha); linha = []; celula = ""; }
      else celula += c;
    }
    if (celula || linha.length) { linha.push(celula); out.push(linha); }
    // Algumas exportações (ABSA) põem a linha inteira entre aspas quando o descritivo tem uma vírgula: voltar a separar
    const largura = Math.max(0, ...out.slice(0, 50).map((l) => l.length));
    return out.map((l) => (l.length === 1 && largura > 2 && l[0].includes(sep) ? lerCsv(l[0])[0] || l : l));
  }

  // ---------- Excel ----------
  async function lerExcel(buffer) {
    const XLSX = await importar(LIBS.xlsx);
    const livro = XLSX.read(new Uint8Array(buffer), {type: "array", cellDates: false, dense: true});
    // A folha com mais linhas que parecem movimentos
    let melhor = null;
    for (const nome of livro.SheetNames) {
      const matriz = XLSX.utils.sheet_to_json(livro.Sheets[nome], {header: 1, raw: true, defval: "", blankrows: true});
      const r = C.matrizParaMovimentos(matriz);
      if (!melhor || r.movimentos.length > melhor.resultado.movimentos.length) melhor = {folha: nome, matriz, resultado: r};
    }
    return melhor || {folha: "", matriz: [], resultado: C.matrizParaMovimentos([])};
  }

  // ---------- PDF ----------
  async function abrirPdf(buffer) {
    const pdfjs = await importar(LIBS.pdf);
    pdfjs.GlobalWorkerOptions.workerSrc = LIBS.pdfWorker;
    return pdfjs.getDocument({data: new Uint8Array(buffer), isEvalSupported: false}).promise;
  }
  // Agrupa os fragmentos de texto em linhas pela coordenada vertical (tolerância de 3 pontos).
  function agruparLinhas(itens) {
    const ordenados = itens.filter((i) => i.str.trim()).sort((a, b) => b.y - a.y || a.x - b.x);
    const linhas = [];
    for (const it of ordenados) {
      const u = linhas[linhas.length - 1];
      if (u && Math.abs(u.y - it.y) <= 3) u.itens.push(it); else linhas.push({y: it.y, itens: [it]});
    }
    return linhas;
  }
  async function textoDoPdf(doc, aoProgresso) {
    const linhas = [];
    let caracteres = 0;
    for (let n = 1; n <= doc.numPages; n++) {
      aoProgresso?.("A ler o PDF: página " + n + " de " + doc.numPages + "…");
      const pagina = await doc.getPage(n);
      const conteudo = await pagina.getTextContent();
      const itens = conteudo.items.filter((i) => typeof i.str === "string").map((i) => ({x: i.transform[4], y: i.transform[5], w: i.width, str: i.str}));
      caracteres += itens.reduce((s, i) => s + i.str.trim().length, 0);
      // Cada página abaixo da anterior: desloca o y para as linhas não se misturarem
      for (const l of agruparLinhas(itens)) linhas.push({y: l.y - n * 10000, itens: l.itens, pagina: n});
    }
    return {linhas, caracteres};
  }
  function interpretarLinhas(linhas, ano) {
    const tabela = C.linhasPdfParaMatriz(linhas);
    let resultado = tabela ? C.matrizParaMovimentos(tabela.matriz, {mapa: tabela.mapa, linhaCabecalho: 0, ano}) : null;
    const textos = linhas.map((l) => [...l.itens].sort((a, b) => a.x - b.x).map((i) => i.str).join(" "));
    if (!resultado || !resultado.movimentos.length) {
      const alt = C.textoParaMovimentos(textos, {ano});
      if (!resultado || alt.movimentos.length > resultado.movimentos.length) resultado = {...alt, mapa: null, linhaCabecalho: -1, saldoFinal: null};
    }
    return {resultado, matriz: tabela?.matriz || null, textos};
  }

  // ---------- OCR ----------
  let trabalhadorOcr = null;
  async function ocr(fontes, aoProgresso) {
    const T = await script(LIBS.tesseract, "Tesseract");
    if (!trabalhadorOcr) {
      aoProgresso?.("A preparar o OCR (português)… na primeira vez pode demorar.");
      trabalhadorOcr = await T.createWorker("por", 1, {logger: (m) => { if (m.status === "recognizing text") aoProgresso?.("OCR: " + Math.round(m.progress * 100) + "% da página…"); }});
    }
    const linhas = [];
    for (let k = 0; k < fontes.length; k++) {
      aoProgresso?.("OCR: página " + (k + 1) + " de " + fontes.length + "…");
      const {data} = await trabalhadorOcr.recognize(fontes[k]);
      for (const l of data.lines || []) {
        const itens = (l.words || []).filter((w) => w.text && w.text.trim()).map((w) => ({x: w.bbox.x0, w: w.bbox.x1 - w.bbox.x0, y: -w.bbox.y0, str: w.text}));
        if (itens.length) linhas.push({y: -l.bbox.y0 - (k + 1) * 100000, itens});
      }
    }
    return linhas;
  }
  async function paginasParaImagens(doc, aoProgresso) {
    const imagens = [];
    for (let n = 1; n <= doc.numPages; n++) {
      aoProgresso?.("A preparar a página " + n + " de " + doc.numPages + " para OCR…");
      const pagina = await doc.getPage(n);
      const vista = pagina.getViewport({scale: 2.5}); // ~180 dpi: bom compromisso entre precisão e tempo
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(vista.width); canvas.height = Math.ceil(vista.height);
      await pagina.render({canvasContext: canvas.getContext("2d"), viewport: vista}).promise;
      imagens.push(canvas);
    }
    return imagens;
  }

  // ---------- Entrada única ----------
  // Devolve {metodo, nome, matriz (para mapear colunas, quando existe), resultado:{movimentos, saldoInicial, saldoFinal, avisos, mapa}}
  async function lerFicheiro(ficheiro, {aoProgresso, forcarOcr = false} = {}) {
    if (!C) throw Error("Motor de conciliação não carregado.");
    if (ficheiro.size > 25 * 1024 * 1024) throw Error("Ficheiro demasiado grande (máximo 25 MB).");
    const ext = extensao(ficheiro.name), buffer = await ficheiro.arrayBuffer();
    if (ext === "csv" || ext === "txt") {
      aoProgresso?.("A ler o CSV…");
      const matriz = lerCsv(decodificar(buffer));
      return {metodo: "CSV", nome: ficheiro.name, matriz, resultado: C.matrizParaMovimentos(matriz, {ano: anoProvavel(matriz.flat())})};
    }
    if (["xlsx", "xls", "xlsm", "ods"].includes(ext)) {
      aoProgresso?.("A ler o Excel…");
      const r = await lerExcel(buffer);
      return {metodo: "Excel" + (r.folha ? " (folha " + r.folha + ")" : ""), nome: ficheiro.name, matriz: r.matriz, resultado: r.resultado};
    }
    if (ext === "pdf") {
      aoProgresso?.("A abrir o PDF…");
      const doc = await abrirPdf(buffer);
      const {linhas, caracteres} = await textoDoPdf(doc, aoProgresso);
      // PDF digitalizado: quase sem texto por página
      if (!forcarOcr && caracteres / doc.numPages > 80) {
        const ano = anoProvavel(linhas.flatMap((l) => l.itens.map((i) => i.str)));
        const r = interpretarLinhas(linhas, ano);
        if (r.resultado.movimentos.length) return {metodo: "PDF (texto)", nome: ficheiro.name, matriz: r.matriz, resultado: r.resultado, paginas: doc.numPages};
      }
      const imagens = await paginasParaImagens(doc, aoProgresso);
      const linhasOcr = await ocr(imagens, aoProgresso);
      const r = interpretarLinhas(linhasOcr, anoProvavel(linhasOcr.flatMap((l) => l.itens.map((i) => i.str))));
      r.resultado.avisos = ["Lido por OCR: confirme datas, valores e sinais antes de conciliar.", ...(r.resultado.avisos || [])];
      return {metodo: "OCR (PDF digitalizado)", nome: ficheiro.name, matriz: r.matriz, resultado: r.resultado, paginas: doc.numPages};
    }
    if (["png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"].includes(ext)) {
      const linhasOcr = await ocr([ficheiro], aoProgresso);
      const r = interpretarLinhas(linhasOcr, anoProvavel(linhasOcr.flatMap((l) => l.itens.map((i) => i.str))));
      r.resultado.avisos = ["Lido por OCR: confirme datas, valores e sinais antes de conciliar.", ...(r.resultado.avisos || [])];
      return {metodo: "OCR (imagem)", nome: ficheiro.name, matriz: r.matriz, resultado: r.resultado};
    }
    throw Error("Formato não suportado: ." + (ext || "?") + ". Use Excel, CSV, PDF ou imagem.");
  }

  async function exportarExcel(linhas, nome) {
    const XLSX = await importar(LIBS.xlsx);
    const folha = XLSX.utils.json_to_sheet(linhas);
    const livro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(livro, folha, "Conciliação");
    XLSX.writeFile(livro, nome);
  }

  // Folhas em bruto (listagens do PRIMAVERA): [{folha, matriz}] e o SHA-256 do ficheiro (detecção de reimportação).
  async function lerFolhas(ficheiro) {
    if (ficheiro.size > 25 * 1024 * 1024) throw Error("Ficheiro demasiado grande (máximo 25 MB).");
    const ext = extensao(ficheiro.name), buffer = await ficheiro.arrayBuffer();
    let hash = "";
    try { hash = [...new Uint8Array(await crypto.subtle.digest("SHA-256", buffer))].map((b) => b.toString(16).padStart(2, "0")).join(""); } catch {}
    if (ext === "csv" || ext === "txt") return {hash, folhas: [{folha: "CSV", matriz: lerCsv(decodificar(buffer))}]};
    if (!["xlsx", "xls", "xlsm", "ods"].includes(ext)) throw Error("Use o ficheiro Excel (.xlsx/.xls) exportado do PRIMAVERA.");
    const XLSX = await importar(LIBS.xlsx);
    const livro = XLSX.read(new Uint8Array(buffer), {type: "array", cellDates: false, dense: true});
    return {hash, folhas: livro.SheetNames.map((folha) => ({folha, matriz: XLSX.utils.sheet_to_json(livro.Sheets[folha], {header: 1, raw: true, defval: "", blankrows: true})}))};
  }

  root.TES_IMPORTAR = Object.freeze({lerFicheiro, lerFolhas, lerCsv, exportarExcel, LIBS});
})(globalThis);
