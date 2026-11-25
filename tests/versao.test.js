// Todos os .js/.css locais do index.html com a mesma ?v= (evita misturar versões em cache depois de publicar).
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
test("index.html: scripts e estilos locais com a mesma versão", () => {
  const html = fs.readFileSync(require("node:path").join(__dirname, "..", "index.html"), "utf8");
  const refs = [...html.matchAll(/(?:src|href)="\.\/([\w-]+\.(?:js|css))(\?v=[^"]*)?"/g)];
  assert.ok(refs.length >= 10, "referências encontradas: " + refs.length);
  const semVersao = refs.filter((m) => !m[2]).map((m) => m[1]);
  assert.deepEqual(semVersao, [], "sem ?v=: " + semVersao.join(", "));
  assert.equal(new Set(refs.map((m) => m[2])).size, 1, "versões diferentes: " + [...new Set(refs.map((m) => m[2]))].join(" "));
});
