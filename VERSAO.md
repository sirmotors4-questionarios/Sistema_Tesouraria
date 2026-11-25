# Versão dos ficheiros (cache do navegador)

O GitHub Pages deixa o navegador guardar cada ficheiro até 10 minutos. Logo a seguir a uma publicação, a página nova
pode correr com JavaScript ou CSS antigos e mostrar mensagens que já não existem.

Para evitar isso, o `index.html` carrega os ficheiros com `?v=AAAAMMDD-n`. **Em cada alteração a `.js` ou `.css`,
mude o número em todas as referências do `index.html`**:

```bash
V=$(date +%Y%m%d)-1; sed -i -E "s#\?v=[0-9]{8}-[0-9]+#?v=$V#g" index.html
```

O teste `tests/versao.test.js` falha se alguma referência local ficar sem `?v=` ou com versões diferentes.
