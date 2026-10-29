# Gateway de consulta segura — implementação preparada

O GitHub Pages é público e **não pode** guardar a URL secreta de um accionador HTTP do Power Automate.

## Infraestrutura necessária

1. Criar serviço de backend protegido por Microsoft Entra ID (por exemplo Azure Functions com Easy Auth), exigindo utilizador autenticado e verificação de grupo/role autorizada.
2. Configurar `TES_FLOW_CONSULTAR_URL` apenas no servidor, através de definições protegidas/Key Vault; nunca guardar na pasta pública.
3. Expor `GET /api/despesas` para enviar os filtros ao Power Automate e receber `{registos:[...],total:n}`.
4. Configurar CORS somente para a origem exacta do site GitHub Pages da Tesouraria e validação de token com audience e tenant correctos.
5. O fluxo `TES_API_Consultar` terá de ser alterado: o accionador manual actual **não** recebe pedidos HTTP. Configurar accionador HTTP protegido pelo Entra, com resposta JSON, ou usar intermediário com credenciais de aplicação controladas. Uma alteração de accionador pode exigir licença premium e permissões.
6. Definir em `config.js` `apiEnabled:true` e `apiBaseUrl` apenas após a publicação do backend; implementar `TES_AUTH.getAccessToken()` com MSAL e app registration próprios da empresa.
7. Testar exaustivamente num ambiente de cópia antes de activar escrita. Este gateway é exclusivamente para leitura.

**Estado:** contrato preparado, backend não publicado, nenhuma conexão real activada. Não colocar segredos neste repositório.
