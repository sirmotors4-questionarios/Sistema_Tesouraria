# Sistema Web de Tesouraria — Grupo SIR

Primeira implementação do sistema de Tesouraria, baseada nos formulários VBA de Despesas e Despachos e na arquitectura de interface do projecto RH.

## Estado
- **Despesas:** interface inicial implementada, com pesquisa, filtros, novo registo, edição em modo de demonstração, dashboard e exportação CSV.
- **Despachos:** secção inicial de consulta; gravação/associação real depende do modelo Excel.
- **Recebimentos:** deliberadamente suspenso até integração PRIMAVERA.
- **Power Automate:** adaptador preparado, **sem ligação configurada**. Nenhuma escrita real no SharePoint foi activada.
- **Segurança:** não publicar URLs de gatilhos HTTP, tokens, credenciais ou dados financeiros no GitHub. O modo demo guarda exclusivamente no navegador.

## Experiência local
Abrir `index.html`. O aviso superior indica **Modo demonstração**. Dados criados ficam somente no armazenamento local deste navegador e podem ser eliminados em «Limpar dados locais».

## Para ligar a produção
1. Confirmar o livro Excel do SharePoint, nomes das tabelas, colunas e tipos.
2. Definir IDs imutáveis e controlo de versões para despesas e despachos.
3. Implementar endpoints autenticados com autorização do lado do servidor e auditoria; **não** colocar URL secreto de Power Automate em JS público.
4. Acrescentar paginação, idempotência, teste de concorrência e validações de pagamento.
5. Configurar sincronização PRIMAVERA *apenas* após confirmar ERP, licença e mecanismo de leitura.
6. Testar primeiro num ambiente isolado; não usar em operações financeiras até validação.

## Referências de migração
VBA: `Lançamento.frm` → lançamentos; `Despachos.frm` → despacho; `C_Docs.frm` → documentos; `Recebimento.frm` → suspenso. O Excel original não foi fornecido neste ZIP, logo nomes das colunas ainda têm de ser conferidos.

Interface: `index.html`, `styles.css`, `app.js`; configuração operacional em `config.js`. O projecto RH é referência de organização e experiência, não uma fonte de dados partilhada.
