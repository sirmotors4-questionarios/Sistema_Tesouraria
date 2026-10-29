# Contrato de integração — Despesas (versão 0.1)
**Estado:** pronto para testes de transformação local; **ainda não existe API Web autenticada activa**.

## Fonte de verdade
- Excel SharePoint: tabela `Lançamentos`, folha `T. DESPESAS`.
- Formulários VBA: `Lançamento.frm`, `Despachos.frm`, `C_Docs.frm`, `F_Banco.frm` e `F_TPag.frm`.
- Frontend: `app.js` encontra-se em `mode: "demo"`. Os dados demonstrativos nunca devem ser confundidos com lançamentos reais.
- Recebimentos: suspenso até validação da integração PRIMAVERA.

## Conversão
A função `TES_DATA.fromExcel` transforma uma linha da tabela `Lançamentos` em propriedades do Web; `TES_DATA.toExcel` faz a conversão de campos de despesas para uma operação de escrita **ainda não activada**. `TES_DATA.contractCheck` compara o cabeçalho real com as colunas esperadas.
As variáveis do site/ficheiro no Power Automate são `sir_sir_TES_SiteSharePoint`, `sir_sir_TES_Biblioteca`, `sir_sir_TES_CaminhoExcel`. Os identificadores do SharePoint precisam ser resolvidos/testados no conector.

## Regras para produção
1. Gerar `ID_Despesa` imutável em migração controlada e **nunca** usar Nr como chave técnica; conservar Nr e Nr de Requisição.
2. Bloquear escrita por duas sessões simultâneas, validar versão e chave, implementar idempotência, trilho de auditoria, permissões.
3. Não expor endpoints HTTP secretos em GitHub Pages nem realizar escrita Excel directamente no browser.
4. Validar divergências de nomes reais das 32 colunas e estados históricos antes de executar importação.
5. Requer autenticação do lado do serviço e contrato de resposta: `{registos:[...],total:n}`.
6. Dados monetários: a camada de processamento deverá preservar precisão decimal; a conversão JS é apenas para apresentação/validação inicial, não contabilidade transaccional.
7. Execução integral do VBA **não foi confirmada**: requer Excel Desktop no ambiente operacional, com cópia de testes.

## Testes automatizados
Executar `node tests/contrato.test.js` num ambiente com Node.js. Os testes verificam round-trip, valores, datas, estado desconhecido e campo obrigatório.
