# Integração dos cadastros — TES_API_Consultar

A base Excel original contém as tabelas abaixo na folha **TABELAS**. O fluxo HTTP actualmente em uso devolve `Lançamentos`; para disponibilizar os cadastros oficiais, deve acrescentar a leitura destas tabelas ao mesmo fluxo (sem criar novos gatilhos).

| Tabela exacta | Colunas exactas | Utilização |
|---|---|---|
| `Descrit_Empresas` | Empresa | Lista de empresas |
| `Descrit_Fornecedores` | Fornecedor | Cadastro de fornecedores |
| `Fornecedor` | Fornecedor | Segunda lista de fornecedores existente; verificar se é a fonte operacional preferida |
| `Descrit_categoria_despesa` | Categoria, Classe, Conta, Descrição da conta principal | Cadastro e classificação de categorias |
| `Tabela5` | Sector, Empresa | Sectores por empresa |
| `Tabela514` | Departamento, Sector | Departamentos por sector |
| `Descrit_tcusto` | Tipo de Custo | Tipos de custo |
| `Regulariadade_pgto` | Regularidade | Regularidades |
| `T_Requitado` | Requisitado Por | Requisitantes |
| `Prioridades` | Numeração, Descrição, Codigo de Corres, Column1 | Prioridades |
| `Bancos` | Banco, Empresa, Nº de Conta | Bancos e contas |
| `Tbl_Sta` | Status | Estados |
| `T_Ano` | Ano | Anos de referência |

## Contrato JSON de consulta

Devolver uma resposta HTTP `200` com o mesmo campo `registos` e acrescentar `tabelas`:

```json
{
  "registos": [{"Nr de Requisição": "EXEMPLO", "Empresa": "EXEMPLO", "Data": 45292}],
  "tabelas": [
    {"tabela": "Descrit_Empresas", "linhas": [{"Empresa": "EXEMPLO"}]},
    {"tabela": "Descrit_categoria_despesa", "linhas": [{"Categoria": "EXEMPLO", "Classe": ""}]}
  ]
}
```

**Nota:** os exemplos acima não são dados reais. O site já interpreta tanto `tabelas:[{tabela,linhas}]` como `tabelas:{nome:[linhas]}` e preserva o fallback com valores distintos presentes nos lançamentos quando ainda faltam cadastros.

## Power Automate — passos restantes

Em `TES_API_Consultar`, acrescentar acções **List rows present in a table** apontadas ao mesmo livro SharePoint para cada cadastro, com paginação adequada. Em `Resposta_HTTP`, montar `tabelas` com as saídas `body('NOME_DA_ACAO')?['value']` das respectivas acções (referenciar os nomes reais criados). Evitar acções de escrita: a integração de cadastros é somente de leitura.

A fonte definitiva entre `Descrit_Fornecedores` (65 linhas de intervalo no ficheiro original) e `Fornecedor` (303 linhas de intervalo) deve ser determinada antes de configurar os menus de registo em produção. A implementação Web actual usa `Descrit_Fornecedores` e complementa com valores dos lançamentos, sem alterar a base.

O formulário Web continua em modo consulta quando conectado; não é possível assumir que o cadastro esteja totalmente sincronizado enquanto o HTTP não devolver a propriedade `tabelas`.
