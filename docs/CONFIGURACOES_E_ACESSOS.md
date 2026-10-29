# Integração de cadastros e permissões — 29/10/2026

## Origem de dados
SharePoint: SIR MOTORS INTERNATIONAL
Biblioteca: Documentos
Pasta: General/Tesouraria e Facturação
Ficheiro: Sistema de Tesouraria 1.2.xlsx

## Tabelas normalizadas esperadas
- tbl_Fornecedores: ID_Fornecedor, Nome_Fornecedor, Nome_Comercial, NUIT, Tipo_Fornecedor, Categoria_Fornecedor, Contacto, Email, Endereço, País, Condições_Pagamento, Estado, Data_Cadastro, Data_Actualizacao, Observacoes
- tbl_Categorias: ID_Categoria, Categoria, Descricao, Classe_Contabilistica, Conta_Contabilistica, Descricao_Conta, Tipo_Custo, Estado
- tbl_Empresas: ID_Empresa, Nome_Empresa, Nome_Abreviado, NUIT, Endereco, Contacto, Email, Estado
- tbl_Sectores: ID_Sector, Sector, ID_Empresa, Estado
- tbl_Departamentos: ID_Departamento, Departamento, ID_Sector, Estado
- tbl_Requisitantes: ID_Requisitante, Nome, ID_Empresa, ID_Departamento, Funcao, Email, Estado
- tbl_Prioridades: ID_Prioridade, Ordem, Descricao, Cor, Estado
- tbl_Bancos: ID_Banco, Nome_Banco, Estado
- tbl_ContasBancarias: ID_Conta, ID_Banco, ID_Empresa, Numero_Conta, Moeda, Estado

## Contrato HTTP
TES_API_Consultar deve responder:
{"registos":[], "tabelas":[{"tabela":"tbl_Fornecedores","linhas":[]},{"tabela":"tbl_Categorias","linhas":[]}]}
Manter os registos de despesas existentes na propriedade registos.
A API deve preservar a paginação e enviar apenas os campos necessários para cada perfil.

## Onde ficam os módulos
Configurações: Empresas, Sectores, Departamentos, Prioridades, Bancos, Contas e Utilizadores/Permissões.
Cadastros: Fornecedores, Categorias, Requisitantes, Tipos de custo e Regularidades.

## Permissões recomendadas
- Superadministrador: toda a gestão e permissões.
- Administrador Financeiro: consulta e CRUD dos cadastros, excepto contas bancárias e permissões.
- Gestor de Tesouraria: consulta e CRUD operacionais, sujeito a regras de aprovação.
- Técnico de Tesouraria: consulta e propostas de criação/alteração; sem escrita directa.
- Auditor: consulta autorizada sem edição.
- Dados bancários: apenas perfis financeiros explicitamente autorizados.

## Integração de escrita pendente
Criar TES_API_Cadastros_Actualizar com autorização efectiva no servidor, validação do perfil contra tabela de utilizadores, lista branca de tabelas/campos, ID permanente, registo de auditoria e detecção de conflitos. O front-end GitHub Pages não deve determinar permissões por si só.

## Estado
A página Cadastros foi criada no menu Web. Enquanto o fluxo HTTP actual não fornecer o campo tabelas com as novas bases, a página não recebe os cadastros reais. Criar/editá-los permanece bloqueado, correctamente, até estar activa a API de gravação autenticada.
