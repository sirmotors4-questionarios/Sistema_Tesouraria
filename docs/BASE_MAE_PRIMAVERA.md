# Base mãe — Facturação, Recebimentos e Cobranças a partir do PRIMAVERA

A **base mãe** são tabelas no livro Excel da Tesouraria (SharePoint) com as facturas, notas de crédito e recibos
importados das listagens do PRIMAVERA, **já validados**. O PRIMAVERA continua a ser a origem; a base mãe é o que o
sistema usa para reconciliar e é actualizada só depois de cada importação ser comparada com o que já lá está.

```
Listagem Excel do PRIMAVERA ──► leitura e validação no site ──► comparação com a base mãe ──► gravação (fluxo V9 + Office Script)
                                     (rejeitados com motivo)     (novos / alterados / iguais /        tbl_Faturas, tbl_Recibos,
                                                                  na base mas não no ficheiro)        tbl_Liquidacoes, tbl_Importacoes
Base mãe ──► Facturação (saldo de cada documento) ──► Cobranças (antiguidade por cliente)
         └─► Recebimentos (recibos ↔ facturas, excepções) ──► Conciliação Bancária (recibos ↔ entradas do extracto) ──► tbl_RecebimentosBanco
```

## 1. Ficheiros a exportar do PRIMAVERA (Excel)
| Módulo | Listagem | Colunas reconhecidas (nomes habituais; outros podem ser indicados em «Indicar colunas») |
|---|---|---|
| Facturação | Documentos de venda: FA, FT, FS, FR, VD, ND, NC | Documento (`FA 2026/101`) **ou** Tipo Doc + Série + Número · Data · Data Venc. · Entidade/Cliente · Nome · NUIT · Moeda · Total · Anulado/Estado · Empresa (opcional) |
| Recebimentos | Recibos (RE), **uma linha por documento liquidado** | Recibo/Documento · Data · Cliente · Nome · Modo Pagamento · Referência · Documento Liquidado · Valor Liquidado · Total Recibo |

- Se a listagem não tiver a coluna Empresa, indica-se a empresa no ecrã de importação (as listagens são por empresa).
- Linhas de título, totais e subtotais são ignoradas. A folha com mais documentos válidos é escolhida automaticamente.
- Chaves: `EMPRESA|TIPO|SÉRIE|NÚMERO` (ex. `SIR MOTORS|FA|2026|101`). FR, VD e FS contam como pagas no acto; NC reduz o saldo do cliente.

### 1.1 Relatórios PRIMAVERA reais (formatos reconhecidos sem configuração)
- **«Documentos Emitidos de Vendas»** (relatório Crystal exportado para Excel): empresa na 1.ª linha (com «Pág.»),
  cabeçalho `Entidade | Nome | Documento | Data | Total Merc. | Total IVA | Total Desc. | Total Outros | Impostos | Total Doc.`.
  Todos os valores são importados (mercadoria, IVA, desconto, outros, impostos e total). As linhas «Total <Tipo>» e
  «Total Geral» são usadas como **totais de controlo**: o ecrã mostra, por tipo, o total do relatório, a soma importada e
  se conferem. Rodapés (©, Pág.) são ignorados.
- **«Recebimentos»**: `Documento` (só o tipo) + `Série` + `Número` são combinados; importa também Valor Desconto, Filial,
  Moeda, Estorno e Motivo, e as colunas «Tipo/Série/Número Doc. Original» quando vêm preenchidas (liquidações explícitas).
- **Empresa**: detectada pelo cabeçalho do relatório; se faltar, pelo nome do ficheiro (ex. `Recebimentos_Metrobus.xlsx`);
  se faltar, pelos clientes já existentes na base (≥ 40 % dos clientes de uma só empresa). O campo é preenchido e pode ser
  alterado. Sem empresa identificável, o ecrã pede-a antes de gravar.
- **Tipos**: FA, FT, FS, FR, VD, VDS (venda a dinheiro SchoolBus), ND, NC, DV (devolução a dinheiro); recibos RE, VDR
  (recibo de VD/VDS) e NPC (pagamento a cliente que liquida DV/NC — saída, não vai ao banco como entrada). Tipos
  desconhecidos são importados com aviso e o sinal é tirado do valor.
- Códigos de cliente com «O» em vez de «0» (ex. `CLO0101`) são corrigidos com aviso; recibos de valor 0 e documentos
  estornados entram na base (estado «Estornado») mas não contam para saldos.

## 2. Validação de cada importação (antes de gravar)
- **Rejeitados** (com a linha do Excel e o motivo): documento não reconhecido, data ou total inválidos, cliente em falta,
  documento repetido no ficheiro, recibo no ficheiro de facturas e vice-versa.
- **Novos**: não existem na base. **Alterados**: existem com outro conteúdo, com a lista de campos (antes → depois).
  **Iguais**: não são reenviados. **Na base mas não no ficheiro**: documentos da mesma empresa, tipos e período que o
  PRIMAVERA já não lista — a confirmar (possível anulação); **nunca são apagados automaticamente**.
- O SHA-256 do ficheiro é guardado: importar o mesmo ficheiro outra vez gera um aviso.

## 3. Reconciliação interna
- **Facturas ↔ recibos**: liquidado e saldo de cada documento; estados Em aberto, Vencido, Parcial, Parcial vencido, Liquidado, Crédito.
- **Liquidações inferidas** (quando o PRIMAVERA não exporta a ligação recibo → documento), sempre marcadas «(inferida)»
  com a regra usada: VDR → VD/VDS e NPC → DV/NC do mesmo cliente e valor até 31 dias; documento único com saldo igual ao
  recibo; vários iguais → o mais antigo (marcado «rever»); pago sem IVA; combinação única de até 4 documentos.
- **Excepções**: recibos por alocar classificados (anteriores ao período importado/adiantamento, pagamento parcial com
  candidatos, por alocar); NPC sem DV; devolvido acima do documento; liquidação de documento que não existe na base; recibo sem documento liquidado (adiantamento/por alocar);
  soma das liquidações diferente do recibo; liquidado acima do total.
- **Cobranças**: por cliente, em aberto por antiguidade (não vencido, 1–30, 31–60, 61–90, > 90 dias) à data de referência,
  créditos por regularizar (NC e recibos por alocar), saldo líquido, maior atraso e último recibo.
- **Recibos ↔ banco** (módulo Conciliação Bancária): as entradas do extracto são emparelhadas com os recibos por valor
  exacto, data e referência/cliente; recibos em numerário ficam de fora. As ligações confirmadas gravam-se em
  `tbl_RecebimentosBanco` e o recibo passa a «ligado».

## 4. Tabelas da base mãe (criadas pelo script na primeira gravação, em folhas `BM_*`; colunas novas são acrescentadas automaticamente a tabelas existentes)
| Tabela | Colunas |
|---|---|
| `tbl_Faturas` | ID_Documento, Empresa, Tipo_Documento, Serie, Numero, Documento, Data, Data_Vencimento, Codigo_Cliente, Cliente, NUIT, Moeda, Total, Total_Mercadoria, Total_IVA, Total_Desconto, Total_Outros, Impostos, Sinal, Tipo_Nome, Estado, Hash, ID_Importacao, Data_Importacao |
| `tbl_Recibos` | ID_Recibo, Empresa, Tipo_Documento, Serie, Numero, Documento, Data, Codigo_Cliente, Cliente, Valor, Modo_Pagamento, Referencia, Conta_Bancaria, Valor_Desconto, Filial, Moeda, Motivo, Estorno, Tipo_Nome, Estado, Liquidacoes, Hash, ID_Importacao, Data_Importacao |
| `tbl_Liquidacoes` | ID_Liquidacao, ID_Recibo, ID_Documento, Documento_Liquidado, Valor, ID_Importacao |
| `tbl_Importacoes` | ID_Importacao, Tipo, Empresa, Ficheiro, Hash_Ficheiro, Utilizador, Data, Linhas_Ficheiro, Novos, Alterados, Iguais, Rejeitados, Id_Pedido |
| `tbl_RecebimentosBanco` | ID_Ligacao, ID_Recibo, Data_Banco, Valor, Descritivo_Banco, Conta, Utilizador, Data |

## 5. Configuração no Power Automate (uma vez)
1. **Office Script**: no livro da Tesouraria → Automatizar → Novo script → colar `office-scripts/TesourariaBaseMae.ts` →
   guardar com o nome **TesourariaBaseMae**.
2. **Fluxo de escrita V9** (o mesmo endpoint já usado pelas despesas): acrescentar, antes dos ramos existentes, uma
   condição `startsWith(triggerBody()?['operacao'], 'base.')`:
   - Sim → **Executar script** (Excel Online Business): ficheiro *Sistema de Tesouraria*, script *TesourariaBaseMae*,
     parâmetro `pedido` = `string(triggerBody())` → **Resposta** 200 com o corpo = `outputs('Executar_script')?['body/result']`
     (contém `sucesso`, `codigo`, `mensagem`, `novos`, `alterados`).
   - O site envia lotes de até 1 000 registos; cada chamada fica bem abaixo dos 120 s.
3. **Fluxo de consulta** (`TES_API_Consultar`): acrescentar *List rows present in a table* para `tbl_Faturas`,
   `tbl_Recibos`, `tbl_Liquidacoes`, `tbl_RecebimentosBanco` e `tbl_Importacoes` (com paginação, limite adequado ao
   volume) e incluí-las em `tabelas` (`[{"tabela":"tbl_Faturas","linhas":[...]}]`). Na primeira utilização, depois da
   primeira gravação (que cria as tabelas), adicionar estas acções ao fluxo.
4. Permissões: a operação `base.*` deve ser limitada no fluxo aos perfis financeiros autorizados.

Enquanto o endpoint de escrita não estiver configurado, o site permite «Usar só nesta sessão» para testar a leitura,
a comparação e as vistas sem gravar nada.

## 6. Testes
`node --test tests/faturacao.test.js tests/base-mae-script.test.mjs` — leitura das listagens, comparação com a base,
reconciliação, antiguidade, recibos ↔ banco e o próprio Office Script num Excel simulado (criação das tabelas,
upsert, substituição das liquidações, validação antes de escrever, reimportação sem duplicar).
