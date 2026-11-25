# Bancos e Extractos — importação para a base mãe e conciliação de despesas

O site lê o extracto no navegador (Excel, CSV, PDF com texto, PDF digitalizado ou fotografia por OCR) e grava-o na
**base mãe**, que fica no livro *Sistema de Tesouraria* no SharePoint. A gravação passa sempre pelo **Power Automate**:
o fluxo de escrita V9 executa o Office Script `TesourariaBaseMae`. Pela mesma via grava-se a **conciliação** de cada
movimento do banco com as despesas (saídas) e os recibos.

## 1. Regras que garantem que a base mãe nunca tem duplicados

| Regra | Onde é garantida |
|---|---|
| O mesmo ficheiro não entra duas vezes. Cada ficheiro é identificado pelo SHA-256 do seu conteúdo (`ID_Extrato`). | Office Script: `EXTRATO_JA_IMPORTADO`. O site também bloqueia antes de enviar. |
| O mesmo movimento não entra duas vezes, mesmo que venha de outro formato (CSV, Excel ou PDF), de um extracto com período sobreposto ou de um lote repetido. | Office Script: só **acrescenta** movimentos cuja chave `ID_Movimento` ainda não existe. |
| Um extracto só é gravado na conta certa: o nº de conta indicado no ficheiro tem de coincidir com a conta seleccionada. | Site (bloqueia) e Office Script (recusa movimentos de outra conta). |
| Cada documento (despesa ou recibo) fica conciliado com um único movimento. Cada movimento fica conciliado com um único conjunto de documentos, cuja soma tem de ser igual ao valor do banco. | Office Script: `DOCUMENTO_JA_CONCILIADO`, `MOVIMENTO_JA_CONCILIADO` e `VALIDACAO`. |
| Só se concilia o que está na base mãe. | Office Script: `MOVIMENTO_INEXISTENTE` («grave primeiro o extracto»). |

**Chave do movimento:** `BANCO|CONTA|AAAA-MM-DD|cêntimos|ocorrência`, por exemplo `BCI|2308710007|2026-01-14|-7147764|1`.
- A chave não usa o descritivo, a referência nem o saldo, porque estes mudam entre formatos ou vêm cortados no export.
- A «ocorrência» numera, pela ordem do ficheiro, os movimentos iguais do mesmo dia (por exemplo, duas comissões de 167,99).

Esta chave foi validada com os extractos reais:
- o ABSA Jan–Jul em bruto tem os seus 86 movimentos contidos no normalizado;
- o `BCI 23.07` tem 4 de 4 contidos;
- o `trial1.csv` tem 66 movimentos em comum, mais 329 de Ago–Dez/2025 que são novos.

**Nome padronizado:** `BANCO_CONTA_AAAA-MM-DD_AAAA-MM-DD.ext`, com o 1.º e o último dia do extracto, por exemplo
`BCI_2308710007_2026-01-08_2026-10-31.csv`.
- O Office Script atribui o nome final. Se outro ficheiro, com conteúdo diferente, da mesma conta e período já usou o nome, recebe `_v2`, `_v3`…
- O ficheiro original não é alterado. O nome padronizado fica registado em `tbl_Extratos`.
- O ficheiro pode ser guardado no SharePoint com esse nome pelo fluxo (passo 3.3) ou descarregado pelo botão «Descarregar com nome padronizado».

## 2. Tabelas criadas pelo script na primeira gravação (folhas `BM_*`)

| Tabela | Chave | Colunas |
|---|---|---|
| `tbl_Extratos` | ID_Extrato (SHA-256) | ID_Extrato, Nome_Original, Nome_Padronizado, Banco, ID_Conta, Conta, Formato, De, Ate, Movimentos_Ficheiro, Novos, Ja_Existentes, Saldo_Inicial, Saldo_Final, Estado (`Parcial`/`Completo`/`Sem movimentos novos`), Utilizador, Data_Importacao, Id_Pedido |
| `tbl_MovimentosBancarios` | ID_Movimento | ID_Movimento, ID_Conta, Banco, Conta, Data_Movimento, Data_Valor, Referencia, Descricao, Debito, Credito, Valor, Saldo, ID_Extrato, Data_Importacao |
| `tbl_ConciliacaoBancaria` | ID_Conciliacao (`despesa\|ID` ou `recibo\|ID`) | ID_Conciliacao, ID_Movimento, Tipo_Documento, ID_Documento, Documento, Valor, Data_Banco, Valor_Banco, Descritivo_Banco, Conta, Pontos, Motivos, Metodo (`automatico`/`confirmado`/`manual`), Utilizador, Data_Importacao |

As conciliações com recibos também são gravadas em `tbl_RecebimentosBanco`, usada pelos módulos de Recebimentos e
Cobranças.

As despesas são identificadas pelo `ID_Despesa` da consulta ou, na falta dele, pelo `Nr` (`NR-<Nr>`). Uma despesa
sem nenhum dos dois pode ser conciliada no ecrã, mas não fica gravada.

## 3. Configuração no Power Automate (uma vez)

### 3.1 Office Script
1. Abra o livro *Sistema de Tesouraria 1.2.xlsx* no SharePoint.
2. Vá a **Automatizar**, abra o script **TesourariaBaseMae** (ou crie-o) e substitua o código pelo de
   `office-scripts/TesourariaBaseMae.ts`.
3. Guarde o script **com o mesmo nome**.

O script mantém as operações `base.*` da Facturação e acrescenta as operações `banco.*`.

### 3.2 Fluxo de escrita V9 (o mesmo endpoint já configurado no site)
Acrescente a condição antes dos ramos das despesas. Se já existir a condição `base.`, alargue-a:

```
or(startsWith(triggerBody()?['operacao'], 'base.'), startsWith(triggerBody()?['operacao'], 'banco.'))
```

No ramo **Sim**:
1. **Compor** (nome `Pedido`), com a expressão `removeProperty(triggerBody(), 'ficheiro')`.
   O ficheiro original nunca vai para o script.
2. **Executar script** (Excel Online Business):
   - Localização: o site *SIR Comercio Internacional*;
   - Biblioteca: Documentos;
   - Ficheiro: *Sistema de Tesouraria 1.2.xlsx*;
   - Script: *TesourariaBaseMae*;
   - Parâmetro `pedido`: `string(outputs('Pedido'))`.
3. **Resposta**: estado 200, corpo `outputs('Executar_script')?['body/result']`.
   O corpo traz `sucesso`, `codigo`, `mensagem`, `novos`, `duplicados`, `estado`, `ligacoes` e `nome_padronizado`.
   O site só considera gravado quando `sucesso` é `true`.

### 3.3 Guardar o ficheiro original no SharePoint (opcional, recomendado)
Esta parte só é usada quando o utilizador marca «Guardar também o ficheiro no SharePoint». O site envia então, no
último lote do extracto, `ficheiro: {nome, pasta, conteudo_base64}`, com um máximo de 10 MB por ficheiro.

Depois de **Executar script**, acrescente uma condição:

```
and(equals(outputs('Executar_script')?['body/result/sucesso'], true), not(empty(triggerBody()?['ficheiro'])))
```

No ramo **Sim**, use **Criar ficheiro** (SharePoint):
- Endereço do site: *SIR Comercio Internacional*;
- Caminho da pasta: `/Shared Documents/General/Tesouraria e Facturação/` seguido de `triggerBody()?['ficheiro']?['pasta']`.
  A pasta fica, por exemplo, `02_Importacoes/Bancos/BCI`; confirme o caminho da biblioteca no seu site;
- Nome do ficheiro: `outputs('Executar_script')?['body/result/nome_padronizado']`.
  Use o nome devolvido pelo script e não o que vem do site: o script já garantiu que o nome é único;
- Conteúdo do ficheiro: `base64ToBinary(triggerBody()?['ficheiro']?['conteudo_base64'])`.

Só é criado quando o script aceitou o extracto: um ficheiro repetido é recusado pelo script e não é guardado de novo.

### 3.4 Fluxo de consulta (`TES_API_Consultar`)
1. Acrescente **Listar linhas presentes numa tabela** para `tbl_Extratos`, `tbl_MovimentosBancarios` e
   `tbl_ConciliacaoBancaria`.
   - Ative a **paginação**, com um limite acima do nº de movimentos; hoje são cerca de 6 800 nas 8 contas.
   - Use o formato de data **número de série**, como nas outras tabelas.
2. Inclua as três tabelas em `tabelas` (`[{"tabela":"tbl_MovimentosBancarios","linhas":[...]}]`).

As tabelas só existem depois da primeira gravação. Grave o primeiro extracto e depois acrescente estas acções.

Sem estas tabelas na consulta, o site funciona na mesma: o script verifica os duplicados ao gravar. Perdem-se três
coisas:
- o aviso prévio «já na base»;
- a vista «Bancos e Extractos» com o histórico;
- a opção «Usar movimentos já gravados na base mãe».

### 3.5 Limites e permissões
- O site envia os movimentos em lotes de 1 000 e as conciliações em lotes de 1 000. Cada chamada fica bem abaixo do
  limite de 120 s do *Executar script*.
- Um extracto de 5 008 movimentos (ACCESS POS) gera 6 chamadas. Conte com isto no limite diário de execuções de
  scripts por utilizador.
- Se um lote falhar a meio, repetir é seguro: os lotes já gravados não duplicam. O extracto fica `Parcial` até ao
  último lote.
- Limite as operações `banco.*` e `base.*` aos perfis financeiros autorizados.

## 4. Utilização

1. **Conciliação Bancária → Extracto bancário:** escolha o ficheiro.
   - A conta é seleccionada automaticamente pelo nº indicado no ficheiro (e pelo banco, quando o ficheiro o diz).
   - Se a conta seleccionada for outra, a gravação fica bloqueada.
2. **Movimentos lidos:** confira os saldos e as linhas.
   - O painel da base mãe mostra o nome padronizado, quantos movimentos são novos e quantos já existem, e se o ficheiro já foi importado.
   - **Gravar extracto na base mãe** só grava os movimentos novos.
   - Com avisos de saldo, é preciso marcar «Gravar mesmo com avisos de saldo».
3. **Conciliação com a Tesouraria:**
   - Clique em **Conciliar**. Os movimentos e os documentos já conciliados na base ficam de fora; veja o separador «Já conciliados na base».
   - Confirme as sugestões e use «Associar…» para os casos manuais.
   - Depois clique em **Gravar conciliação**.
4. Para corrigir uma conciliação gravada, use **Desfazer** em «Já conciliados na base». Isto corre `banco.desfazer` e o
   movimento volta a estar por conciliar.
5. **Usar movimentos já gravados na base mãe:** concilia uma conta sem voltar a ler ficheiros. Usa os movimentos
   ainda não conciliados.
6. **Bancos e Extractos:** mostra as contas com movimentos (período, último saldo, conciliados e por conciliar) e o
   registo de todos os extractos importados.

## 5. Contrato dos pedidos (para testar o fluxo à mão)

```json
{"operacao":"banco.importar_extrato","id_pedido":"…","final":true,
 "extrato":{"ID_Extrato":"<sha256>","Nome_Original":"BCI_2308710007.csv","Nome_Padronizado":"BCI_2308710007_2026-01-08_2026-10-31.csv",
            "Banco":"BCI","ID_Conta":"…","Conta":"2308710007","Formato":"CSV","De":"2026-01-08","Ate":"2026-10-31",
            "Movimentos_Ficheiro":107,"Saldo_Inicial":574977.64,"Saldo_Final":-1},
 "movimentos":[{"ID_Movimento":"BCI|2308710007|2026-01-08|101500000|1","ID_Conta":"…","Banco":"BCI","Conta":"2308710007",
                "Data_Movimento":"2026-01-08","Data_Valor":"2026-01-08","Referencia":"771389574","Descricao":"TRF : TMCEL…",
                "Debito":0,"Credito":1015000,"Valor":1015000,"Saldo":1589977.64}],
 "ficheiro":{"nome":"…","pasta":"02_Importacoes/Bancos/BCI","conteudo_base64":"…"}}

{"operacao":"banco.conciliar","id_pedido":"…","ligacoes":[{"ID_Movimento":"BCI|2308710007|2026-01-14|-7147764|1",
 "Tipo_Documento":"despesa","ID_Documento":"D1","Documento":"REQ-101","Valor":71477.64,"Data_Banco":"2026-01-14",
 "Valor_Banco":-71477.64,"Descritivo_Banco":"Transferência…","Conta":"2308710007","Pontos":100,"Motivos":"valor igual; mesma data","Metodo":"automatico"}]}

{"operacao":"banco.desfazer","id_pedido":"…","movimentos":["BCI|2308710007|2026-01-14|-7147764|1"]}
```

## 6. Testes
- `node --test tests/*.test.js tests/*.test.mjs`: inclui `tests/bancos-script.test.mjs`, que corre o próprio Office
  Script num Excel simulado. Cobre o ficheiro repetido, os períodos sobrepostos, os lotes, as contas trocadas, a
  conciliação idempotente, os conflitos, o desfazer e os nomes únicos.
- O ensaio no navegador corre os extractos reais (BCI, BCI em bruto, ACCESS com 5 008 movimentos) contra um fluxo
  simulado que executa o mesmo script. Gravar, reimportar, conciliar, desfazer, conciliar a partir da base e recarregar
  a página não produziram nenhum duplicado.
