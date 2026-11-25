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
| Duas gravações ao mesmo tempo (dois utilizadores ou dois separadores) nunca se misturam: as operações que gravam reservam o livro. Quem chega com outra gravação em curso recebe `EM_USO`, sem nada gravado, e o site repete sozinho de 5 em 5 s. | Office Script: reserva em `BM_Controlo!A2:C2`, que expira ao fim de 150 s se o script for interrompido. |
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
- O ficheiro pode ser guardado no SharePoint com esse nome pelo fluxo TES_BASE_MAE (secção 3) ou descarregado pelo botão «Descarregar com nome padronizado».

**Conta com numeração diferente da do banco.** Quando o banco muda de sistema, o nº no extracto pode deixar de
coincidir com o do cadastro. É o caso do Vista Bank depois da mudança de gestão: o extracto diz 3368 e o cadastro tem
o nº novo.
- O site não obriga a escolher o nº do ficheiro. Mostra os dois números e a opção **«Confirmo que este extracto é desta
  conta»**.
- Depois de confirmar e gravar, a associação (nº e NIB do extracto → conta do cadastro) fica guardada **neste
  navegador**. Na vez seguinte, a conta é escolhida automaticamente e aparece «associada a esta conta (confirmado
  antes)».
- Os movimentos ficam sempre registados com o nº do **cadastro**, por isso extractos do sistema antigo e do novo não
  se duplicam.

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

A integração é feita por um fluxo próprio, **TES_BASE_MAE**, entregue como pacote de importação. O pacote foi feito a
partir do export do fluxo `FIN_IA_CONCILIACAO`, por isso usa o mesmo site, a mesma biblioteca e a mesma ligação do
Excel. O fluxo recebe os pedidos `base.*` e `banco.*` do site, corre o Office Script e, se for pedido, guarda o
extracto original no SharePoint. As despesas continuam no fluxo V9, que não é alterado.

O pacote não está neste repositório público, porque contém os identificadores do site e da biblioteca. Está na pasta
`05_Scripts/Power Automate` do SharePoint da Tesouraria, junto com o script.

### 3.1 Office Script
1. Abra *Sistema de Tesouraria 1.2.xlsx* no Excel para a Web e vá a **Automatizar → Novo script**.
2. Cole `office-scripts/TesourariaBaseMae.ts` e guarde com o nome **TesourariaBaseMae**. Se o script já existir,
   substitua o código.

### 3.2 Importar o fluxo TES_BASE_MAE
O pacote **não traz a acção «Executar script»**. O identificador do script só existe na conta de cada utilizador, e o
Power Automate valida-o ao importar: com um marcador no lugar dele, recusa com «Unable to parse script reference». Por
isso o pacote traz um bloco **Script** com um marcador, e o resto do fluxo lê o resultado com
`last(result('Script'))`, seja qual for o nome da acção.

1. Em make.powerautomate.com, vá a **Os meus fluxos → Importar → Importar pacote (legado)** e escolha
   `TES_BASE_MAE_pacote.zip`.
2. Escolha **Criar como novo** e a ligação do **SharePoint**. Clique em **Importar**.
3. Abra o fluxo para editar e vá a **Validar_Operacao → Verdadeiro → Script**.
4. Por baixo do marcador `SUBSTITUIR_por_Executar_script`, acrescente **Excel Online (Business) → Executar script**:
   - Localização: *SIR MOTORS INTERNATIONAL*;
   - Biblioteca: *Documentos*;
   - Ficheiro: `General/Tesouraria e Facturação/01_Bases_de_Dados/Sistema de Tesouraria 1.2.xlsx`;
   - Script: **TesourariaBaseMae**;
   - `pedido`: a expressão `string(outputs('Pedido'))`.
5. Elimine o marcador e **guarde** o fluxo.
6. Se a acção **Criar ficheiro** pedir, escolha o site `SIR MOTORS INTERNATIONAL`.
7. Copie o **URL HTTP POST** do accionador e cole-o no site, em **Conectar base de dados → Endpoint HTTPS do
   TES_BASE_MAE**. Fica guardado só nesse navegador e nunca vai para o GitHub.

Enquanto faltar o passo 4, o fluxo responde `SCRIPT_POR_CONFIGURAR` e não grava nada.

O script só aparece na lista se tiver sido criado com a **mesma conta da ligação do Excel**. Se não aparecer, use essa
conta para o criar, ou partilhe o script com o livro (Excel → script → … → *Partilhar com outras pessoas neste livro*).

Notas sobre o fluxo:
- O accionador aceita **Qualquer pessoa** (URL com `sig=`), como os outros fluxos do site. Se a política do tenant o
  impedir, peça ao administrador uma excepção para este fluxo.
- **Sem controlo de simultaneidade no accionador:** o Power Automate não o permite com resposta HTTP síncrona (erro
  `InvalidConcurrencyConfiguration`). A exclusão mútua é feita pela **reserva do livro** no Office Script: quem chega
  com outra gravação em curso recebe `EM_USO` e o site repete sozinho.
- O URL com `sig=` dá acesso ao fluxo. Use-o só em dispositivos de confiança e, se for exposto, regenere-o.

### 3.3 O que o fluxo faz
```
Quando um pedido HTTP é recebido
└─ Validar_Operacao: a operação começa por «banco.» ou «base.»?
   ├─ Sim
   │  ├─ Pedido = removeProperty(corpo, 'ficheiro')      → o ficheiro original nunca vai para o script
   │  ├─ Script (bloco): Executar script TesourariaBaseMae (pedido), acrescentado depois de importar
   │  ├─ Resposta 200 com o resultado do script     → o site não espera pelo SharePoint
   │  ├─ Guardar_Ficheiro: sucesso, importar_extrato e ficheiro enviado?
   │  │  └─ Criar ficheiro (SharePoint) com o nome_padronizado devolvido pelo script
   │  └─ Resposta_Erro_Script 502 se o script falhar ou exceder o tempo
   └─ Não → Resposta 400 OPERACAO_NAO_SUPORTADA
```
- **Consultar:** ao ligar, o site envia `base.consultar` e o script devolve as tabelas da base mãe, em formato
  compacto. Não é preciso mexer no fluxo de consulta das despesas.
- **Ficheiro original:** só é criado quando o script aceitou o extracto. Um ficheiro repetido é recusado pelo script e
  não é guardado de novo.

### 3.4 Alternativa sem fluxo novo
Se preferir usar o fluxo V9:
1. Acrescente-lhe o ramo da secção 3.3. A condição é `or(startsWith(...,'base.'), startsWith(...,'banco.'))`.
2. Deixe vazio o campo do TES_BASE_MAE no site: as operações da base mãe passam a ir para o endpoint V9.
3. Para o site ver o histórico, o fluxo de consulta tem de devolver `tbl_Extratos`, `tbl_MovimentosBancarios` e
   `tbl_ConciliacaoBancaria`. Use *Listar linhas* com paginação, ou chame `base.consultar`.

### 3.5 Rapidez e limites
O tempo de cada pedido depende quase só do *Executar script*. Cada execução tem um custo fixo de arranque e cada
leitura ao livro é uma ida ao servidor do Excel.

No **script**:
- cada tabela é aberta com 2 leituras (cabeçalho e endereço);
- da tabela de movimentos só se lê a coluna `ID_Movimento`;
- cada bloco novo é escrito de uma vez (formatos e valores) e a tabela é redimensionada no fim;
- `base.consultar` lê cada tabela numa só leitura.

| Operação | Leituras antes | Leituras agora |
|---|---|---|
| Importar um lote | 20–33 | 1–8 |
| Ficheiro repetido | 7 | 3 |
| Conciliar | 12 | 3 |
| Consultar a base | 12 | 3 |

Nas escritas, um lote de 1 000 movimentos passou de cerca de 20 para 6.

No **fluxo**:
- responde ao site logo que o script termina e só depois guarda o ficheiro original no SharePoint;
- o *Executar script* tem no máximo uma nova tentativa, ao fim de 5 s, para não passar o tempo de espera do site.

No **site**:
- envia os movimentos em lotes de 2 000 (o script aceita até 2 500) e só os que ainda não estão na base;
- as conciliações seguem em lotes de 1 000 e as listagens do PRIMAVERA também em lotes de 1 000;
- um extracto de 5 008 movimentos gera 3 execuções em vez de 6.

Limites a ter em conta:
- Cada lote conta para o limite diário de execuções de scripts por utilizador.
- O limite do *Executar script* é 120 s. Se um lote falhar, repetir é seguro: o extracto fica `Parcial` até ao último lote.
- A **reserva do livro** garante que duas gravações nunca correm ao mesmo tempo. O preço é que dois utilizadores a gravar
  ao mesmo tempo esperam um pelo outro; o site repete de 5 em 5 s, até 2 minutos. Custa 1 leitura e 2 escritas por pedido.
- `base.consultar` devolve as tabelas inteiras. Hoje são cerca de 7 000 movimentos, entre 1 e 2 MB. Com vários anos
  será preciso filtrar por período.

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
