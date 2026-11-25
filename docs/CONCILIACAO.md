# Conciliação Bancária — módulo (fase 1: leitura, revisão e relatório)

**Estado:** activo no menu *Bancos e Controlo → Conciliação Bancária*. Lê o extracto, emparelha-o com as saídas da
Tesouraria e exporta o relatório. A gravação do extracto e da conciliação na base mãe (pelo Power Automate) está descrita em
[BANCOS_EXTRATOS.md](BANCOS_EXTRATOS.md).

## Ficheiros
| Ficheiro | Papel |
|---|---|
| `conciliacao-core.js` | Motor puro (sem DOM): números/datas, colunas, PDF→tabela, validação de saldos, emparelhamento, relatório. Testado em Node. |
| `conciliacao-import.js` | Leitura no navegador: Excel/CSV, PDF com texto e OCR. Carrega as bibliotecas a pedido. |
| `conciliacao-ui.js` | Interface da secção `#conciliacao`. |
| `app.js` | Só expõe `window.TES_CONTEXTO` (leitura das despesas e contas já carregadas) e o evento `tes:contexto`. |
| `tests/conciliacao.test.js` | `node --test tests/conciliacao.test.js` |

## Formatos de extracto
| Formato | Como é lido | Precisão |
|---|---|---|
| Excel (`.xlsx/.xls/.ods`) e CSV do internet banking | SheetJS / leitor CSV próprio (`;` `,` tab, UTF-8 ou Windows-1252). Procura o cabeçalho nas primeiras 40 linhas. | Exacta |
| PDF gerado pelo banco | pdf.js: texto e posição de cada palavra; as colunas vêm das posições dos títulos (Data, Descritivo, Débito, Crédito, Saldo…). | Exacta |
| PDF digitalizado ou fotografia | OCR Tesseract.js (português) sobre a página renderizada a ~180 dpi; mesma reconstrução por posições; sem cabeçalho, deduz débito/crédito pela variação do saldo. | Rever sempre |

- Cabeçalhos reconhecidos em português e inglês (Data Mov., Data Valor, Descritivo/Detalhes, Nº Doc, Débito/Saídas, Crédito/Entradas, Montante, Saldo…). Se um banco usar outros nomes, **Indicar colunas** permite escolher as colunas; o mapeamento fica guardado no navegador por conta.
- Valores: `1.234,56`, `1 234,56`, `1,234.56`, `(1 234,56)`, `1.234,56-`, `1.234,56 D/C/DR/CR`, prefixo MZN/MT.
- Datas: dd/mm/aaaa, dd-mm-aa, dd.mm.aaaa, ISO, série Excel, `05 Nov 2026`, `05-DEZ-26`, `05/11` (ano deduzido do extracto).
- Linhas de continuação do descritivo são juntadas ao movimento anterior; *Saldo anterior/final* são lidos à parte.

**Validação** antes de conciliar: saldo de cada linha = saldo anterior ± movimento; saldo inicial + movimentos = saldo final; movimentos repetidos. As linhas com problema ficam destacadas e podem ser corrigidas na tabela (data, descritivo, valor, trocar débito/crédito, remover).

## Privacidade
O extracto é lido **no navegador** e fica só na memória da página; não é enviado a nenhum servidor nem guardado.
Do CDN descarregam-se apenas as bibliotecas (versões fixas): SheetJS 0.20.3 (`cdn.sheetjs.com`), pdf.js 4.10.38,
Tesseract.js 5.1.1 e o dicionário de português (`cdn.jsdelivr.net`). No navegador guarda-se apenas o mapeamento de colunas.

## Regras de emparelhamento
Entram as despesas **Pago** e **Aguardando Saída** com valor positivo, na conta escolhida (por `ID_Conta` quando a
despesa o tem — despachos V9 — senão pelo número de conta) e com data de saída no período do extracto ± tolerância.

| Critério | Pontos |
|---|---|
| Valor igual ao cêntimo | **obrigatório** (50) |
| Mesma data / diferença de dias (até à tolerância, 5 por omissão) | +25 / +20 − 3 por dia |
| `Ref. Pagamento` no descritivo do banco | +30 |
| Nº de despacho no descritivo | +25 |
| Nº de requisição no descritivo | +20 |
| Palavras do fornecedor no descritivo | até +15 |

- **1:1** pela melhor pontuação global. **Automático** com ≥ 75 pontos e sem outro candidato com a mesma pontuação; caso contrário fica em **Sugestões** para confirmar ou rejeitar.
- **N:1** (vários pagamentos num só débito, ex. lote de salários): soma exacta de 2 a 5 pagamentos próximos; sempre sugestão.
- **Associação manual** de um débito a um ou mais pagamentos cuja soma seja igual ao débito.
- Decisões (confirmar, rejeitar, desfazer, associar) mantêm-se ao conciliar novamente.
- **Créditos ↔ recibos**: as entradas do extracto são emparelhadas com os recibos da base mãe (ver `docs/BASE_MAE_PRIMAVERA.md`); «Gravar ligações dos recibos» grava as ligações confirmadas em `tbl_RecebimentosBanco`.

## Relatório
CSV (`;`, decimais com vírgula, protegido contra fórmulas) ou Excel: uma linha por movimento do banco (estado,
pontuação, motivos, requisição, despacho, referência, fornecedor, data de saída, valor) e uma por pagamento sem movimento.

## Fase 2 — gravação no Excel (por implementar)
Pelo mesmo endpoint de escrita V9, com `id_pedido` para idempotência e validação de perfil no fluxo:

```json
{"operacao":"conciliacao.registar","id_pedido":"uuid",
 "extracto":{"id_conta":"C7","de":"2026-11-01","ate":"2026-11-30","saldo_inicial":850000,"saldo_final":731999.5,"ficheiro":"extracto.pdf","hash":"sha256"},
 "movimentos":[{"chave":"2026-11-03|12500000|trf alfa","data":"2026-11-03","descritivo":"…","referencia":"","valor":-125000,"saldo":725000}],
 "pares":[{"chave_movimento":"…","ids_despesa":["D-0001"],"tipo":"1:1","pontuacao":100,"motivos":"…"}]}
```

Tabelas sugeridas: `tbl_Extractos` (ID_Extracto, ID_Conta, Periodo_De, Periodo_Ate, Saldo_Inicial, Saldo_Final, Ficheiro,
Hash, Importado_Por, Data_Importacao), `tbl_MovimentosBancarios` (ID_Movimento, ID_Extracto, Data, Data_Valor, Descritivo,
Referencia, Valor, Saldo, Chave, Estado) e `tbl_Conciliacoes` (ID_Conciliacao, ID_Movimento, ID_Despesa, Tipo, Pontuacao,
Motivos, Estado Confirmado/Anulado, Utilizador, Data). O `hash` do ficheiro e a `chave` do movimento impedem importar o
mesmo extracto/movimento duas vezes. Uma despesa conciliada pode passar de *Aguardando Saída* a *Pago* com a data do banco.


## Cheques (nº do cheque dos despachos)

Nos despachos (tabela `Lançamentos`), o meio de pagamento fica em **«F. de Pagamento»** e o nº do cheque em **«Ref. Pagamento»**.
No extracto, o nº aparece no descritivo («Chq.Caixa 45564506», «Ch.Comp.45564409», «Pagamento de Cheque 1294827») ou na referência.

| Situação | Resultado |
|---|---|
| Nº igual e valor igual | **Automático**. O cheque pode ser debitado até 90 dias depois da data de saída (quando é apresentado ao banco) e até 3 dias antes. |
| Vários despachos com o mesmo nº de cheque cuja soma é igual ao débito | **Automático**, um débito ligado a todas as despesas, sem limite de quantidade. |
| Nº abreviado no despacho («4517» = 45564417: início + fim) | **Automático**, com a nota «registar o nº completo». |
| Nº parecido (um algarismo trocado, a mais ou a menos) | **Sugestão**, com a nota «verificar digitação». Os dois números aparecem para corrigir o despacho. |
| Nº diferente com o mesmo valor | Nunca é automático. |
| O mesmo cheque e valor em dois despachos | **Sugestão**, com a nota «possível despacho duplicado». |

Resultado com os extractos reais das contas BCI e ABSA e a base de despachos (25/11/2026):
- a conciliação automática passou de **150 para 188 despesas** e as sugestões desceram de 50 para 38;
- os dados mostraram 24 cheques registados abreviados, 6 com provável erro de digitação, 4 com nº diferente, 5 cheques que pagam várias despesas e 1 despacho possivelmente duplicado.

**Recomendação para o registo dos despachos:** escrever sempre o nº completo do cheque, tal como está impresso, por exemplo 45564417 e não 4517.

## PDFs impressos do browser e OCR

- **PDF com texto (impresso do internet banking):** é sempre lido pelo texto e nunca vai para o OCR sem o utilizador pedir. Se a tabela não for reconhecida, o site diz isso e pede que se indiquem as colunas.
- **Títulos de colunas cortados na margem** (por exemplo, o Vista Bank impresso pelo Edge, onde «CREDIT» e «BALANCE» ficam fora da página): as colunas em falta são deduzidas pelo alinhamento à direita dos valores.
- **Saldo cortado sem cêntimos** («2,306,722.»): o saldo exacto é reconstruído a partir dos débitos e créditos. Cada saldo cortado limita os cêntimos, e a intersecção de todos dá o valor. Um saldo mal lido é assinalado e não invalida os outros.
- **Contagem de movimentos:** «Results 1 - 354 of 354» é usado para confirmar que não falta nenhuma linha.
- **Conta no PDF:** a conta é lida de «Account Number : 3368» e o NIB do IBAN moçambicano (MZ + 21 algarismos).
- **OCR (PDF digitalizado ou imagem):**
  - a página é tratada como **um bloco único** (PSM 6), com os espaços preservados e as páginas a cerca de 300 dpi;
  - as confusões típicas em números são corrigidas (`]`, `|` e `l` → 1, `O` → 0);
  - o texto nunca vai para colunas de valores.
  - Medido com o extracto Vista convertido em imagem (3 páginas, 113 movimentos):

    | | Movimentos lidos |
    |---|---|
    | Antes | 0 |
    | Agora | **113, todos com data, valor e sinal certos** |

- **Melhor prática:** sempre que o banco permitir, exportar em Excel/CSV. Se for PDF, imprimir em **paisagem** ou com **«Ajustar à página»**, para nenhuma coluna ficar cortada.
