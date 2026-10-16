# Resultado TLM (P&L) — briefing

> Criado em 14/08/2026, a partir da explicação do Caetano + leitura das macros e dos links do
> `Resultado TLM - Act Jul '26.xlsm`. Área DAC_TLM, empresa/área de contabilidade de custos **0981**.
> Rede: `\\SRV054-01BR.group.pirelli.com\DAC_TLM\<ano>\`
> Irmão do projeto **Faturamento TLM** (`briefings/FATURAMENTO_TLM.md`), que trata das NFs.

## O que é

O P&L mensal da TLM. Cada mês nasce **duas vezes** — **Flash** (prévia) e **Act** (definitivo) — e o
ano tem ainda os **rounds de forecast R0x**. Uma pasta por versão:

```
\DAC_TLM\2026\1. FCST\R07\            ← forecast round 7
\DAC_TLM\2026\2. Flash\07.Julho\      ← Flash de julho
\DAC_TLM\2026\3. Actual\Act_07.Julho\ ← Act de julho
\DAC_TLM\2026\Bases\                  ← bases intermediárias baixadas
\DAC_TLM\2026\KSB1\                   ← KSB1 do mês
```

## O arquivo vivo e o painel de controle

`Resultado TLM - <Versão> <Mês> '26.xlsm` — 33 abas, **5 módulos VBA com código** (os 33 `SheetN.cls`
estão vazios). Toda macro lê a aba **`Parametros`**, cada uma numa **linha fixa**, em laço
`While Cells(i,4) <> ""` (dá para empilhar execuções em linhas seguidas).

| Macro | Linha | O que faz |
|---|---|---|
| `Module1.ATUALIZAR_INTERMEDIARIA` | **L3** | **YKCCP013_LT** (0981 · ano · versão · ex. base), `R_EFET` + `radALL` → recalcula a base intermediária no SAP |
| `Module2.Rodar_Mdo` — 4 subs | **L6/L9/L12/L15** | 1º **YKCCA003_LT** (0981, MM.AAAA) · 2º **YKCCA004_LT** (gera batch input) · 3º **SM35** pasta `HRCO_0981` em modo Erro · 4º **KSU5** rateio (período, ano, ciclo `HR<MMAA>`, ex. `HR0726`), em batch e depois com lista detalhada |
| `Module3.Rodar_Intermediaria_Flash` — 2 subs | **L18** | **YKCCP014_LT** (set `gestoriais`) duas vezes: `R_FIX` e `R_VAR` → salva o par na pasta `Bases` |
| `Module4.Base_Desepesas_KSB1` | **L21** | OKKS **0981** → **KSB1** (data de/até, layout **`/XBD_TLM`**) → exporta para a pasta `KSB1`, fecha o wb e carimba "OK" |
| `Module5.Select_Sheets` | L24 | Só agrupa as 17 abas de bloco (gravador de macro) |

## Ordem real de execução (Flash e Act) — declarada pelo Caetano

```
[gate] OK da contabilidade
  1. MDO 1º passo  (YKCCA003_LT)
  2. MDO 2º passo  (YKCCA004_LT)
  3. MDO 3º passo  (SM35 HRCO_0981)
  4. MDO 4º passo  (KSU5, ciclo HR<MMAA>)
       ↑ UM DE CADA VEZ — espera o anterior acabar. Nunca pelo `Rodar_Mdo`,
         que dispara os 4 seguidos (o 3º depende do batch input do 2º).
  5. ATUALIZAR_INTERMEDIARIA (YKCCP013_LT) — "para pegar os valores atuais":
     recalcula com o que a contabilidade lançou + a MDO já rateada.
  6. Baixar a intermediária (YKCCP014_LT) com o NOME DA VERSÃO que está sendo feita.
  7. Baixar o KSB1 — serve de CHECK: o dado da intermediária tem de ser o mesmo do KSB1.
```

O confronto do passo 7 já existe montado: abas **`Interm.vsKSB1`** e `Interm.vsKSB1 (2)` no
`Base Total_Act_Flash_Fct.xlsx` e **`Conciliação Mdo`** no `Base_KSB1.xlsx`.

## Como o forecast (R0x) é montado

Insumos na pasta `1. FCST\R<nn>\`:
- **toneladas** — área técnica (`<data> - Volume Movimentado R07.xlsx`)
- **despesas** — coletadas por área no `<AAAAMM> - R07 TLM.xlsm` (28 macros `AtualizarCaixas_*`:
  HSE · RH · Lean · Manutenção · Qualidade · Planejamento · Transporte · PMO · HC — só marcam caixas
  na aba `Dados PMO`, nenhuma toca o SAP)
- **valores e HC** — RH (`HC R07.xlsx`, `Layoff R7.xlsx`, `Labour Cost TLM - R07 2026.xlsx`)

Tudo desemboca em **`Carga FCST R<nn>.xlsx`**, layout
`EMPRESA | ANO | VERSÃO | CENTRO DE CUSTO | CLASSE DE CUSTO | DESCRIÇÃO | VAR. | ORIG. | JAN…DEZ`.

**A carga sobe pela transação `YKCCP003_LT` — "Carga base dados PDG - centro custos"**, tela
**"Carga Excel em base intermediária (dados por centro custo)"**:

```
Tipo de Carga:   Empresa · Ano · Versão        (ex.: 0981 · 2026 · F07)
                 ( ) TOTAL   ( ) Fixo   ( ) Variável
Dados do arquivo:( ) Re/Carga Versão   ( ) Atualizar dados Versão
                 Selecionar Mês De [1] Até [12]
                 Caminho do Arquivo: C:\tempsap\
Layout exigido:  Empresa/Ano/Versão/CentCusto/ClasseCusto/Nome/Variabilidade/Origem/YMES01/YMES0…
```

O layout da tela é **exatamente** o cabeçalho do `Carga FCST R<nn>.xlsx`
(`EMPRESA|ANO|VERSÃO|CENTRO DE CUSTO|CLASSE DE CUSTO|DESCRIÇÃO|VAR.|ORIG.|JAN…DEZ`).
**Regra da versão: `F<nn>` corresponde ao round `R<nn>`** — F07↔R07, F09↔R09, F11↔R11. A versão vai
numa **coluna do próprio Excel** carregado, além do campo da tela.
**O arquivo é lido de caminho LOCAL (`C:\tempsap\`)**, não da rede.

> Esta é a ÚNICA escrita do fluxo. As 5 macros do Resultado só recalculam, baixam e conciliam.
> Qualquer automação daqui em diante passa pelo gate de escrita em SAP (CLAUDE.md §6.1/§6.5).

## Convenção de nomes das bases

`Base_interm_<Cenário>_<MM.AAAA>_<Fix|Var>.XLSX`, **sempre em par Fix+Var**, em `\2026\Bases`.
Cenários vistos: `Flash`, `Actual`, `Fcst_R02…R07`, `MP2026`. KSB1: `KSB1_<MM.AAAA>.XLSX`.
O nome é **digitado à mão** nas células H18/I18 do `Parametros` — daí os órfãos
`Base_interm_FlasX_02.2026` e `Base_interm_Fcxt_R03V2_03.2026`, gerados com nome torto.

## Arquitetura do workbook (como o dado vira P&L)

**A macro não alimenta a planilha** — quem alimenta são os **24 links externos**. Os dois que
sustentam o resultado:
- `[2] Base Total_Act_Flash_Fct.xlsx` → tabela **`Intermediaria`** (coluna por mês `[July]`; chaves
  `WH` · `Variabilidade` · `Desc.Resultado` · `Cenario` · `Mês/Ano`)
- `[24] Base_KSB1.xlsx` → tabela **`KSB1_`** (chaves `Ano/Mês` · `Variabilidade` · `WH` · `Desc.Resultado`)

**A chave de tudo é a linha 6** de cada aba de bloco (`='P&L - TLM'!H6`):

```
=IF(H$6="Flash_KSB1", -SUMIFS(KSB1_…)/1000, -SUMIFS(Intermediaria[July]…)/1000)
```

Mesma linha de resultado, **duas fontes possíveis, e o cenário do mês decide qual vale**.
Em julho/26: B6:H6 = `Actual` (jan-jul) e I6:J6 = `Fcst_R07` (ago em diante).
`Parametros!J3` manda o rótulo do cenário para `P&L - TLM!B1`.

Cascata:
```
17 abas de WH (BAADM · BALog · SBLog · CPLog · CBLog · GVLog · FSLog · CB_AG ·
CPFab · FSFab · SBTransp · CPEng · BAMP · Hevea · Fitted · BAQA · Campneus · FSCheck · CPCheck)
   ↓ 1.994 SUMIFS cada, layout idêntico linha a linha
P&L Log · P&L Fab · P&L Outras Atividades   (+ CBTotal · BATotal)
   ↓
P&L - TLM → Resumo · Resumo de Resultado · Resumo Resultado · Justificativa
Faturamento (4.848 fórmulas; alimenta as linhas 8-12 dos blocos) · Other Income · Delta Armazém
```

## Fragilidades já visíveis (não julgadas ainda com o Caetano)

1. **Nome de arquivo digitado à mão** → dois arquivos órfãos já existem na pasta `Bases`.
2. **HC lê fontes diferentes por coluna**: a linha 19 do bloco puxa ora `[21] HC_R06`, ora
   `[22] Headcount TLM - 2026.05`, ora `[23] Headcount TLM - 2026.07`.
3. **Nenhuma macro confere resultado**: não valida sessão SAP, não confirma que o arquivo saiu
   (só o KSB1 carimba "OK"), não trata erro.
4. Muitos `definedNames` herdados apontando para `#REF!`; `Carga FCST R07.xlsx` tem abas de
   round anterior, uma delas (`(2)`) inteira em `#REF!`.

## Pendências / a confirmar

> Sessão de 14/08/2026 interrompida aqui (o Caetano precisou sair). **Retomar por estas perguntas** —
> o mapa do fluxo está completo; falta a operação fina.

Sobre a carga (YKCCP003_LT) — é a única ESCRITA do fluxo:
- `TOTAL` × `Fixo` × `Variável`: carrega em qual? (a intermediária depois é *baixada* em Fix e Var)
- `Re/Carga Versão` × `Atualizar dados Versão`: qual usa, e a Re/Carga apaga a versão inteira antes?
- `Mês De 1 Até 12`: ano cheio sempre, ou limita ao range do round (ex.: 7-12, preservando o realizado)?
- Empresa: sempre `0981`? (o print da tela estava com `0580`)

Sobre o resto:
- No **Flash**, a `Versão` das L3/L18 muda de `ACT` para quê? E no **FCT**, é `F<nn>`?
- Como o Caetano sabe que um passo da MDO acabou (mensagem da transação, SM37, log da SM35)?
- O que é, na prática, o "OK da contabilidade" (e-mail, aviso, conferência no SAP)?
- O que ele faz quando **Intermediária × KSB1 não bate** — causa típica e tratamento.
- O que muda entre a rodada Flash e a rodada Act além do dado estar mais completo.