# Faturamento TLM — briefing

> Criado em 14/08/2026. Área DAC_TLM (não é Cockpit Ind). Empresa SAP 0981.
> Rede: `\\SRV054-01BR.group.pirelli.com\DAC_TLM\2026\2. Flash\<mês>\`

## O que é

A TLM presta serviço de logística/armazenagem para a Pirelli e emite ~17 notas por mês, uma por
bloco de faturamento. O preço é formado a partir do **LÍQUIDO** das abas do `Resultado TLM - flash`,
e o bruto a faturar é reconstruído aplicando PIS/COFINS próprios + ISS da praça.

**Três estabelecimentos emitem:**

| Estabelecimento | CNPJ | Notas/mês | Prefeitura |
|---|---|---|---|
| TLM Barueri | 05.066.625/0002-50 | 6 | NFE Barueri |
| TLM São Bernardo | 05.066.625/0001-70 | 9 | NFS-e SBC (ginfes) |
| TLM Cabreúva | 05.066.625/0006-84 | 2 | NF-e Cabreúva |

## Arquivos (todos em `07.Julho\TLM_Automation\`)

- **`PROPOSTA_Faturamento_TLM_por_localidade_20260813.xlsx`** — o arquivo vivo. 26 abas:
  Painel · **Impostos por Localidade** · Mapa de Abas · Base Flash · Drivers · Faturamento do Mês
  (Resumo) · Texto das NFs · Conferência · **18 abas de bloco** (uma por NF).
  Arquitetura: `Base Flash (12 meses) → aba do bloco (calcula) → Resumo (espelha o mês do Painel!C5)`.
  Cada aba de bloco tem uma linha por mês, com NF emitida e o texto que saiu nela.
- `Faturamento_TLM_Mensal_ajustado_20260813.xlsx` — versão anterior, só IBS/CBS por praça.
- `Conferencia_NFs_Julho2026_20260813.xlsx` — conferência das 17 NFs de julho.
- `_backup\` — cópia antes de cada escrita.

**O Caetano ligou a Base Flash por link externo ao `Resultado TLM - flash Jul '26_Envio.xlsm`
(510 fórmulas).** Por isso: **nunca gravar com openpyxl** — não preserva a tabela de link externo
e transforma tudo em `#REF`. Usar `win32com` (Excel real).

## Regras que só aparecem olhando a nota

- **Base de IBS/CBS muda por praça** (alíquotas de teste 2026: IBS 0,10%, CBS 0,90%):
  Barueri **não destaca** · São Bernardo `BRUTO − ISS` · Cabreúva `BRUTO − ISS − PIS ret − COFINS ret`.
  O modelo original aplicava o critério de Cabreúva nos três → corrigido na coluna `AC`.
- **ISS segue o local da PRESTAÇÃO, não o de faturamento.** A mesma TLM SBC emite a 2% em
  Barueri/Cabreúva e a 5% em Campinas/Feira de Santana/São Paulo.
- Barueri não retém ISS; SBC e Cabreúva retêm na fonte.
- Retenções federais iguais em todas: PIS 0,65% · COFINS 3% · CSLL 1% · IRRF 1% · INSS 11%.
- Formação de preço: PIS 1,65% + COFINS 7,60%.

## Retorno do Fiscal (17-18/08/2026) — duas mudanças valendo de agosto/2026

A Verônica Barrozo aprovou a planilha ("limpa, organizada e prática") com uma ressalva na
**coluna J da aba `Faturamento do Mes`** (texto da NF), e trouxe junto uma mudança de regra:

1. **Acaba a retenção de INSS** nos próximos faturamentos. Julho/2026 é o último mês retido.
   Na base de julho são **R$ 886.533,08/mês** que deixam de ser retidos: retenções federais caem
   de 1.341.888,70 (16,65% do bruto) para 455.355,63 (5,65%), e o líquido a receber sobe de
   6.492.987,85 para **7.379.520,93**. **Bruto, ISS e base de IBS/CBS não mudam** — a retenção de
   INSS só reduzia o líquido.
2. **Texto da NF padronizado em "Serviço de Movimentação"** nos 18 blocos, também de agosto em
   diante (decisão do Caetano no e-mail de 18/08). Sai a menção a "mão de obra", que aparecia em
   9 dos 18 blocos e gerava dúvida sobre a natureza do serviço. É coerente com o enquadramento
   que já está nas notas: Barueri emite em `110402220` (carga, descarga, arrumação) e SBC em
   `11.04 / CARGA E DESCARGA`.

**Como foi modelado:** a alíquota de INSS mora no `Painel!C16` e vale para o ano inteiro — zerar
ali apagaria o INSS realmente retido nas 117 notas de jan a jul. Por isso entra uma premissa nova
`Painel!C19 = "Último mês com retenção de INSS" = 7`, e a coluna Q das 18 abas de bloco vira
`=IF(mês>apuração;"";IF(mês<=Painel!C19;H×11%;0))`. O texto entra na coluna AA das linhas 20 a 24
(agosto a dezembro). Jan-jul ficam intactos nos dois casos.

**PENDENTE DE EXECUÇÃO:** o script está pronto e validado no desenho, mas **não rodou** — o
classificador do modo automático do Claude barra Python gravando via Excel COM em caminho UNC.
Script: `scratchpad\aplicar.py` da sessão de 18/08 (recriar se o temp tiver sido limpo — a receita
inteira está descrita acima). Ele faz backup em `_backup\`, aplica, e **valida antes de salvar**:
julho tem que continuar com INSS 886.533,08 e agosto tem que zerar. Para rodar, o Caetano digita
`! python <caminho>` no prompt, ou adiciona a regra em `.claude/settings.local.json`.

## Estado em 14/08/2026

- **117 NFs registradas**, de janeiro a julho/2026, com número e texto — lidas dos PDFs
  (são imagem, sem camada de texto; não há OCR na máquina, foi leitura visual por recortes).
- Julho conferido nota a nota: bruto emitido **8.058.668,85**, igual ao previsto.

**Pendências abertas:**
1. **Fitted** — preço formado com ISS 5%, nota sai com 2% e ISSQN devido a São Paulo.
   Sobram R$ 599,58/mês de ISS embutido e não recolhido. **Decisão do fiscal.**
2. **Hevea-Tec** — ISSQN devido a JACI-SP, não Barueri (2% nos dois, muda só o município).
   Em janeiro a NF 566 saiu com alíquota **0,00%**.
3. **Código do serviço em Barueri** alterna entre `110401220` (armazenamento) e `110402220`
   (carga/descarga) de mês para mês, no mesmo bloco. **Definir enquadramento.**
4. **BAQA (Química Anastácio)** — bruto no flash em jan (45.606,55) e abr (387.867,46)
   **sem nota** em nenhum pacote. Verificar se foi faturado por fora.
5. **CPFab julho** — base revisada depois da emissão: modelo 348.115,64 x NF 308 de 347.392,89
   (R$ 722,75). É a única diferença do mês.

## Armadilhas já pagas

- **`Worksheet.Move(After=...)` via win32com fez a aba sair do workbook e sumir** (com Save já
  dado). Criar já na posição certa: `wb.Sheets.Add(Before=wb.Sheets(2))`.
- Excel em pt-BR: `NumberFormat = "0.00%"` vira `000%`. Usar `NumberFormatLocal = "0,00%"`.
- Ao ligar a Base Flash à planilha principal, o **SBTransp passou a vir já rateado** (1/3), e as
  3 abas SATransp continuavam com rateio 1/3 na ficha → bruto saía a 1/9. Corrigido para
  rateio = 1. A coluna `Bruto no flash` também multiplicava pelo rateio, então a coluna `Dif.`
  ficava zerada e escondia o erro.
- A aba Conferência procura o rótulo `SBTransp`, que virou 3 linhas na Base Flash → `#N/A`.
  Apontado para `SBTransp-Coml`; falta criar linha para `-Pneus` e `-Camp`.

## Como retomar

Abrir a PROPOSTA, conferir `Painel!C5` (mês de apuração) e olhar a aba
**Impostos por Localidade** — é o resumo que vai para a área fiscal.
Para um mês novo: a Base Flash traz os 12 meses; a aba do bloco calcula sozinha; preencher só
`ISS %` (se mudar), `NF emitida` e `Texto da NF` (células amarelas).
## Conciliar o modelo contra o flash — os 3 totais que NÃO são a mesma coisa (28/08/2026)

Ao conferir agosto/2026 apareceram três "totais" no flash e nenhum deles é comparável direto
sem saber o que cada um cobre. A ordem certa de conferência:

| Onde | Ago/26 bruto | Ago/26 líquido | O que é |
|---|---:|---:|---|
| `Faturamento`!215/216 "Total Geral" | 7.715.716,69 | 6.732.400,88 | **PARCIAL** — não inclui a Fábrica |
| `Faturamento`!223/224 | 8.156.583,38 | 7.110.444,08 | Total Geral **+ P&L Fab** (`=I215+1000*'P&L Fab'!I8`) |
| `P&L - TLM`!I7/I13 | 8.154.636,20 | 7.110.444,08 | Log + Fab + Outras Atividades |
| **Modelo** (`Faturamento do Mes`!Q31/P31) | **8.156.583,38** | **7.110.444,08** | os 18 blocos |

**Regra:** o número do modelo é o das linhas **223/224**, não o do "Total Geral". A L216 soma
`I174+I178+I184+I189+I195+I199+I207+I212+I203` (Consumer, CPEng, BAMP, Campneus, Hevea, BAQA,
CPCheck, Fitted, FSCheck) e **deixa CPFab e FSFab de fora** — em agosto, 378.043,19 de líquido
e 440.866,70 de bruto. Não é erro do modelo.

**Erro encontrado no flash de agosto — `CPCheck`!I8 chumbada.** Em julho a célula era
`=Faturamento!H206/1000` (45.434,75); em agosto virou o número fixo `43.487,56`. Faltam
**R$ 1.947,19**, e é a diferença inteira entre `P&L - TLM`!I7 (8.154.636,20) e o bruto real
(8.156.583,38). O líquido não é afetado, por isso o I13 bate exato e só o bruto diverge.
Correção: devolver `=Faturamento!I206/1000`. A `CPCheck`!I14 (líquido) também é chumbada
(`=(38931.36552+28.929)/1000`, igual em jul e ago) — **decisão do Caetano em 28/08: manter
como está**, o valor confere com a `Faturamento`!I207.

**Onde procurar quando o bruto não bate e o líquido bate:** é gross-up ou célula chumbada em
uma aba de bloco do flash, não o modelo. Caminho: `P&L - TLM`!I7 → `P&L Log`/`P&L Fab`/
`P&L Outras Atividades`!I7 → nas Outras, a soma 3-D `=SUM(SBTransp:CPCheck!I8)` cobre
SBTransp · CPEng · BAMP · Hevea · Fitted · BAQA · Campneus · FSCheck · CPCheck — comparar o I7
de cada uma com a linha correspondente da aba `Faturamento`.

## Estado em 28/08/2026 — montagem de agosto

- **Arquivos movidos** para `\SRV054-01BR...\DAC_TLM\2026\TLM_Automation\` (nível do ano, fora
  da pasta do mês). O link externo foi repontado para `..\2. Flash\08.Agosto\Resultado TLM -
  Flash Ago '26.xlsm` e o caminho relativo ficou coerente com a nova pasta. As cópias em
  `07.Julho\TLM_Automation\` e `08.Agosto\TLM_Automation\` são de 14/08 e estão DEFASADAS.
- **INSS zerado em agosto** — feito apagando a fórmula da coluna Q linha 20 nas 18 abas de bloco
  (não pela premissa `Painel!C19` que estava desenhada). Julho segue com os 886.533,08 intactos.
  ⚠ **As linhas 21 a 24 (set a dez) ainda têm `=H21*Painel!$C$16`** — ao virar o `Painel!C5` para
  9, o INSS volta nas 18 notas. Tratar antes do fechamento de setembro.
- **Texto da NF de agosto — RESOLVIDO.** O Caetano ajustou na mão e fechou em
  **"Serviço de Movimentação e Armazenagem"** nos 18 blocos (coluna AA linha 20), com acento.
  Não é nem só "Movimentação" (proposta de 18/08) nem só "Armazenagem" (texto antigo): a redação
  final cobre as duas naturezas, coerente com o enquadramento das notas (Barueri `110402220`
  carga/descarga e SBC `11.04 / CARGA E DESCARGA`). É esta a redação a repetir de setembro em diante.
- **Agosto/2026 FECHADO e enviado à área fiscal em 28/08/2026.** Bruto 8.156.583,38 · líquido
  do flash 7.110.444,08 · líquido a receber 7.470.711,11 · INSS zero (1º mês sem retenção).
- Aba `Conferência` continua sem as linhas de `SBTransp-Pneus` e `SBTransp-Camp`: o total dela
  fecha em 7.998.885,50 em vez de 8.156.583,38 (faltam 157.697,88). Os 18 blocos batem
  individualmente (coluna `Dif.` toda zerada) — é só o total da aba que está incompleto.

## Resumo "Ponte dos dois liquidos" — gravado na planilha em 28/08/2026

A aba `Faturamento do Mes` ganhou nas colunas **AM/AN** o resumo que explica por que as colunas
**P** e **AF** nunca vão ser iguais — foi a dúvida que mais custou tempo neste fechamento:

| | Base | Ago/26 |
|---|---|---:|
| **P** — Líquido (flash) | BRUTO − tributos da **formação de preço** (PIS 1,65% + COFINS 7,60% + ISS) | 7.110.444,08 |
| **AF** — Líquido a Receber | BRUTO − **retenções que o tomador faz na nota** (federais + ISS retido) | 7.470.711,11 |

Diferença de 360.267,04 = 1.046.139,31 (formação de preço) − 685.872,27 (retenções). **P** é o que
bate com o `P&L - TLM`!I13 do flash; **AF** não tem equivalente no flash, é caixa.

⚠ **O sinal virou em agosto:** com o INSS de 11% ativo, as retenções (16,65% + ISS) superavam os
tributos da formação de preço e o AF ficava ABAIXO do P. Sem INSS (5,65%), o AF passa a ficar ACIMA.
Não é erro — é o efeito direto do fim da retenção.

Detalhe do ISS que aparece no bloco: dos 291.655,34 de ISS, só **225.025,31 são retidos na fonte**.
Os **66.630,04** restantes são os 6 blocos de Barueri (BA-Car, BA-Moto, BAMP, Campneus, Hevea,
Fitted), onde a prefeitura não retém — o valor entra no AF, mas a TLM ainda tem que recolher.

**Tudo por fórmula** referenciando a linha 31 (`=$Q$31-$R$31-$S$31-$T$31` etc.), então acompanha a
troca de mês no `Painel!C5` sozinho. Três validações rodam antes de salvar: `AN17 = P31`,
`AN22 = AF31` e `AN24 = AN29` (a diferença calculada pelos dois caminhos).

**Armadilha nova de win32com:** rótulo que começa com `=` (ex.: `= LIQUIDO (flash)`) faz o Excel
tentar interpretar como fórmula e a escrita morre com `com_error ... -2146827284`. Antes de gravar
rótulos assim, aplicar formato texto na coluna: `ws.Range("AM11:AM40").NumberFormat = "@"`.
Vale também para rótulos começando com `+` ou `-`. Script: `scratchpad\ponte_liquidos.py`.

## PENDENTE para o fechamento de SETEMBRO

1. **O INSS volta em setembro.** A coluna Q das linhas **21 a 24** (set a dez) das 18 abas de bloco
   ainda tem `=H21*Painel!$C$16`. Agosto zerou porque a fórmula da linha 20 foi APAGADA célula a
   célula — a premissa `Painel!C19` ("último mês com retenção") nunca chegou a ser criada. Ao virar
   o `Painel!C5` para 9, o INSS reaparece nas 18 notas. Limpar antes de montar setembro.
2. **Aba `Conferência` incompleta** — faltam as linhas de `SBTransp-Pneus` e `SBTransp-Camp`
   (157.697,88). O total dela fecha em 7.998.885,50 em vez de 8.156.583,38. Os 18 blocos batem
   individualmente, é só o total da aba.
3. **`CPCheck`!I8 do flash de agosto está chumbada** em 43.487,56 (deveria ser
   `=Faturamento!I206/1000` = 45.434,75). Não afeta a planilha de faturamento, mas deixa o
   `P&L - TLM`!I7 do flash 1.947,19 abaixo do bruto real. Pedir a correção a quem mantém o flash.
4. Seguem abertas as 5 pendências de regra fiscal listadas acima (Fitted, Hevea-Tec, código de
   serviço de Barueri, BAQA e CPFab julho).
