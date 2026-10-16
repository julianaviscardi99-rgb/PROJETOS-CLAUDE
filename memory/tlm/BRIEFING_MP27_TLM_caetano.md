# MP27 / PDG27 — Master Plan 2027 da TLM (0981) — briefing

> Criado em 28/09/2026 (sessão de mapeamento, nada tocado na rede). Área DAC_TLM, empresa 0981.
> Irmão de `RESULTADO_TLM.md` (o P&L mensal + rounds R0x) — o MP é a versão ANUAL que nasce em out/nov
> para o ano seguinte e é feita para todas as empresas do grupo. Rede: `\SRV054-01BR.group.pirelli.com\DAC_TLM\MP27\`

## Como o MP26 foi montado (referência = `\DAC_TLM\MP26\`)

Duas rodadas, uma pasta por rodada: `1ª Versão MP26 vs R09` (out/25) e `2ª Versão MP26 vs R10` (nov/25).
Subpastas: `Bases 2026` · `Intermediaira` (sic) · `Mão de Obra` · `Volume`. Versão SAP do MP = **`J01`**
(ano 2026) — não é F/R como nos rounds. Carga em 3 blocos, mesmo layout da `YKCCP003_LT`:

| Bloco | Arquivo MP26 | Linhas | R$ mil | De onde vem |
|---|---|---|---|---|
| Despesas (classes N…) | `Carga Desp MP26.xlsx` | 376 | 45.738 | `20251114 - MP'26 TLM.xlsx`: aba `Base contas` (R09 por WH·conta gestorial·item·CC, com Responsável COA / Responsável MP / Observação) → aba `Base contas_Inflação 4,5%` (×1,045 uniforme) |
| Fixo RH (benefícios, N…) | `Carga Fixo RH.xlsx` | 69 | 5.541 | RH |
| Mão de obra (classes 4…) | `Carga_MDO_RH.xlsx` | 767 | 56.817 | `Pirelli - Labour Cost Report_MP26_TLM_<data>.xlsx` (RH) + `HC_TLM_MP26.xlsx` + `LayOFF MP26.xlsx` |
| **Carga total** | `Carga Custo MP26.xlsx` | 1.313 | **103.142** | união dos três (+ Trab/IFRS16/repasse) |

Fora do SAP: **toneladas** (`Volume/<data> - Volume Movimentado MPvN.xlsx`, área técnica — aba `01. Premissas`
lista os fluxos por montadora/cliente), **premissas de contrato** (`Premissas Campneus_Hevea e Materia prima.xlsx`:
tarifa R$/m², R$/t, reajustes), **faturamento/tarifa** (abas `Tarifa`, `Tarifa por B.U`, `Faturamento` do Resultado).

Saída: intermediária baixada (`Intermediaria_Fix/Var_<dd.mm>.XLSX`) → `Base Total MP26.XLSX` → `Resultado TLM - MP26 vs R09.xlsx`
(32 abas, mesmo esqueleto do Resultado mensal) → `Resumo Resultado MP26 vs R10'25 - TLM FY.xlsx` → PPT.
O `Resumo` do `20251114` explica o delta R09→MP26 por "Principais Impactos" (ex.: Ganhos AG Cabreúva -8,3 M).

**Método MP26, em uma linha:** despesas = R09 ano cheio × 1,045 ± ajustes item a item validados pelo
responsável da área; MDO = RH (não é índice); volume = técnica; IFRS16/depreciação = contrato.
MP26/R09'25 anualizado ficou em ×1,29 Other e ×1,05 RH — o índice é o piso, o grosso é ajuste de negócio.

## O que já existe para o MP27 (28/09/2026)

- `MP27\Macroeconomics 1st draft 10_09.pdf` (PIB + inflação S&P/Pirelli) e `MP27\Budget 2027 ForEx 1st draft 10_09.pdf`.
  **Brasil:** inflação média MP2027 **4,4 %** (S&P 4,5; 2026 = 4,6) · fim de período 4,3 % · PIB Pirelli 1,5 % (S&P 2,3)
  · **BRL/USD 5,25** (2026 R9 5,16 → -1,6 %) · BRL/EUR 6,14 (2026 R9 6,02).
- R09 '26 fechado em 21/09: `2026\1. FCST\R09\` — `Carga FCST R09.xlsx`, `Resultado TLM - R09 '26.xlsm`, `HC R09 TLM.xlsx`,
  `Labour Cost TLM - R09 2026.xlsx`, `Layoff R9.xlsx`, `20260916 - Volume Movimentado R09 - Oficial.xlsm`.
- Intermediária do R09 já baixada: `2026\Bases\Base_interm_Fcst_R09_09.2026_Fix|Var.XLSX` (1.183 linhas, set–dez = 33.161 k).

⚠ **Armadilhas do R09 achadas na leitura:**
1. Na `Carga FCST R09.xlsx` a aba **`R9` é a cópia de 2025** (ANO 2025, idêntica ao `2025\FCST\R09\Carga FCST R9 - 2.xlsx`).
   A carga válida de 2026 é a aba **`0981_Sem valor Zero`** (ANO 2026 · F09 · 1.355 linhas · set–dez 33.161 k = intermediária).
2. **R09 só tem set–dez.** Jan–ago do ano é Actual (cenário `Actual` na tabela `Intermediaria` do `Base Total_Act_Flash_Fct.xlsx`).
   Ano cheio 2026 = Act jan–ago + F09 set–dez — é essa a base do MP27, não set–dez × 3.
3. `VAR.` vem vazio na carga do R09 (F/V só na intermediária); o MP26 carregou F/V preenchido em 456 linhas.
4. Mini-fábricas na intermediária: 101 G2C Transitório · 102 Sto André · 103 Barueri · 104 Campinas · 106 Gravataí ·
   107 F.Santana · 108 Cabreúva · 204 Fab Campinas · 207 Fab F.Santana.

## Proposta de draft MP27 (a confirmar com o Caetano — gate de abordagem)

`Base 2026 ano cheio (Act jan–ago + F09 set–dez) por CC × classe` × **índice por conta** → `Carga Desp MP27` (versão J01, ano 2027).
Tabela de índice por conta gestorial é a decisão-chave — proposta inicial:

| Grupo de contas | Índice proposto | Fonte |
|---|---|---|
| Serviços, materiais, aluguéis de terceiros, transporte, vigilância, limpeza, alimentação | inflação 4,4 % | Macroeconomics MP2027 |
| Contratos com reajuste próprio (Campneus, Hevea, aluguel Barueri) | reajuste do contrato | `Premissas Campneus_Hevea…` |
| Itens em USD/EUR (se houver) | FX 5,25 / 6,14 | ForEx MP2027 |
| Mão de obra (classes 4…), benefícios, PLR | **não indexar** — vem do RH (Labour Cost / HC / dissídio) | RH |
| IFRS16 (N1430100RT / N4200500RT), depreciação, repasse CSC | contrato / zero | contabilidade |
| Créditos PIS/COFINS (N143110034…) | proporcional à base | derivado |

## Pendências / perguntas para o Caetano

- Confirmar "aplicar macro **por conta**": índice por conta gestorial (tabela acima) ou 4,4 % uniforme como no MP26?
- Base = ano cheio 2026 (Act+F09) ou só set–dez anualizado? (recomendo ano cheio — sazonalidade real)
- Versão SAP do MP27: `J01` / ano 2027, e a carga sobe só na 1ª rodada oficial (out) ou já no draft?
- Estrutura de pasta: replicar `MP27\1ª Versão MP27 vs R09\{Bases 2027, Intermediaria, Mão de Obra, Volume}`?
- Quem entrega HC/Labour Cost MP27 (RH) e Volume MP27 (técnica) e quando — são insumos externos, sem eles o draft é só despesas.
- Pasta `MP26` da raiz e `2026\MP26`: a segunda tem só 2 arquivos (cópia). Padrão para 2027: manter na raiz.

## Perímetro do MP27 (Caetano, 28/09/2026)

- **Sai a unidade de Cabreúva** (mini-fábrica 108, WH `CBLog` + `CB_AG`, CCs 7500–7575/7712 e 7700–7707). No R09'26 pesa
  1.706 k em set–dez (~5 % do custo), mais faturamento, HC e toneladas próprios. Excluir tudo do bloco; confirmar se o
  volume migra para Barueri/Campinas e se o fechamento é ainda em 2026 (aí a base 2026 também precisa expurgar meses).
- **Sai o faturamento de matéria-prima de Barueri** (bloco `BAMP`, CCs 7219/7220). As premissas de 2026 já zeravam o volume
  a partir de mai/26; custo residual set–dez = 83 k. Excluir receita, tarifa e custo do bloco.
- Regra gravada em `ontology/tlm_logic.json › master_plan.perimetro_mp27`.

## Entregue em 28/09/2026 — draft de despesas (1ª rodada)

Decisões do Caetano: índice **por conta** · base **ano cheio** (Act jan–ago + F09 set–dez) · versão **J01/2027** ·
**Cabreúva não existe em 2027** (CBLog + CB_AG fora, base 2026 fica intacta) · **faturamento MP Barueri (BAMP) fora**.

- Estrutura criada na rede: `MP27\1ª Versão MP27 vs R09\{Bases 2027, Intermediaria, Mão de Obra, Volume, TLM_Automation}`.
- Script: `scripts/tlm/mp27_draft.py` (roda com `python`, lê as intermediárias de `2026\Bases`, o `Centros TLM - 2026.xlsx`
  e a `Carga Custo MP26.xlsx`; grava no padrão azul via `antidumping/caderno.py`).
- Saída: `TLM_Automation\Draft MP27 Despesas - 20260928.xlsx` — abas Premissas · Indice_por_grupo · Resumo_Conta ·
  Resumo_WH · Base_2026 · Draft_MP27 · Carga_J01 · Fora_do_layout.

| Número | R$ mil |
|---|---|
| Base 2026 ano cheio (intermediária Fix+Var) | 90.483 |
| Sai do perímetro (Cabreúva 5.979 + BAMP 888) | 6.867 |
| Draft MP27 (todas as contas) | 85.190 |
| Carga_J01 (só classes que subiram no MP26: 606 linhas) | 41.384 |
| Fora_do_layout (MDO/rateio, 1.674 linhas — RH substitui) | 43.806 |

Grupos de índice: INFLACAO 4,4 % (33 contas de despesa) · RH 0 (28 contas de folha/benefício, RH substitui) ·
CONTRATO 0 (IFRS16 4257001/4255002, depreciação 4255000) · REVISAR 0 (4266000 Ajuste de Custo 2,5 M, 4200200 transitória).

**Aprendido na montagem:** a intermediária é o RESULTADO do SAP — traz classes de simulação de folha (`_HR`, `PF/PI`,
`N2…`, `N61…`) que nascem do rateio e nunca entram na YKCCP003_LT. Por isso a carga só leva classes que subiram no MP26.

**Próximos passos:** Caetano revisa `Resumo_Conta` (grupo/índice por conta e a 4266000) → pedir HC/Labour Cost MP27 ao
RH e Volume MP27 à técnica → premissas de contrato 2027 → faturamento/tarifa → montar `Resultado TLM - MP27 vs R09`.

## Entregue em 28/09/2026 — planilha de preenchimento do faturamento (pasta Volume)

Pedido do Caetano: "na pasta Volume, dados do R09 do faturamento como base e colunas para preenchimento do MP27".
- Script: `scripts/tlm/mp27_faturamento_base.py` → `MP27\1ª Versão MP27 vs R09\Volume\Faturamento MP27 - base R09 - 20260928.xlsx`.
- Fonte: `2026\1. FCST\R09\Análise Fat. R09 vs R07.xlsx`, aba `Analise Faturamento vs Fct` — por unidade (Barueri · Campinas ·
  Cabreúva · Feira de Santana) × BU (Car · Moto) × linha (Tarifa R$/t · Tons k · Valor Liq.Var · Ecolabel · Valor Liq.Fix),
  mais Outras Atividades (Comercial: SB.Transp, Campneus · Pneus: SB.Transp, R&D CB, MP BA, Campneus BA, HEVEA, Q.A, CP.Fab,
  FS.Fab, FS.Check, CP.Check, Fitted). R$ mil e k ton. Total 2026 = 92.844 (Log 76.181 + Outras 16.663) — bate com a fonte.
- Layout: base 2026 à esquerda · MP27 jan–dez à direita (amarelo = digitar; Valor Liq.Var = Tarifa × Tons por fórmula;
  totais e Delta por fórmula; Tarifa total = média ponderada Valor/Tons). Cabreúva e MP BA ficam como referência, cinza, sem entrada.
- ⚠ Parser: a fonte tem depois de Feira de Santana os blocos consolidados "Total CAR / Total Moto" sem unidade — o bloco da
  unidade fecha no "Total <unidade>", senão Feira dobra (erro pego no teste: 84.099 em vez de 7.918).
## Entregue em 28/09/2026 — Resultado TLM - MP27 vs R09 (TLM_Automation)

Pedido: "Resultado TLM com as fórmulas pegando da base de MP27; pode usar o Flash de setembro, que já tem o LF R09".
- Script: `scripts/tlm/mp27_resultado.py` (precisa de Excel — win32com; openpyxl quebraria os 24 links). Faz dois passos:
  1. `TLM_Automation\Base Total MP27.xlsx` — tabela `Intermediaria` (2.331 linhas, 85.190 k) com as colunas que os SUMIFS das abas
     de WH consomem (WH · Variabilidade · Desc.Resultado · Cenario · Mês/Ano · January..December), alimentada pela aba `Draft_MP27`.
     Desc.Resultado = VLOOKUP na aba `Parametros` do `Base Total_Act_Flash_Fct.xlsx`, senão a descrição gestorial (100 % das linhas caem
     numa linha do P&L). Cenario = Mês/Ano = `MP27`.
  2. `TLM_Automation\Resultado TLM - MP27 vs R09.xlsm` — cópia do `2. Flash\09.Setembro\Resultado TLM - Flash Set '26.xlsm`; via Excel:
     `ChangeLink` da Base Total_Act_Flash_Fct → Base Total MP27; `'P&L - TLM'!B5:M6` = MP27 (as abas de WH apontam para lá); `Parametros!J3`
     = MP27; aba **Faturamento** ligada por fórmula à `Volume\Faturamento MP27 - base R09` (Tons ×1000, Tarifa, Fixo ×1000, Ecolabel ×1000
     por unidade/BU + Outras Atividades — 25 linhas); Cabreúva (82–127) e MP BA (181/184) zerados. Comparativo vs R09 continua pelo link [24].
- Checks do arquivo gravado: rótulo MP27 · custos variáveis -27.563 · fixos -56.733 · **TOTAL COST -84.297** (= base 85.190 − Other Income 893,
  que o P&L não leva a custo, igual no Flash) · faturamento 0 até a planilha da pasta Volume ser preenchida (só R&D CB e CP Check, que são
  cost-plus). Erros restantes são derivados disso: ROS (linha 133 das WH) e `Resumo de Resultado` em #DIV/0! por faturamento zero; linha 150
  `VSMP` #VALUE! (comparação com MP26, link [3]); `Resumo Resultado` com 39 #REF! já existia no Flash.
- ⚠ Três armadilhas do Excel aprendidas: (a) referência estruturada a tabela externa (`[1]!Intermediaria[...]`) só resolve com a fonte ABERTA
  — abrir a Base antes de recalcular; (b) SUMIF/SUMIFS em pasta externa FECHADA devolve #VALUE! — o script abre todas as fontes dos links
  (25) antes do `CalculateFull`, como estavam quando o Flash foi salvo; (c) `Workbooks.Open` devolve `None` para fonte já aberta.
- HC (link [21]) e Layoff ([2]) continuam apontando para o R09 até o RH entregar o MP27.
- Passos seguintes: preencher a planilha da pasta Volume (tons da técnica, tarifas) → abrir o Resultado com a Base Total MP27 e a
  planilha abertas e recalcular → RH: Labour Cost/HC/Layoff MP27 → repontar [21]/[2].
## Entregue em 08/10/2026 — raciocínio da tarifa no faturamento + Resumo Resultado MP27

Pedido do Caetano: "no arquivo de faturamento, o raciocínio de como montamos a tarifa fixa e variável (base MP26), de onde
tiramos a despesa, comparativo contra R09, e o % de BU com uma conta feita em base à tonelada carregada" + "o mesmo
arquivo de resumo de resultado que usamos para as versões de flash/fct".
- Regra do MP26 lida nas fórmulas (`Resultado TLM - MP26 vs R09.xlsx`, abas Tarifa / Tarifa por B.U) — gravada em
  `ontology/tlm_logic.json › master_plan.tarifa_mp27`:
  **Tarifa V** = custo variável da WH (linha 21) × % BU / tons da BU × (1 + markup: Barueri 2 %, Campinas/Feira 0 %);
  **Tarifa F** = (custo fixo linha 50 − questões trabalhistas linha 83 − Ecolabel [só Barueri]) / m² do ano, igual Car e Moto;
  bruto = líquido / (1 − 11,25 % BA | 14,25 % CP/FS). % BU = coeficiente de empenho de M.O (BA 76/24, CP 50/50).
- `Volume\Faturamento MP27 - base R09 - 20261008.xlsx` (o de 28/09 fica; o Resultado passa a ler o novo): abas novas
  **Custos** (linhas 21/50/83 de BALog·BAADM·CPLog·FSLog do Resultado MP27, MP27 e R09, com a célula de origem) ·
  **Area_m2** (m² R09 por mês + MP27 amarelo = dez/R09) · **BU** (coeficiente × % pela tonelada, critério em C4, tarifa
  pelos dois) · **Tarifa** (passo a passo 1–19 por unidade) · **Comparativo_R09** (R09 praticado × R09 pela mesma conta × MP27).
  Na aba Faturamento, Tarifa e Valor Liq.Fix MP27 viraram VERDES = fórmula da aba Tarifa. Sem toneladas: 50.013 k (só fixo).
- ⚠ % pela tonelada deixa a tarifa por tonelada IGUAL em Car e Moto (Campinas Moto 247 → 75 R$/t). Padrão = coeficiente MP26;
  trocar só com decisão do Caetano (BU!C4).
- Teste com as toneladas do R09 (custo MP27): Barueri 200/255 R$/t e fixa 39,8 R$/m² · Campinas 46,8/191,8 e 28,2 · Feira 47,1 e 22,0.
  R09 praticado tem margem Barueri +1.589 k, Campinas −522 k, Feira −323 k — a tarifa a custo do MP27 zera Campinas/Feira.
- Resumo: `TLM_Automation\Resumo Resultado TLM - MP27 vs R09.xlsx` (cópia do resumo do R09, link → Resultado MP27;
  C = MP27, G = R09, I = MP26). Script `scripts/tlm/mp27_resumo.py`. Coluna Flash e linhas digitadas do R09 limpas.
- Ordem para refazer: `mp27_faturamento_base.py` → abrir/recalcular no Excel → `mp27_resultado.py` → `mp27_resumo.py`.

### Ainda 08/10/2026 — simulação com a curva do R09 e Resumo fechado
- Pedido: "como simulação, repetir as tons de R09 para 27" + "deixar os valores de 2027 carregados para o Resultado pegar".
  Toneladas, Ecolabel e Outras Atividades MP27 = R09 mês a mês (LARANJA, `SIMULAR_TONS_R09 = True`). Trocar pelo Volume MP27
  da técnica e pelas propostas das atividades quando chegarem (é só digitar por cima ou desligar a chave e rodar de novo).
- **Resultado MP27 vs R09**: faturamento líquido 86.724 k · custo -84.687 k · **PBIT 2.037 k** (R09 3.437 · MP26 189).
  Bloco de comparação MP26 (AD:AP) voltou: o Resultado do MP26 não é mais aberto no recálculo (5 erros restantes, eram 1.497).
- **Resumo Resultado TLM - MP27 vs R09.xlsx**: Resumo FY C = MP27 86.724 / PBIT 2.037, G = R09, I = MP26. YTD/YTG por WH (INDIRECT)
  só resolvem com o Resultado aberto, como no Flash/FCT. M55/M57 #REF! vêm do modelo R09.
- Excel: o `DispatchEx` subia "ocupado" depois dos taskkill (cada um vira fechamento inesperado → recuperação lenta).
  `mp27_tarifa.abrir_excel()` espera 45 s e tenta de novo, sem matar a instância presa. Erro registrado em
  `memory/errors/2026-10-08_14-20-00_contagem_erros_celula_a_celula_com.md`.

## 14/10/2026 — base de custos das áreas (Rafaela) substitui o draft por índice

E-mail "MP27 | Base custos TLM" (Rafaela, 09/10, para Caetano e Juliana): anexo `20261009 - MP'27 Custos TLM_envio.xlsx`,
copiado para `MP27\1ª Versão MP27 vs R09\Arquivos Recebidos\` (era `Bases 2027`, movido em 15/10). Aba `Base`: duas versões empilhadas — `MP27` (32,8 M) e
`MP27 Inf` (×1,044 = 34,2 M, **a que vale**). Despesas por CC final, item a item com comentário; Cabreúva já fora.
**Não traz vozes de RH nem alocadas** (ela avisou que ainda aguarda). No e-mail: FY R09 x MP27 −7,55 M (efeito CB −3,76 M ·
aluguel BA −5,04 M = some a `N4200500RT` · alimentação transcin./contratações +0,64 M · transporte interno CP...).
- Decisões do Caetano: (1) VAR. vem do R09/Actual 2026 do mesmo CC × classe (a base vem toda F e a tarifa variável depende disso);
  (2) as 3 linhas IFRS16 repetidas ficam como vieram; (3) pergunta à Rafaela — rascunho de resposta salvo no Outlook do Caetano
  (N4221000RT CC 7000 · N4221100RT e N1430600RT CC 7289; líquido −127 k, bruto dobrado).
- Script `scripts/tlm/mp27_base_areas.py` → `TLM_Automation\Draft MP27 Despesas - 20261014 - base areas.xlsx`
  (despesas não-RH do draft 39,2 M → base das áreas 34,2 M; RH 46,0 M mantido; total 80,3 M). O `mp27_resultado.py` pega esse
  como último draft. Saídas de 08/10 guardadas em `TLM_Automation\_Versoes_anteriores\20261008_draft_por_indice\`.
- Ordem para refazer: `mp27_base_areas.py` → `mp27_resultado.py` (custo novo) → `mp27_faturamento_base.py` (tarifa a custo lê o
  Resultado) → `mp27_resultado.py` de novo → `mp27_resumo.py`. Sem a 2ª volta a tarifa fica com o custo velho (PBIT saiu 7.036 errado).
- **Resultado 14/10** (tons = R09 simulado): faturamento líquido 82.189 k · custo −79.688 k · **PBIT 2.501 k** (08/10: 86.724 / −84.687 / 2.037;
  R09 3.437 · MP26 189). Faturamento do dia: `Volume\Faturamento MP27 - base R09 - 20261014.xlsx`.
- Pendente: resposta da Rafaela (duplicidade) · RH e alocadas MP27 · Volume MP27 da técnica.
- **Comparativo R09 x MP27 (despesas sem RH)** no mesmo caderno (abas Ponte_R09 · R09_x_MP_WH · R09_x_MP_Conta · R09_x_MP_WH_Conta):
  R09 FY 41.713 k (Act+F09, A+P) → −3.762 Cabreúva (= "Efeito CB" do e-mail, exato) −370 BAMP → 37.580 → variação de base −4.776
  → inflação +1.443 → MP27 34.247 (Δ −7.466; e-mail −7.554, diferença de 88 k não localizada). BAQA (Q.A., 1.235 k no R09) não vem na
  base — provável parte das "alocadas" pendentes. Serviços 4236020/4236050 caem 5,6 M (CPLog 4236020 sozinho −2,6 M) — confirmar com ela.
- **14/10 — e-mail enviado** pelo Caetano à Rafaela e à Juliana (cc Jefferson), com o `Comparativo Despesas MP27 x R09 - 20261014.xlsx`
  e 5 pontos: IFRS16 duplicado · aluguel BA −4,3 x −5,04 M · serviços −5,6 M (CPLog) · BAQA 1,2 M fora da base · 88 k de diferença.
  Pedido à Ju: ajudar no acompanhamento e na vinda de RH e alocadas. **PARADO aguardando resposta** — ao retomar, ler a resposta
  no Outlook (assunto "MP27 | Base custos TLM"), ajustar a base e rodar a sequência de refazer.

## 15/10/2026 — pasta `Arquivos Recebidos`

Tudo que a TLM manda para o MP27 fica em `MP27\1ª Versão MP27 vs R09\Arquivos Recebidos\` (decisão do Caetano):
`20261009 - MP'27 Custos TLM_envio.xlsx` (Rafaela, despesas das áreas) · `MP27_env_14.10_RH.xlsx` (Iana, RH) · `Labour Cost TLM - MP27.xlsx` (MDO por CC × conta, salvo pelo Caetano em 14/10).
O `mp27_base_areas.py` passou a ler dali. `Bases 2027` ficou vazia. `TLM_Automation` = só o que nós geramos.

## 15/10/2026 — resposta do Jefferson (RE: MP27 | Base custos TLM, 10:08) e base única

Respostas em destaque dentro do e-mail do Caetano de 14/10; anexo `20261014 - Base FY (R09).xlsx` (base R09 que usaram, abas
Pivot/Base/Base_Conta) salvo em `Arquivos Recebidos`.
1. IFRS16 repetidas: "seguimos conforme dados atuais no SAP" + print da classe N4221000RT CC 7000 (−54.276,45 duas vezes,
   +54.276,45 e +63.894,84; saldo 9.618,39). O print mostra estorno de um dos lançamentos — a repetição no MP27 continua suspeita.
2. Aluguel BA (−4,3 x −5,04 M): sem resposta.
3. Serviços CPLog −5,6 M: MdO avulsa de CP (ACT mais caro) vira contratação TLM no MP27 — o custo migra para a MDO (Labour Cost).
4. BAQA 1,2 M: é a provisão dos ganhos DAC, que o Caetano vem realizando — não vem na base delas.
5. 88 mil: FECHADO com a base dele (41.801.020 x 41.712.616) = CC 5072 N14206S002 PROV COM Benef.sociais (−74.150, está na
   intermediária e não na base deles) + CC HR_DUMMY (−14.223 líquido da simulação de folha) + ~30 de arredondamento.

Base única (`mp27_base_areas.py`, decisão Caetano 15/10): áreas (Rafaela 34,25 M) + custo fixo RH 0981 (Iana 2,94 M) + MDO
Labour Cost Pyplan (54,44 M) = 91,63 M, SOMADAS mesmo onde CC x classe repete. Colunas Fonte / Enviado por; carga J01 em
`TLM_Automation\Carga Custo MP27 - <data>.xlsx`. Aprendizes da Iana em classe de fábrica do Pneus (N14215M001/N14219M000, 696 mil)
ficam como vieram, sem gestorial, fora do Resultado até o RH corrigir. **Regra (Caetano 15/10): nada entra que não
tenha vindo das bases recebidas** — classe, valor, VAR. e ORIG. como vieram (VAR. das áreas pelo R09 é a exceção de 14/10). Regra em `ontology/tlm_logic.json › master_plan.base_unica_mp27`.
- **E-mail enviado 15/10 23:00** — "MP27 TLM | Possíveis duplicidades entre a base de custos das áreas e as bases do RH", para
  Rafaela e Jefferson (Cc Ju), com a lista (`scripts/tlm/mp27_sobreposicao.py`: 23 CC x classe, 6,81 M), a base única e os 3
  originais. **Aguardando** retorno item a item; até lá as três bases seguem somadas. Ao retomar: ler a resposta no Outlook.
