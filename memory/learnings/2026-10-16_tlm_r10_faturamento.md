# 2026-10-16 — TLM R10'26: diagnóstico do faturamento

Pasta: `\\SRV054-01BR.group.pirelli.com\DAC_TLM\2026\1. FCST\R10\` (tudo de 14/10, Caetano).
Mensalização = `Resultado TLM - R10 '26.xlsm` (mesmo esqueleto do R09: abas Faturamento, P&L - TLM, WHs).

## Estado em 16/10
- **Despesas R10 carregadas** (intermediária `2026\Bases\Base_interm_Fcst_R10_10.2026_Fix|Var.XLSX`); Out–Dez de custo ≠ R09.
- **Set = Actual** e o faturamento de Set no R10 bate com `3. Actual\Act_09.Setembro\Resultado TLM - Act Set '26.xlsm` (zero diferença na aba Faturamento).
- **Faturamento Out–Dez = idêntico ao R09** (líquido 6.987 / 6.887 / 6.836 k; tons 25.562 / 25.681 / 24.425).
  Motivo: aba Faturamento, colunas K–M:
  - Tons → `'[17]09. Fluxos_R09'!N..` = `R09\20260916 - Volume Movimentado R09 - Oficial.xlsm` (Rafaela/técnica). **Não existe Volume R10 na rede.**
  - Tarifas → colunas AP–AR (bloco MP26) menos descontos fixos digitados (ex.: Barueri Car `=AP5-45-40-15-15`, Moto `-45-40-40-15-15`, Feira `-20`).
  - Ecolabel/m²/Outras atividades → também do bloco MP26/R09 (=AP.., =AA..); Hevea Out digitado 670 t / 500 m².
- P&L R10 FY: Fat. líq. 93.065 k (R09 93.024) · custo −88.817 k (R09 −89.587) · **EBIT 4.248 k (R09 3.437)**.
- No FCST a tarifa NÃO é a custo (diferente do MP): é tarifa MP26 − descontos. Então custo novo não mexe na receita; o que falta é volume (e eventual revisão de tarifa/m²/ecolabel).
- Padrão do ciclo R09: Volume oficial da técnica + `Análise Fat. R09 vs R07.xlsx` (abas "Analise Faturamento vs Fct" / "vs MP").

## Troca do link de volume R09 → R10 (16/10)
- Volume R10 recebido: `R10\20261015 - Volume Movimentado R10 - Draft.xlsm` (aba nova `10. Fluxos_R10`; o arquivo ainda traz uma aba `09. Fluxos_R09` antiga, 374 células ≠ do R09 Oficial).
- No Resultado R10 o link do volume ([17]) só aparece na aba Faturamento: K–M das linhas 4, 33, 56, 67, 82, 97, 137 (tons Out–Dez) + linha 219 J–M (só check). Jan–Set (efetivo) vem dos arquivos de faturamento, não é afetado.
- **"Alterar fonte" sozinho NÃO serve:** as fórmulas continuariam na aba `09. Fluxos_R09` (dado velho, sem erro aparente). E a aba R10 não tem a coluna Set: tons Out/Nov/Dez = **L/M/N** (no R09 eram N/O/P). CP Moto TOTAL foi da linha 87 para a **88**. Cabreúva (linhas 24/69) = 0 no R10. As linhas de troca de NF (128/129) e o total 116 não existem mais; o total de tons agora está em L114:N114.

## v2 criada (16/10, pedido da Juliana — original do Caetano intocado)
- `R10\Resultado TLM - R10 '26_v2.xlsm` = cópia do original (salvo 14/10 15:39) + 24 fórmulas da aba Faturamento (K–M das linhas 4,33,56,67,82,97,137 e check 219) apontando para `10. Fluxos_R10` col L/M/N. Excel via DispatchEx, macros desligadas, só Calculate (links fechados ficam com valor em cache).
- Conferência: tons Out–Dez = 23.661 / 23.043 / 21.226 (= total L114:N114 do volume, check 219 = 0); Jan–Set e custos sem mudança.
- Resultado FY: Fat. líq. 93.065 → 92.886 (−179) · EBIT 4.248 → 4.070 (R09 3.437). Out/Nov/Dez líq.: 6.942 / 6.832 / 6.758.
- Pendências: Resumo R10 ainda liga no Resultado original; Cabreúva AG (linhas 106/111/119/124) e m²/tarifas seguem do R09/MP26; volume é "Draft".

## Conferência da v2 + Resumo v2 (16/10)
- Diff completo original × v2: só as 24 fórmulas trocadas (as "diferenças" de array formula do openpyxl são falso positivo — texto idêntico). Jan–Set e custos sem mudança em nenhuma aba.
- Reconciliação: Δ fat. líq. de cada unidade/mês = Δ tons × tarifa var. da própria linha, diferença 0,00; total −178.615 = Δ da linha 216.
- `R10\Resumo Resultado TLM - R10 '26_v2.xlsx`: cópia do Resumo + ChangeLink para o Resultado v2. Resumo FY: R10 92.886 / PBIT 4.070 (R09 3.437). YTD igual (51.254), YTG −179.
- Erros que JÁ existem no original (não mexi): M55/M57 #REF! (Resumo FY/Month) e 962 #REF! nas abas Resumo YTD/YTG, colunas por WH — o INDIRECT está fixo em `'[Resultado TLM - R09 ''26.xlsm]...'`, ou seja, ainda aponta para o Resultado do **R09** (sobra da virada R09→R10) e só resolve com esse arquivo aberto.
