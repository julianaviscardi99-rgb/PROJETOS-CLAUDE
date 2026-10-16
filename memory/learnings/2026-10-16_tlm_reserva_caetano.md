# TLM — onde o Caetano coloca a "Reserva" (16/10/2026)

- Lançada à mão na **Carga FCST** do round (`2026\1. FCST\R10\Carga FCST R10.xlsx`, aba `0981_Sem valor Zero`, linha 1367, coluna W = "Reserva").
- Chave: **CC 7071** (WH **BAQA**, Log_Barueri), classe **N15713S000** "COM outras despesas grupo" (gestorial 4236050), Fixo → Desc.Resultado **Outras Despesas**.
- Caminho: Carga → `Base_interm_Fcst_R10_10.2026_Fix` → Base Total (Intermediaria) → aba BAQA linha 105 → P&L Outras Atividades → **P&L - TLM linha 105 (Outras Despesas)**. Não existe linha "Reserva" no P&L.
- Valores (R$): R09 = 406.207 em Out/Nov/Dez (1.218.620). R10 = Nov 812.413 (Out empurrado para Nov) + Dez 1.206.207 (`=406.206,61+800.000`) = 2.018.620 → +800 k vs R09.
- Histórico no mesmo CC/classe: Actual Dez/2025 900 k + 900 k; R07 600 k em Jul.

## Uso da reserva como alavanca (pedido do Caetano, relatado pela Juliana em 16/10)
- Regra: **FY do EBIT não muda**, mas **outubro tem que mostrar PBIT de 2 mi BRL**. A reserva (linha 1367 da Carga, colunas R/S/T = Out/Nov/Dez) é o que se move entre os meses.
- R10 v2 (volume Draft): PBIT Out = −1.196,3 k (FY 4.069,5 k). Para Out = +2.000 k → Out da reserva = **−3.196.267** (crédito) e +3.196.267 em Nov/Dez.
- A carga sobe no SAP (YKCCP003_LT, versão F10) → baixar de novo a intermediária Fcst_R10 Fix/Var → Base Total → Resultado. O ajuste é o último passo do round (qualquer mudança de volume/custo muda o PBIT de outubro).
- 16/10: Juliana confirmou que o **Volume R10 Draft virou o oficial** (não precisa trocar link). O Resultado NÃO lê a Carga: só a Base Total (Power Query da pasta `2026\Bases`). Único link a uma Carga = `P&L - TLM` linha 166 → Carga R06 (memória fora do P&L).
