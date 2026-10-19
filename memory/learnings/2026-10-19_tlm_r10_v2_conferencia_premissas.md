# 2026-10-19 — TLM R10: conferência da v2 x premissas de 16/10

- Link [27] da v2 = `R10\20261016 - Volume Movimentado R10 - Draft.xlsm`; as 21 células de tons (Faturamento K–M, linhas 4/33/56/67/82/97/137) batem 100% com a aba `10. Fluxos_R10` do arquivo 16/10 (linhas já deslocadas: 14/63/45/90/25/71/57). Check da linha 219 (Out–Dez) = 0.
- O Volume 16/10 NÃO tem premissa de custo: a aba `01. Premissas` é idêntica à do 15/10 (só fluxos/rotas por cliente).
- Custo da v2 vem da Base Total [16] ← intermediária `Base_interm_Fcst_R10_10.2026_Fix|Var` (14/10). Nenhum arquivo de custo na rede `DAC_TLM\2026` foi alterado depois de 15/10.
- Ainda no R09 dentro do Resultado R10: Lay-off (P&L - TLM linha 126, abaixo do PBIT, entra no Operating Income) → `R09\Layoff R9.xlsx` (P17 = 171,2 k; o `R10\Layoff R10.xlsx` tem 196,0 k); HC (linha 19 das abas WH) → `R09\HC R09 TLM.xlsx` (só memória).
- PBIT Out na v2 (cache) = −1.149,7 k.

## Volume 15/10 × 16/10 (aba `10. Fluxos_R10`, por chave da coluna B)
- Só outubro mudou: esvaziamento de Cabreúva. CB>BA Car 530,0 t (42.808 pçs), CB>CP Car 47,1 t (3.800), CB>BA Moto 17,1 t (3.525) = 594,3 t saindo de CB, contadas na saída (CB) e na entrada (BA/CP) → total movimentado +1.188,5 t (23.661,4 → 24.849,8). Nov/Dez sem mudança; FS sem mudança.
- Chaves renomeadas (zeradas nas duas versões): CBCarOUTCB>GV → CB>CP; BAMotoINImport>CB>BA → CB>BA.
- Impacto na v2 (backup antes da troca × atual), outubro: Fat. bruto +52,6 k, líquido +46,6 k (Pneus +18,8 / Comercial +27,8), MC/PBIT/Operating Income +46,6 k (PBIT −1.196,3 → −1.149,7). Custo não muda (não depende de tons).

## Cópia "_ENVIO GUILHERME" (sem links) + Resumo v2 (19/10)
- `data/processed/tlm/r10/Resultado TLM - R10 '26_v2_ENVIO GUILHERME.xlsm`: cópia da v2 via Excel COM (macros desligadas) com `BreakLink` nas 27 fontes + remoção de 2 hiperlinks de pasta na aba Parametros (G18/I21 → @GMT…/Bases e /KSB1). Verificado no zip: 0 externalLinks, 0 rels externos, 0 fórmulas com [n]; 344.960 células comparadas com a v2 = 0 diferenças. Nomes definidos com #REF! (legado `\A`, `_123Graph…`) já existiam, não são links.
- Dica: depois de BreakLink, `LinkSources(1)` devolve array vazio que o PowerShell conta como 1 — não é link; conferir no zip.
- Resumo v2 (`Resumo Resultado TLM - R10 '26_v2.xlsx`): link [1] aponta para o Resultado v2 (caminho de rede) e o cache dos 23 abas lidas bate 100% com a v2 (0 diferenças em ~20 mil células). Link [2] = Base_KSB1.xlsx.

## Faturamento R10 v2 × R09 (k R$, P&L - TLM linha 13 Líquido)
- FY líquido 92.933 vs 93.024 = −91 (bruto −104). Jan–Jul 0; Ago −53 (CPFab, Pneus); Set +94 (Actual vs fcst: BA +80, CP +17, FS +20, CB +8, Fab −48, Fitted +17); Out +1; Nov −55; Dez −78.
- Tons FY −5.362 t, quase tudo Cabreúva (−5.071 t: −1.160 Out, −1.968 Nov, −2.029 Dez); BA −202, CP +379, FS −468. Líquido de CB −116 FY.
- Mix: Pneus −145, Comercial +54. Campneus/Hevea/Terceiros iguais.
