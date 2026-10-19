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

## Lay-off corrigido na v2 + auditoria R9 (19/10)
- Pedido da Juliana: o ENVIO GUILHERME era só para ele ver faturamento; faturamento fechado → não precisa mais manter esse arquivo.
- v2 (rede) ajustada: ChangeLink `R09\Layoff R9.xlsx` → `R10\Layoff R10.xlsx` (P&L - TLM K126:M126: −171,2/−286,8/−171,2 → −196,0/−286,8/−196,0). PBIT inalterado (−1.149,7 / −1.667,3 / −2.908,3); Operating Income Out −1.345,7, Nov −1.954,1, Dez −3.104,4. Backup: `data/processed/tlm/r10/Resultado TLM - R10 '26_v2_backup_antes_layoff_1910.xlsm`.
- Custos Out–Dez (TOTAL COST −8.139/−8.499/−9.666) batem exatamente com `Fcst_R10` da Base Total (14/10) → custo está em R10.
- Ainda R09/antigo: HC (linha 19 das abas WH, Out–Dez ← `HC R09 TLM.xlsx`; não existe HC R10 na pasta; informativo); P&L linha 155 (check do lay-off ← Layoff R6, linha 156 vira diferença); linha 166 ← Carga R06 (zeros). Comparativos "vs R09/MP26" (Faturamento col R+, P&L P126:R126, AD126) são intencionais.
- Pendente do round: reserva do Caetano (Out PBIT −1.149,7 vs alvo +2.000 k; ver learning 2026-10-16_tlm_reserva_caetano). Resumo v2 não reflete o novo Lay-off até ser atualizado.

## HC R10 aplicado na v2 (19/10)
- Novo arquivo `R10\20261009 - HC_R10.xlsx`: aba `Base dados` com todas as versões (MP26…09. R09, 10. R10) e só Out/Nov/Dez; sem coluna de chave por armazém. Chave da v2 (WH2) vem do CC → mapa do `HC R09 TLM.xlsx` (col V, VLOOKUP no 'Centros de Custos TLM.xls'); os 177 CCs da versão R10 mapeiam 100%.
- HC Out/Nov/Dez total: R09 505/505/505 → R10 509/489/489 (BALog +3 só Out; CPLog +2; FSLog −1; CBLog −17 Nov/Dez = phase out). Bate com a tabela "Revisões HC TLM" do arquivo.
- Na v2: K19:M19 das abas BALog, CBLog, CPLog, FSLog trocados por valores (constantes); demais abas inalteradas (método reproduz os valores atuais). Jan–Set continuam ligados ao HC R09/R07/Act Set. PBIT sem mudança (HC é informativo). Backup: `..._v2_backup_antes_hc_1910.xlsm`.

## Carga FCST R10_v2 — aluguel Barueri (19/10)
- Cópia editável: `R10\Carga FCST R10_v2.xlsx` (original intocado). Juliana digitou **R1365 = −1.400.000** (CC 7287 BALog, N14301S001 "COM aluguel armazém", linha "Passagem de custo do Aluguel para Pirelli Pneus"; só Out; Nov/Dez vazios). Diff vs original = 1 célula. Total da carga Out/Nov/Dez = 6.738,6 / 8.499,0 / 9.666,4 k (intermediária atual 8.139/8.499/9.666 → Out −1.400 exato).
- Efeito esperado (só isso): PBIT Out −1.149,7 → ~+250,3 k; FY 4.116,1 → ~5.516,1 k; linha "Aluguel Barueri" do P&L ≈ +1.389 k. Para manter FY em 4.116,1 e Out = +2.000 k é preciso reserva (CC 7071) com Out ≈ −1.749,7 k e +1.749,7 k em Nov/Dez (mais o −1.400 de FY se houver contrapartida).
- Backup das intermediárias antes da troca: `data/processed/tlm/r10/Base_interm_Fcst_R10_10.2026_Fix|Var_ANTES_aluguel.XLSX`.
- Excel dela deixado com Resultado v2 + 22 fontes abertos (read-only); NÃO salvar o Resultado v2 nesse estado.

## ERRO: carga F10 duplicou a versão no SAP (19/10 13:07)
- Intermediárias baixadas às 13:07 (`Base_interm_Fcst_R10_10.2026_Fix|Var`) vieram com o DOBRO de linhas: 793→1.586 (Fix) e 401→802 (Var), cada linha 2x — uma com usuário `JOSECA001` (carga antiga do Caetano) e outra `SILVEJU001` (carga nova da Juliana). Out/Nov/Dez dobrados (Fix Out 8.997 k, Var Out 5.880 k).
- A parte SILVEJU001 sozinha está certa: Fix Out 3.798,7 (= 5.198,7 − 1.400) + Var Out 2.939,9 = 6.738,6 = total da Carga v2; linha 7287/N14301S001 Out = −1.400.000.
- Causa provável: carga em modo "Atualizar dados Versão" (soma por usuário) em vez de "Re/Carga Versão" (substitui a versão). Lição: antes de subir carga no SAP, confirmar o modo; a chave inclui o usuário.
- NÃO atualizar a Base Total com esses arquivos (dobraria o custo). Backups das intermediárias de 14/10 em data/processed/tlm/r10/*_ANTES_aluguel.XLSX.

## Workaround 19/10 13:30: intermediárias sem duplicata (SAP ainda duplicado)
- Mesmo após Re/Carga ("1.194 incluídos, recarga substituição") e novo download (13:15/13:16), as bases vieram com 793+793 (Fix) e 401+401 (Var), `JOSECA001`+`SILVEJU001`. Causa não confirmada (download antes da substituição valer, ou intermediária precisa de recálculo — botão 1 ATUALIZAR_INTERMEDIARIA, mas Parametros!D3 está com versão 'ACT'; confirmar com o Caetano).
- Com OK da Juliana ("isso, faz isso por favor"): geradas as bases sem as linhas JOSECA001 (793 / 401 linhas, só SILVEJU001; Out/Nov/Dez 6.738,6 / 8.499,0 / 9.666,4 k = Carga v2). **AINDA NÃO COPIADAS para `2026\Bases`**: os arquivos estavam abertos (lock ~$) no Excel dela; copiar depois que ela fechar. Cópias: `data/processed/tlm/r10/dedup/` (corrigidas), `..._DUPLICADA_1316.XLSX` (como baixadas), `..._ANTES_aluguel.XLSX` (14/10).
- Próximo: atualizar a Base Total (Refresh All) e conferir EBIT (esperado PBIT Out ≈ +250 k; FY ≈ 5.516 k).
