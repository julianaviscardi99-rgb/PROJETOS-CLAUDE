# 2026-10-20 (noite) — TLM MP27: depreciação, aba Custos v2 e tarifas desalinhadas

## 1. "Melhora" de R$ 5 mi em Depreciação = custo faltando (armazém Barueri CC 7287)
- P&L - TLM linha 106: MP27 −244 x R09 −5.253. 100% no BALog fixo (−124 x −5.170).
- R09 tem N4200500RT "Amortização imóveis alugados (IFRS16)" CC 7287, ~R$ 840–850 mil/mês jan–jun (Actual) e
  o ganho do término antecipado do contrato em jul (M2306700RT −1.490, Outras Despesas).
- Armazém segue em uso: set/26 tem provisão de aluguel 975 (N14301S056, mapeada errada em "Loc. e Manut. de
  Empilhadeira") e o estorno IFRS16 N1430100RT −975 continua → líquido 0.
- MP27 (Rafaela, MP27 Inf) e Fcst R09 set–dez NÃO trazem aluguel nem amortização do CC 7287 → faltam ~R$ 975 mil/mês
  (~R$ 11,7 mi/ano) ou ~0,8–0,85 mi/mês se o contrato novo for IFRS16. PBIT MP27 superestimado.
- Pendente: perguntar à Rafaela/Jefferson sobre contrato novo, valor, IFRS16 ou não. Nada alterado na base.
  Detalhe: `2026-10-20_tlm_mp27_depreciacao_aluguel_barueri_7287.md`.

## 2. Aba Custos da planilha de preenchimento atualizada para o Resultado v2
- Ela salvou `Resultado TLM - MP27 vs R09_v2.xlsm` (19:00). A aba `Custos` de
  `Volume\Faturamento MP27 - base R09 - 20261020_v2.xlsx` guarda VALORES COPIADOS (não fórmulas), origem na coluna R.
- Reescrevi as 24 linhas via COM no Excel aberto (GetActiveObject; GetObject com caminho de rede dá "Invalid syntax").
  Só 4 mudaram: BAADM fixo −8.244,6→−8.025,4; CPLog fixo −17.231,9→−17.675,9; FSLog fixo −4.343,6→−4.377,8; BALog var −7.824,3→−7.823,3.
- Ela salvou às 19:11 e 19:44. `mp27_faturamento_base.py` (RESULTADO_MP) ainda aponta para o Resultado SEM _v2:
  se rodar, a aba Custos volta ao valor antigo.

## 3. Tarifas: de/para planilha x Resultado v2
- Planilha (aba Tarifa) antes → agora: var BA Car 208,54→208,52; BA Moto 270,40→270,36; fixa BA 35,86→35,49;
  fixa CP 24,34→24,97; fixa FS 21,79→21,96; var CP e FS sem mudança.
- Faturamento fixo MP27 ano (R$ mil): BA Car 19.144→18.947; BA Moto 2.151→2.129; CP Car 11.180→11.468;
  CP Moto 6.052→6.208; FS 4.344→4.378. O Resultado aberto já tem esses valores e as tarifas VARIÁVEIS novas.
- NÃO alinhado: na aba Faturamento do Resultado, linhas de m² e Tarifa fixa MP27 (9/10, 38/39, 61/62, 72/73, 142/143) são
  números digitados do R09 (=AG10, =R62...): fixa BA 55,42, CP 22,93, FS 24,27 (planilha 35,49 / 24,97 / 21,96).
  m² só difere em CP Car (42.680 jan–jun no Resultado x 38.280 na planilha). `LIGACOES_FATURAMENTO` em `mp27_resultado.py`
  não liga m² nem tarifa fixa.
- P&L do Resultado aberto: fat. líq. FY 88.973 e PBIT 3.301 (antes 86.980 / 1.308). Planilha nova explica pouco
  (BA −200, CP/FS +330 no fixo); o resto são mudanças dela no arquivo (não rastreadas).
- Proposta pendente: ligar m² (Area_m2) e tarifa fixa (aba Tarifa) no Resultado v2 e em LIGACOES_FATURAMENTO, com ela
  salvando e fechando o Resultado antes.
