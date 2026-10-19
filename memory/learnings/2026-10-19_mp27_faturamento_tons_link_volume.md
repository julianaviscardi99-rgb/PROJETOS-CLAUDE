# 2026-10-19 — MP27 TLM: toneladas do Faturamento ligadas ao Volume Movimentado MP27

- Volume da técnica chegou: `MP27\1ª Versão MP27 vs R09\Volume\20261016 - Volume Movimentado MP27.xlsm` (aba `03. Fluxos_MP27`, linhas `<WH><BU>TOTALTOTALMP27`, tons nas colunas U:AF).
- "Tons Mov" do Resultado = IN + OUT (linha TOTAL) — mesma lógica do R10 (`'10. Fluxos_R10'!L14/L45/L63...`).
- Novo `Faturamento MP27 - base R09 - 20261019.xlsx`: 5 linhas Tons (BA Car/Moto, CP Car/Moto, FS Car; S:AD) = fórmula com link externo / 1000 (células azuis). Script `scripts/tlm/mp27_faturamento_link_volume.py` (origem 15/10 intocada).
- Tons MP27 (k t) vs R09: BA Car 29,8 (28,8) · BA Moto 7,1 (7,2) · CP Car 134,4 (122,4) · CP Moto 31,3 (29,9) · FS 82,7 (82,7). Moto CAM (31 t em BA) não entra.
- **Insight:** com tarifa a custo, tons NÃO mexem na receita variável (tarifa = custo / tons). Total TLM MP27 continua 53,3 mi (vs R09 92,8) porque a aba Custos é cópia estática do Resultado MP27 quebrado (MO Pyplan fora; custo variável CPLog/FSLog = 0 → tarifa var. 0). Resolver o Resultado antes de confiar no faturamento.
- Ainda simulação (laranja): Ecolabel e Outras Atividades; m² (Area_m2 = dez/R09).

## Rodada completa 19/10 (tarde) — nova Mão de Obra + correção do de-para
- Juliana salvou `Mão de Obra\Labour Cost TLM - MP27.xlsx` (14:04): 127 linhas mudaram, total 54,44 → 53,39 mi (−1,05 mi; provisões férias/13º horista −0,90, mensalista −0,14). Cópia em `Arquivos Recebidos` (anterior em `_Versoes_anteriores\..._14.10.xlsx`).
- `mp27_resultado.montar_base`: conta fiscal fora do mapa → Desc.Resultado da base 2026 pela conta GESTORIAL (R$ 44,85 mi de MO voltam ao Resultado). `guardar_versao` (em mp27_tarifa) antes de regravar Base Total / Resultado / Resumo. `mp27_faturamento_base` não sobrescreve mais (_v2). `caderno._largura` corrigido p/ pandas 3 (astype(str) mantém NaN). `mp27_base_areas.py` direto chama Outlook: usar `python -c "import mp27_base_areas as m; m.montar()"`.
- Ordem rodada: base_areas → resultado → faturamento_base (_v2) → link_volume (_v3) → resultado → resumo.
- **Resultado MP27 (tons da técnica, Ecolabel/Outras Atividades ainda R09):** fat. líq. 90.358 k (R09 93.024) · custo var. −34.117 · fixo −56.146 · TOTAL COST −90.263 · **PBIT 95 k** (R09 3.437; MP26 189). Previdência 4247100 mantida "como veio" (decisão dela pendente).
- Resumo: 5 erros em FY/Month/Resumo de Resultado; YTD/YTG 962 cada (INDIRECT, esperado).
