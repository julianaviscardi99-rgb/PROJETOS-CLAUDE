# 2026-10-19 — MP27 TLM: toneladas do Faturamento ligadas ao Volume Movimentado MP27

- Volume da técnica chegou: `MP27\1ª Versão MP27 vs R09\Volume\20261016 - Volume Movimentado MP27.xlsm` (aba `03. Fluxos_MP27`, linhas `<WH><BU>TOTALTOTALMP27`, tons nas colunas U:AF).
- "Tons Mov" do Resultado = IN + OUT (linha TOTAL) — mesma lógica do R10 (`'10. Fluxos_R10'!L14/L45/L63...`).
- Novo `Faturamento MP27 - base R09 - 20261019.xlsx`: 5 linhas Tons (BA Car/Moto, CP Car/Moto, FS Car; S:AD) = fórmula com link externo / 1000 (células azuis). Script `scripts/tlm/mp27_faturamento_link_volume.py` (origem 15/10 intocada).
- Tons MP27 (k t) vs R09: BA Car 29,8 (28,8) · BA Moto 7,1 (7,2) · CP Car 134,4 (122,4) · CP Moto 31,3 (29,9) · FS 82,7 (82,7). Moto CAM (31 t em BA) não entra.
- **Insight:** com tarifa a custo, tons NÃO mexem na receita variável (tarifa = custo / tons). Total TLM MP27 continua 53,3 mi (vs R09 92,8) porque a aba Custos é cópia estática do Resultado MP27 quebrado (MO Pyplan fora; custo variável CPLog/FSLog = 0 → tarifa var. 0). Resolver o Resultado antes de confiar no faturamento.
- Ainda simulação (laranja): Ecolabel e Outras Atividades; m² (Area_m2 = dez/R09).
