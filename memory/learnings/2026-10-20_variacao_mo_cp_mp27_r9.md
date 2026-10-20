# Variação MO Circuito Panamericano: MP27 v1/v2 x R9 x run-rate — 2026-10-20

- Fontes: `Labour Cost CP - MP27.xlsx` (v1) e `_v2.xlsx` na pasta `...\Circuito Panamericano\2027\MP'27`; R9 do CP = `...\Circuito Panamericano\2026\09 - Sep\01. Forecast\Labour Cost CP - R09 2026.xlsx`.
- A Juliana indicou a pasta Fitted (`GFU_DAC\...\09_Sep_Forecast\Labour Cost FITTED - R09 2026.xlsx`) como R9, mas ela só tem centros 8xxx da Fitted (0 linhas de CP). Usei o R9 do CP e avisei.
- CP tem um único CC (4000), categorias Gerente/Horista, 22 contas, mesmas do Fitted → mesma tabela de grupos.
- Resultado (R$ mil): R9 ano 833,9 | run-rate (set-dez x3) 982,0 | MP27 v1 1.003,1 | v2 986,9 → v2 +18,3% vs R9 ano, +0,5% vs run-rate; v2 −16,2 vs v1 (só provisão de férias gerente −11 e horista −5).
- R9 ano é baixo por entrada tardia de horistas (jul R$ 39k) e estornos; por isso o run-rate é a referência para o presidente.
- Armadilha: os xlsx de Labour têm linha de TOTAL abaixo dos dados (SAP Code vazio); filtrar `SAP Code.notna()` ou o total dobra.
- Caminho UNC com `MP'27` falha em Python/PowerShell digitado; resolver a pasta com `Get-ChildItem | Where Name -like 'MP*27'` e copiar pelo PowerShell.
- Script: `scripts/sap/circuito_panamericano/variacao_mo_cp_mp27_r9.py` (args: v1, v2, r9, saída). Saída em `data/processed/` + cópia em MP'27.
- Já existia `Delta Labour Cost CP - MP27 vs R9.xlsx` (15/10, só v1, por conta, com comentários).
