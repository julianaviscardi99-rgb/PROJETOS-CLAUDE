# 2026-10-06 — Comparativo de Detalhe de Despesas e auditoria de fórmulas

- "Voz" = Gestorial + Descrição Gestorial; "arquivo da MP" para despesas = **Detalhe_Despesas_Fitted Units_*.xlsx** (não o MENS FITTED). Perguntar/confirmar qual arquivo antes de montar, a usuária usa os nomes de forma curta.
- No Detalhe, o valor final está em "Total C/ Curva" (aba `DataBase_Detail`); as abas MP26 e R9 têm o mesmo layout (CQ:DB), MP27 tem layout reduzido (BC:BN). Conferir sempre contra `Resumo Custos!P4` (R$ mil).
- Arquivos na rede mudam durante o dia (usuária edita e salva): copiar para o scratchpad, anotar `LastWriteTime` e avisar se o número mudou depois (V3 foi de 42.486 para 37.886 em horas).
- Auditoria de erros: `.xlsx` → openpyxl (fórmulas + `data_only` para valores em cache); `.xls` → Excel COM em instância própria (`DispatchEx`, `UpdateLinks=0`, `ReadOnly=True`), ler `UsedRange.Formula`/`Value2` em bloco; erros vêm como inteiros negativos (-2146826265 = #REF!). Fechar só a instância própria (`xl.Quit()`).
- Ao normalizar fórmulas por linha, cuidado com falsos positivos: texto como `"N41410JA06"` contém "06" e parece referência à linha 6.
- Python via PowerShell enxerga a rede; imprimir com `PYTHONIOENCODING=utf-8` (símbolos como Δ quebram no cp1252).
