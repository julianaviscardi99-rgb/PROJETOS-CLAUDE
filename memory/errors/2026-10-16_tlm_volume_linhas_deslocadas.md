# TLM R10 — volume novo com linhas inseridas quebra os links do Resultado (16/10/2026)

- `ChangeLink` do Volume 15/10 → 16/10 apontou para as MESMAS linhas, mas a Rafaela inseriu fluxos novos (ex.: BACarINCB>BAR10) na aba `10. Fluxos_R10` e tudo desceu 1–2 linhas.
- Sintoma: tonelagem Out–Dez −54 mil t, PBIT FY 4.070 → 90, check `Faturamento!K219:M219` ≠ 0 (−8.819). **Sempre conferir o check 219 depois de trocar o volume.**
- Correção: localizar as linhas pela chave da coluna A (`<WH><BU>TOTALTOTALR10`) e o TOTAL geral (BA/CB/CP/FS/TOTAL, t/mês). 16/10: 13→14, 24→25, 43→45, 55→57, 61→63, 69→71, 88→90, 114→116.
- Erro meu no caminho: regex em PowerShell com `$?` dentro de aspas duplas → exceção → o loop gravou fórmula vazia em 24 células. Arquivo NÃO salvo; restaurado a partir do backup (`data/processed/tlm/r10/`). Lição: em PowerShell, regex sempre em aspas simples, e nunca gravar se a transformação falhar.
