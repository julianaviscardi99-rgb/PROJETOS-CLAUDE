# TLM — Resultado dá #VALUE! (ex.: N124) até abrir outro arquivo (2026-10-19)

**Queixa da Juliana:** ao abrir "Resultado TLM v2" (ou qualquer Resultado TLM), células como a N124 ficam com #VALUE!; só funciona se ela abrir outro arquivo antes.

## O que foi verificado
- `Resultado TLM - MP27 vs R09.xlsm` (rede `DAC_TLM\MP27\1ª Versão MP27 vs R09\TLM_Automation\`) tem **27 links externos**. Um deles é `Base Total MP27.xlsx` (mesma pasta) por **referência estruturada** (nome de tabela) — só resolve com a fonte ABERTA, senão #VALUE!. `mp27_resultado.py:183` já abre essa base antes de recalcular por isso.
- Links secundários que também podem dar #VALUE!: `Resultado TLM - R09 '26.xlsm` e `Resultado TLM - MP26 vs R10.xlsm` (comentário em `mp27_resultado.py:218-219`: Resultado aberto com fontes fechadas se recalcula e devolve #VALUE! no bloco AD:AP). Outras fontes: Base_KSB1, Base Total_Act_Flash_Fct, HC R09/R07, Faturamento MP27 v3, Layoffs.
- `Resultado TLM - R10 '26_v2_ENVIO GUILHERME.xlsm` (local, data/processed/tlm/r10) **não tem externalLinks**; N124 de `P&L Outras Atividades` = `SUM(B117:M117)` (valor em cache OK). Na aba `Resumo` há 921 fórmulas com INDIRECT (nome de aba em A) — recálculo volátil na abertura. Se o erro ocorre no R10 v2, a causa é outra (recálculo/INDIRECT, ou o arquivo certo é outro).
- Já documentado antes (BRIEFING 16/10): para abrir R10 v2 sem #VALUE!, abrir antes Base Total_Act_Flash_Fct, Base_KSB1 e os 3 HC.

## Hipótese NÃO confirmada
Culpado principal = `Base Total MP27.xlsx`. Não rastreei a cadeia de precedentes do N124 célula a célula.

## Próximos passos
1. Teste: abrir `Base Total MP27.xlsx` antes do Resultado — N124 normaliza?
2. Perguntar a ela: qual arquivo abre (MP27 ou R10 v2) e se o erro é no N124 de todas as abas ou só uma.
3. Eliminar a dependência: converter o link da Base Total em valores ou trocar a referência por uma que não exija o arquivo aberto.
