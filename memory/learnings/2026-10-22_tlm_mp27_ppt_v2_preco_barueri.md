# 2026-10-22 - TLM MP27: PPT v2 com novo preço de Barueri

- Ela trocou o preço (faturamento) de Barueri no `Resultado TLM - MP27 vs R09_v2.xlsm` (salvo 16:55). Recalculado em instância separada do Excel (`mp27_extrair_resumo_slide.py`, fontes: Base Total_Act_Flash_Fct, Base_KSB1, HC R09, HC R07, Headcount 2026.09, Base Total MP27). Só BATotal Líquido mudou: 27.958 -> 30.285 (+2.327). Total: fat. líq. 88.176, PBIT 2.504 (era 177), R9 3.437, Δ -933. Custo igual (-85.672).
- Com o Headcount 2026.09 aberto, o HC do R9 (AB19) vira #VALUE!; usei os valores de R09 do JSON anterior (`data/processed/tlm/slide/pl_unidades.json`) onde a nova extração dava erro. JSON final: `pl_unidades_v2.json`.
- Template do gerador: `data/processed/circuito/20261020_v8/Bridge_CP_v8.pptx`. Saída: `TLM_Economics_MP27_v2.pptx` na rede (v1 intacto). Comentários = notas do orador de cada slide (efeitos ordenados por tamanho + roteiro no slide Total).
- Textos: desconto de tarifa em Barueri = decisão de eficiência (TLM busca menor resultado); D&A +4.946 = fim IFRS16 do armazém CC 7287 (jul/26) sem aluguel/amortização na base MP27.
- Barueri Other costs -3.036 vs R9: Outras Despesas -1.387 (R9 tinha crédito +1.297), Prestados -610, Aluguéis -587, Outros TLM -324, Treinamento -304, empilhadeira -239, Vigilância -185, Limpeza -114, Despesas Sociais +360. Motivo por lançamento NÃO verificado (falta abrir Base Total MP27 x base R9).
- Shell: heredoc `<<EOF` não funciona no PowerShell; caminhos de rede com barra invertida em bash quebram -> usar PowerShell.
