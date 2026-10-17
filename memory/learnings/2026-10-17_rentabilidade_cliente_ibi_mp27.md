# 2026-10-17 — Rentabilidade por cliente de Ibirité (MP27) a partir da MENS FITTED

- "Arquivo de mensalização da Fitted" = `EO_FITTED\BU FITTED\Forecast\MP\MP 2027\MENS FITTED MP27_v3_teste daniel.xlsx` (NÃO a pasta GFU_DAC\Management Plan\MP 2027; lá não há MENS). Ela teve dificuldade de colar o caminho de rede no prompt: procurar o arquivo com `find` nas duas raízes (GFU_DAC e EO_FITTED).
- Aba IBI: linhas 53–55 volume (Fiat/Iveco/CNH), 65–67 faturamento, 75–77 material de consumo por cliente; demais custos (19–23, 32–37) só no total. Meses em E:P, custos em R$ mil, faturamento/material em R$.
- Frete (linha 79) = `Fiat vol × (−419×0,9075/180)` → é 100% Fiat (confirmado). Rateio fixo entre unidades fica na aba oculta `Rateio Fixo`.
- Método: direto (faturamento, material) + rateio por volume mensal; driver por linha em Premissas (Volume / Faturamento / Fiat 100%). Check contra aba IBI: faturamento 39.834, EBIT 5.391.
- Hipótese em aberto: CNH/Iveco têm custos específicos não abertos no arquivo (Rafaela).
- Técnica: gerar com openpyxl (fórmulas) e recalcular/salvar via Excel COM (`SaveAs` formato 51) para o arquivo já abrir com valores em cache; heredoc longo no Bash truncou o script → usar a ferramenta Write.
