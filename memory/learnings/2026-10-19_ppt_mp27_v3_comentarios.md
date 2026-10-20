# 2026-10-19 — PPT MP27 Fitted v3 (comentários enxutos, sobre a v2 dela)

- Ela salvou a v2 (21:34) com **as tabelas recoladas como imagem** ('Picture 3' nos slides 2,5,7,8,9) e reordenou: 1 capa, 2 Brasil, 3 Mix, 4 Resumo, 5 IBI, 6 Clientes, 7 SJP, 8 GOI, 9 RES, 10 MO, 11 Perguntas.
- Feedback dela: comentários "truncados, não é nível de presidente" → **uma ideia por bullet, ≤ ~45 caracteres, sem parênteses encadeados**; bridges **sem azul** nas barras positivas (todas as variações em cinza claro DFE5EF); MO **sem run-rate** (ela não confia que seja "certo" na frente do presidente) → R9 x MP'27 + composição do aumento (sal+enc 81%, assist. médica 19%, +38%).
- Gráfico 100% empilhado (python-pptx, percentStacked) **renderiza fora de proporção** no Export do PowerPoint (Fiat 93% aparecia ~40%) apesar do cache correto → trocado por barras de formas nativas. Ela leu isso como "rentabilidade não atualizada"; os dados (Rentab v3 20:03) batiam com MENS e Base Book 21:34 (EBIT IBI 6.075).
- Base Book 21:34 = mesmos valores do de 20:03 (diff célula a célula zero).
- Script: `scripts/fitted/gerar_ppt_mp27_v3.py` (importa os dados/helpers do rev2 — montagem do rev2 agora em `main()`); edita a v2 dela no lugar, não toca nas imagens (checado por hash). Assert: nenhum "run-rate" no texto dos slides.
- Saída: `EO_FITTED\...\MP 2027\Fitted_Units_Meeting_Budget'27_v3.pptx` (21:46) + data/processed.

## Ajustes 22:14 (v3 sobrescrita a pedido dela; backup da v3 dela de 22:02 em data/processed/..._v3_backup_2202.pptx)
- Script `scripts/fitted/gerar_ppt_mp27_v3_ajustes.py` (9 args: + volumes MP'27_ADJ e pasta de logos `data/raw/logos/`, logos baixados do Wikimedia Commons).
- Rentabilidade por cliente v3 de 22:00: ela trocou o rateio para **volume** → Fiat ROS 10,9% (78% do EBIT IBI), Iveco 41,6%, CNH 38,4%; Fiat = 43% do EBIT Brasil.
- IBI: premissa DELA — 3º turno Fiat = MO 2,5 + demais despesas 2,0 mi/ano; ANTT 1,5 custo extra; reajuste Fiat 5,65 cobre 94%.
- MO: bridge por driver (R9 25,5 → estrutura 2026 ano cheio +2,7 [IBI 1,1, RES 1,2, demais 0,3] → reajuste 2027 +1,7 [sal+enc 1,4; assist. 0,3] → 29,9), eixo truncado em 22.
- Novos slides: New Business (Nissan RES 373k, Tukan SJP 357k, CNH 16k = 747k / 14%; saídas Jaguar e Sorocaba) e Risks & Opportunities (−7,4 / +1,8).
- python-pptx: `percentStacked` renderiza errado no Export do PowerPoint → usar formas nativas para participação.
