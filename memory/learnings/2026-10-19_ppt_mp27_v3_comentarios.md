# 2026-10-19 — PPT MP27 Fitted v3 (comentários enxutos, sobre a v2 dela)

- Ela salvou a v2 (21:34) com **as tabelas recoladas como imagem** ('Picture 3' nos slides 2,5,7,8,9) e reordenou: 1 capa, 2 Brasil, 3 Mix, 4 Resumo, 5 IBI, 6 Clientes, 7 SJP, 8 GOI, 9 RES, 10 MO, 11 Perguntas.
- Feedback dela: comentários "truncados, não é nível de presidente" → **uma ideia por bullet, ≤ ~45 caracteres, sem parênteses encadeados**; bridges **sem azul** nas barras positivas (todas as variações em cinza claro DFE5EF); MO **sem run-rate** (ela não confia que seja "certo" na frente do presidente) → R9 x MP'27 + composição do aumento (sal+enc 81%, assist. médica 19%, +38%).
- Gráfico 100% empilhado (python-pptx, percentStacked) **renderiza fora de proporção** no Export do PowerPoint (Fiat 93% aparecia ~40%) apesar do cache correto → trocado por barras de formas nativas. Ela leu isso como "rentabilidade não atualizada"; os dados (Rentab v3 20:03) batiam com MENS e Base Book 21:34 (EBIT IBI 6.075).
- Base Book 21:34 = mesmos valores do de 20:03 (diff célula a célula zero).
- Script: `scripts/fitted/gerar_ppt_mp27_v3.py` (importa os dados/helpers do rev2 — montagem do rev2 agora em `main()`); edita a v2 dela no lugar, não toca nas imagens (checado por hash). Assert: nenhum "run-rate" no texto dos slides.
- Saída: `EO_FITTED\...\MP 2027\Fitted_Units_Meeting_Budget'27_v3.pptx` (21:46) + data/processed.
