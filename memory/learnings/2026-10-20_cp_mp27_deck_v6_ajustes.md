# CP MP'27 — ajustes na v6 sobre a versão dela (20/10/2026, ~10:58)

**Pedido:** atualizar a própria `Bridge CP_v6.pptx` (NÃO criar v7). Ela colou a tabela nova no slide 1 (imagem) às 10:48.
**Script:** `scripts/circuito/atualizar_ppt_cp_mp27_v6.py` (args: v6 dela, Base Book, MENS, Faturamento, Base de Custo v2, saída). Mantém a imagem da tabela do slide 1 (shape_type 13), redesenha o resto; slides 2–5 são limpos (menos o título) e refeitos.
**Backup da v6 dela (10:48):** `data/processed/circuito/20261020_fontes/Bridge CP_v6_BACKUP_antes_ajuste_1048.pptx`.

## Fonte de verdade
- A **MENS (aba LOCAL)** é a fonte da tabela colada: EBIT 2.845 (+359, +14,4%), custo fixo 10.692 (+952), ROS 21,0%. O Base Book em disco ainda estava com o número antigo (10.745 / 2.791) → para MP'27 e R9 ler a MENS; Base Book só para o histórico.
- A mensalização de eventos mudou: MENS E8:P8 = 140,196,224,252,252,252,168,196,308,336,280,196 (2.800). O Faturamento tem 3.200 em "Volume MP27" (não bate com a tabela) → usar a MENS.

## Abertura do custo fixo (dada por ela)
postergações 411 | novos itens obrigatórios HSE 300 | labour 138 | inflação 4,4% 213 | redução de despesas −150 | portaria 40 = 952.
Postergações: lista 2026 = 611; cortados 317 (8 itens); revisões de valor +117 (Sansão 170→180, restauração 72→167, elétrica 23→35) → 411.

## Pontos em aberto
- "Redução de despesas −150": não sei quais itens; ficou genérico no deck.
- "Mensagens-chave" do slide 1 virou "Resumo executivo".
