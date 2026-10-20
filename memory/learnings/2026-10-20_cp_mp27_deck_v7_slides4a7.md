# CP MP'27 — deck v7: slides 4 a 7 (20/10/2026, manhã/meio-dia)

**Arquivo:** `...\Circuito Panamericano\2027\MP'27\Bridge CP_v7.pptx` — é a versão DELA; ela edita em paralelo (v7 de 11:26 → 11:41 → 12:01, e já tem 8 slides às 12:01). Nunca partir da minha cópia: copiar da rede, comparar hash antes de gravar e só trocar os slides pedidos.
**Backup do estado final (12:01, 8 slides):** `data/processed/circuito/20261020_backup_final/` (v6, v7 e as fontes MENS / Base de Custo / Base Book / Faturamento).

## Scripts (todos no repo, `scripts/circuito/`)
- `atualizar_ppt_cp_mp27_v6.py` — v6 com a nova rentabilidade (slides 1–5; fonte MENS aba LOCAL).
- `atualizar_ppt_cp_mp27_v7_slides5a7.py` — slide 5 (linhas de custo por voz), 6 (riscos e oportunidades), 7 (investimentos, só estrutura). Só mexe nos slides 5–7.
- `refazer_slide4_ppt_cp_mp27_v7.py` — refaz o slide 4 lendo a tabela que ela ajustou (fonte = a tabela dela). Só mexe no slide 4.

## Regras de negócio / números
- Fonte de verdade do MP'27 e R9: **MENS CP (aba LOCAL)**: EBIT 2.845, custo fixo 10.692 (+952), receita 13.537. O Base Book na rede demorou a refletir; para o histórico (FY24/FY25/MP26) usa-se o Base Book.
- Abertura do custo fixo (dada por ela): postergações 411, HSE 300, labour 138, inflação 4,4% 213, redução de despesas −150, portaria 40.
- Postergações: lista 2026 = **611 = 411 mantidos + 200 cortados** (pintura 70 "realizado 2026", reajuste contratos 52, viagens 35, sanitários 21, TI 20, pluvial 2). NÃO há "revisão de valores": meu +117/−317 estava errado (os valores da lista já eram os atualizados por ela).
- EBITDA 6.718 = EBIT 2.845 + D&A 3.873 (49,6%); cálculo meu, não está na tabela dela; "geração de caixa" é simplificação.
- R9 da Base de Custo (9.791) ≠ R9 da MENS (9.740): por isso o slide 5 não mostra deltas por linha vs R9.
- Slide 5, "Contratos e locações de apoio" (370) = gerador (locação 123,7 + diesel 5,0), climatização 77,5, poços artesianos (60,0 + 32,6), container portaria 21,5, combate a incêndio 14,0, elevador 13,2, rádios 6,9, purificadores 6,8, compressores 5,5, máquina de café 3,2. Agrupamento por natureza, escolha minha.
- Slide 6: riscos/oportunidades são SENSIBILIDADES que propus (eventos ±10% = ±269; contrato R&D sem reajuste −817; inflação ±1 p.p. ≈ 48; OEM ±10% ≈ 41; labour +1 p.p. ≈ 10; escalonar postergações até +411). A validar com ela. Título do slide 6 trocado de "Investimentos" para "Riscos e Oportunidades".
- Faturamento CP "Volume MP27" tem eventos 3.200 brutos, MENS/tabela usam 2.800 (usado o da MENS). Aberto: o que é a "redução de despesas −150".

## Técnico
- python-pptx: em `a:tcPr`, `a:lnB` tem que vir ANTES do preenchimento (insert(0)); com SubElement depois do fill a borda some.
- Remover gráfico de slide: `slide.part.drop_rel(rId)` além de remover o elemento.
- Copiar shape: deepcopy + novo `cNvPr id` + `addnext`.
