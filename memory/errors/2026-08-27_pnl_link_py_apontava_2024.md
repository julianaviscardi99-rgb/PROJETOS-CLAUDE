# Erro encontrado: link de comparação PY do P&L apontava para 2024, não 2025

**Data:** 2026-08-27
**Onde:** arquivo `<MM>_P&L Fitted Units_Actual_<Mês>-26.xlsx` (Fitted Units, todos os meses de 2026), aba "Resumo Resultado Ano", coluna BJ (rotulada "Actual 2025", bloco de comparação com o ano anterior/PY).

## O que estava errado

A coluna BJ (e o bloco de colunas AW:BH que alimenta o `SUM` de BJ) está linkada (link externo do Excel) em:
`...\Resultados Fitted\2024\12_December_Actual\12_P&L Fitted Units_Actual_December-24.xlsx`

Deveria apontar para:
`...\Resultados Fitted\2025\12_December_Actual\12_P&L Fitted Units_Actual_December-25.xlsx` (existe na rede, fechado em 09/01/2026).

## Onde foi confirmado

Checado nos 7 arquivos P&L Actual de 2026 já fechados (Jan-Jul) — **todos com o mesmo link errado**, incluindo Janeiro (o mês em que o link deveria ter sido "rolado" pra frente na virada do ano). Ou seja, o erro nasceu em Janeiro/2026 (o rollover de PY não aconteceu) e se arrastou mês a mês porque cada arquivo novo nasce de uma cópia do anterior.

## Impacto (quantificado, teste feito em cima do arquivo real de Julho/2026, só leitura — nada foi salvo)

Usando `Workbook.ChangeLink` via COM (Excel isolado) pra repontar o link pro arquivo certo (2025) e recalculando, sem tocar em nenhuma fórmula:

| Métrica (Julho/2026) | Com link errado (2024) | Com link certo (2025) | Diferença |
|---|---|---|---|
| EBIT PY | 18.317,80 mil | 22.513,66 mil | +4.195,86 mil |
| ROS% PY | 19,59% | 23,44% | +3,85 p.p. |

Ou seja, todo P&L enviado pra consolidação em 2026 até agora mostrou o EBIT do ano anterior subestimado em ~R$ 4,2 milhões (nesse mês específico), distorcendo a leitura de "evolução vs. ano passado".

## Decisão (confirmada com a usuária)

- **Jan-Jul/2026 não serão corrigidos retroativamente** (já fechados/enviados) — decisão explícita da usuária.
- **A partir de Agosto/2026**, o link será corrigido (mesmo mecanismo `ChangeLink`, mantendo a célula como link vivo, não vira valor fixo) quando o arquivo de Agosto for criado a partir da cópia de Julho.
- Como PY só muda 1x/ano (Janeiro), a correção de Agosto deve bastar pro resto de 2026 — só precisa rolar de novo em Jan/2027 (pra "Actual 2026").
- O link "MP26" (Management Plan, coluna AU) foi conferido e está correto — só o PY estava quebrado.

## Mecanismo técnico validado (reaproveitar na automação do P&L, Passo 7)

`wb.LinkSources(1)` lista os links Excel do workbook (tipo `xlExcelLinks=1`); `wb.ChangeLink(nome_antigo, nome_novo, 1)` troca o destino mantendo a fórmula de link viva; `excel.CalculateFullRebuild()` força recalcular depois da troca. Testado abrindo o arquivo real com `ReadOnly=True`, fechando com `SaveChanges=False` — zero risco ao arquivo de produção.

Ver também `memory/DECISOES.md` (entrada 2026-08-27) e `memory/PROJECT_MAP.md` (Fitted Units Despesas — próximo passo, P&L).
