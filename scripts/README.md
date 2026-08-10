# Scripts — Índice por Módulo
> Última atualização: 2026-08-10

## Fitted Units
| Script | Descrição |
|---|---|
| — | ainda não criado |

## Circuito Panamericano
| Script | Descrição |
|---|---|
| — | ainda não criado |

## Scripts de Infraestrutura
| Script | Descrição |
|---|---|
| `check_session_length.py` | Hook (PostToolUse) que conta as ações da sessão e, ao atingir 45, dispara `session_transition.py` e alerta o Claude/usuária |
| `session_transition.py` | Arquiva snapshot do `BRIEFING.md` em `memory/long_term/` e faz `git add + commit + push` automaticamente |

## SAP
| Script | Descrição |
|---|---|
| `sap/test_conexao_sap.py` | Testa a conexão via SAP GUI Scripting com a sessão SAP já aberta e logada. Pré-requisito para qualquer automação de SAP. |

Sempre que criarmos um novo script, adicionar aqui antes de encerrar a sessão.
