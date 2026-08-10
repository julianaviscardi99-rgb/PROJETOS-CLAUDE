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
| `sap/inspecionar_tela.py` | Lista os IDs técnicos e textos de todos os campos da tela atual do SAP (não clica em nada) — usado para mapear telas novas. |
| `sap/extrair_ksb1.py` | Preenche os filtros da KSB1 (BU, período, agrupamento gestoriais), executa e exporta o resultado para `.xlsx` em `data/raw/`. Para Fitted Units, também copia o arquivo para a área de rede da Pirelli (pasta do mês correspondente). Só leitura/exportação, nada que altere dados no SAP. Uso pessoal (Juliana, via Claude Code). |
| `sap/atualizar_ksb1_gui.py` | Versão com janela gráfica (sem terminal), fixa em Fitted Units, extrai Gestoriais + Sem Agrupamento de uma vez e salva direto na área de rede. Autossuficiente — feita para ser copiada para a rede e usada por qualquer pessoa com acesso (ex: estagiária). |
| `sap/ATUALIZAR KSB1.bat` | Atalho que roda o `atualizar_ksb1_gui.py` — checa/instala Python e pywin32 se precisar. Publicado também na área de rede (`00.Extração Base KSB1/ATUALIZAR KSB1.bat`) para uso compartilhado. |

Sempre que criarmos um novo script, adicionar aqui antes de encerrar a sessão.
