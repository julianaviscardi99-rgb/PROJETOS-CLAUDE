# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## Resumo do dia 2026-08-10 — Fitted Units Despesas (KSB1)
- **Instrução da usuária para as próximas sessões:** focar só em Fitted Units por enquanto. Circuito Panamericano fica pausado até ela pedir de novo — não avançar nele sem ela pedir.
- **O que foi entregue e validado hoje** (sub-projeto batizado por ela de "Fitted Units Despesas" — ver `memory/PROJECT_MAP.md`):
  1. `scripts/sap/atualizar_ksb1_gui.py` — GUI que extrai a KSB1 da Fitted Units (Gestoriais + Sem Agrupamento) direto do SAP via GUI Scripting e salva na área de rede. Nunca sobrescreve arquivo existente (versiona `_v2`, `_v3`...). Mês/ano padrão: ano atual dinâmico, mês = mês anterior ao atual. Navegação entre as duas extrações usa o botão "Voltar" da toolbar do SAP (mais robusto que F3). Sem popups manuais de confirmação (as notificações de segurança de scripting já estão desativadas no SAP GUI da usuária).
  2. `scripts/sap/check_agrupamentos_ksb1.py` — conferência automática: Check 1 lista as contas contábeis sempre ignoradas (6 contas fixas, mesmo com valor 0, mais qualquer conta iniciada em "B" encontrada no mês); Check 2 lista contas do Sem Agrupamento sem vínculo no agrupamento gestorial (pra mandar pra controladoria central); Resumo compara o total das duas bases. Tudo numa aba só, arquivo `Check de agrupamentos - MM.YYYY.xlsx` salvo na pasta do mês.
  3. As duas funções ficam na mesma janela/app, acessada por um único atalho na rede: `ATUALIZAR KSB1.lnk` (ícone de pneu Pirelli), com os botões "Extrair KSB1" e "Gerar Check de Agrupamentos".
  4. Regras de negócio documentadas em `ontology/fitted_units.json` (→ `classificacao_despesas.check_de_agrupamentos`); decisões e motivos em `memory/DECISOES.md`; novas regras gerais em `memory/REGRAS_RAPIDAS.md` (#11 respostas em português, #12 nunca sobrescrever arquivo — versionar).
  5. `requirements.txt` atualizado com `openpyxl` (necessário pro check).
- Histórico detalhado passo a passo de hoje (todos os bugs e correções, em ordem cronológica): `memory/long_term/2026-08-10_*_briefing_snapshot.md`.
- Tudo commitado e sincronizado com o GitHub (branch `main`, sem pendência de commit/push).

---
## Próximos passos (retomar amanhã — só Fitted Units)
- Usuária ainda precisa rodar "Gerar Check de Agrupamentos" com um mês real (Gestoriais + Sem Agrupamento do mesmo mês) pela primeira vez desde o último ajuste (aba única + Check 1 sempre listando as 6 contas fixas) e confirmar se bate com o esperado.
- Próxima etapa do processo de despesas da Fitted Units (ver `ontology/fitted_units.json` → `processo_recorrente`, passos 3-5): montar a base intermediária a partir da KSB1 já extraída, depois o rateio dos custos da Gerência (GER) pras demais unidades, depois carregar no arquivo de P&L.
- `memory/PROJECT_MAP.md`: Original Equipment ainda não detalhado (fora do escopo por ora, já que o foco agora é só Fitted Units).
- Circuito Panamericano: não mexer até a usuária pedir explicitamente.

---
## Contexto permanente do projeto
- Esta pasta (`C:\Users\silveju001\Projetos Claude`) está estruturada seguindo o "Guia de Onboarding — Como Trabalhar com o Claude de Forma Profissional" (maio 2026, baseado no projeto Cockpit Ind — Pirelli Planning & Control).
- Objetivo real deste projeto: automatizar controladoria (Fitted Units e Circuito Panamericano) hoje feita em Excel — resultado, despesas, faturamento, EBIT, P&L mensal. Detalhes completos em `CLAUDE.md`. **Foco atual (a partir de 2026-08-10): só Fitted Units.**
- Repositório Git já configurado com backup remoto no GitHub; backup automático diário às 18h (Agendador de Tarefas do Windows) além do backup por sessão longa (a cada 45 ações).
- GUI compartilhável da KSB1 (`scripts/sap/atualizar_ksb1_gui.py`, atalho único `ATUALIZAR KSB1.lnk` na rede) e o Check de Agrupamentos (`scripts/sap/check_agrupamentos_ksb1.py`) já em uso — ver `memory/DECISOES.md` para o histórico completo de decisões sobre eles.
