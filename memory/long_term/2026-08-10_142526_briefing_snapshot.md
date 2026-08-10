# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## Sessão atual
- Data: 2026-08-10
- O que foi feito:
  - Confirmado que a estrutura de pastas da Fase 2 já existia (`memory/`, `ontology/`, `data/raw`, `data/processed`, `scripts/`) e que o repositório Git já estava conectado ao GitHub (origin: `julianaviscardi99-rgb/PROJETOS-CLAUDE`).
  - Coletado o contexto pessoal da usuária (cargo, empresa, objetivo do projeto).
  - Criado `CLAUDE.md` na raiz com identidade da usuária, contexto de negócio e as regras de carga de contexto / registro de conhecimento / autonomia / qualidade.
  - Criado `memory/REGRAS_RAPIDAS.md` (Fase 4) com as 10 regras críticas + gate pré-execução.
  - Confirmada Fase 1 (instalação/ambiente) como concluída.
  - Criadas as ontologias `ontology/fitted_units.json` e `ontology/circuito_panamericano.json` (unidades, sistemas SAP — KSB1/ZLFIB/FBL5N, classificação de despesas por gestoriais, processo recorrente Flash/Actual). Original Equipment fica para depois.
  - Feito o primeiro commit + push de tudo (branch `main` agora rastreia `origin/main`).
  - Criado `Auto_Backup_GitHub.bat` (add + commit + push automático) e agendada task diária no Windows (`Backup_Projeto_Claude`, 18h).

---
## Próximos passos
- Testar o `Auto_Backup_GitHub.bat` manualmente (duplo-clique) para confirmar que roda certo fora deste chat.
- Criar `ontology/shared_entities.json` (conceitos comuns às BUs: gestoriais, Flash/Actual, KSB1, FBL5N).
- Detalhar Original Equipment quando a usuária quiser.
- Entender o processo atual em Excel (Fitted Units e Circuito Panamericano: faturamento, despesas, EBIT, P&L mensal) para planejar a automação.
- Preencher `memory/PROJECT_MAP.md` (ainda vazio).

---
## Contexto permanente do projeto
- Esta pasta (`C:\Users\silveju001\Projetos Claude`) está estruturada seguindo o "Guia de Onboarding — Como Trabalhar com o Claude de Forma Profissional" (maio 2026, baseado no projeto Cockpit Ind — Pirelli Planning & Control).
- Objetivo real deste projeto: automatizar controladoria (Fitted Units e Circuito Panamericano) hoje feita em Excel — resultado, despesas, faturamento, EBIT, P&L mensal. Detalhes completos em `CLAUDE.md`.
- Repositório Git já configurado com backup remoto no GitHub.
