# CLAUDE.md — Regras e Identidade do Projeto

> Lido sempre no início de cada sessão. Define quem é a usuária, o que este projeto faz e como o Claude deve trabalhar aqui.

---

## Quem é a usuária

Juliana é **Controller na Pirelli**, responsável por três frentes:

1. **Original Equipment (mercado)** — controller comercial de toda a parte de equipamento original vendido no mercado.
2. **Fitted Units** — business onde a Pirelli monta e sequencia o conjunto pneu + roda e envia montado para o cliente. A unidade normalmente fica fisicamente dentro da planta montadora do cliente, para otimizar transporte e tempo. Juliana é controller dessa BU: constrói o resultado, analisa despesas, classifica, analisa faturamento e monta o EBIT. Envia um arquivo de P&L mensalmente. Hoje esse trabalho é feito em Excel.
3. **Circuito Panamericano** — maior complexo de testes da América Latina. Juliana exerce as mesmas atividades de controladoria que na Fitted Units (resultado, despesas, faturamento, EBIT, P&L mensal).

## Objetivo do projeto

Automatizar com o Claude o que hoje é feito em Excel nessas frentes (especialmente Fitted Units e Circuito Panamericano): construção de resultado, classificação de despesas, faturamento, EBIT e geração do P&L mensal. Metas: reduzir erros, encontrar erros existentes nas planilhas/processos atuais, e trazer ganhos (tempo e qualidade) para a companhia.

---

## Carga de contexto (obrigatório ao iniciar cada sessão)

1. Ler `memory/REGRAS_RAPIDAS.md` — regras críticas
2. Ler `memory/BRIEFING.md` — o que está pendente
3. Ler `memory/PROJECT_MAP.md` — quais projetos estão conectados
4. Ler `memory/DECISOES.md` — decisões tomadas e seus motivos
5. Ler as ontologias relevantes à tarefa (`ontology/*.json`)

## Registro de conhecimento (obrigatório durante a sessão)

- Ao aprender algo que funcionou → salvar em `memory/learnings/` com data no nome
- Ao encontrar um erro → salvar em `memory/errors/` com data no nome
- Ao tomar decisão importante → registrar em `memory/DECISOES.md` com o motivo
- Ao explicar conceito de negócio → registrar em `ontology/<dominio>.json`

## Autonomia

- Ações seguras (criar arquivos, analisar, escrever scripts) → executar direto
- Ações irreversíveis (enviar email, modificar dados, deletar) → confirmar antes

## Qualidade

- Nunca hardcodar caminhos absolutos do usuário nos scripts
- Scripts devem funcionar em qualquer máquina com `pip install -r requirements.txt`
- Todo output salvar em `data/processed/`, nunca sobrescrever arquivos originais
