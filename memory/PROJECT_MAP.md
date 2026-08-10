# PROJECT_MAP — Mapa de conexões entre projetos

---

## Projeto 1: Fitted Units — Automação de Controladoria
- **Domínio:** montagem/sequenciamento de pneus e rodas dentro de plantas montadoras de clientes. Controladoria: despesas, faturamento, EBIT, P&L mensal (Flash e Actual).
- **Ontologias:** `ontology/fitted_units.json`
- **Sistemas externos:** SAP (transações KSB1, ZLFIB, FBL5N), Excel

### Sub-projeto: Fitted Units Despesas
> Nome dado pela usuária em 2026-08-10 para identificar essa frente específica dentro de Fitted Units (a parte de despesas — extração da KSB1 e conferência de agrupamento gestorial). As outras frentes de Fitted Units (faturamento, EBIT, P&L) ainda não têm automação — ver "Próximos passos" no `PROJECT_MAP` e no `BRIEFING.md`.
- **O que já existe (concluído em 2026-08-10):**
  - `scripts/sap/atualizar_ksb1_gui.py` — GUI que extrai a KSB1 (Gestoriais + Sem Agrupamento) direto do SAP e salva na rede, com versionamento automático de arquivo (nunca sobrescreve).
  - `scripts/sap/check_agrupamentos_ksb1.py` — conferência automática: verifica se toda conta contábil do Sem Agrupamento está vinculada a um agrupamento gestorial, gera `Check de agrupamentos - MM.YYYY.xlsx` na pasta do mês.
  - Atalho único na rede (`ATUALIZAR KSB1.lnk`, ícone de pneu) com os dois botões ("Extrair KSB1" e "Gerar Check de Agrupamentos") na mesma janela.
  - Regra de negócio completa registrada em `ontology/fitted_units.json` → `classificacao_despesas`.
- **Próximas etapas do processo de despesas (ainda não construídas, ver `ontology/fitted_units.json` → `processo_recorrente`):**
  3. Carregar as informações da KSB1 em Excel pra formar a base intermediária.
  4. Carregar a base intermediária no arquivo de rateio das despesas da Gerência (GER) pras demais unidades.
  5. Carregar a base intermediária (já rateada) no arquivo de P&L.

---

## Projeto 2: Circuito Panamericano — Automação de Controladoria
- **Domínio:** complexo de testes (Elias Fausto), modelo de aluguel de espaço (Pirelli R&D e terceiros). Controladoria: despesas, EBIT, P&L mensal (Flash e Actual). Mesmo processo de KSB1 da Fitted Units, sem etapa de rateio da Gerência; faturamento recebido já fechado.
- **Scripts principais:** ainda não criados.
- **Ontologias:** `ontology/circuito_panamericano.json`
- **Sistemas externos:** SAP (transações KSB1, FBL5N), Excel

---

## Projeto 3: Original Equipment — (a detalhar)
- Ainda não detalhado com a usuária. Adicionar quando ela quiser explicar o domínio.

---

## Regra de carregamento cruzado

Ao iniciar qualquer tarefa, verificar neste arquivo quais projetos estão conectados e carregar as ontologias correspondentes.
