# PROJECT_MAP — Mapa de conexões entre projetos

---

## Projeto 1: Fitted Units — Automação de Controladoria
- **Domínio:** montagem/sequenciamento de pneus e rodas dentro de plantas montadoras de clientes. Controladoria: despesas, faturamento, EBIT, P&L mensal (Flash e Actual).
- **Scripts principais:** ainda não criados. Vão tratar a extração da KSB1, montagem da base intermediária, rateio da Gerência e carga no P&L.
- **Ontologias:** `ontology/fitted_units.json`
- **Sistemas externos:** SAP (transações KSB1, ZLFIB, FBL5N), Excel

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
