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

### Sub-projeto: Fitted Recuperação
> Nome dado pela usuária em 2026-08-11. Objetivo: detectar lançamento/pagamento a fornecedor em duplicidade na KSB1 (Fitted Units), período 01.01.2026 a 31.07.2026, excluindo documentos estornados.
- **Scripts (fora do fluxo mensal recorrente, específicos deste estudo):**
  - `scripts/sap/extrair_ksb1_periodo.py` — extrai a KSB1 (Sem Agrupamento) para um período arbitrário (não só um mês). Salva em `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Estudos\Estudo Duplicidade Pagamento\`.
  - `scripts/sap/analisar_duplicidade_pagamento.py` — identifica pares de estorno (mesmo fornecedor/valor com sinal oposto) e aplica os critérios de duplicidade, gerando `Análise Duplicidade Pagamento.xlsx` na mesma pasta.
- **Critérios de duplicidade confirmados pela usuária:** Fornecedor+Valor+Documento de compras, e Fornecedor+Valor+Data de lançamento.
- **Resultado da primeira rodada (01.01-31.07.2026):** 175.310 linhas no extrato, 7.147 com fornecedor, 50 pares de estorno excluídos. 228 grupos duplicados por Documento (R$ 823.535,87) e 649 grupos por Data (R$ 1.533.184,43) — triagem heurística, precisa revisão manual.
- **Tentativa de refinar com Nº de NF — NÃO deu certo, não repetir sem novo dado:** a coluna "Nº doc.de referência" da KSB1 não é a Nota Fiscal (confirmado pela usuária). Tentei cruzar esse número na transação ZLFIB (tem campos "Nr Documento" e "Nota Fiscal" de verdade), mas nenhuma busca (por número de documento, nem por fornecedor+período) trouxe resultado confiável — a própria usuária confirmou que não sabe como esse cruzamento funcionaria. **Pausado até alguém do time funcional/TI do SAP confirmar como ligar o documento de referência da KSB1 a uma Nota Fiscal na ZLFIB.** Não tentar de novo por tentativa e erro sem essa confirmação.

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
