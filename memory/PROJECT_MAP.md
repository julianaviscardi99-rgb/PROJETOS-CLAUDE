# PROJECT_MAP — Mapa de conexões entre projetos

---

## Projeto 1: Fitted Units — Automação de Controladoria
- **Domínio:** montagem/sequenciamento de pneus e rodas dentro de plantas montadoras de clientes. Controladoria: despesas, faturamento, EBIT, P&L mensal (Flash e Actual).
- **Ontologias:** `ontology/fitted_units.json`
- **Sistemas externos:** SAP (transações KSB1, ZLFIB, FBL5N), Excel

### Estrutura de pastas dos scripts (reorganizada em 2026-08-13)
> Antes de 2026-08-13, todos os scripts de Fitted Units (Despesas + Recuperação) ficavam soltos numa pasta só (`scripts/sap/`), misturados. Reorganizado a pedido da usuária, por sub-projeto:
```
scripts/sap/fitted_units/
  _shared/                    <- codigo compartilhado entre sub-projetos
    ksb1_core.py               (connect_session, navegacao KSB1, nome_com_versao, BU/REDE_BASE/MESES_*)
    ferramentas/                (diagnostico, nao ligados a nenhum sub-projeto especifico)
      inspecionar_tela.py, test_conexao_sap.py, diagnosticar_popup.py, "DIAGNOSTICAR POPUP.bat"
  fitted_units_despesas/       <- sub-projeto "Fitted Units Despesas"
    atualizar_ksb1_gui.py, check_agrupamentos_ksb1.py, gerar_base_intermediaria.py,
    extrair_ksb1.py (legado), "ATUALIZAR KSB1.bat", atualizar_ksb1_launcher.vbs,
    criar_atalho_ksb1.ps1, assets/pirelli_tire.ico
  fitted_recuperacao/          <- sub-projeto "Fitted Recuperação"
    extrair_ksb1_periodo.py, analisar_duplicidade_pagamento.py, analisar_zlfib_duplicidade.py,
    verificacao_mensal_zlfib.py, watcher_mensal_zlfib.bat
```
**Pontos de integração externos que dependem desses caminhos (atualizados na mesma reorganização):**
- Tarefa agendada do Windows `Verificacao_ZLFIB_Duplicidade_Mensal` → aponta pra `fitted_recuperacao/watcher_mensal_zlfib.bat`.
- Atalho `ATUALIZAR KSB1.lnk` na rede → aponta (via `atualizar_ksb1_launcher.vbs`) pra `fitted_units_despesas/atualizar_ksb1_gui.py`. Se os arquivos dessa pasta forem movidos de novo no futuro, é preciso rodar `criar_atalho_ksb1.ps1` de novo pra regenerar o `.lnk` (ele fica fora do Git, é local na rede).

### Sub-projeto: Fitted Units Despesas
> Nome dado pela usuária em 2026-08-10 para identificar essa frente específica dentro de Fitted Units (a parte de despesas — extração da KSB1 e conferência de agrupamento gestorial). As outras frentes de Fitted Units (faturamento, EBIT, P&L) ainda não têm automação — ver "Próximos passos" no `PROJECT_MAP` e no `BRIEFING.md`.
- **O que já existe (concluído em 2026-08-10):**
  - `scripts/sap/fitted_units/fitted_units_despesas/atualizar_ksb1_gui.py` — GUI que extrai a KSB1 (Gestoriais + Sem Agrupamento) direto do SAP e salva na rede, com versionamento automático de arquivo (nunca sobrescreve).
  - `scripts/sap/fitted_units/fitted_units_despesas/check_agrupamentos_ksb1.py` — conferência automática: verifica se toda conta contábil do Sem Agrupamento está vinculada a um agrupamento gestorial, gera `Check de agrupamentos - MM.YYYY.xlsx` na pasta do mês.
  - Atalho único na rede (`ATUALIZAR KSB1.lnk`, ícone de pneu) com os dois botões ("Extrair KSB1" e "Gerar Check de Agrupamentos") na mesma janela.
  - Regra de negócio completa registrada em `ontology/fitted_units.json` → `classificacao_despesas`.
- **Próximas etapas do processo de despesas (ainda não construídas, ver `ontology/fitted_units.json` → `processo_recorrente`):**
  3. Carregar as informações da KSB1 em Excel pra formar a base intermediária.
  4. Carregar a base intermediária no arquivo de rateio das despesas da Gerência (GER) pras demais unidades.
  5. Carregar a base intermediária (já rateada) no arquivo de P&L.

---

### Sub-projeto: Fitted Recuperação
> Nome dado pela usuária em 2026-08-11. Objetivo: detectar lançamento/pagamento a fornecedor em duplicidade na KSB1 (Fitted Units), período 01.01.2026 a 31.07.2026, excluindo documentos estornados. **Desde 2026-08-13, também inclui checagem mensal automática de NF duplicada na ZLFIB (ver abaixo).**
- **Scripts (fora do fluxo mensal recorrente da Despesas, específicos deste estudo — todos em `scripts/sap/fitted_units/fitted_recuperacao/`):**
  - `extrair_ksb1_periodo.py` — extrai a KSB1 (Sem Agrupamento) para um período arbitrário (não só um mês). Salva em `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Estudos\Estudo Duplicidade Pagamento\`.
  - `analisar_duplicidade_pagamento.py` — identifica pares de estorno (mesmo fornecedor/valor com sinal oposto) e aplica os critérios de duplicidade, gerando `Análise Duplicidade Pagamento.xlsx` na mesma pasta.
  - `analisar_zlfib_duplicidade.py` — checa NF duplicada na ZLFIB (Direção=Entrada, exclui Tipo NF 'R8'/transferência de material). Exporta a grade pra arquivo real em vez de ler via COM (ver `memory/errors/2026-08-13_zlfib_getcellvalue_dados_incorretos.md`).
  - `verificacao_mensal_zlfib.py` + `watcher_mensal_zlfib.bat` — **automação mensal ativa desde 2026-08-13**: tarefa agendada do Windows (`Verificacao_ZLFIB_Duplicidade_Mensal`, de hora em hora) que, no 1º dia útil do mês, assim que detecta o SAP logado, roda `analisar_zlfib_duplicidade.py` do mês anterior e manda e-mail (`juliana.silveira@pirelli.com`, automático, com anexo) só se achar duplicidade real. Detalhe completo em `memory/DECISOES.md` (entradas de 2026-08-13).
- **Critérios de duplicidade confirmados pela usuária (estudo KSB1):** Fornecedor+Valor+Documento de compras, e Fornecedor+Valor+Data de lançamento.
- **Resultado da primeira rodada do estudo KSB1 (01.01-31.07.2026):** 175.310 linhas no extrato, 7.147 com fornecedor, 50 pares de estorno excluídos. 228 grupos duplicados por Documento (R$ 823.535,87) e 649 grupos por Data (R$ 1.533.184,43) — triagem heurística, precisa revisão manual.
- **Resultado do estudo ZLFIB (mesmo período, corrigido em 2026-08-13):** nenhuma duplicidade real de NF nas 4 filiais (SJP, IBI, SOR, GOI) depois de corrigir o bug de leitura e excluir as notas de transferência de material (Tipo NF 'R8').
- **Tentativa de refinar com Nº de NF — NÃO deu certo, não repetir sem novo dado:** a coluna "Nº doc.de referência" da KSB1 não é a Nota Fiscal (confirmado pela usuária). Tentei cruzar esse número na transação ZLFIB (tem campos "Nr Documento" e "Nota Fiscal" de verdade), mas nenhuma busca (por número de documento, nem por fornecedor+período) trouxe resultado confiável — a própria usuária confirmou que não sabe como esse cruzamento funcionaria. **Pausado até alguém do time funcional/TI do SAP confirmar como ligar o documento de referência da KSB1 a uma Nota Fiscal na ZLFIB.** Não tentar de novo por tentativa e erro sem essa confirmação.

---

### Sub-projeto: Energia Elétrica Fitted
> Iniciado pela usuária em 2026-08-13. Objetivo: dois itens — (1) checar se todos os lançamentos de energia elétrica foram feitos corretamente/sem esquecimento, (2) mais importante, conferir se os créditos de PIS/COFINS/ICMS sobre energia estão sendo lançados (usuária acredita que não). Pasta: `scripts/sap/fitted_units/energia_eletrica_fitted/`.
- **Conta contábil identificada:** `N17002S001` ("COM Fix - Energia El[étrica]") — única conta de energia elétrica no extrato da KSB1, usada por 6 centros de custo diferentes.
- **Status:** exploração inicial feita (jan-jul/2026), achado concreto de lançamento faltando (centro de custo 8303/Fiat/Goiana sem nenhum lançamento em fev e jun/2026) e confirmado que nenhuma linha mostra detalhamento de PIS/COFINS/ICMS. Ver `memory/BRIEFING.md` (2026-08-13) para a tabela completa por centro de custo.
- **Mapeamento Centro de custo → unidade RESOLVIDO em 2026-08-18** (usuária mandou o print da lista completa de centros de custo por unidade — ver `ontology/fitted_units.json` → `centros_de_custo_por_unidade`): 8296→IBIRITE, 8290 e 8289→SJP, 8269 e 8292→SOROCABA, 8303→GOIANA.
- **Ainda não fechado:** conta contábil dos créditos de PIS/COFINS/ICMS (se existir separada), como a energia da Resende (RES) é lançada (não apareceu no extrato com os fornecedores conhecidos até agora).
- **Fornecedores de energia conhecidos (código de 10 dígitos, mesmo padrão da KSB1/ZLFIB):** CEMIG `4211308770`, CPFL `4211324097`, COPEL `4211333301`, FIAT AUTOMOVEIS S/A `4211330756` (revende energia pra Goiana, 2 notas — transmissão e repasse). Ignorar SERENA GERAÇÃO S.A `4211333021` (é rateio).

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
