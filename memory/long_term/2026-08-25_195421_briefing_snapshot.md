# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## Continuação 2026-08-25 — "Rateio de Custos" (Passo 5): PRIMEIRA VERSÃO FUNCIONANDO, testada contra Julho/Actual real, ainda falta polimento visual + integrar no cockpit

**Script criado:** `scripts/sap/fitted_units/fitted_units_despesas/gerar_rateio_custos.py` (roda standalone por linha de comando por enquanto: `--mes --ano --ciclo --pasta-saida`). Config do rateio: `ontology/rateio_gerencia.json` (entradas por `vigente_desde`, nunca hardcoded no script).

**Achados importantes desta sessão (todos por leitura, nada de produção foi alterado):**
1. **Fonte de verdade da classificação Variável/Fixo:** a Base Intermediária tem um campo "Tp.Custo" (coluna H da aba Intermediária) que às vezes vem em branco (achado real: TODAS as 161 linhas do SJP em Julho vieram em branco). A usuária explicou que isso é uma fórmula que só resolve quando o arquivo é aberto de verdade no Excel (busca no arquivo `Base_Contas_Contábeis_Fitted_22.xlsx`, rede, já documentado em `ontology/fitted_units.json` → `fonte_contas_gestoriais_e_centros`). **Solução implementada:** o script usa a Base Intermediária como fonte principal (pedido explícito da usuária) e, só quando o campo vem em branco, cai no Base_Contas_Contábeis como fallback (`carregar_var_fallback_por_conta`). Se nem assim resolver, a linha vai pra uma categoria visível "Não Classificado" — nunca some silenciosamente.
2. **Bug real encontrado e corrigido:** "Labour" existe como subcategoria tanto em Variable Cost quanto em Fixed Cost — a primeira versão do código somava os dois no mesmo balde (mesma chave de dicionário), duplicando o valor no Total Costs. Corrigido usando `(tipo, subcategoria)` como chave em vez de só a subcategoria. **Validação forte:** depois da correção, o total geral calculado pro Julho/Actual (R$ 6.655,04 mil) bateu EXATAMENTE com o valor já validado em sessões anteriores (ver `long_term/` ou histórico — "Actual R$ 6.655.041,91"), e o "Check" (TOTAL do quadro sem rateio vs. com rateio) fechou em 0,00.
3. **Mapeamento de unidades confirmado com a usuária:** SJP=0490, IBI=0491, GOI=0481, RES=0483 (mini-fábrica), Gerência=0499 — todas ativas. Sorocaba/Camaçari/Itatiaia = encerradas (mesma fonte que `gerar_base_intermediaria.py` já usa, `ontology/fitted_units.json` → `centros_de_custo_por_unidade`).
4. **Regra de resíduo de unidade encerrada implementada:** se aparecer custo numa unidade encerrada diferente de Sorocaba, soma no total da Gerência antes do rateio (e aparece também como linha extra "Resíduo Unidades Encerradas" no quadro sem rateio, só na coluna GER, pra o Check continuar fechando em 0). Sorocaba fica de fora do rateio (reclassificação pra custo não-recorrente). Em qualquer um dos casos, aparece um aviso detalhado (unidade/conta/descrição/valor) no arquivo gerado.
5. **~13 contas** (Materiais de Limpeza, Vigilância, Alimentação, Depreciação, etc.) aparecem com tipo Variável em alguns lançamentos reais, mesmo o arquivo antigo só listando elas como Fixo — caem em "Other Variable" (catch-all já esperado pra esse tipo de caso, não é erro).

**Estrutura do arquivo gerado (confirmada com a usuária antes de implementar):** tabela do rateio vigente no topo, aviso de resíduo de encerradas (se houver), quadro "sem rateio" (SJP\|IBI\|GOI\|RES\|GER\|TOTAL) e quadro "com rateio" (SJP\|IBI\|GOI\|RES\|TOTAL, com linha própria "Rateio Gerência"). Fórmulas de verdade no Excel pra Variable Cost/Fixed Cost/Total Costs/coluna TOTAL (não são só valores estáticos).

**Testado 2x contra a Base Intermediária REAL de Julho/Actual** (só leitura, saída na pasta de teste local `data/processed/fitted_units_despesas/rateio_custos_teste/`, nada de rede tocado) — 2 arquivos enviados pra usuária revisar (o segundo já com o bug do Labour corrigido).

**Pendências pra próxima sessão:**
1. **Usuária ainda não confirmou visualmente** se o arquivo gerado está com a formatação boa (cores/fontes como o mockup) — enviei mas não recebi feedback ainda nesta sessão.
2. **Integrar no cockpit** (`atualizar_ksb1_gui.py`): novo Passo 5, aba "⑤ Rateio de Custos", botão "Abertura e Rateio de Custo" — ainda não feito, o script só roda por linha de comando.
3. **Lembrete automático de Janeiro** (rateio geralmente muda nessa época) — combinado com a usuária, ainda não implementado.
4. Salvar o arquivo final no racional de sempre (`resolver_pasta_ciclo`) — a função já existe no script (`nome_com_versao` + pasta certa), só falta ligar no cockpit apontando pra rede em vez da pasta de teste.
5. Considerar testar com Ciclo Flash também (só Actual foi testado até agora).

**Nada commitado desta parte ainda** (script novo + `ontology/rateio_gerencia.json` não commitados).

---
## Continuação 2026-08-25 — NOVO PROJETO: "Rateio de Custos" (Passo 5) — em análise/scoping, ainda não implementado

**Pedido da usuária:** automatizar o arquivo `_Abertura custos Fitted Units <Mês> <Ciclo> <Ano>.xlsx` (hoje 100% manual — troca de link externo + mês toda vez). Vai virar um **Passo 5 novo no cockpit**, aba "⑤ Rateio de Custos", botão **"Abertura e Rateio de Custo"**.

**Engenharia do arquivo antigo, já mapeada (leitura, nada foi alterado):**
- 11 abas, maioria oculta (uma por unidade: SJP, Ibirité, Camaçari, Sorocaba, Goiana, Itatiaia + variantes "Sem Rodas"), todas larguíssimas (até coluna IG).
- Aba visível "Resumo Fitted Units": `H8` é o seletor de mês/Ciclo (ex: "Jul*") — é o que a usuária troca manualmente.
- Link externo de verdade (Excel "Edit Links") pra outro arquivo com abas "Relatório Faturamento" (não usado) e "Base_Cenários" — é o link que ela troca manualmente todo mês.
- Achado o mapeamento completo **conta contábil → categoria** (Variable Cost: Labour/Handling/Direct Materials/Transportation/Other Variable; Fixed Cost: Labour/Depreciation/IFRS16/Rents/Condominio/Other Fixed) — mais de 100 contas, extraídas linha a linha do arquivo real (ver histórico de comandos desta sessão se precisar re-extrair).
- Achado o mecanismo de rateio: `Base_Cenários!G55:H61` guarda a % por unidade (`%Rateio 1`/`%Rateio 2`); fórmula "R$ Rateio Staff" pega o custo da **Gerência** (mini-fábrica `0499`) do mês × essa %.
- Base Intermediária (`Intermediária!A:H`): colunas relevantes = Conta Gestorial(A)/Conta Fiscal(C)/Centro de Custo(E)/**Mini-Fábrica(F)**/Tp.Custo V-ou-F(H). Meses a partir da coluna I (January).

**Mapeamento de unidades (mini-fábrica → status), confirmado com a usuária:**
- SJP=0490, IBI=0491, GOI=0481, Gerência=0499 — todas **ativas**.
- Sorocaba=0496, Camaçari=0498, Itatiaia=0482 — **encerradas** (mesma lista já usada em `ontology/fitted_units.json` → `centros_de_custo_por_unidade`, chave por Centro de Custo, não por Mini-Fábrica).
- **RESENDE é ativa mas ainda sem custo real (unidade nova)** — **PENDENTE: preciso que a usuária me diga o código de Mini-Fábrica dela** (não achei em nenhum arquivo ainda, Julho não tinha nenhuma linha com os Centros de Custo dela: 8333/8348/8349/8350).

**Regras de negócio confirmadas com a usuária (importantes, não óbvias):**
1. % de rateio muda o mais comum em **Janeiro**, mas pode mudar fora de época também (ex: quando a RES entrou, GOI cedeu 4 pontos pro IBI+RES) — vai ficar guardada num arquivo de config próprio (não hardcoded), que a usuária avisa quando mudar. **Combinado criar um lembrete automático no cockpit**: se for Janeiro e o rateio salvo não tiver sido confirmado pra esse ano, avisar na tela.
2. Unidade encerrada **nunca** recebe rateio (sem faturamento).
3. Se aparecer custo residual numa unidade encerrada **diferente de Sorocaba**: soma no custo da Gerência antes de aplicar o rateio (Sorocaba está em reclassificação pra não-recorrente, fica de fora dessa soma).
4. Em qualquer caso de resíduo (Sorocaba ou não), o arquivo novo precisa mostrar um **aviso visível com as linhas/contas/valores** — não só logar.
5. **Não mexer em nada da Base Intermediária** (`gerar_base_intermediaria.py` fica intocado) — o script novo só lê.

**Estrutura do arquivo novo, confirmada com a usuária:**
1. Tabela do rateio vigente (% por unidade, com nota "vigente a partir de ...").
2. Quadro "por unidade, sem rateio" — colunas SJP | IBI | GOI | RES | **GER** | TOTAL.
3. Quadro "por unidade, com rateio" — colunas SJP | IBI | GOI | RES | TOTAL, com linha própria "Rateio Gerência" antes de Total Costs, e um "Check" mostrando que o TOTAL do quadro 2 bate com o TOTAL do quadro 3.
4. Mesma formatação visual do arquivo/mockup que a usuária desenhou (faixa azul escura, negrito, itálico, coluna TOTAL destacada, nota "'000 BRL").
5. Salva no mesmo racional de sempre: `resolver_pasta_ciclo(REDE_BASE/ano/MESES_PASTA[mes], mes, ciclo)` — Actual e Flash sem diferença nenhuma de lógica.

**Estado atual:** nada implementado ainda (só análise/leitura do arquivo real e do Base Intermediária, nenhum dado alterado). Pasta de teste local criada: `data/processed/fitted_units_despesas/rateio_custos_teste/`. Próximo passo assim que a usuária responder o código da Resende: escrever o script (provavelmente `gerar_rateio_custos.py`, mesma pasta dos outros passos), testar isolado, depois integrar no cockpit como Passo 5.

**Nada commitado desta parte ainda** (só a sessão anterior, popup SAPGUI, já está commitada — ver entrada acima/anterior deste arquivo).

---
## Sessão 2026-08-25 — Popup "Segurança SAPGUI"/travamento do Excel no Passo 1: FECHADO, confirmado ao vivo de ponta a ponta

**Pedido original da usuária:** o popup nativo do SAP que pede autorização a cada pasta nova ("Segurança SAPGUI") a incomoda. Tentativa de generalizar a regra em `saprules.xml` foi bloqueada 2x pelo classificador de segurança do Auto mode (arquivo fora da pasta do projeto), mesmo com autorização explícita dela. **Ideia da própria usuária, implementada:** SAP sempre exporta pra uma pasta fixa (`.../00.Extração Base KSB1/Temporario/`), e o código move o arquivo de lá pra pasta certa (`resolver_pasta_ciclo`) depois — como a pasta de destino do SAP nunca muda, só pede autorização 1x (não mais 1x/mês).

**Problema real encontrado ao testar ao vivo (não previsto):** o SAP abre o arquivo recém-exportado automaticamente no Excel (visível, não é instância isolada) — isso trava o arquivo (`WinError 32`) na hora de mover pra pasta final. Resolvido em várias rodadas de ajuste, todas em `ksb1_core.py`:
1. Retentativa (30x, 1s cada) ao mover, em `extrair_um`.
2. `fechar_excel_se_aberto` — fecha só a aba certa via COM, buscando direto na Running Object Table (`pythoncom`) pelo caminho do arquivo (não por `GetObject(Class=...)`, que é ambíguo com múltiplas instâncias do Excel abertas). Testado isolado com 2 instâncias simultâneas — fecha só a certa, sem tocar na outra.
3. Limpeza da sobra na pasta `Temporario` (antes de exportar de novo) também virou resiliente a arquivo travado.
4. `limpar_excel_orfao` — fecha o aviso residual "Sorry, we couldn't find..." (efeito colateral inofensivo: o Excel que o SAP abriu fica "órfão" porque o código moveu o arquivo dali) e a janela vazia que sobra. **Bug real encontrado testando ao vivo:** a versão moderna do Office desenha esse aviso numa janela classe `NUIDialog` (não o `#32770` clássico), e o botão "OK" é só um desenho dentro de um único controle `NetUIHWND` — não dá pra "clicar" nele feito botão de verdade. Corrigido fechando a janela inteira via `WM_CLOSE` (mesmo efeito do X), em vez de procurar um botão.

**CONFIRMADO AO VIVO pela usuária, ponta a ponta (Agosto/2026, Ciclo Flash):** rodou "Extrair KSB1" sem clicar em nada — os 2 avisos do Excel fecharam sozinhos, a janela vazia fechou sozinha, e os 2 arquivos (Gestoriais + Sem Agrupamento) caíram certinho em `08_Aug_Flash` (nome limpo, sem versionamento espúrio). **Item fechado.**

**Cuidado descoberto durante o debug, importante pra próximas sessões:** o `Get-ChildItem` do PowerShell, quando eu (Claude) uso pra checar essa pasta de rede específica, às vezes retorna resultado desatualizado ("pasta vazia" quando na real tinha arquivo lá) — me confundiu 2x durante essa sessão, achando que o `move` tava falhando quando na real já tinha funcionado. `[System.IO.Directory]::GetFiles(...)` via .NET direto deu resultado correto nas duas vezes que testei em paralelo. Se for checar essa pasta de novo, preferir `.GetFiles()` .NET ou pedir confirmação visual da usuária (print do Explorer) em vez de confiar só num `Get-ChildItem` isolado.

**Arquivos de teste que sobraram do debug, ainda na rede:** `08_Aug_Actual` ficou com `_v2`/`_v3`/`_v4` de uma rodada anterior de testes (dados reais de Agosto/Actual, mas duplicados) — a usuária disse ter limpado a pasta em algum momento da sessão, mas vale conferir/confirmar com ela se sobrou algo antes do fechamento real de Agosto.

**Detalhe técnico completo em `memory/DECISOES.md` → "2026-08-25 — Popup Segurança SAPGUI".**

**ALERTA DE SESSÃO LONGA disparou nesta sessão (45 ações) — backup automático já rodou.** Usuária pode fechar esta janela e abrir uma nova pra continuar; contexto já está salvo aqui.

**Pendência pra próxima vez que a usuária rodar o Passo 1 de verdade:** confirmar que (a) o arquivo final cai certo na subpasta do Ciclo e (b) o popup só aparece 1x (pra autorizar a pasta `Temporario`) ou nem aparece mais.

---
## Sessão 2026-08-25 — Teste pendente (extração real contra o SAP) FECHADO com sucesso

**Retomada:** pendência única deixada em 2026-08-24 — testar contra o SAP real se a extração (Passo 1) passa a salvar dentro da subpasta de Ciclo (`<MM>_<Mês3>_<Ciclo>`) em vez de solta na pasta do mês.

**Rodado pela usuária pela GUI real (cockpit), Agosto/2026, Ciclo Flash.** Confirmado via PowerShell na rede:
- `.../2026/00.Extração Base KSB1/08 - Aug/08_Aug_Flash/KSB1 - Fitted Units 08.2026 - Gestoriais - Flash.XLSX` (25/08/2026 09:50:30)
- `.../08_Aug_Flash/KSB1 - Fitted Units 08.2026 - Sem Agrupamento - Flash.XLSX` (25/08/2026 09:51:11)
- Raiz `08 - Aug/` sem nenhum arquivo solto (formato antigo) — confirmado limpo.
- `08_Aug_Actual/` seguiu vazia (só o Flash foi extraído nesse teste).

**Item fechado.** A mudança do Passo 1 (subpasta de Ciclo, commit `218e72c`) está validada de ponta a ponta: unitário (4 cenários, 2026-08-24) + agora contra SAP real. Nenhum dado de outro mês/Ciclo foi tocado.

**Pendência residual, sem urgência:** validar o mesmo pro Ciclo Actual (só Flash foi testado ao vivo até agora) — não bloqueante, a lógica é idêntica (mesma função `resolver_pasta_ciclo`), só não foi exercitada nesse Ciclo especificamente.

---
## RESUMO DO DIA 2026-08-24 — Watchdog de travamento fechado; Passo 1 (extração) passa a usar subpasta de Ciclo

**Retomada do dia:** único item pendente deixado de 2026-08-22 (watchdog/timeout do Excel). Fechado logo no início do dia; o resto da sessão foi um pedido novo (estrutura de pastas do Passo 1).

### 1. Watchdog de travamento (Excel isolado, 12 min) — implementado, corrigido e validado ao vivo
- Se uma operação com Excel (Passos 3/4) passar de 12 min sem terminar, avisa e oferece forçar o encerramento só da instância isolada (PID próprio, não afeta outro Excel da usuária). SAP (Passo 1) e Check (Passo 2) só avisam, sem oferecer matar nada — decisão explícita pra não arriscar fechar outras sessões do SAP.
- Bug real achado no próprio teste: o aviso Sim/Não usava botões em inglês (`messagebox.askyesno` padrão do Tk) — corrigido com diálogo customizado em português (`_perguntar_sim_nao`).
- Validado com simulação de travamento real (fora do repo, sem tocar dado real): confirmado por monitoramento de processo que só a instância isolada do teste foi encerrada — o Excel real da usuária, aberto no fundo o tempo todo, nunca foi tocado.
- Commits `77b6295` e `9063c30`, ambos no GitHub.

### 2. Passo 1 (extração) passa a salvar em subpasta de Ciclo, com fallback pro formato antigo
- A pedido da usuária, a extração (Passo 1) passou a salvar dentro de `<MM>_<Mês3>_<Ciclo>/` (mesmo padrão já usado pelos Passos 3/4), em vez de solta direto na pasta do mês.
- Meses já extraídos (Jan/Fev/Mar/Jun/Jul, soltos) **não foram reorganizados** — decisão explícita da usuária. Nova função `localizar_extracao_ksb1` procura primeiro na subpasta nova e cai pro formato antigo se não achar. Check (Passo 2) e leitura do Passo 3 atualizados pra usar essa mesma busca.
- Testado com 4 cenários em pasta temporária local (formato antigo, formato novo, nenhum dos dois, os dois coexistindo) — todos passaram. **Ainda não testado contra o SAP real.**
- As 24 subpastas (12 meses × Actual/Flash) já foram criadas vazias na rede, a pedido da usuária — confirmado com PowerShell em 2 meses (Julho e Dezembro), nada existente foi tocado.
- Commit `218e72c`, no GitHub.

### Pendência pra amanhã (2026-08-25) — ver bloco "PRÓXIMA SESSÃO" acima
**Testar a extração de verdade contra o SAP**, assim que a usuária conectar, pra confirmar que o arquivo cai na subpasta certa. Sugestão: Agosto/2026 (mês ainda vazio).

### Nenhum dado real de produção foi alterado incorretamente hoje
Todas as validações rodaram em pasta temporária local ou contra uma instância Excel isolada de teste; a única escrita real na rede foi a criação das 24 subpastas vazias (sem tocar nenhum arquivo existente), feita a pedido explícito da usuária.

---
## Continuação 2026-08-24 (mesmo dia) — Passo 1 (extração) passa a usar subpasta de Ciclo, com fallback pro formato antigo — implementado, testado isolado, falta validar ao vivo

- **Pedido da usuária:** hoje o Passo 1 (extração KSB1) salva os arquivos soltos direto na pasta do mês (`00.Extração Base KSB1/<MM - Mês>/`). Ela pediu pra usar dentro de cada pasta de mês subpastas de Ciclo (Flash/Actual), **"utilizando a mesma premissa"** já usada pelos Passos 3/4 (`<MM>_<Mês3>_<Ciclo>`, ex: `07_Jul_Actual` — ver `resolver_pasta_ciclo` em `ksb1_core.py`).
- **Confirmado com a usuária antes de implementar:** os arquivos JÁ extraídos (soltos, sem subpasta — todos os meses com dado real hoje: Jan/Fev/Mar/Jun/Jul) **não são movidos/reorganizados**. Só as extrações NOVAS passam a usar a subpasta; o código passa a procurar primeiro na subpasta e, se não achar, cai pro formato antigo (mesma pasta do mês) — igual ao padrão de tolerância já usado em outras partes do sistema (ex: `resolver_pasta_ciclo` pra pastas com mês por extenso, `encontrar_arquivo_ksb1` pra nome sem Ciclo).
- **Implementado:**
  - `ksb1_core.py`: nova função `localizar_extracao_ksb1(pasta_mes, bu_nome, mes, ano, agrup_label, ciclo)` — tenta achar o arquivo primeiro em `resolver_pasta_ciclo(pasta_mes, mes, ciclo)` (a subpasta do Ciclo); se não achar nada lá, tenta a `pasta_mes` direto (formato antigo); se não achar em nenhum dos dois, erro claro citando os dois caminhos tentados.
  - `atualizar_ksb1_gui.py` (`extrair_um`, Passo 1): a extração agora salva em `resolver_pasta_ciclo(pasta_mes, mes, ciclo)` em vez de direto em `pasta_mes` (com `mkdir(parents=True)` — a subpasta é criada na primeira extração daquele mês/Ciclo).
  - `check_agrupamentos_ksb1.py` (`gerar_check`, Passo 2): trocado `encontrar_arquivo_ksb1` direto por `localizar_extracao_ksb1` (acha nos dois formatos). O arquivo de saída "Check de agrupamentos" agora é salvo **junto da extração que ele conferiu** (`arquivo_gest.parent` — cai na subpasta pros meses novos, na pasta do mês pros antigos), em vez de sempre na pasta do mês.
  - `gerar_ksb1_mensal.py` (`decidir_fonte_e_ler_linhas`, Passo 3): mesma troca, `encontrar_arquivo_ksb1` → `localizar_extracao_ksb1`.
- **Testado:** `py_compile` + import real dos 4 arquivos editados — sem erro. Teste isolado (pasta temporária local, não é o repo nem a rede) com 4 cenários: (1) só formato antigo (solto) → acha via fallback; (2) só formato novo (subpasta) → acha direto; (3) nenhum dos dois existe → erro claro citando os 2 caminhos tentados; (4) os dois coexistem (hipotético, ex: alguém migrou um mês na mão) → prefere a subpasta nova. **Todos os 4 passaram.**
- **NÃO testado ao vivo ainda** — não rodei a extração de verdade contra o SAP/rede com essa mudança, nem o Check/Passo 3 lendo um mês novo de verdade. Nenhum dado real/rede foi tocado (só leitura da estrutura de pastas real, feita antes de implementar, e testes em pasta temporária local).
- **Continuação (mesmo dia):** a usuária foi conferir a rede antes de rodar qualquer extração nova e perguntou por que não achou a subpasta — esclarecido que o código só cria a subpasta na hora que a extração roda de verdade (não retroativo). Ela pediu explicitamente pra **criar as 24 subpastas vazias agora** (12 meses × Actual/Flash) na rede, adiantando a estrutura. Feito com um script pontual que importa `REDE_BASE`/`MESES_PASTA`/`resolver_pasta_ciclo` direto de `ksb1_core.py` (garante nome idêntico ao que o código de produção vai procurar depois) e chama `.mkdir(parents=True, exist_ok=True)` pra cada `<MM>_<Mês3>_<Ciclo>/` dentro das 12 pastas de mês. **Confirmado com PowerShell (`Get-ChildItem`) em 2 meses (Julho e Dezembro):** as 24 subpastas foram criadas certinho, vazias, e nenhum arquivo/pasta existente foi tocado (os arquivos antigos soltos em Julho continuam lá, intactos).
- **Próximo passo:** quando a usuária for extrair o próximo mês de verdade (ou quiser um teste dirigido), confirmar que a extração escreve dentro da subpasta certa e que o Check/Passo 3 acham o arquivo novo sem problema. Meses antigos (Jan/Fev/Mar/Jun/Jul) devem continuar funcionando exatamente como antes (fallback pro formato solto).
- **Sessão atingiu o limite de 45 ações de novo — backup automático já rodou** (`session_transition.py`). Código ainda não commitado nesta parte (as 24 subpastas na rede já estão criadas, independente do commit).

---
## Sessão 2026-08-24 — Watchdog/timeout do Excel travado: implementado, corrigido e TESTADO AO VIVO — item fechado

- **Retomada:** único item pendente deixado explicitamente pra hoje em 2026-08-22 — hoje, se o Excel travar de verdade (hang, não erro) durante uma automação COM (Passos 3/4), o processo COM fica preso indefinidamente, sem timeout nem aviso, janela fica com botões desabilitados pra sempre.
- **Design confirmado com a usuária antes de implementar (via pergunta direta):**
  1. Timeout de aviso: **12 minutos** sem a operação terminar.
  2. Quando o aviso aparecer: **avisa + oferece forçar o encerramento só da instância isolada do Excel** (processo próprio, capturado por PID, não afeta outro Excel aberto por ela). No Passo 1 (SAP, sem Excel) e Passo 2 (Check, só openpyxl, sem COM): **só aviso, sem oferecer matar nada** — decisão explícita de não automatizar encerrar o SAP GUI, porque mataria TODAS as sessões abertas dele, não só a desta automação.
- **Implementado:**
  - `ksb1_core.py`: nova função `abrir_excel_isolado(log, pid_callback=None)` — centraliza a abertura do Excel isolado (`DispatchEx`, `Visible=False`) que já existia repetida em 4 lugares, e captura o PID do processo via `win32process.GetWindowThreadProcessId(excel.Hwnd)` (funciona mesmo invisível — a janela existe, só não aparece). Se `pid_callback` for passado, é chamado com o PID assim que capturado.
  - `gerar_ksb1_mensal.py` (1 ponto) e `gerar_base_intermediaria.py` (3 pontos: `lancar_provisoes`, `atualizar_provisoes`, `atualizar_base_intermediaria`) — todos os `win32com.client.DispatchEx(...)` diretos substituídos por `abrir_excel_isolado(log, pid_callback)`; as 4 funções de topo ganharam parâmetro opcional `pid_callback=None`, repassado adiante.
  - `atualizar_ksb1_gui.py` — `rodar_em_thread` reescrita: `func` agora recebe `(log, pid_callback)` (todos os 6 `ao_clicar_*` atualizados); um relógio (`time.monotonic()`) mede quanto tempo a operação está rodando; se passar de `TIMEOUT_AVISO_SEGUNDOS` (12 min) ainda rodando, mostra `_avisar_travamento` (dedup por intervalo de 12 min, não fica repetindo a cada 150ms). Se a operação usa Excel (`permite_forcar_excel=True`, é o padrão) e já temos o PID capturado, o aviso vira uma pergunta Sim/Não — Sim mata o processo (`psutil.Process(pid).terminate()`, com checagem de que o processo ainda se chama `EXCEL.EXE` antes de matar, proteção contra PID reciclado) e libera a janela na hora (`estado["abandonado"]=True`); Não só fecha o aviso e continua esperando. Passo 1 (SAP) e Passo 2 (Check) passam `permite_forcar_excel=False` explicitamente — só aviso informativo, nunca oferece matar nada.
  - Cuidado de sincronização: como a thread trava pra sempre em daemon (Python não mata thread à força), depois de "abandonar" a operação a `checar()` para de fazer polling da UI, mas a thread pode eventualmente terminar sozinha (quando a chamada COM trava percebe que o processo morreu e levanta erro) — nesse caso só loga uma linha discreta, não reabre popup de conclusão/erro duplicado.
  - `psutil` adicionado a `requirements.txt` e instalado no ambiente local.
- **Bug real achado e corrigido no próprio teste:** o aviso Sim/Não usava `messagebox.askyesno`, que no Tk sempre mostra botões em **inglês** ("Yes"/"No") — quebrava a regra do projeto de manter tudo em português (REGRAS_RAPIDAS #11). Corrigido com diálogo customizado `_perguntar_sim_nao` (Toplevel próprio, botões "Sim"/"Não" de verdade, mesmo estilo visual do resto da GUI).
- **Testado AO VIVO, ponta a ponta, sem tocar rede/dados reais:** script de teste isolado (fora do repo, scratchpad) trocou `gerar_ksb1_mensal.gerar_ksb1_mensal` por uma versão falsa que abre uma instância Excel isolada REAL (mesma chamada de produção) e trava de propósito; `TIMEOUT_AVISO_SEGUNDOS` reduzido pra 15s só no teste; clique automático (travessia da árvore de widgets Tk) simulou clicar "Atualizar Pivot KSB1" e depois "Sim" no aviso, sem precisar de interação manual. Confirmado por monitoramento direto de processos: apareceu um 2º `EXCEL.EXE` (isolado) que foi encerrado sozinho ~15s depois, **o Excel real da usuária (aberto no fundo o tempo todo, arquivo de trabalho dela) nunca foi tocado** (mesmo PID do início ao fim), a janela do cockpit voltou ao normal (botões/cursor recuperados) sem erro, e o diálogo em português foi encontrado e clicado corretamente pela automação. Detalhe completo em `memory/DECISOES.md` → "2026-08-24 — Watchdog de travamento".
- **Item fechado.** Único ponto pendente de 2026-08-22 (watchdog/timeout) está implementado, corrigido e validado. Nenhum dado real/rede foi tocado em nenhuma parte deste trabalho — só edição de código e testes isolados locais/scratchpad.
- Commit `77b6295` (implementação inicial) já enviado ao GitHub. Fix do diálogo em português (`_perguntar_sim_nao`) ainda não commitado nesta sessão — próximo passo imediato.

---
## Continuação 2026-08-22 (mesmo dia) — Cockpit: barra de progresso com pneuzinho girando, cursor de mãozinha e cursor "ocupado" — testado e aprovado pela usuária

- **Motivação:** testando o Ciclo Flash/Forecast pela GUI de verdade (pela primeira vez, ver seção anterior), a usuária notou a janela "travando" durante operações longas (Excel via COM) — na verdade não travava, só não tinha feedback visual (mesmo problema mencionado em sessões anteriores, nunca resolvido).
- **Pedido da usuária:** barra de progresso "na parte de baixo", depois refinado pra "não precisa de %" (indeterminada tá bom) + "um ícone rodando" (spinner) + cursor "ocupado" durante processamento + cursor de mãozinha ao passar o mouse sobre qualquer botão.
- **Mudança de arquitetura em `atualizar_ksb1_gui.py`:** as 6 operações (Extrair KSB1, Check, Lançar/Atualizar Provisões, Atualizar Pivot KSB1, Finalização) agora rodam numa `threading.Thread` separada (com `pythoncom.CoInitialize()`/`CoUninitialize()` pra COM do SAP/Excel funcionar isolado por thread) via um helper novo, `rodar_em_thread(descricao, func, ao_concluir)` — mantém a janela respondendo durante a operação. `messagebox` só é mostrado de volta na thread principal (`ao_concluir`, agendado via `root.after`), nunca dentro da thread de trabalho.
  - `rodar()` (Passo 1/Extração) refatorada pra não chamar `messagebox` direto (não dava pra rodar em thread) — segue o mesmo padrão das outras funções (recebe `log`, levanta exceção; nova classe `ErroComTitulo` preserva o título específico de cada erro pro messagebox).
  - Log agora é thread-safe via fila (`queue.Queue`) — a função `log()` só enfileira, quem escreve no widget de verdade é `_drenar_fila()`, chamada em loop (`root.after(80, ...)`), sempre na thread principal.
- **Indicadores visuais adicionados:**
  1. Cursor da janela vira "watch" (ocupado) durante qualquer operação, volta ao normal ao terminar.
  2. Cursor "hand2" (mãozinha) em todos os botões, sempre (não só durante processamento).
  3. Texto de status com ícone girando (spinner Unicode, ex: "⠋ Processando: Finalizando a Base Intermediária...") — **movido pro cabeçalho** (canto superior direito, área escura fixa), não no corpo — o corpo tem altura variável (descrição de cada aba + log expansível) e o texto ficava espremido pra fora da janela visível, foi o motivo da usuária não ter visto ele nos primeiros testes.
  4. Barra de progresso indeterminada (sem %, só anima) — também movida pra logo abaixo do trim amarelo do cabeçalho, **sempre visível** (não usa mais `pack`/`pack_forget`, só `start()`/`stop()`), pelo mesmo motivo do item 3.
  5. `root.update_idletasks()` forçado antes de iniciar a thread, pra garantir que o estado "processando" apareça mesmo se a operação terminar rápido (ex: Excel já "aquecido" de uma rodada anterior).
- **Bug real encontrado e corrigido durante os próprios testes desta sessão:** a primeira versão (status/barra dentro do corpo, embaixo do log) simplesmente não aparecia pra usuária — o log widget tem `expand=True` e a janela (800px de altura) não tinha espaço sobrando quando a descrição da aba era longa, empurrando o indicador pra fora da área visível. Corrigido movendo os dois pro cabeçalho/topo da janela (área que nunca encolhe).
- **Pneuzinho com calota (pedido final da usuária):** o indicador de "processando" não é o bloco padrão do `ttk.Progressbar` — é um `Canvas` (`ALTURA_BARRA_PROGRESSO = 32`) com um ícone de pneu desenhado via PIL em runtime (`_gerar_frames_pneu`, banda preta com sulcos + calota prateada com 6 raios + miolo escuro com detalhe amarelo claro), 12 frames pré-rotacionados, animando (gira + desliza de um lado a outro da barra) via `root.after(40, ...)`. **Pillow adicionado a `requirements.txt`** (nova dependência). **Aprovado pela usuária** ("ficou maravilhoso") depois de 2 ajustes: calota (pedido explícito, versão inicial era só um círculo liso) e barra maior (32px em vez de 20px, pra ficar mais visível).
- **Metodologia de teste usada (sem tocar rede):** script fora do repo (`lancar_cockpit_teste.py`, scratchpad da sessão) importa `atualizar_ksb1_gui`, troca só o atributo `REDE_BASE` desse módulo (não o de `ksb1_core.py` nem o de `gerar_base_intermediaria.py`, que ficam intocados) pra uma pasta local (`data/processed/fitted_units_despesas/gui_teste_rede/`) antes de chamar `main()` — assim toda ESCRITA da GUI cai na pasta local, mas as LEITURAS de fonte (KSB1, Fast Provisão, Forecast, template do mês anterior) continuam vindo da rede real, porque `gerar_base_intermediaria.py` importou seu próprio `REDE_BASE` separadamente. Testado com sucesso o fluxo "Atualizar Pivot KSB1" (Julho, Flash) rodando pela GUI de verdade, log real aparecendo, janela responsiva durante o processamento.
- **Commitado nesta sessão** (junto com o quadro de comparação Forecast). **Nenhum dado real de produção foi tocado** — todo teste rodou na pasta local via o script de teste com REDE_BASE trocado.

### Continuação (mesmo dia) — Fallback de Agosto (R8→R7) testado e confirmado pela GUI real, sem tocar rede

- Como Agosto/2026 ainda não tem nenhum dado real na rede (fechamento real ainda não aconteceu), o teste "emprestou" o KSB1 e a Base Intermediária reais de julho (só leitura), rotulados como agosto via monkeypatch pontual (só intercepta a busca do KSB1 — a busca do Forecast R8/R7 continuou 100% real). Detalhe completo em `memory/DECISOES.md` → "2026-08-22 (continuação) — Fallback de Agosto".
- **Resultado: popup apareceu certinho** — "Fechamento de August/2026: não existe Forecast R8 — usei o Forecast de July (R7) como comparação."
- **Item de teste fechado.** O quadro de comparação Forecast (implementado mais cedo hoje) está validado em todas as frentes: lógica de valores, rótulos, GUI ao vivo, e agora o fallback R8→R7.
- Arquivos de teste (`gui_teste_rede/`) removidos depois da validação — nada é dado oficial, nada foi escrito na rede.
- **Pendências que seguem em aberto** (nenhuma tocada nesta parte): Janeiro/Budget-MP, Faturamento (linha 25), inserção automática de linha colorida se provisões excederem capacidade.

### Continuação (mesmo dia) — Janeiro/Budget-MP implementado e validado — fecha o quadro de comparação Forecast por completo

- **Confirmado com a usuária:** Budget = MP (mesma coisa). Caminho: `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Management Plan\MP <ano>\P&L Fitted Units_Budget<AA>_.xlsx` — **mesmo ano** que está sendo fechado (não o anterior; o MP é preparado no fim do ano anterior mas arquivado sob o ano que cobre — ex: MP27 é feito no fim de 2026). Estrutura idêntica ao Forecast (mesma aba, mesmas linhas 19/20/30/31/38), confirmado inspecionando ao vivo.
- **Implementado:** `localizar_arquivo_budget` + `localizar_forecast_para_comparacao` reescrita (mes=1 vira caso especial, direto pro Budget). Rótulo do quadro vira "Budget" (não "Forecast") quando mes=1.
- **Validado pela GUI real, sem tocar rede** — Janeiro/2026 já tinha fechamento de verdade, então copiei a Base Intermediária Flash REAL de Janeiro pra pasta de teste local (não precisou emprestar dado de outro mês, diferente do teste de Agosto). Popup apareceu certinho, rótulo "Budget" e valor (R$ 2.983.872,25) bateram exatamente com o que a usuária já tinha calculado à mão, na época, com o rótulo "MP'26". Detalhe completo em `memory/DECISOES.md` → "2026-08-22 (continuação) — Janeiro/Budget-MP".
- **Achado à parte, sem impacto:** pasta `01_Jan_Forecast` na rede tem arquivos com "R1" no nome, mas nenhum no padrão que a automação procura — não interfere.
- **Corrigido na ontologia:** uma suposição errada anterior ("Budget/MP do ano anterior") foi trocada por "mesmo ano", confirmado explicitamente pela usuária.
- **Todos os cenários do quadro de comparação Forecast agora estão fechados e testados:** Actual, Flash normal, fallback R8/R12, e Janeiro/Budget.
- **Pendências que seguem em aberto** (sem relação com este quadro): Faturamento (linha 25), inserção automática de linha colorida se provisões excederem capacidade.

### Continuação (mesmo dia) — Faturamento adiado por decisão da usuária; inserção automática de linha amarela implementada + bug real corrigido

- **Faturamento (linha 25):** usuária decidiu explicitamente deixar de lado por enquanto — "hoje quem me envia é o time da Fitted mesmo, eu sou da controladoria... pretendo primeiro otimizar meu processo e depois rever isso." Não é mais prioridade da lista de pendências, só retomar se ela pedir.
- **Bug real achado (não relacionado ao pedido original) investigando as cores das linhas coloridas:** o cálculo de "capacidade" de provisões e a limpeza de conteúdo (`limpar_provisoes`) contavam TODA a área colorida (amarelo+verde+roxo) como se fosse capacidade de provisão, e `limpar_provisoes` apagava até a fórmula "molde" da linha roxa antes de tentar copiá-la. Não causou dano real ainda (verde/roxo estavam vazios em todos os testes), mas ia aparecer no primeiro mês com reclassificação de verdade. Corrigido com detecção de cor real via `Interior.Color` (amarelo=65535), não `Pattern`/`ColorIndex` (não discriminam bem). Detalhe completo em `memory/DECISOES.md` → "2026-08-22 (continuação) — Bug real corrigido + inserção automática".
- **Inserção automática implementada:** quando as provisões do mês excedem as linhas amarelas disponíveis, insere linhas amarelas novas automaticamente (nunca verde/roxo), sempre antes da primeira verde, via `Rows().Insert()` nativo do Excel (não copia/cola manual). Confirmado com a usuária: sempre amarela, frequência imprevisível mas real ("pode acontecer todos os meses"), sempre antes das verdes, roxa mantém só 1 linha de propósito como molde.
- **Validado com sucesso numa cópia local do arquivo real de julho** (regressão sem forçar inserção + teste forçando 3 linhas novas) — verde/roxo deslocaram certinho sem perder nada, e a fórmula molde da roxa se auto-ajustou sozinha (confirma que usar `Insert()` nativo, não manipulação manual, foi a escolha certa).
- **Nenhum dado real de produção foi tocado** — todos os testes rodaram em cópias locais.

### Continuação (mesmo dia) — Linhas verdes deprecadas (decisão final da usuária)

- **Decisão:** bloco verde (reclassificações) nunca mais será usado — ela já reclassifica direto no SAP antes do fechamento, então o mecanismo ficou obsoleto.
- **Não vamos automatizar remoção física das linhas** — proposto e aceito por ela: já que a automação ignora 100% esse bloco (só desloca posição quando insere amarela nova), não vale o risco de mexer em estrutura de arquivo por um ganho cosmético. Se ela quiser, apaga manualmente uma vez no arquivo-modelo atual.
- **Roxa continua sendo mantida** (é o "molde" de fórmula, papel diferente da verde).
- Detalhe em `memory/DECISOES.md` → "2026-08-22 (continuação) — Linhas verdes deprecadas".
- **Todos os itens da lista de pendências de Fitted Units Despesas estão fechados por ora** (Faturamento adiado por decisão dela, linha colorida com inserção automática implementada, verde deprecada). Próximo passo fica em aberto pra usuária decidir.

### Continuação (mesmo dia) — Revisão completa do processo (Actual + Flash), pedida pela usuária — EM ANDAMENTO

- **Pedido da usuária:** revisar TODO o processo de fechamento (Teste, Actual, Flash), achar fragilidades/erros, testar, deixar tudo certo — "preciso que isso esteja funcionando no fechamento."
- **Achado real (em correção agora):** 3 pontos do código buscam arquivo pelo NOME EXATO (sem considerar `_v2`/`_v3`), inconsistente com o resto do sistema que sempre pega a versão mais recente — `localizar_ksb1_actual_anterior` (gerar_ksb1_mensal.py), `localizar_base_ksb1_do_mes` e `localizar_base_intermediaria_mes_anterior` (gerar_base_intermediaria.py). Risco real: se algum passo for rerodado pra corrigir algo (gera `_v2`), os passos seguintes (mesmo mês ou mês seguinte) continuariam lendo a versão antiga em silêncio.
- **Corrigido:** nova função `encontrar_arquivo_mais_recente(pasta, nome_base)` em `ksb1_core.py`, aplicada nas 4 leituras que tinham o problema (KSB1 mês anterior, BASE_KSB1 do mês, Base Intermediária mês anterior, Base Intermediária Flash do mesmo mês). Testada isoladamente.
- **Testado de ponta a ponta, Julho/2026, Actual E Flash, sem tocar rede:** Pivot KSB1 (Actual e Flash), Finalização (Actual e Flash), Lançar/Atualizar Provisões — tudo rodou sem erro, todos os valores bateram exatamente com o que já era conhecido (Actual R$ 6.655.041,91 / Flash R$ 6.760.751,46, quadros de comparação idênticos aos validados antes). Log confirmou que a correção de cor das linhas (sessão anterior) está ativa.
- **Não foi possível testar o Passo 1 (extração SAP) de ponta a ponta** — depende de sessão SAP ao vivo. Revisão de código não achou problema óbvio ali.
- **Risco conhecido, não resolvido:** sem timeout/vigia se o Excel travar de verdade (hang, não erro) durante uma automação — já documentado desde 2026-08-14, baixa probabilidade.
- Detalhe completo em `memory/DECISOES.md` → "2026-08-22 (continuação) — Revisão completa do processo de fechamento". **Commitado.**

---
## PRÓXIMA SESSÃO (retomar na segunda-feira) — usuária pediu explicitamente

**Primeira coisa a fazer:** resolver o risco conhecido do "vigia"/timeout — hoje, se o Excel travar de verdade (hang, não erro) durante qualquer automação (Passo 3 ou 4, Actual ou Flash), o processo COM fica preso indefinidamente, sem timeout nem aviso. A usuária confirmou explicitamente no fim desta sessão ("vamos fazer isso, me notifica da próxima vez que eu voltar pra resolver esse risco conhecido") que quer resolver isso — não implementado ainda, é o próximo item da lista.

**Estado do projeto ao encerrar esta sessão (2026-08-22):** todo o processo de fechamento (Passos 1-4, Actual e Flash) foi revisado, testado de ponta a ponta e está funcionando — ver resumo completo logo abaixo. Nenhuma pendência funcional aberta além do watchdog acima (Faturamento foi adiado por decisão dela, linhas verdes deprecadas por decisão dela, Janeiro/Budget-MP e inserção automática de linha implementados e validados).

**Tudo commitado e enviado ao GitHub** (commit `2493fef`, push confirmado pela usuária via `! git push`).

---
## RESUMO DO DIA 2026-08-22 — Quadro de comparação Forecast/Budget completo, cockpit com feedback visual, bug de linhas coloridas corrigido, revisão geral aprovada

**Sessão longa (vários alertas de 45 ações) — resumo consolidado de tudo, na ordem em que aconteceu. Detalhe técnico completo de cada item em `memory/DECISOES.md`, todas datadas 2026-08-22.**

### 1. Quadro de comparação (linhas 18/19) completo pro Ciclo Flash
- Implementado `atualizar_comparacao_forecast`: pro Ciclo Flash, compara contra o Forecast mais recente (nomenclatura Pirelli R+mês=REFRESH) em vez de outro Flash. Fallback pro mês anterior quando R<mês> não existe (R8/R12), com popup avisando qual foi usado.
- **Janeiro (sem R1) resolvido:** usa direto o Budget/MP do mesmo ano (`\\...\Management Plan\MP <ano>\...`) — confirmado que MP é preparado no fim do ano anterior mas arquivado sob o ano que cobre.
- Validado com dado real: julho (R7) e janeiro (Budget) bateram exatamente com os valores que a usuária já tinha calculado manualmente. Fallback de Agosto (R8→R7) testado e confirmado pela GUI real.

### 2. Cockpit: feedback visual durante processamento
- Operações passaram a rodar em thread separada (janela não trava mais durante Excel/SAP via COM).
- Indicadores: cursor "ocupado", cursor de mãozinha nos botões, e um pneuzinho Pirelli (com calota, desenhado via PIL) girando/deslizando numa barra sempre visível no cabeçalho — aprovado pela usuária ("ficou maravilhoso").

### 3. Bug real corrigido: linhas coloridas (provisões)
- Capacidade de provisões contava TODA a área colorida (amarelo+verde+roxo) por engano, e "Atualizar Provisões" apagava a fórmula "molde" da roxa antes de copiá-la. Sem dano real até agora (verde/roxo sempre vazios), mas ia quebrar no primeiro mês com reclassificação de verdade.
- Corrigido com detecção de cor real via `Interior.Color`. Implementada inserção automática de linha amarela quando as provisões excedem a capacidade (nunca mexe em verde/roxo) — testado inserindo linhas numa cópia real, tudo deslocou certinho.
- **Linhas verdes deprecadas** por decisão da usuária (reclassificações já acontecem no SAP antes do fechamento) — não automatizamos remoção, só ficam ignoradas.
- **Faturamento (linha 25) adiado** por decisão da usuária (time da Fitted envia, ela quer otimizar o processo primeiro).

### 4. Revisão completa do processo, pedida pela usuária
- Achado e corrigido 1 bug real: 4 leituras de arquivo ignoravam versões `_v2`/`_v3`, podendo silenciosamente usar dado desatualizado se algum passo fosse rerodado.
- Testado de ponta a ponta (Julho, Actual E Flash, sem tocar rede): tudo bateu exatamente com os valores já validados.

### Estado final: processo de fechamento revisado, testado e funcionando
- Passos 1-4 (Actual e Flash) prontos pra uso real. Nenhum dado de produção foi alterado incorretamente hoje.

### Pendência pra próxima sessão (segunda-feira)
1. **Watchdog/timeout** — se o Excel travar de verdade (hang) durante uma automação, hoje não há timeout. Usuária pediu explicitamente pra resolver isso na próxima sessão.

---
## Continuação 2026-08-22 — Quadro de comparação (linhas 18/19) implementado pro Ciclo Flash: fonte vira Forecast (R<mês>), não outro Flash

- **Retomada:** pendência #1 do fim de 2026-08-21 — o quadro de comparação `atualizar_comparacao_flash` só rodava pro Actual (comparando contra o Flash do mês). Usuária explicou a lógica do Ciclo Flash: comparar contra o **Forecast mais recente** (nomenclatura Pirelli: R+número do mês = REFRESH, ex R7=refresh de julho; não existem R1/R8/R12).
- **Implementado e validado** (detalhe técnico completo em `memory/DECISOES.md` → "2026-08-22"): `localizar_forecast_para_comparacao` + `ler_forecast_despesas_mao_de_obra` + `atualizar_comparacao_forecast`, tudo em `gerar_base_intermediaria.py`. Busca R<mês fechando> na pasta do mês; se não achar (R1/R8/R12), cai pro mês anterior (ex: Agosto usa R7) e **avisa via popup na GUI** qual R foi usado. Janeiro (sem R1, usa Budget/MP) fica pendente — fonte ainda não mapeada, quadro não preenchido nesse caso, com aviso.
- **Estrutura da fonte mapeada inspecionando ao vivo o arquivo real de julho (R7)** — aba "Resumo Resultado Ano": linha 19=Variable Cost, 20=Labour(var), 30=Fixed Cost, 31=Labour(fixo), 38=Total Costs (confirmado = 19+30 nos 12 meses, com check em runtime que avisa se não bater, pedido explícito da usuária). Valores em '000 BRL negativos — fórmula converte pra BRL absoluto positivo.
- **Validação forte:** o valor calculado pra Julho bateu EXATAMENTE com o que a usuária já tinha colado à mão no arquivo real (`Base Intermediária Fitted July Flash 2026.xlsx`: Despesas R$ 3.940.062,77 / Mão de Obra R$ 2.380.392,91). Rodado de ponta a ponta (Lançar Provisões → Finalização) contra pasta de teste local, mesmo resultado.
- **Achado extra confirmado com a usuária:** rótulos de texto do quadro (herdados errados do template Actual do mês anterior) também são corrigidos agora — linha 15/A "Actual"→"Flash", linha 18/A "Flash"→"Forecast", cabeçalho linha 24 H/I "Flash"/"Actual"→"Forecast"/"Flash".
- **Registrado em `ontology/fitted_units.json`** → `ciclos.forecast` (nomenclatura R, cenários que não existem, estrutura de pastas/arquivo) e `quadro_comparacao_pivot` (lógica completa dos dois casos, Actual e Flash).
- **Nenhum dado real de produção foi alterado** — todo teste rodou contra pasta de teste local (`data/processed/fitted_units_despesas/base_intermediaria_teste/`), lendo arquivos reais da rede só como fonte (ReadOnly).
- **Pendências que seguem em aberto (herdadas de 2026-08-21, nenhuma delas foi tocada nesta sessão):**
  1. Fonte de Budget/MP pra Janeiro (sem R1) — retomar perto do fechamento de Janeiro.
  2. Faturamento (linha 25 do quadro amarelo) — ainda manual nos dois ciclos.
  3. Inserção automática de linha colorida se as provisões excederem a capacidade existente (hoje para com erro claro).

---
## Continuação 2026-08-21 — Fitted Units Despesas: implementada a opção 1 (Ciclo já na extração), decisão pendente de 2026-08-19 fechada

- **Retomada rápida no início da sessão:** recapitulei pra usuária o estado de 2026-08-19 (Passo 3 validado com julho Actual, GUI cockpit aprovada visualmente, mas ainda não em produção) e a decisão pendente das 3 opções pra resolver o risco de o Passo 3 pegar a extração errada (Flash vs Actual) do mesmo mês. Ela confirmou direto: **opção 1 (recomendada) — marcar o Ciclo já na extração.**
- **Implementado nesta sessão** (detalhe técnico completo em `memory/DECISOES.md` → "2026-08-21"):
  - `scripts/sap/fitted_units/_shared/ksb1_core.py`: novas funções `prefixo_arquivo_ksb1`, `nome_arquivo_ksb1` e `encontrar_arquivo_ksb1` (busca pelo Ciclo pedido, com fallback pra arquivos antigos sem Ciclo no nome — meses de jan-jul/2026 continuam funcionando sem re-extrair nada).
  - `atualizar_ksb1_gui.py` (Passo 1, `extrair_um`/`rodar`): grava o Ciclo no nome do arquivo bruto (ex: `KSB1 - Fitted Units 07.2026 - Gestoriais - Actual.XLSX`), lendo o Ciclo do painel compartilhado (Mês/Ano/Ciclo) da GUI.
  - `check_agrupamentos_ksb1.py` (Passo 2, `gerar_check`): agora recebe Ciclo, busca os arquivos certos e nomeia a saída incluindo o Ciclo (`Check de agrupamentos - MM.AAAA - Ciclo.xlsx`).
  - `gerar_ksb1_mensal.py` (Passo 3, `decidir_fonte_e_ler_linhas`): busca a extração do Ciclo pedido em vez da mais recente por data de modificação — **é exatamente o bug que motivou a decisão de 2026-08-19.**
- **Testado (sem tocar rede/SAP):** `python -m py_compile` nos 4 arquivos + import real de todos os módulos (assinaturas conferidas) + 5 cenários de teste unitário do `encontrar_arquivo_ksb1` em pasta temporária (só arquivo antigo sem Ciclo, Flash+Actual novos coexistindo, antigo+novo Actual coexistindo, mês sem nenhum arquivo, e o caso mais delicado — arquivo antigo ambíguo + arquivo novo de OUTRO Ciclo coexistindo, pra garantir que o fallback não pega o Ciclo errado por engano). Todos passaram.
- **`extrair_ksb1.py` (script standalone antigo, raiz de `fitted_units_despesas/`) não foi tocado** — confirmado (grep) que nenhum `.bat`/`.vbs`/atalho em uso chama ele; o fluxo real de produção é 100% via `atualizar_ksb1_gui.py`. Fora do escopo desta mudança.
- **Nenhum dado real foi alterado** — só código (4 arquivos `.py`) e documentação (`DECISOES.md`, este `BRIEFING.md`). Nada commitado ainda nesta sessão.
- **Pendente pra fechar de vez a promoção do cockpit pra produção** (item #2 já identificado em 2026-08-19, ainda em aberto):
  1. Trocar o botão "Atualizar KSB1 Pivot" da GUI pra chamar `gerar_ksb1_mensal.py` (hoje ainda chama o script antigo/revertido `gerar_base_intermediaria.py`).
  2. Apontar a pasta de saída do Passo 3 pra rede oficial (`<REDE_BASE>/<ano>/<MM - Mês>/<MM>_<Mês3>_<Ciclo>/`) — hoje ainda escreve em `data/processed/fitted_units_despesas/base_ksb1_teste/`.
  3. Passo 4 (ler `Pivot_Inter.` e colar nas linhas brancas da `Intermediária`, já excluindo unidades encerradas) continua não escrito.
- **Pergunta em aberto pra usuária:** ela quer seguir agora pros itens 1-2 acima (colocar o Passo 3 em produção de fato), ou prefere revisar/testar mais a mudança do Ciclo antes?

### Continuação (mesmo dia) — Teste ao vivo Jan-Jul/2026 (Ciclo Actual): lacuna de dados achada e preenchida, todos os 7 meses OK agora

- **Usuária pediu teste ao vivo** só do cenário Actual, todos os meses, garantindo que "o valor volte" jan-jul/2026. Primeira rodada: Jan/Fev/Mar/Jun deram erro "arquivo não encontrado" — **não é bug da mudança do Ciclo**, é porque a pasta `00.Extração Base KSB1/<mês>` desses 4 meses estava vazia (o fluxo de extração automatizada só passou a salvar ali a partir de abril/2026; Jan/Fev/Mar foram fechados antes disso, por fora). Detalhe completo em `memory/DECISOES.md` → "2026-08-21 (continuação)".
- **Usuária autorizou extrair ao vivo via SAP** — feito com sucesso (`extrair_um`, mesma função de produção da GUI), 8 arquivos novos (Gestoriais + Sem Agrupamento × 4 meses), já com o Ciclo no nome.
- **Obstáculo tratado:** popup nativo "Segurança SAPGUI" apareceu uma vez por pasta nova (não dá pra fechar via script, é proteção por design). Usuária resolveu clicando "Permitir" + "Memorizar minha decisão" em cada uma — mesma solução já usada em 2026-08-13 para outras pastas. **Ela perguntou se dava pra "corrigir pra não aparecer mais"** — expliquei que é um aviso de segurança nativo do SAP GUI (não do nosso script), que "Memorizar minha decisão" já resolve permanentemente por pasta (foi o que ela fez), e que não tentei desabilitar isso globalmente nas configurações do SAP GUI porque reduziria uma proteção de segurança real — não decidido/mudado nada nessa frente, só expliquei.
- **Resultado final:** reteste completo Jan-Jul/2026 (Ciclo Actual) — todos os 7 meses retornam valor sem erro agora (tabela completa em `DECISOES.md`). Confirma disponibilidade de dado, não é reconciliação de valor contra a Base Intermediária real (só Julho tem essa validação formal, de 2026-08-19).
- **Nada commitado ainda desta parte** (só a implementação do Ciclo, commit `fe4f098`, foi commitada antes do teste).

---
## Continuação 2026-08-19 — Retomada Fitted Units Despesas: Flash vs Actual, check de agrupamentos confirmado, comparação July em andamento (compare contra Actual, não Flash)

- **Recap pedido pela usuária no início:** confirmado que passo 1 (extração KSB1 + check de agrupamentos) está OK/estável desde 2026-08-10, sem pendência. Já o arquivo `KSB1 July` gigante (passo 3, BASE_KSB1 + Pivot) **não estava validado** — o teste às cegas de 2026-08-14 (`data/processed/fitted_units_despesas/base_ksb1_teste/KSB1 July Flash 2026 - TESTE VALIDAÇÃO.xlsx`, 51.039 linhas) nunca teve o resultado da comparação conferido (o script exploratório `comparar_julho.py` daquela sessão era só de scratchpad, se perdeu, e a sessão encerrou antes de rodar).
- **Explicação da usuária sobre Flash vs Actual (novo conhecimento de negócio, ainda não totalmente registrado no ontology):**
  - **Flash** = prévia do fechamento (1º dia útil). As linhas coloridas da aba `Intermediária` (ver `ontology/fitted_units.json` → `intermediaria_linhas_coloridas_flash`) existem só no Flash porque são provisões solicitadas manualmente pra cobrir o que a contabilidade ainda não lançou no SAP a tempo.
  - **Actual** = fechamento efetivo (~dia 5). Já tem as contas de PIS/COFINS integradas (a usuária chama de contas "PC") — essas contas **não aparecem no Flash** porque a integração ainda não rodou. Isso explica a observação antiga do ontology (`ciclos.actual.descricao`: "precisa integrar todos os impostos").
  - **Ainda não registrado formalmente no ontology** — fazer isso na próxima oportunidade (adicionar as contas PC/PIS-COFINS como um conceito formal em `ciclos.actual` ou em `base_intermediaria`).
- **Passo 2 (check de agrupamentos) — confirmado que já faz o que a usuária queria, sem precisar mudar código:** `check_agrupamentos_ksb1.py` já escreve as contas sem vínculo gestorial na aba do Check (seção "Check 2", conta a conta com valor) e já sinaliza no resumo (`Situação` vira "ATENÇÃO - valores não batem, ver Check 2") e no log/console (`"AVISO: N conta(s) sem vínculo..."`). Perguntei se ela queria algo mais chamativo (cor de destaque no Excel, popup na GUI) — **ela respondeu "está ótimo, perfeito"**, ou seja, comportamento atual já atende, não precisa de reforço visual.
- **Passo 3/validação July — retomado, com uma correção importante da usuária:** comparar contra **Actual, não Flash**. Investigando o motivo, faz sentido: os arquivos brutos de extração usados no teste de 2026-08-14 (`00.Extração Base KSB1/07 - Jul/KSB1 - Fitted Units 07.2026 - Gestoriais/Sem Agrupamento` v4) foram puxados do SAP em **10-11/08/2026**, ou seja, DEPOIS do fechamento Actual de julho (~05/08) — então já devem conter as contas PC/PIS-COFINS, o que torna a comparação contra o Flash real (que não tem PC) inválida por natureza (não seria erro da automação, seria diferença de dado-fonte). Comparar contra o Actual real é o teste correto.
  - **Confirmado que os arquivos reais de July Actual já existem na rede** (fechamento de julho já ocorreu): `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\2026\07 - Jul\07_Jul_Actual\KSB1 July Actual 2026.xlsx` e `...\Base Intermediária Fitted July Actual 2026.xlsx`.
  - **Confirmado inspecionando ao vivo:** a `Base Intermediária Fitted July Actual 2026.xlsx` tem 601 linhas de dados na aba `Intermediária`, **nenhuma colorida** — bate com o que o ontology já dizia (linhas coloridas só existem no Flash).
  - **Estrutura do `Pivot_Inter.` mapeada** (arquivo de teste): colunas A-H = chave (Gestorial, Descrição, Classe de custo, Denom., Centro custo, MF, Centro de Montagem, Variabilidade), colunas I em diante = "Mês" (I=1/Janeiro ... O=7/Julho), cada uma com a soma de `Valor/MR`. Chave de casamento com a `Intermediária` real = (Classe de custo/Conta Fiscal, Centro custo).
  - **Script de comparação escrito** em `C:\Users\SILVEJ~1\AppData\Local\Temp\claude\...\scratchpad\comparar_julho_actual.py` (fora do repo, mesmo padrão exploratório do `comparar_julho.py` anterior — **vai se perder ao fechar a sessão/scratchpad, considerar mover pro repo se for repetir esse tipo de comparação no futuro**). Lê `Pivot_Inter.` do arquivo de teste (mês=7) e a coluna "July" da `Intermediária` real do Actual, agrupa por (Conta Fiscal, Centro de Custo) e reporta diferenças.
  - **RESULTADO (fechado ainda nesta sessão):** 587 combinações comparadas, 571 bateram exatamente. As 16 restantes (R$ 112.275,58) são 100% explicadas por uma regra de negócio nova que a usuária confirmou: **unidades "encerradas" aparecem na KSB1/BASE_KSB1 (residual retroativo) mas NÃO são coladas na `Intermediária`/EBIT — o residual é estornado contra uma provisão de "Não Recorrente" que cobre custos de encerramento.** 14 das 16 diferenças eram do centro 8269 (Sorocaba), 1 do 8292 (Sorocaba), 1 do 8247 (Sorocaba, confirmado pela usuária). **Passo 3 (BASE_KSB1 + Pivot) está validado.** Detalhe completo em `memory/DECISOES.md` (2026-08-19) e regra em `ontology/fitted_units.json` → `regra_unidades_encerradas_no_ebit`.
  - **Pendente pra próxima sessão:** (1) perguntar à usuária se já pode promover `gerar_ksb1_mensal.py` da pasta de teste local pra pasta de rede oficial e trocar o botão da GUI (regra de "testar isolado → validar → promover" de 2026-08-14, validação está feita agora); (2) implementar o passo 4 (ler `Pivot_Inter.` e colar nas linhas brancas da `Intermediária`), já considerando a exclusão de unidades encerradas confirmada nesta sessão.
- **Nenhum dado real foi alterado nesta sessão** — só leitura (arquivos reais abertos com `data_only=True`, nunca salvos) e um script de comparação em pasta de scratchpad, fora da rede e fora do repositório.

### Continuação (mesmo dia) — Redesign visual da GUI (`atualizar_ksb1_gui.py`): tema "cockpit" escuro + abas por passo, EM ANDAMENTO

- **Pedido da usuária:** deixar a janela da GUI ("aquele quadrinho") com cara de cockpit — tela escura, logo da Pirelli, um rastro de pneu em algum lugar. Depois de ver a primeira versão (janela pequena, 440x600), pediu pra aumentar bastante e organizar em **abas seguindo a ordem dos passos do processo** (Extração → Check de Agrupamentos → Base Intermediária). Depois de ver essa segunda versão, pediu um ajuste fino: manter o cabeçalho escuro como está, só trocar o texto "COCKPIT KSB1" por **"COCKPIT FECHAMENTO FITTED"**, e trocar o **corpo abaixo da linha vermelha pra fundo branco com letras pretas** (em vez de escuro).
- **Estado atual do código** (`scripts/sap/fitted_units/fitted_units_despesas/atualizar_ksb1_gui.py`): cabeçalho escuro (`BG_ROOT`) com o logo Pirelli (badge amarelo, embutido em base64, sem transparência — confirmado inspecionando os pixels) + título "COCKPIT FECHAMENTO FITTED" + subtítulo "Fitted Units · Despesas", uma linha de trim vermelho (Pirelli) embaixo do cabeçalho, e todo o corpo (painel Mês/Ano/Ciclo, `ttk.Notebook` com 3 abas numeradas — ①Extração ②Check de Agrupamentos ③Base Intermediária —, console de log) agora em fundo branco/texto preto (`BG_PAINEL`/`BG_CARD`/`LOG_BG` = branco, `TEXTO_CLARO`/`LOG_FG` = preto). Rodapé com um "rastro de pneu" desenhado via `Canvas` (blocos repetidos simulando sulco de pneu, sem depender de imagem externa) — fundo do rodapé também branco agora, tread escuro (fica com cara de marca de pneu no chão).
  - Cada aba tem: título do passo, descrição curta do que ele faz, e o botão de ação (mantido vermelho Pirelli/texto preto — não foi pedido pra mudar). O Mês/Ano/Ciclo ficam num painel compartilhado acima das abas (usado pelos 3 passos).
  - **Bug de contraste corrigido durante a implementação:** o subtítulo do cabeçalho usava a variável `TEXTO_SECUNDARIO`, que foi redefinida pra cor escura (tema branco do corpo) — deixaria o texto quase invisível num fundo escuro. Corrigido hardcodando a cor clara (`#9a9da2`) só nesse label do cabeçalho, já que ele fica sobre `BG_ROOT` (escuro), não sobre o corpo branco.
  - **Achado à parte, não relacionado ao pedido:** a tela do ambiente onde a Juliana está rodando isso é só 1280x720 — a primeira versão "grande" (900x720) ficava mais alta que a tela inteira, sem espaço pra barra de título/barra de tarefas. Reduzido pra 860x640 (com `minsize` 780x560, redimensionável) pra caber com folga.
- **Testado abrindo a janela de verdade (não só lida no código) e capturando screenshot** pra conferir visualmente antes de mostrar — prática nova nesta sessão para mudanças de UI, vale repetir em qualquer ajuste visual futuro.
- **Sessão atingiu o limite de 45 ações de novo — backup automático já rodou** (`session_transition.py`).

### Continuação (mesmo dia) — GUI cockpit: aprovada visualmente, dois bugs reais achados e corrigidos, decisão de arquitetura pendente pra amanhã

- **Dois bugs reais achados testando a janela ABERTA de verdade (não só lendo o código) — ambos corrigidos e reconfirmados com screenshot:**
  1. **Texto borrado:** o processo Python não avisava o Windows que sabia lidar com DPI/escala de tela sozinho, então o Windows "esticava" a janela inteira como bitmap pra bater com o zoom da tela — borrando o texto. Corrigido chamando `ctypes.windll.shcore.SetProcessDpiAwareness(1)` (com fallback pra `user32.SetProcessDPIAware()`) logo no topo do script, antes de qualquer janela do Tk ser criada.
  2. **Subtítulo do cabeçalho espremido contra a linha vermelha:** depois da correção de DPI, a fonte do título passou a renderizar no tamanho real (maior) dentro de um cabeçalho com altura TRAVADA em pixels (`height=78` + `pack_propagate(False)`) — não sobrava espaço pro subtítulo. Corrigido removendo a altura fixa, deixando o cabeçalho se ajustar ao conteúdo.
- **Aprovado pela usuária** o resultado visual: cabeçalho escuro com logo Pirelli + "COCKPIT FECHAMENTO FITTED" + trim vermelho, corpo branco/letras pretas com painel Mês/Ano/Ciclo + 3 abas numeradas na ordem dos passos + console de log, rastro de pneu no rodapé. GUI deixada aberta pra usuária mexer com calma.
- **Pergunta da usuária, ainda SEM decisão fechada — retomar amanhã por aqui:** ela confirmou que o campo Ciclo (Actual/Flash) deve **decidir** (não só nomear) tanto o arquivo final quanto a pasta de destino na rede (padrão já existente `<MM>_<Mês3>_<Ciclo>/`, ex: `07_Jul_Actual`). Isso já era o plano combinado em 2026-08-14, só que ainda não implementado (`gerar_ksb1_mensal.py` hoje só escreve numa pasta de teste local fixa, ver `data/processed/fitted_units_despesas/base_ksb1_teste/`).
  - **Problema técnico real que isso expôs (não é só sobre a GUI):** hoje o Passo 3 (`decidir_fonte_e_ler_linhas` em `gerar_ksb1_mensal.py`) escolhe a extração bruta do Passo 1 **mais recente por data de modificação** (`encontrar_arquivo`, ordena por `mtime`), não uma extração vinculada ao Ciclo escolhido. Se a usuária extrai a KSB1 2x no mês (Flash dia 1, Actual dia ~5, confirmando o padrão real — ver `processo_recorrente.frequencia` no ontology) e depois tenta gerar/regerar o arquivo do Flash já com a extração do Actual mais nova disponível, o script pegaria a extração errada (dado do Actual) só com o nome/pasta do Flash — o mesmo tipo de inconsistência encontrada na validação de julho desta sessão.
  - **Propus 3 opções pra resolver isso, ela pediu pra esclarecer antes de escolher, mas decidiu encerrar a sessão nesse ponto — RETOMAR AMANHÃ perguntando o que ela queria esclarecer:**
    1. *(Recomendada)* Marcar o Ciclo já na extração (Passo 1 ganha o mesmo seletor Ciclo; nome do arquivo bruto passa a incluir o Ciclo; Passo 3 busca pelo Ciclo, não mais pela mais recente).
    2. Escolher manualmente qual extração usar (Passo 3 lista as extrações disponíveis do mês com data/hora pra usuária escolher).
    3. Extrair de novo na hora, sempre, ao gerar o arquivo do Passo 3 (elimina o risco, mas exige SAP aberto toda vez e é mais lento).
- **Ainda pendente (itens antigos que continuam abertos):**
  1. A decisão acima (Ciclo → extração correta) precisa ser fechada antes de promover `gerar_ksb1_mensal.py` pra pasta de rede oficial.
  2. O botão "Atualizar KSB1 Pivot" da GUI cockpit **ainda chama o script antigo/revertido** `gerar_base_intermediaria.py` (arquitetura pré-reversão de 2026-08-11), não o `gerar_ksb1_mensal.py` validado nesta sessão — troca ainda não feita.
  3. Passo 4 (ler `Pivot_Inter.` e colar nas linhas brancas da `Intermediária`) continua não escrito, já considerando a exclusão de unidades encerradas confirmada hoje (`ontology/fitted_units.json` → `regra_unidades_encerradas_no_ebit`).
- **Nenhum dado real foi alterado nesta sessão** (extrações/arquivos de rede só foram lidos, nunca escritos) — só a GUI (código) e a documentação (`BRIEFING.md`, `DECISOES.md`, `ontology/fitted_units.json`) foram alteradas.
- **Fim de sessão a pedido da usuária:** guardar backup + briefing, retomar amanhã.

---
## Continuação 2026-08-18 — Ajuste de frequência do Agendador da checagem ZLFIB (tela cmd piscando de hora em hora)

- **Pedido da usuária:** ela notou que uma tela cmd preta abre e fecha de hora em hora, todo santo dia, e queria saber o que era.
- **Diagnóstico:** é a tarefa agendada do Windows `Verificacao_ZLFIB_Duplicidade_Mensal` (criada em 2026-08-13, ver sessão de 2026-08-13 mais abaixo), que roda `watcher_mensal_zlfib.bat` → `verificacao_mensal_zlfib.py`. O gatilho original era um único `TimeTrigger` com repetição de 1h **sem restrição de dia**, então disparava de hora em hora todos os dias do mês — o script só se auto-encerra rápido (sem fazer nada) nos dias que não são o 1º dia útil, mas a janela cmd ainda piscava.
- **Correção aplicada — só no Agendador de Tarefas do Windows, nenhuma mudança no `verificacao_mensal_zlfib.py`:** tarefa reconfigurada com dois gatilhos (`schtasks /create /xml` com XML customizado, já que os cmdlets `New-ScheduledTaskTrigger` desta máquina não suportam `-Monthly`):
  1. `CalendarTrigger` mensal nos dias **1, 2 e 3** (cobre qualquer 1º dia útil deslocado por fim de semana), repetindo de hora em hora das **7h às 18h** — é essa a janela real de polling esperando o login no SAP.
  2. `CalendarTrigger` diário às **9h**, todo dia — cobre os outros ~29 dias do mês com uma única checagem rápida (auto-encerra, "hoje não é o dia certo") em vez de hourly.
- **Confirmado com a usuária antes de aplicar** (perguntei via AskUserQuestion, ela escolheu explicitamente esse formato híbrido) e validado depois com `schtasks /query /tn Verificacao_ZLFIB_Duplicidade_Mensal /v` — os dois `CalendarTrigger` aparecem certinhos (Monthly 1-3 com repetição 1h/11h, Daily 09:00).
- **Nenhum dado real foi tocado** — mudança é só na configuração do Agendador do Windows. Detalhe também salvo na memória global (`memory/project_fitted_recuperacao_zlfib.md` fora do repo, pasta `.claude/projects/.../memory/`).
- **Nada pendente nisso** — só acompanhar no próximo 1º dia útil do mês (setembro) se o gatilho dispara certinho nos dias 1-3.

- **Continuação (mais tarde) — duas fontes novas de referência guardadas na memória, a pedido da usuária:**
  1. **Arquivo mestre de contas contábeis/gestoriais e centros de custo:** `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Custos Efetivos\_Acompanhamentos e Controles\Base_Contas_Contábeis_Fitted_22.xlsx` (mantido pela controladoria central, ativo — última modificação 16/07/2026). Inspecionado ao vivo: aba `Contas` (Classe de custo → Conta Gestorial, Fixo/Variável, MO/DG) e aba `Centros` (Centro de custo → unidade/Sigla, Fixo/Variável). Registrado em `ontology/fitted_units.json` → `fonte_contas_gestoriais_e_centros`.
  2. **Lista completa de centros de custo por unidade:** usuária mandou print de Tabela Dinâmica com todos os centros de custo agrupados por unidade (Camaçari, Faturamento, Gerência, Goiana, Ibirité, Itatiaia, Juiz de Fora, Resende, Santo André, SJP, Sorocaba). Registrado em `ontology/fitted_units.json` → `centros_de_custo_por_unidade`. **Grupos Camaçari, Itatiaia, Juiz de Fora e Santo André vieram com destaque amarelo no print — significado não confirmado, não presumir.** Também apareceram 3 grupos que não existem ainda na lista `unidades` do ontology (Camaçari, Faturamento, Juiz de Fora, Santo André) — perguntar à usuária o que são antes de tratar como plantas operacionais.
  - **Bônus:** essa lista resolveu de quebra a pendência de mapeamento Centro de custo → unidade do sub-projeto Energia Elétrica Fitted (8296→IBIRITE, 8290/8289→SJP, 8269/8292→SOROCABA, 8303→GOIANA) — ver `memory/PROJECT_MAP.md`.
  - Nada commitado/enviado ainda pro GitHub nesta parte — só salvo localmente, pendente do próximo backup.
  - **Confirmação da usuária logo em seguida:** Camaçari, Itatiaia e Santo André **já fecharam, como a Sorocaba** (status "encerrada") — adicionadas em `ontology/fitted_units.json` → `unidades` (códigos CAM, ITA — corrigido nome de "Itaiaia" pra "Itatiaia" —, STA) e marcadas em `centros_de_custo_por_unidade`. **Faturamento pode ser ignorado** (não é unidade operacional). **Juiz de Fora ficou em aberto** — tinha o mesmo destaque amarelo das 3 encerradas no print original, mas a usuária não a citou ao confirmar; não presumir que também fechou, perguntar na próxima sessão.
  - **Sessão atingiu o limite de 45 ações — backup automático já rodou** (`session_transition.py`). Commit `4bab931` (fonte de contas + centros por unidade) enviado ao GitHub pela usuária via `! git push` (o `git push` direto do Claude foi bloqueado pelo classificador do modo Auto — ficar atento a esse padrão em backups futuros: `git add`/`git commit` funcionam normalmente, só `git push` é bloqueado, precisa a usuária rodar com `!`).
  - **Confirmação final da usuária: Juiz de Fora também encerrou.** As 4 unidades com destaque amarelo no print (Camaçari, Itatiaia, Santo André, Juiz de Fora) estão todas encerradas, junto com a Sorocaba (já estava documentada) — total de 5 unidades encerradas em `ontology/fitted_units.json` → `unidades`.
  - **Ajuste de texto pedido pela usuária:** a observação de todas as unidades encerradas (ITA, SOR, CAM, STA, JDF) foi padronizada para "mas ainda pode aparecer algum lançamento retroativo (custo residual)" — termo mais preciso que só "custo residual".

---
## Continuação 2026-08-14 — Passo 3 (BASE_KSB1 + Pivot nativo) automatizado e testado, validando com July Flash

- **Objetivo da sessão:** implementar o passo 3 do processo recorrente (Fitted Units Despesas) — copiar o KSB1 Actual do mês anterior, colar as linhas novas do mês na aba `BASE_KSB1`, arrastar as fórmulas S:AI, refresh das Pivot Tables nativas via COM. Plano já vinha fechado desde 2026-08-11 (retomada à tarde), mas nunca tinha sido codificado até hoje.
- **Script novo: `scripts/sap/fitted_units/fitted_units_despesas/gerar_ksb1_mensal.py`** — roda de ponta a ponta com sucesso (confirmado com julho/2026 Flash, ver abaixo).
  - `decidir_fonte_e_ler_linhas(mes, ano, log)`: decide Gestoriais vs Sem Agrupamento (mesma regra do `check_agrupamentos_ksb1.py`) e lê as linhas de detalhe completas (18 colunas, A-R do BASE_KSB1).
  - `localizar_ksb1_actual_anterior(mes, ano)`: acha o KSB1 Actual do mês anterior na rede — **não depende do Ciclo** (Actual/Flash sempre partem do Actual do mês anterior); Ciclo só entra no nome do arquivo final. Confirmado com a usuária que isso está certo.
  - `copiar_para_teste(...)`: copia (nunca sobrescreve, `nome_com_versao`) pra pasta de saída.
  - `remover_flag_somente_leitura_recomendada(...)`: remove do XML da cópia (nunca do original) a flag `<fileSharing readOnlyRecommended="1"/>` que o arquivo carrega — ver bug #1 abaixo.
  - `colar_linhas_e_atualizar_pivots(...)`: abre uma instância **isolada e oculta** do Excel (`win32com.client.DispatchEx`), cola as linhas novas, usa `Range.AutoFill` pra "arrastar" as fórmulas S:AI, `wb.RefreshAll()` pra atualizar as Pivot Tables (`Pivot_Inter.`, `Pivot_Detalhes`), salva. Tem checagens ativas que lançam erro de verdade (`wb.ReadOnly`, última linha antes/depois de colar, `wb.Saved`) em vez de falhar silenciosamente.
  - `gerar_ksb1_mensal(mes, ano, ciclo, pasta_saida, sufixo_nome, log)`: orquestra tudo.
- **Escopo combinado com a usuária:** rodar isso primeiro como **validação às cegas** com julho/2026 Flash (mês já fechado manualmente por ela em 03/08 — `KSB1 July Flash 2026.xlsx` e `Base Intermediária Fitted July Flash 2026.xlsx` já existem na rede) antes de virar processo oficial. Saída de teste em `data/processed/fitted_units_despesas/base_ksb1_teste/` (local, fora da rede) — **combinado explicitamente: só vamos apontar pra pasta de rede oficial numa próxima sessão**, depois de ver os números da comparação.
- **Regra nova confirmada pela usuária, vale para TODOS os projetos:** todo script novo que vai gerar/atualizar arquivo oficial na rede primeiro roda numa pasta apartada até ela validar; só depois o destino muda pra rede e qualquer GUI/atalho ligado ao processo antigo é atualizado. Registrado em `memory/DECISOES.md` (2026-08-14) e na memória global (`feedback_collaboration_style.md`).
- **Descoberta paralela:** o botão "Atualizar KSB1 Pivot" da GUI antiga (`atualizar_ksb1_gui.py`, tem seletor "Ciclo: Actual/Flash") ainda está ligado ao script **antigo/revertido** `gerar_base_intermediaria.py` (arquitetura pré-reversão de 2026-08-11, pulava o BASE_KSB1). Ação combinada: trocar esse botão pra chamar `gerar_ksb1_mensal.py` assim que a validação for aprovada.
- **Dois bugs reais encontrados e corrigidos durante os testes de hoje** (ambos em `gerar_ksb1_mensal.py`, ambos confirmados corrigidos — arquivo final validado com 51.039 linhas na `BASE_KSB1`, 44.976 originais + 6.063 novas de julho):
  1. **Salvamento silencioso sem gravar nada, sem erro nenhum:** o `KSB1 June Actual 2026.xlsx` (fonte) carrega a flag interna `<fileSharing readOnlyRecommended="1"/>` no XML. Com `DisplayAlerts=False`, o Excel abria a cópia em modo leitura sem avisar, e `wb.Save()` virava um no-op. `IgnoreReadOnlyRecommended=True` no `Workbooks.Open()` **não resolveu** (parâmetro não confiável via win32com). Corrigido removendo a flag direto do XML (zipfile) da cópia antes de abrir no Excel.
  2. **Processos ocultos do Excel ficando travados/zumbis** depois de uma tentativa com bug (não fechavam sozinhos). Pedi confirmação da usuária antes de encerrar via `Stop-Process` (bloqueado por padrão pelo classificador de permissão) — ela autorizou, resolvido.
- **Risco identificado e comunicado à usuária (ainda não mitigado):** o peso do arquivo (hoje ~20MB/51 mil linhas, deve chegar a ~35-40MB/90-100 mil linhas até dezembro) não é problema técnico pro Excel. O risco real é a automação travar silenciosamente sem aviso se aparecer um popup que `DisplayAlerts=False` não suprime — foi exatamente o que aconteceu com os dois bugs acima. **Antes de qualquer versão agendada/recorrente (como a automação mensal do ZLFIB), falta adicionar um "vigia" com timeout que mata o processo e reporta erro em vez de travar pra sempre.** Não implementado ainda.
- **Status da comparação (task #5, tocada nesta sessão):** rodando `comparar_julho.py` (script exploratório em scratchpad, não commitado — lê o `Pivot_Inter.` do arquivo gerado, agrupa por (Centro custo, Classe de custo) somando julho, compara com a coluna "July" da `Base Intermediária Fitted July Flash 2026.xlsx` real, só nas 686 linhas sem cor — as 67 linhas coloridas, ver 2026-08-13/2026-08-14 anterior, são ajustes manuais fora do escopo desta automação). **Ainda não tinha retornado resultado até o fim desta sessão — retomar conferindo se terminou e registrar o resultado (bateu ou não bateu, com números).**
- **Combinado explicitamente com a usuária:** depois de ver o resultado da comparação, ainda vamos para a próxima sessão antes de apontar a automação pra pasta de rede oficial — nada foi escrito na rede nesta sessão, só na pasta de teste local.

---
## Continuação 2026-08-13 (mais tarde) — Novo sub-projeto "Energia Elétrica Fitted"

- **Escopo confirmado pela usuária:** dois objetivos —
  1. Checar se todos os lançamentos de energia elétrica foram feitos corretamente (sem esquecimento, tudo lançado).
  2. **Mais importante:** conferir se os créditos de PIS, COFINS e ICMS sobre energia elétrica estão sendo lançados (usuária acredita que não).
- **Dados de partida que ela deu:** conta fiscal de referência `N17002S001` ("COM Fix"), e fornecedores por unidade (mesmo padrão de código de fornecedor da ZLFIB/KSB1, 10 dígitos):
  - CEMIG DISTRIBUIÇÃO S/A — `4211308770` (provável IBI)
  - COMPANHIA PAULISTA DE FORÇA E LUZ (CPFL) — `4211324097` (provável SJP)
  - COPEL DISTRIBUIÇÃO S.A — `4211333301` (provável SOR ou outra — confirmar)
  - FIAT AUTOMOVEIS S/A — `4211330756` — caso especial: a Fiat **revende** energia pra Goiana (GOI), em **duas notas**: uma de transmissão e uma de repasse.
  - **Ignorar:** SERENA GERAÇÃO S.A — `4211333021` — é rateio, não entra nesta análise.
  - **Ainda falta confirmar:** qual filial exatamente usa CEMIG vs CPFL vs COPEL (a usuária deu a lista mas não amarrou 1:1 com SJP/IBI/SOR ainda — não presumir, perguntar antes de aplicar).
- **Dois problemas técnicos encontrados e resolvidos durante a exploração inicial na KSB1:**
  1. **Erro "Selecionar uma das alternativas indicadas"** ao rodar a KSB1: a tela de seleção tinha `Centro de custo` (8204) e `Classe de custo` (N17002S000) preenchidos ao mesmo tempo que os campos de **grupo** (`Grupo de centros de custo`/`Grupo de classes de custo`) — são pares alternativos (rótulo "ou" entre eles no layout), o SAP não aceita os dois lados preenchidos. **Regra confirmada pela usuária: sempre usar o mesmo layout/parâmetros já mapeados (`BU['kstgr']`=0495, `BU['disvar']`='/DESPFITTED', `KOAGR` em branco ou "gestoriais") e nunca preencher Centro de custo/Classe de custo direto.** Script novo (`explorar_conta_energia.py`) já segue essa regra.
  2. **Popup "Segurança SAPGUI"** pedindo autorização pra criar arquivo numa pasta nova — apareceu porque eu estava salvando exports de teste numa pasta temporária efêmera (muda a cada sessão do Claude Code). **Corrigido: passar a salvar sempre em `data/processed/`** (regra que já existia no `CLAUDE.md` e que eu não tinha seguido nesse caso específico). Cada pasta nova ainda pede aprovação uma vez (usuária marca "Memorizar minha decisão"), mas usando sempre a mesma pasta estável isso só acontece uma vez, não a cada sessão.
- **Novo script:** `scripts/sap/fitted_units/energia_eletrica_fitted/explorar_conta_energia.py` — extrai KSB1 Sem Agrupamento de 2026 (01.01-31.07) pra `data/processed/energia_eletrica_fitted/`. Sub-projeto ainda não tem pasta própria documentada no `PROJECT_MAP.md` — fazer isso quando o escopo fechar mais.
- **Status no fim da sessão:** rodando a extração pela primeira vez com o método corrigido — travou de novo esperando a usuária aprovar o popup de segurança da nova pasta `data/processed/energia_eletrica_fitted`. **Retomar confirmando se a extração terminou e o arquivo foi gerado**, depois abrir o Excel e localizar a conta `N17002S001` e os fornecedores de energia pra entender a estrutura antes de decidir a lógica de checagem (não foi definida ainda — só o objetivo de negócio).
- **Extração concluída e inspecionada com sucesso** (`data/processed/energia_eletrica_fitted/KSB1 - Fitted Units 2026 - Sem Agrupamento (energia).xlsx`, jan-jul/2026): a conta `N17002S001` ("COM Fix - Energia El[étrica]") é a **única** conta contábil de energia elétrica em todo o extrato (nenhuma outra variante `N17002*`) — 58 lançamentos no total, cobrindo **6 centros de custo diferentes**:
  | Centro custo | Fornecedor | Cobertura jan-jul/2026 |
  |---|---|---|
  | 8296 | CEMIG (`4211308770`) | Todo mês, sem falha aparente — provável IBI |
  | 8290 | COPEL (`4211333301`) | Todo mês, sem falha aparente — provável SJP |
  | 8289 | COPEL (`4211333301`) | Só abril (1 lançamento) |
  | 8269 | CPFL (`4211324097`) | Só fev-mar, nada depois — padrão bate com SOR (unidade encerrada, "custo residual") |
  | 8292 | CPFL (`4211324097`) | Só março (5 lançamentos, mesmo mês) |
  | 8303 | FIAT/GOI (`4211330756`) | **Falta fevereiro e junho inteiros** — achado concreto pro objetivo 1 |
  - Achado objetivo 1 (lançamentos faltando): **centro 8303 (Fiat/Goiana) sem nenhum lançamento em fev/2026 e jun/2026.** Também tem pares de valores que se cancelam no mesmo mês (ex: 8269 fev tem -12.623,38 e +12.623,38) — parecem estorno/correção, não tratados como erro sem confirmar com ela.
  - Achado objetivo 2 (créditos PIS/COFINS/ICMS): **nenhuma das 58 linhas mostra esse detalhamento** — só um valor líquido (`Valor/MR`) por fatura. Compatível com a suspeita dela, mas falta confirmar se existe uma conta contábil SEPARADA pra esses créditos (perguntado, ainda sem resposta).
  - **Resende (RES/Nissan) não apareceu em nenhum lançamento** dessa conta com os fornecedores conhecidos — não sabemos ainda como a energia da Resende é lançada/qual fornecedor.
- **Pendências pra amanhã (usuária confirmou que retoma amanhã):**
  1. Ela vai mandar a lista de qual Centro de custo pertence a qual unidade (SJP/IBI/SOR/GOI/RES/GER) — **não presumir o mapeamento acima, são só palpites por geografia da concessionária, aguardar confirmação dela.** Ela quer ver o resultado organizado por unidade.
  2. Perguntar de novo sobre a conta contábil separada pros créditos de PIS/COFINS/ICMS (não respondeu ainda).
  3. Entender como a energia da Resende é lançada (não apareceu no extrato).
  4. Depois de mapear tudo, decidir a lógica formal de checagem (completude mensal + verificação dos créditos) e provavelmente criar um script recorrente (parecido com os outros sub-projetos).
- Scripts exploratórios desta etapa (ainda ad-hoc, não é o script final do sub-projeto): `scripts/sap/fitted_units/energia_eletrica_fitted/explorar_conta_energia.py`, `inspecionar_energia.py`, `resumo_cobertura.py`.

---
## Fechamento da sessão 2026-08-13 (continuação final)
- **Reorganização de pastas concluída:** `scripts/sap/` virou `scripts/sap/fitted_units/{_shared, fitted_units_despesas, fitted_recuperacao}/`, a pedido da usuária (queria os sub-projetos separados em pastas antes de mais sub-projetos, ex. Circuito Panamericano, serem criados). Detalhe técnico completo (o que moveu pra onde, imports corrigidos, pontos externos atualizados) em `memory/DECISOES.md` → "2026-08-13 — Reorganização de scripts/sap/". Estrutura nova também documentada em `memory/PROJECT_MAP.md`.
- **Tudo testado depois da reorganização, nada ficou quebrado:** imports Python confirmados em runtime (não só sintaxe), tarefa agendada `Verificacao_ZLFIB_Duplicidade_Mensal` corrigida pro novo caminho (`schtasks /change`), atalho `ATUALIZAR KSB1.lnk` da rede regenerado (`criar_atalho_ksb1.ps1`) e conferido via PowerShell.
- **Memória global do Claude Code também iniciada nesta sessão** (fora da pasta do projeto, em `C:\Users\silveju001\.claude\projects\...\memory\`, pedido explícito da usuária "guarde na memória"): 3 arquivos novos — referência pro sistema de memória do próprio projeto, estilo de colaboração da usuária (rigor com números financeiros, confirmar antes de automação irreversível, PT-BR), e um resumo do sub-projeto Fitted Recuperação/ZLFIB. Não duplica o que já está em `memory/BRIEFING.md`/`DECISOES.md` — só pontos que atravessam sessões/projetos.
- **Pendência mínima:** `ATUALIZAR KSB1.bat` (o `.bat` legado dentro de `fitted_units_despesas/`) não foi testado ao vivo depois do move — deveria funcionar sem mudança (usa `%~dp0`, caminho relativo), mas vale confirmar da próxima vez que for usado.
- Sessão bateu no limite de 45 ações **três vezes** — backup automático rodou a cada vez, tudo commitado e sincronizado no GitHub.

---
## Resumo do dia 2026-08-11 — Fitted Units Despesas (popup SAP + estrutura de pastas + regra e automação da Base Intermediária)
- **Popup "Definir área contab.custos" do SAP:** usuária relatou que, ao sair/voltar a entrar no SAP, um popup pede a área contábil (0580) antes de liberar qualquer transação. Primeira tentativa de fechar automaticamente adivinhou o campo errado e causou erro em cascata no SAP. Criado `scripts/sap/diagnosticar_popup.py` (+ atalho `DIAGNOSTICAR POPUP.bat`) pra inspecionar o popup ao vivo; descobriu-se que o campo fica dentro de um subscreen (`usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE`), não direto em `usr`. Corrigido com busca recursiva por `GuiCTextField` em `atualizar_ksb1_gui.py` e `extrair_ksb1.py`. **Testado pela usuária e funcionando.** Detalhe completo em `memory/errors/2026-08-11_popup_area_contabil_ao_reentrar_sap.md`.
- **Regra de negócio confirmada — passo 3 do processo (montar a base intermediária):**
  - Depois de extrair a KSB1 (Gestoriais + Sem Agrupamento) e rodar o Check: se o total do Gestoriais bate com o Sem Agrupamento, usa-se o Gestoriais; se não bater (porque existem contas fora do agrupamento gestorial), usa-se o Sem Agrupamento, mas ainda excluindo as contas do Check 1 (fixas + qualquer conta iniciada em "B", que são bens de investimento e nunca entram como despesa).
  - Isso acontece 2x/mês (Flash e Actual). Forecast existe mas está **fora do escopo por enquanto**.
  - Cada fechamento parte do arquivo do Actual do mês anterior (que já acumula o efetivo histórico) — ex: fechando julho, parte-se do arquivo "June Actual" e inserem-se as linhas de julho nele, sem apagar os meses anteriores.
  - Estrutura descoberta nos arquivos de referência (`.../2026/<mês>/<ciclo>/`): `KSB1 <Mês> <Ciclo> <Ano>.xlsx` → aba `BASE_KSB1` (tabela acumulada com todos os meses, colunas derivadas como Gestorial/MF/DG-MO são fórmulas que a usuária arrasta) → aba `Pivot_Inter.` alimenta `Base Intermediária Fitted <Mês> <Ciclo> <Ano>.xlsx` (matriz Conta Gestorial × mês, com Total Ano).
  - **Ainda não implementado em script** — só a regra de negócio foi confirmada e documentada. Automatizar esse passo é o próximo trabalho pesado.
- **Reorganização de pastas na rede** (`\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\2026\`): antes cada ciclo ficava solto na raiz (`01_Jan_Actual`, `01_Jan_Flash`...); agora cada mês tem sua pasta própria (`01 - Jan` a `12 - Dec`, mesmo padrão já usado em `00.Extração Base KSB1`) com Actual/Flash/Forecast dentro. Jan-Jul: pastas movidas sem alterar nenhum arquivo (uma pasta de julho — Forecast — ficou presa por um arquivo aberto no Excel, resolvido depois que a usuária fechou o arquivo). Ago-Dez: subpastas Actual/Flash/Forecast já criadas vazias, prontas pro uso.
- Pendente registrar formalmente em `ontology/fitted_units.json` a nova estrutura de pastas e a regra da base intermediária (só ficou no BRIEFING por enquanto — fazer isso na próxima sessão antes de começar a programar o passo 3).
- **Ajuste final da reorganização de pastas (mesma sessão, depois do alerta de sessão longa):** as subpastas de Ago-Dez tinham ficado só como `Actual`/`Flash`/`Forecast` (sem prefixo). Usuária pediu para seguir o mesmo padrão dos meses já usados (`MM_Mon_Ciclo`, ex: `07_Jul_Actual`). Renomeadas para `08_Aug_Actual/Flash/Forecast`, `09_Sep_...`, `10_Oct_...`, `11_Nov_...`, `12_Dec_...` — dentro das pastas `08 - Aug` a `12 - Dec`. Estrutura de pastas de 2026 agora está 100% padronizada e completa.

- **Automação do passo 3 (continuação da mesma sessão, depois de retomar):**
- **Decisão de arquitetura (aprovada pela usuária):** para alimentar a aba `Intermediária` do arquivo `Base Intermediária Fitted <Mês> <Ciclo> <Ano>.xlsx`, **não é preciso** abrir/crescer o arquivo acumulado `BASE_KSB1` (~45 mil linhas, com Tabelas Dinâmicas nativas `Pivot_Inter.`/`Pivot_Detalhes` e links externos). Motivo: o BASE_KSB1 só tem 2 usos — alimentar a Base Intermediária e fazer o "check do agrupamento" — e o segundo já é feito hoje direto dos extratos brutos por `check_agrupamentos_ksb1.py`, sem depender do BASE_KSB1. A classificação Fixo/Variável de cada `(Conta Fiscal, Centro de Custo)` já está fixada na própria `Intermediária` (coluna H), não precisa ser recalculada. Link externo `RHFitted February Actual 2026_.xlsx` dentro do BASE_KSB1 é confirmado **lixo/link errado pela usuária — ignorar, não corrigir**.
- **Criado `scripts/sap/gerar_base_intermediaria.py`:** soma o `Valor/MR` do extrato bruto do mês (Gestoriais ou Sem Agrupamento, mesma regra de decisão de 2026-08-11) agrupado por `(Centro de Custo, Conta Fiscal)`, exige que o Check de Agrupamentos do mês já exista (gera se não existir) como trava de qualidade, casa cada combinação com a linha única correspondente na aba `Intermediária` e preenche a coluna do mês. Nunca sobrescreve o arquivo original — salva sempre uma cópia nova versionada (`... - gerado.xlsx`) com uma aba extra `Pendências` (combinações sem linha correspondente ou com chave ambígua — linha duplicada na Intermediária).
- **Validado contra junho/2026 Actual real** (extratos antigos em `Bases SAP/`, fora do fluxo automatizado, usados só para o teste — não editam nada real): de 440 combinações, **409 bateram exatamente** com o valor já colado manualmente na coluna June. As 31 restantes se dividem em: 16 linhas que na Intermediária estavam com 0 mas deveriam ter valor (a automação teria preenchido corretamente — não é erro da lógica), 14 combinações com **linha duplicada** na Intermediária (mesma conta+centro em 2 linhas — ficam como pendência, corretamente não preenchidas às cegas) e 1 conta/centro (`N151420000`/8297) sem linha correspondente — que bate exatamente com a anotação manual já existente na planilha (linha 752: "REVISAR TODOS OS MESES, TROCAR DE 8297 PARA CC FIXO 8295"). Ou seja: a lógica acertou tudo que tinha uma linha única, e só ficou de fora exatamente o que já era um problema conhecido/estrutural na planilha manual.
- **Pendente nesta sessão (retomar):**
  1. Adicionar terceiro botão "Atualizar KSB1 Pivot" em `scripts/sap/atualizar_ksb1_gui.py` (nome escolhido pela usuária), chamando `atualizar_base_intermediaria` — precisa de um seletor de Ciclo (Actual/Flash) na GUI, que hoje só tem mês/ano.
  2. Registrar em `ontology/fitted_units.json` a regra da base intermediária (estrutura da aba `Intermediária`, chave `Conta Fiscal + Centro de Custo`, casos de duplicidade) e a estrutura de pastas de rede de 2026 — ainda pendente desde o resumo anterior de 2026-08-11.
  3. Registrar a decisão de arquitetura (pular o BASE_KSB1) em `memory/DECISOES.md`.
  4. Rodar o novo botão de ponta a ponta pela primeira vez com um mês real (a partir de agosto, quando a extração automatizada já grava em `00.Extração Base KSB1`).
- Plano completo salvo em `C:\Users\silveju001\.claude\plans\quizzical-imagining-hearth.md`.

- **Retomada à tarde do mesmo dia — reversão de arquitetura:** a usuária pediu pra voltar atrás e não usar mais o atalho direto do extrato (`gerar_base_intermediaria.py`). Motivo confirmado por ela: quer que a automação **fique fiel ao processo manual real** (não é sobre resolver as linhas duplicadas — perguntei diretamente e ela confirmou que não é esse o motivo). Decisão completa e motivo em `memory/DECISOES.md` (entrada "2026-08-11 (retomada à tarde)").
- **Novo plano (substitui o anterior), já com a estrutura real do arquivo inspecionada e confirmada** (sem adivinhar, ver `memory/learnings/2026-08-11_estrutura_real_base_ksb1_e_pivot.md`):
  1. Copiar `KSB1 <mês anterior> Actual <ano>.xlsx` → nova cópia versionada `KSB1 <mês> <ciclo> <ano>.xlsx`.
  2. Colar as linhas novas do extrato bruto do mês no fim do `BASE_KSB1` (colunas A-R, 1:1 confirmado) e replicar as fórmulas das colunas S-AI pras linhas novas (mesmo padrão, só o número da linha muda).
  3. Refresh das Pivot Tables nativas (`Pivot_Inter.`, `Pivot_Detalhes`) via automação COM do Excel (`win32com`) — **usuária confirmou OK até aqui**. Atenção: os arquivos precisam estar fechados no Excel dela enquanto a automação roda.
  4. Ler o `Pivot_Inter.` já atualizado (mesma chave Centro custo + Classe de custo da `Intermediária`) e colar o valor do mês nas linhas **brancas** da `Intermediária` — **esse passo ficou interrompido**, ver item novo abaixo.
- **Tema novo, não fechado — linhas coloridas na `Intermediária` (só existe no Flash, não no Actual):** além das ~753 linhas brancas fixas, a aba tem linhas amarelas (provisões ainda não contabilizadas no SAP — coluna V hoje diz "Reclass" nelas, deveria dizer "Prov"), verdes (reclassificações) e roxas (a apagar). A usuária ainda não ensinou como essas linhas coloridas entram no fluxo automatizado — só disse que existem e o que cada cor significa. Registrado em `ontology/fitted_units.json` → `intermediaria_linhas_coloridas_flash`.
- **Pedido pontual da usuária, ainda NÃO executado** (ela interrompeu a sessão antes de eu confirmar o arquivo-alvo): nas linhas coloridas de um arquivo Flash específico (ela não disse qual mês/ano), copiar as fórmulas que faltam pras linhas coloridas, trocar o texto da coluna V de "Reclass" para "Prov" nas linhas amarelas, e depois apagar as linhas roxas — nessa ordem. **Antes de fazer isso, perguntar qual arquivo/mês/ciclo exatamente** (ela não confirmou, a sessão foi encerrada por cansaço dela nesse ponto).
- **Fim de sessão:** usuária relatou cansaço/"bug na cabeça" e pediu resumo + backup pra retomar no dia seguinte com a mente limpa. Nada foi executado além de investigação (só leitura nos arquivos reais da rede) — nenhum dado real foi alterado nesta sessão.
- **Retomar amanhã, nesta ordem:**
  1. Perguntar em qual arquivo Flash aplicar a correção pontual (Reclass→Prov + apagar roxas + copiar fórmulas nas coloridas) e executar.
  2. Deixar a usuária explicar como as linhas coloridas entram (ou não) no fluxo automatizado do passo 3.
  3. Fechar o passo 4 do novo plano (como exatamente ler o `Pivot_Inter.` e colar nas linhas brancas da `Intermediária`, com a mesma trava de pendência de linhas duplicadas que já existia).
  4. Escrever o novo script (substituindo `scripts/sap/gerar_base_intermediaria.py`) e validar com o teste cego maio→junho (`KSB1 May Actual 2026.xlsx` + extrato bruto de junho, comparando contra o `KSB1 June Actual 2026.xlsx` real — sem sobrescrever nada).

---
## Continuação 2026-08-11 (noite) — Claude Desktop bloqueado pelo TI + novo projeto "Fitted Recuperação"

- **Claude Desktop app:** instalado com sucesso sem admin (Squirrel installer, per-user, `C:\Users\silveju001\AppData\Local\AnthropicClaude`), mas **não conecta** — erro "Sua rede redirecionou esta solicitação para gateway.zscloud.net". Testei via `curl` direto (sem passar pelo app) e claude.ai/api.anthropic.com/etc respondem normalmente, sem redirecionamento — ou seja, não é certificado mal configurado, é o Zscaler da Pirelli bloqueando especificamente o app desktop (provável política de "AI Apps" que libera o uso via navegador mas bloqueia o cliente nativo). **Não é algo que dá pra contornar sem o TI** — usuária vai falar com o TI amanhã (2026-08-12). Enquanto isso, usar claude.ai pelo navegador funciona normalmente.

- **Novo sub-projeto: "Fitted Recuperação"** (nome dado pela usuária) — dentro do domínio Fitted Units, ainda não registrado em `ontology/fitted_units.json`/`PROJECT_MAP.md` (fazer isso assim que o escopo fechar). Objetivo: detectar lançamento/pagamento a fornecedor em duplicidade na KSB1, período **01.01.2026 a 31.07.2026**, excluindo documentos estornados. Critérios de duplicidade confirmados pela usuária: Fornecedor+Valor+Documento de compras, e Fornecedor+Valor+Data. Estorno identificado como par de lançamentos do mesmo fornecedor com valores opostos que se cancelam.
- **Scripts novos criados (fora do fluxo mensal recorrente, são específicos deste estudo):**
  - `scripts/sap/extrair_ksb1_periodo.py` — extrai a KSB1 (Fitted Units, Sem Agrupamento) para um período arbitrário (não só um mês), reaproveitando as funções de conexão/navegação já testadas de `atualizar_ksb1_gui.py`. Salva em `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Estudos\Estudo Duplicidade Pagamento\` (caminho passado pela usuária, fora da estrutura de pastas do fechamento mensal).
  - `scripts/sap/analisar_duplicidade_pagamento.py` — lê o extrato, ignora linhas de subtotal e sem fornecedor, identifica pares de estorno, aplica os dois critérios de duplicidade e gera `Análise Duplicidade Pagamento.xlsx` (abas Resumo, Dup. por Documento, Dup. por Data, Estornos identificados) na mesma pasta.
- **Rodado contra o período real 01.01-31.07.2026:** extrato com 175.310 linhas de detalhe, 7.147 com fornecedor preenchido. 50 pares de estorno identificados e excluídos. Resultado: **228 grupos de duplicidade por Documento de compras** (639 linhas, R$ 823.535,87) e **649 grupos por Data** (1.508 linhas, R$ 1.533.184,43). Deixei claro pra usuária que isso é triagem heurística (pode ter falso positivo), não confirmação de erro.
- **Tentativa de refinar com Nº de NF — pausada, não deu certo:** a coluna "Nº doc.de referência" da KSB1 não é a NF (usuária confirmou). Investiguei a ZLFIB ao vivo (tela real "Pesquisa genérica de Notas Fiscais", campos `S_DOCNUM`/"Nr Documento", `S_NFNUM`/"Nota Fiscal", `S_PARID`/"Parceiro NF", precisa de `P_BUKRS`/"Empresa" = 0580 pra buscar) mas nenhuma busca trouxe resultado confiável, e a própria usuária confirmou que não sabe como esse cruzamento funciona na prática. **Decisão: pausar essa ideia até alguém do time funcional/TI do SAP confirmar o vínculo certo** — não tentar de novo por tentativa e erro. Detalhe completo em `memory/PROJECT_MAP.md` → sub-projeto Fitted Recuperação. Seguimos só com os dois critérios que já funcionam (Documento e Data).
- **Pendência ainda em aberto de mais cedo, não esquecer:** a correção pontual nas linhas coloridas do Flash (Reclass→Prov, apagar roxas, copiar fórmulas) segue sem o nome do arquivo/mês confirmado pela usuária — ver resumo de 2026-08-11 (retomada à tarde) acima.

- **Continuação — duplicidade de NF direto na ZLFIB (mais confiável que cruzar KSB1↔ZLFIB, que foi abandonado):** a usuária mostrou a tela real da ZLFIB e deu as regras certas: usar "Lista por Item" (não "Lista por Nota", que não funciona bem), marcar "Buscar chave de acesso", e os códigos de Filial da Fitted Units são **0031=SJP, 0032=IBI, 0053=SOR, 0054=GOI**. Empresa = 0580.
  - Mapeei a grid real via SAP GUI Scripting (ALV Grid Control, não list clássica — export por menu não funcionou, ler direto por `GetCellValue` via COM foi mais rápido e confiável: ~3,5ms/linha). Colunas técnicas: `BRANCH, DOCNUM, NFNUM, SERIES, NFTYPE, PSTDAT, DOCDAT, PARID, NAME1, NFTOT, ACKEY` (chave de acesso). Cada NF aparece repetida uma vez por item (mesmo `DOCNUM`) — preciso agrupar por `DOCNUM` pra voltar ao nível de NF antes de procurar duplicidade.
  - Critério de duplicidade definido: **notas com chave de acesso** (mercadoria) → mesma `ACKEY` em mais de um `DOCNUM` = duplicidade certa. **Notas de serviço (sem chave de acesso)** → mesmo Parceiro+Nota Fiscal+Série+Valor total em mais de um `DOCNUM` = candidato a duplicidade.
  - Script criado: `scripts/sap/analisar_zlfib_duplicidade.py` (roda as 4 filiais em sequência, reabrindo a ZLFIB do zero a cada uma, e gera `Análise Duplicidade NF (ZLFIB).xlsx` na mesma pasta do estudo).
  - **Rodei parcialmente:** Filial 0031 (SJP) terminou rápido, 3.311 linhas de item. Ao entrar na Filial 0032 (IBI), a busca trouxe **317.042 linhas** — quase 100x mais que SJP. Não confiei nesse número às cegas (poderia ser um bug real, tipo o filtro de Filial não pegando, mas confirmei manualmente que o campo BRANCH da grid realmente mostrava "0032" nas primeiras linhas, então parece genuíno, não bug de filtro) — mas como a usuária pediu pra desligar o computador, **matei o processo antes de terminar de ler** (nada corrompido, só incompleto).
  - **Retomar amanhã:** 1) confirmar com a usuária se um volume de ~317 mil linhas de item pra IBI num período de 7 meses é plausível (IBI atende 3 montadoras diferentes — CNH, Iveco, Fiat — pode ser volume real, mas vale checar antes de rodar de novo às cegas por muito tempo); 2) se fizer sentido, rodar `analisar_zlfib_duplicidade.py` de novo do zero (não dá pra retomar do meio, o script só salva o resultado no final); 3) considerar se vale a pena adicionar algum filtro extra (Cfop, Tipo NF) pra reduzir o volume de IBI antes de ler linha a linha, se o motivo do volume alto for esse.

---
## Resumo do dia 2026-08-10 — Fitted Units Despesas (KSB1)
> Detalhe completo arquivado em `memory/long_term/2026-08-10_*_briefing_snapshot.md`. Entregas principais: `scripts/sap/atualizar_ksb1_gui.py` (extração KSB1 Gestoriais+Sem Agrupamento via GUI Scripting) e `scripts/sap/check_agrupamentos_ksb1.py` (conferência de agrupamento gestorial), acessados por um atalho único na rede (`ATUALIZAR KSB1.lnk`). Regras em `ontology/fitted_units.json` → `classificacao_despesas.check_de_agrupamentos`.

---

## Continuação 2026-08-13 — Atalho "Conversar" + retomada do Fitted Recuperação (ZLFIB)

- **Atalho de acesso rápido:** criado `Conversar.bat` na Área de Trabalho da usuária (`C:\Users\silveju001\Desktop\Conversar.bat`) — duplo clique entra na pasta do projeto e abre o Claude Code (`claude.cmd`, encontrado em `C:\Users\silveju001\node-v24.19.0-win-x64\`). Não versionado no Git (é local, específico da máquina dela).

- **Retomada do sub-projeto Fitted Recuperação (ZLFIB), com novas regras de negócio da usuária:**
  - **Regra nova confirmada:** só interessam notas de **Entrada** (fornecedor) — excluir Saída. Campo achado ao vivo no SAP: `S_DIRECT` (Direção do movimento de mercadorias), valores possíveis 1=Entrada, 2=Saída, 3=Devolução saída transf. estoque, 4=Devolução entrada transf. estoque. Aplicado como filtro na tela de seleção da ZLFIB. **Testado em SJP/jan-2026: reduz de 381 para 175 linhas de item (~46%)** — deve ajudar bastante o volume alto de IBI (~317 mil linhas) que tinha travado a sessão anterior.
  - **Pedido NÃO resolvido — filtro "operação A24" (transferência de material):** a usuária queria excluir notas dessa operação, mas o código "A24" **não foi localizado** em nenhum campo de filtro da tela de seleção da ZLFIB testado ao vivo: Cfop (rejeita como código inválido) e Tipo NF/NFTYPE (campo de 2 caracteres só, valores reais amostrados são tipo R8/RF/YE/YS — nunca 3 dígitos). A usuária mencionou que normalmente vê esse código num "campo OPERA", mas não confirmou onde (perguntei e ela mudou de abordagem antes de responder — ver decisão abaixo). **Não tentar mais achar esse campo por tentativa e erro sem ela confirmar a tela/transação exata.**
  - **Decisão da usuária (pedido explícito):** seguir por enquanto **sem** o filtro A24/OPERA, e em vez disso **cruzar o resultado da ZLFIB com os fornecedores já identificados como duplicados no estudo da KSB1** (`Análise Duplicidade Pagamento.xlsx`, gerado em 2026-08-11 por `analisar_duplicidade_pagamento.py`) — como sinal cruzado de confiança, já que os dois estudos usam o mesmo código de fornecedor (`PARID` na ZLFIB = `Fornecedor` na KSB1, confirmado formato idêntico ao amostrar os dois arquivos). Pediu pra rodar primeiro só a Filial 0031 (SJP), por ser menor e mais simples de validar antes de ir pras outras.
  - **`scripts/sap/analisar_zlfib_duplicidade.py` reescrito:**
    - Filtro `S_DIRECT-LOW = "1"` (só Entrada) adicionado em `buscar_filial`.
    - `FILIAIS` deixou de ser fixo — `analisar()` agora aceita parâmetro `filiais` (dict), permitindo rodar um subconjunto (`__main__` agora chama só `{"0031": "SJP"}`).
    - Nova função `carregar_fornecedores_duplicados_ksb1()` lê `Análise Duplicidade Pagamento.xlsx` (abas "Dup. por Documento" e "Dup. por Data") e monta o conjunto de códigos de fornecedor já duplicados na KSB1.
    - Nas duas abas de duplicidade da ZLFIB, nova coluna "Fornecedor também duplicado na KSB1" (Sim/vazio, com destaque amarelo quando Sim) e no Resumo, contagem de quantos fornecedores duplicados na KSB1 também aparecem com NF duplicada na ZLFIB.
  - **Resultado da primeira rodada (SJP, jan-jul/2026, só Entrada) — NÚMEROS ERRADOS, ver correção abaixo:** a primeira leitura (via `GetCellValue` célula a célula) relatou 38 notas únicas e 0 duplicidades. **Esse número estava errado** — ver bug crítico abaixo.
  - **BUG CRÍTICO ENCONTRADO E CORRIGIDO — leitura da grid via COM não é confiável:** ao rodar a mesma filial (SJP) sem o filtro de Direção pra comparar, a leitura por COM achou só 27 notas únicas — menos que as 38 "de Entrada", o que é matematicamente impossível (Entrada tem que ser subconjunto de "todas as direções"). Comparando os `Nr Documento` um a um, a leitura por COM estava incompleta/embaralhada. **Correção:** reescrito `analisar_zlfib_duplicidade.py` pra exportar a grade pra arquivo de verdade via menu nativo do SAP (Lista > Exportar > Planilha eletrônica — mesmo mecanismo já usado na extração da KSB1) e ler o arquivo com `openpyxl`, em vez de `GetCellValue`. Validado: SJP realmente tem **565 notas únicas** (não 38). Detalhe completo em `memory/errors/2026-08-13_zlfib_getcellvalue_dados_incorretos.md`. **Lição geral: nunca confiar em leitura de grid ALV via GetCellValue em grades grandes (a partir de ~1.000 linhas) — sempre exportar pra arquivo.**
  - **Resultados corretos, com o método de exportação (jan-jul/2026, só Entrada):**
    - SJP: 1.348 linhas → 565 notas únicas → 0 duplicidades.
    - IBI: 25.369 linhas → 14.867 notas únicas → 0 duplicidades.
    - SOR+GOI (rodadas juntas por conveniência, mas são **plantas diferentes** — Sorocaba e Goiana, usuária corrigiu isso explicitamente): 12.565 linhas → 6.592 notas únicas → **2 grupos duplicados por Chave de Acesso encontrados inicialmente, mas usuária confirmou que são transferência de material (Tipo NF 'R8', parceiro FIAT AUTOMOVEIS S/A, mesma chave de acesso mas valores diferentes) — não é duplicidade de pagamento a fornecedor.** Adicionada exclusão automática de `NFTYPE == 'R8'` em `analisar_zlfib_duplicidade.py` (constante `NFTYPE_EXCLUIDOS`). Esse pode ser exatamente o campo/conceito que a usuária chamava de "OPERA"/"A24" — não confirmado 100%, mas o padrão bate (transferência de material, código específico de Tipo NF).
    - **Conclusão do período jan-jul/2026, com os dados corretos: nenhuma duplicidade real de NF encontrada nas 4 filiais.**
  - **Aviso dado à usuária, importante:** duplicidade zero na ZLFIB (documento fiscal) **não** significa que não pagamos a mesma nota duas vezes — isso é sobre o cadastro da NF, não sobre o lançamento de pagamento. O estudo que cobre isso é o da KSB1 (`analisar_duplicidade_pagamento.py`), que já achou candidatos reais (228 grupos por Documento, 649 por Data) ainda não revisados manualmente. Pra confirmar pagamento duplicado de verdade, o caminho mais direto seria olhar documentos de pagamento do fornecedor (ex: FBL5N/FBL1N) — ainda não explorado.
  - **`analisar_zlfib_duplicidade.py` agora aceita `data_de`/`data_ate` como parâmetros** (antes eram fixos no módulo, hardcoded pro estudo jan-jul) — necessário pra reusar a mesma lógica na checagem mensal (abaixo). `analisar()` agora devolve um dict (`arquivo`, `tem_duplicidade`, `grupos_dup_chave`, `grupos_dup_sem_chave`, `notas_envolvidas`, `valor_total`) em vez de só o Path.

- **Nova automação pedida pela usuária: checagem mensal automática, disparada pelo login no SAP (não por horário fixo).**
  - Pedido: no primeiro dia útil de cada mês, assim que ela logar no SAP, rodar a checagem de duplicidade ZLFIB do **mês anterior** (todas as 4 filiais). Se achar duplicidade real, mandar e-mail pra ela mesma (`juliana.silveira@pirelli.com`, e-mail corporativo — diferente do e-mail pessoal já registrado no contexto do Claude) com o Excel de duplicidade em anexo. **Se não houver duplicidade, não notificar.**
  - Decisões confirmadas com a usuária (perguntei antes de automatizar por envolver e-mail, ação irreversível):
    1. E-mail deve ser **enviado automaticamente** (não fica como rascunho pra revisão).
    2. Se o SAP não estiver logado até um horário limite no primeiro dia útil, mandar um e-mail de aviso (sem anexo) avisando que a checagem não rodou — pra ela lembrar de rodar manualmente.
    3. Frequência do polling: **de hora em hora** (ela pediu explicitamente, eu tinha sugerido 15 em 15 min).
  - **Limitação técnica explicada e aceita pela usuária:** não existe um jeito de "escutar" o evento de login do SAP de fora — a automação não consegue abrir o SAP nem fazer login sozinha (não tem a senha dela). A solução é *polling*: um script roda de hora em hora via Agendador de Tarefas do Windows e checa se o SAP já está logado (heurística validada ao vivo: `session.Info.User` vem preenchido com o usuário SAP dela, `'SILVEJU001'`, só depois do login). Se o computador estiver desligado/ela deslogada, a automação não roda até a próxima checagem depois que ela logar (ou manda o aviso por e-mail se passar do horário limite).
  - **Criado `scripts/sap/verificacao_mensal_zlfib.py`:**
    - `mes_anterior_range(hoje)`: calcula automaticamente o período do mês anterior (datas no formato que o SAP espera).
    - `primeiro_dia_util_do_mes(ano, mes)`: primeiro dia útil = primeira segunda-a-sexta do mês. **Limitação conhecida, não resolvida:** não considera feriados nacionais/municipais — se o 1º dia útil "de calendário" cair num feriado, a rotina roda mesmo assim nesse dia (não pula pro próximo). Avisar a usuária disso; podemos adicionar um calendário de feriados depois se ela quiser mais precisão.
    - `sap_esta_logado()`: tenta conectar numa sessão do SAP GUI já aberta; considera "logado" se `session.Info.User` vier preenchido.
    - Estado persistido em `data/processed/zlfib_mensal_estado.json` (fora do Git, é `data/processed/` — já no `.gitignore`): guarda `ultimo_mes_verificado` (evita rodar de novo no mesmo mês) e `aviso_indisponibilidade_mes` (evita mandar o e-mail de aviso mais de uma vez no mesmo mês).
    - `enviar_email()`: usa automação COM do Outlook (`win32com.client.Dispatch("Outlook.Application")`) — testado ao vivo, Outlook está instalado e acessível (versão 16.0.0.20026). `mail.Send()` envia direto (não `mail.Save()`, conforme decisão da usuária de envio automático).
    - `watcher()`: função principal chamada pelo Agendador — sai rápido se hoje não for o 1º dia útil ou se o mês já foi verificado; senão tenta achar o SAP logado e, se achar, roda `analisar()` do mês anterior (todas as filiais, mesma regra de Entrada + exclusão de R8) e manda e-mail só se `tem_duplicidade` for verdadeiro.
  - **Criado `scripts/sap/watcher_mensal_zlfib.bat`** (usa `%~dp0` — sem caminho absoluto hardcoded) que roda `python scripts\sap\verificacao_mensal_zlfib.py`, com saída redirecionada pra `logs\zlfib_mensal.log` (pasta `logs/` nova, adicionada ao `.gitignore`).
  - **AINDA NÃO REGISTRADO NO AGENDADOR DE TAREFAS DO WINDOWS — pendência para a próxima sessão.** Tentei rodar `schtasks` via Bash (Git Bash) pra criar a tarefa (`schtasks /create ...`, de hora em hora) mas esbarrei num problema de compatibilidade: o Git Bash está convertendo os argumentos `/query`, `/create` etc. como se fossem caminhos de arquivo (mangling de path do MSYS), e mesmo tentando `cmd.exe /c "..."` como wrapper, o comando não retornou saída nenhuma (parece ter aberto um cmd interativo em vez de executar e sair — não travou nada, só não voltou resultado). **Não tentei mais tentativas de contorno nesta sessão** (bati no limite de 45 ações de novo). Próxima sessão: resolver isso — alternativas a testar: (a) usar PowerShell (`powershell.exe -Command "..."` ou `Register-ScheduledTask`) em vez de `schtasks` puro: pode lidar melhor com aspas/args vindos do Git Bash; (b) usar `MSYS2_ARG_CONV_EXCL="*"` como variável de ambiente antes do comando `schtasks` pra desativar completamente a conversão de path do Git Bash; (c) escrever um `.ps1` ou `.vbs` auxiliar que crie a tarefa, e rodar esse arquivo em vez de passar os argumentos direto na linha de comando.
  - Nome sugerido pra tarefa (ainda não criada): `Verificacao_ZLFIB_Duplicidade_Mensal`, gatilho recorrente de hora em hora, ação = rodar `scripts\sap\watcher_mensal_zlfib.bat` a partir da raiz do projeto.
  - **CONCLUÍDO nesta mesma sessão:** tarefa `Verificacao_ZLFIB_Duplicidade_Mensal` registrada com sucesso no Agendador de Tarefas do Windows (rodando de hora em hora, sem hora de término). O problema do Git Bash mangling os argumentos do `schtasks` foi resolvido com a variável de ambiente `MSYS2_ARG_CONV_EXCL="*"` antes do comando (evita que o Git Bash converta `/create`, `/tn` etc. como se fossem caminhos). Lógica de datas testada isoladamente (`primeiro_dia_util_do_mes`, `mes_anterior_range`) em vários meses, inclusive virada de ano — bateu certo. **E-mail de teste enviado via `enviar_email()` (Outlook COM) e confirmado recebido pela usuária.** Rodei o `watcher()` de verdade hoje (não é o 1º dia útil do mês) e confirmei que ele sai sem fazer nada, como esperado. **Essa automação está 100% pronta e ativa** — só falta o mês virar pra ver rodando "pra valer" no cenário real (1º dia útil, aguardando SAP logado).

- **Nota:** nesta sessão não houve avanço no passo 3 do processo recorrente (BASE_KSB1 + Pivot nativo) nem na correção pontual das linhas coloridas do Flash — ambos continuam pendentes, ver resumo de 2026-08-11 acima.

---
## Próximos passos (retomar — só Fitted Units)
- **A automação mensal da ZLFIB (checagem de duplicidade) está concluída e ativa** — nada pendente nela por enquanto, só acompanhar se ela dispara certinho no próximo 1º dia útil do mês.
- **Prioridade:** ver a lista "Retomar amanhã" no resumo de 2026-08-11 (retomada à tarde) acima — primeiro a correção pontual pendente no Flash (Reclass→Prov + linhas roxas), depois fechar e escrever o novo script do passo 3 (BASE_KSB1 + Pivot nativo, não mais o atalho antigo).
- Depois: passos 4-5 do processo (`ontology/fitted_units.json` → `processo_recorrente`) — rateio dos custos da Gerência (GER) pras demais unidades, depois carregar no arquivo de P&L.
- Usuária ainda precisa rodar "Gerar Check de Agrupamentos" com um mês real (Gestoriais + Sem Agrupamento do mesmo mês) pela primeira vez desde o último ajuste (aba única + Check 1 sempre listando as 6 contas fixas) e confirmar se bate com o esperado — se isso já rolou informalmente durante as extrações desta sessão, confirmar com a usuária antes de assumir que ainda está pendente.
- `memory/PROJECT_MAP.md`: Original Equipment ainda não detalhado (fora do escopo por ora, já que o foco agora é só Fitted Units).
- Circuito Panamericano: não mexer até a usuária pedir explicitamente.
- Forecast: fora do escopo por enquanto (confirmado pela usuária em 2026-08-11), mas as pastas de rede já foram deixadas prontas pra quando for retomado.

---
## Contexto permanente do projeto
- Esta pasta (`C:\Users\silveju001\Projetos Claude`) está estruturada seguindo o "Guia de Onboarding — Como Trabalhar com o Claude de Forma Profissional" (maio 2026, baseado no projeto Cockpit Ind — Pirelli Planning & Control).
- Objetivo real deste projeto: automatizar controladoria (Fitted Units e Circuito Panamericano) hoje feita em Excel — resultado, despesas, faturamento, EBIT, P&L mensal. Detalhes completos em `CLAUDE.md`. **Foco atual (a partir de 2026-08-10): só Fitted Units.**
- Repositório Git já configurado com backup remoto no GitHub; backup automático diário às 18h (Agendador de Tarefas do Windows) além do backup por sessão longa (a cada 45 ações).
- GUI compartilhável da KSB1 (`scripts/sap/atualizar_ksb1_gui.py`, atalho único `ATUALIZAR KSB1.lnk` na rede) e o Check de Agrupamentos (`scripts/sap/check_agrupamentos_ksb1.py`) já em uso — ver `memory/DECISOES.md` para o histórico completo de decisões sobre eles.

### Continuação (mesmo dia) — Achado do mecanismo do popup "Segurança SAPGUI" e proposta de regra curinga (AINDA NÃO APLICADA, aguardando confirmação)

- **Usuária perguntou:** já que cada mês vira uma pasta nova (`00.Extração Base KSB1/<mês>/`), o popup "Segurança SAPGUI" vai continuar pedindo confirmação manual todo mês daqui pra frente — pediu ajuda pra resolver isso de vez.
- **Investigação (só leitura até aqui):** o SAP GUI guarda essas aprovações em `C:\Users\silveju001\AppData\Roaming\SAP\Common\saprules.xml` (arquivo local, editado automaticamente pelo próprio SAP GUI toda vez que a usuária clica "Permitir" + "Memorizar minha decisão"). Achado importante: já existe uma regra genérica com **curinga** (`\...\Resultados Fitted\*`, permissão `r`, sem contexto de transação) que cobre TODA a árvore de leitura, independente de ano/mês — é por isso que ler nunca pede confirmação. Já a permissão de **escrita** é sempre uma entrada específica por pasta exata (uma por mês: Jan, Fev, Mar, Abr, Mai, Jun, Jul já estão lá individualmente, cada uma amarrada ao contexto da transação KSB1/tela de export).
- **Backup feito:** `saprules.xml.backup_2026-08-21` (cópia idêntica, mesma pasta) — antes de qualquer edição.
- **Proposta ainda não aplicada (esperando confirmação da usuária):** adicionar uma regra nova, no mesmo formato das já existentes (mesmo contexto: sistema G20, transação KSB1, tela SAPLSLVC_FULLSCREEN/0200, permissão `w`), mas com diretório em curinga (`.../Resultados Fitted/2026/00.Extração Base KSB1/*`) em vez de uma pasta exata — cobriria Ago-Dez/2026 automaticamente, sem popup novo a cada mês. Risco identificado e ainda não resolvido: não há certeza de que o curinga funciona do mesmo jeito pra regras COM contexto de transação (`action=3`, como as de escrita) quanto funciona pra regra SEM contexto (`action=0`, a de leitura que já existe) — é um formato interno não documentado do SAP GUI. Editar esse arquivo errado tem risco real de quebrar TODAS as aprovações já acumuladas (mais de 1 ano de regras, não só do nosso projeto) se o XML ficar malformado.
- **Pendente:** decidir com a usuária se aplica essa edição (com teste depois, tentando extrair um mês futuro pra confirmar que não pede popup) ou se ela prefere continuar clicando "Permitir + Memorizar" manualmente todo mês (baixo custo, ~1x/mês, zero risco).
- **Sessão atingiu o limite de 45 ações — backup automático já rodou** (`session_transition.py`).

### Continuação (mesmo dia) — Cockpit promovido pra produção de vez, validação completa Jan-Jul, bug de pastas corrigido, ajustes finos

- **Decisão do popup SAP fechada:** usuária optou por continuar clicando "Permitir + Memorizar" manualmente (não editar `saprules.xml`) — ver `DECISOES.md`.
- **Cockpit promovido pra produção:** botão "Atualizar Pivot KSB1" (renomeado) agora chama `gerar_ksb1_mensal.py` de verdade, escrevendo na pasta de rede oficial (não mais teste local). Painel: Ano antes do Mês, virou combobox selecionável. Rastro de pneu removido do rodapé (pedido da usuária). Atalho de rede renomeado de "ATUALIZAR KSB1" pra "Fechamento Custo Fitted Units".
- **Bug real achado e corrigido:** pastas de Março/Abril usam mês por extenso (`03_March_Actual`/`04_April_Actual`) em vez da abreviação padrão — quebrava o Passo 3 pra Abril/Maio. Corrigido com `resolver_pasta_ciclo` (tolerante, sem lista hardcoded) em `ksb1_core.py`.
- **Validação completa Jan-Jul/2026 (Actual) contra o valor REAL fechado** (não só "sem erro"): todos os 7 meses batem 100%, toda diferença explicada por unidades encerradas (Sorocaba) + 1 inconsistência manual já conhecida (centro 8295/8297, líquido zero). Detalhe mês a mês em `DECISOES.md`.
- **Confirmado com a usuária (checagem ao vivo, não por memória):** coluna S ("Nº doc.de referência") do extrato bruto já é ignorada de propósito no Passo 3 — só cola A-R (18 colunas), consistente nos dois tipos de arquivo (Gestoriais e Sem Agrupamento).
- **Pendente, a pedido explícito da usuária — NÃO gerar ainda:** documento formal "passo a passo pro estagiário" explicando os 3 botões — ela quer terminar mais etapas no cockpit antes (2º botão da aba ③, "arquivo colorido", ainda sem escopo definido).
- **Sessão atingiu o limite de 45 ações de novo — backup automático já rodou.**

### Continuação (mesmo dia) — Passo 4 (Base Intermediária) construído do zero, testado e ligado na GUI

- **Passo 4 implementado** (`gerar_base_intermediaria.py`, reescrito do zero): copia o mês anterior, apaga a área de dados a partir da 1ª linha sem cor (linhas coloridas 2-67 são reservadas pra reclassificações/provisões, só existem de verdade no Flash, nunca tocadas), cola tudo de novo do `Pivot_Inter.` do BASE_KSB1 — full rebuild a cada rodada, não atualização incremental (decisão explícita da usuária, "pra evitar erros"). Unidades encerradas continuam como linha mas com valor em branco; o excluído vai pra um arquivo mensal separado (`Histórico Unidades Encerradas - ...xlsx`, pra usuária mandar pra contabilidade). Arrasta fórmulas de Total Ano (U) e Y:AJ.
- **Bug real achado e corrigido, importante pra qualquer script futuro que cole muitas linhas via COM:** escrever um array grande (600+ linhas) numa única chamada `Range.Value` via pywin32/Excel corrompe silenciosamente algumas células em erro `#N/A` — só aparece depois de salvar e reabrir, sem relação com o conteúdo. Corrigido colando linha por linha + verificação pós-colagem. Mesma proteção aplicada preventivamente no Passo 3 (`gerar_ksb1_mensal.py`), que tinha o mesmo padrão de risco — confirmado que nenhum arquivo de produção foi afetado até agora (Passo 3 nunca rodou de verdade contra a rede).
- **Achado secundário:** link externo quebrado (RHFitted) no arquivo real de julho, removido pela usuária direto no Excel (backup feito antes).
- **Refresh da Pivot Table** da aba "Pivot" adicionado (`wb.RefreshAll()`) — não acontecia sozinho.
- **Quadro "(+) gain" (comparação Flash x Actual):** implementada `atualizar_comparacao_flash` — traz Despesas/Mão de Obra da aba Pivot do arquivo Flash do mesmo mês (linhas 15/16) pro quadro de comparação (linhas 18/19), automatizando o que a usuária fazia à mão. Também corrige as fórmulas de Custos (H26/I26) do quadro amarelo "Month/Flash/Actual/delta", que ficavam travadas na coluna do último mês editado — passam a apontar sempre pro mês atual. Faturamento (linha 25) fica de fora por enquanto, usuária vai automatizar depois.
- **Validação final (Julho):** tudo bateu exatamente com o arquivo real — Intermediária (587 combinações, zero diferença), quadro de comparação Flash (Despesas R$ 4.334.644,06, Mão de Obra R$ 2.426.107,40).
- **Ligado na GUI:** segundo botão na aba ③ (mesma aba do Passo 3, a pedido da usuária — não virou aba nova), rotulado "Finalização da Base Intermediária". Ainda não testado ao vivo pela GUI (só via linha de comando).
- **Testando agora:** o ajuste de H26/I26 (rodando em segundo plano no momento do alerta de sessão longa) — retomar conferindo o resultado.
- **Pendente:** Ciclo Flash (lógica das linhas coloridas manuais ainda não detalhada pela usuária); Faturamento (linha 25) ainda manual, automatizar depois.
- **Sessão atingiu o limite de 45 ações de novo — backup automático já rodou.**

### Continuação (mesmo dia) — Ciclo Flash implementado (provisões), cockpit reestruturado em 4 passos

- **Ciclo Flash (básico) implementado:** provisões preenchidas automaticamente a partir do "Fast Provisão" (versão mais alta) da pasta de rede, na aba `Ficha de Solicitação` — Conta/Centro/Valor mapeados pras linhas coloridas da Intermediária, fórmulas VLOOKUP (A,B,D,F,G) arrastadas do "molde" (última linha colorida, roxa). Provisões sempre começam do zero a cada mês (contabilidade estorna o mês anterior). Validado contra Julho Flash real: área branca zero diferença; 2 achados nas provisões confirmados como esperados pela usuária (coluna D do arquivo real estava incompleta — meu script está certo; pequenas diferenças de valor/centro porque uso sempre a versão mais alta do Fast Provisão, e o arquivo real foi montado com versão mais antiga).
- **Quadro de comparação Flash x Actual (H26/I26/L25/linhas 18-19)** fica de fora quando o ciclo é Flash — usuária vai detalhar depois. Faturamento (linha 25) continua manual (usuária mostrou um quadro de Faturamento por Centro de Montagem que será automatizado num momento futuro).
- **Cockpit reestruturado a pedido da usuária:** novo Passo 3 "Provisões" (2 botões: "Lançar Provisões" cria+preenche pela primeira vez, "Atualizar Provisões" relê o Fast Provisão mais recente sem duplicar arquivo) inserido antes do antigo Passo 3, que virou Passo 4 "Base Intermediária" (mesmos 2 botões de antes). Trava de segurança: não lê o Fast Provisão se ele estiver aberto no Excel (checagem por arquivo de lock `~$`) — testada ao vivo, funcionou (bloqueou de verdade, depois passou quando a usuária fechou o arquivo).
- **Validado de ponta a ponta:** Lançar Provisões → Finalização rodados em sequência contra Julho Flash real, resultado idêntico ao teste anterior à reestruturação.
- **Cockpit reaberto** com os 4 passos pra usuária testar pela interface (só testado via linha de comando até agora).
- **Pendente:** quadro de comparação Flash x Actual pro caso do Flash (usuária vai detalhar), Faturamento (ainda manual), inserção automática de linha colorida se provisões excederem capacidade (hoje para com erro claro).
- **Sessão atingiu o limite de 45 ações de novo — backup automático já rodou.**

---
## RESUMO DO DIA 2026-08-21 — Fitted Units Despesas: Ciclo implementado, Passo 3/4 em produção, Base Intermediária construída do zero (Actual + Flash básico), cockpit reestruturado

**Sessão muito longa (múltiplos alertas de 45 ações) — este é o resumo consolidado de tudo que foi feito, na ordem em que aconteceu. Detalhe técnico completo de cada item está em `memory/DECISOES.md`, todas datadas 2026-08-21.**

### 1. Ciclo (Actual/Flash) na extração e no Passo 3
- Extração (Passo 1) passou a gravar o Ciclo no nome do arquivo bruto; Check (Passo 2) e Passo 3 (KSB1 Pivot) passam a buscar pelo Ciclo pedido, com fallback pros meses antigos sem Ciclo no nome.
- Lacuna de dados descoberta e preenchida: Jan/Fev/Mar/Jun/2026 não tinham extração bruta salva — extraído ao vivo via SAP.
- Jan-Jul/2026 (Actual) validados 100% contra o valor real fechado — toda diferença explicada por unidades encerradas (Sorocaba) + 1 inconsistência manual já conhecida.
- Bug real achado e corrigido: pastas de Março/Abril usam mês por extenso (`03_March_Actual`) em vez da abreviação padrão — quebrava o Passo 3 pra Abril/Maio. Corrigido com `resolver_pasta_ciclo`, tolerante sem lista hardcoded.
- Confirmado (checagem ao vivo): coluna S do extrato bruto já era ignorada de propósito, não precisou mudar nada.

### 2. Cockpit promovido pra produção
- Botão "Atualizar Pivot KSB1" (Passo 3, depois renumerado Passo 4) passou a chamar `gerar_ksb1_mensal.py` de verdade, escrevendo na pasta de rede oficial (antes era pasta de teste local).
- Rastro de pneu removido do rodapé; atalho de rede renomeado "ATUALIZAR KSB1" → "Fechamento Custo Fitted Units".
- Cores trocadas de vermelho Pirelli pra amarelo claro (`#FFE9A8`) em todo o cockpit.
- Tamanho padrão da janela fixado em 1317x800.

### 3. Bug crítico achado e corrigido: corrupção silenciosa ao colar array grande via COM
- Escrever um array grande (600+ linhas) numa única chamada `Range.Value` via pywin32/Excel corrompe aleatoriamente células em erro `#N/A` — só aparece depois de salvar e reabrir, sem relação com o conteúdo.
- Corrigido em `gerar_base_intermediaria.py` e preventivamente em `gerar_ksb1_mensal.py` (colar linha por linha + verificação pós-colagem).
- Confirmado que nenhum arquivo de produção foi afetado (Passo 3/4 nunca tinham rodado de fato contra a rede antes de hoje).

### 4. Passo 4 (Base Intermediária) construído do zero e validado — Ciclo Actual
- `gerar_base_intermediaria.py` reescrito: full rebuild a cada rodada (apaga e cola tudo de novo do `Pivot_Inter.` do BASE_KSB1), unidades encerradas com valor em branco + arquivo histórico mensal separado (pra e-mail à contabilidade), refresh da Pivot Table, quadro de comparação Flash x Actual (Despesas/Mão de Obra, fórmulas de Custos H26/I26 dinâmicas, câmbio L25 puxado do Flash).
- Validado contra Julho/2026 real: zero diferença em tudo.

### 5. Ciclo Flash (básico) implementado e validado
- Provisões preenchidas automaticamente do "Fast Provisão" (versão mais alta) da pasta de rede — Conta/Centro/Valor mapeados, fórmulas VLOOKUP arrastadas do molde.
- Provisões sempre começam do zero a cada mês (contabilidade estorna o mês anterior).
- Validado contra Julho Flash real: área branca zero diferença; diferenças nas provisões confirmadas como esperadas (arquivo real tinha bug antigo na coluna D; versão do Fast Provisão mais nova no script).
- Quadro de comparação Flash x Actual e Faturamento (linha 25) ficam de fora por enquanto — usuária vai detalhar/automatizar depois.

### 6. Cockpit reestruturado: novo Passo 3 "Provisões"
- Passo 3 (novo) · Provisões: "Lançar Provisões" (cria + preenche pela 1ª vez) e "Atualizar Provisões" (relê o Fast Provisão mais recente, sem duplicar arquivo).
- Antigo Passo 3 (KSB1 Pivot + Finalização) virou Passo 4.
- Trava de segurança: não lê o Fast Provisão se estiver aberto no Excel (arquivo de lock `~$`) — testada ao vivo, funcionou.
- Fluxo completo (Lançar Provisões → Finalização) validado de ponta a ponta contra Julho Flash real.

### Estado final: TUDO EM PRODUÇÃO
- Extração, Check, Provisões (Flash) e Base Intermediária (Actual completo, Flash básico) todos escrevem na rede oficial, wireados no cockpit, testados e validados contra dados reais fechados.
- Nenhum dado real de produção foi alterado incorretamente hoje — todo teste foi feito comparando contra arquivos já fechados, sem sobrescrever nada.

### Pendências pra próxima sessão
1. Quadro de comparação Flash x Actual pro caso do próprio Flash (usuária vai detalhar).
2. Faturamento (linha 25) — ainda manual; usuária mostrou um quadro de Faturamento por Centro de Montagem (Forecast/Flash/Delta) que será automatizado depois.
3. Inserção automática de linha colorida se as provisões um dia excederem a capacidade existente (hoje para com erro claro, não tenta sozinho).
4. Segundo botão original da aba de Base Intermediária ("arquivo colorido") — acabou sendo a "Finalização"/Provisões, já resolvido; não há mais pendência nesse ponto.
