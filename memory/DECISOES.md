# DECISOES — Decisões importantes e seus motivos

---

## 2026-08-10 — Backup automático agendado (não manual)

**Decisão:** o backup do projeto para o GitHub roda sozinho, todo dia às 18h, via Agendador de Tarefas do Windows (task `Backup_Projeto_Claude`), executando `Auto_Backup_GitHub.bat`.

**Motivo:** a usuária escolheu não depender de lembrar de rodar o backup manualmente.

**Como funciona:** `Auto_Backup_GitHub.bat` roda `git add -A` + `git commit` + `git push` na pasta do projeto. Se não houver mudanças, ele não faz commit vazio.

**Para alterar/remover:** `schtasks /change /tn "Backup_Projeto_Claude" ...` ou `schtasks /delete /tn "Backup_Projeto_Claude" /f`.

---

## 2026-08-10 — Nomes de arquivo da ontologia sem espaços

**Decisão:** os arquivos de domínio em `ontology/` usam nomes em snake_case (`fitted_units.json`, `circuito_panamericano.json`) em vez do nome literal do negócio com espaço.

**Motivo:** evitar problemas de path em scripts Python/Excel que forem ler esses arquivos futuramente (regra de qualidade do `CLAUDE.md`: nunca hardcodar caminhos, e nomes com espaço exigem escaping extra).

---

## 2026-08-10 — Detecção automática de sessão longa (limite de 45 ações)

**Decisão:** um hook `PostToolUse` (`.claude/settings.json`) roda `scripts/check_session_length.py` após cada ação do Claude. Ao atingir 45 ações na mesma sessão, ele dispara `scripts/session_transition.py` (arquiva snapshot do `BRIEFING.md` em `memory/long_term/` + `git add/commit/push`) e alerta o Claude para atualizar o `BRIEFING.md` e avisar a usuária que pode abrir uma nova janela.

**Motivo:** sessões muito longas aumentam o risco de o contexto ficar pesado/perdido; o alerta garante que o progresso seja salvo antes de trocar de janela, sem a usuária precisar lembrar de fazer isso manualmente.

**Como funciona:** o contador é por `session_id`, guardado em `.claude/session_state/` (gitignorado, é estado transitório). Zera automaticamente após disparar o alerta.

**Limitação conhecida:** o hook só passa a valer depois que a usuária rodar `/hooks` uma vez ou reiniciar a sessão — o Claude Code só observa arquivos de configuração que já existiam quando a sessão começou, e o `.claude/settings.json` foi criado no meio desta sessão.

---

## 2026-08-10 — Dados reais (SAP/financeiro) nunca vão para o Git

**Decisão:** `data/raw/`, `data/processed/` e qualquer arquivo `.xlsx`/`.csv` foram adicionados ao `.gitignore`. Nada disso é versionado, mesmo o repositório sendo privado.

**Motivo:** o primeiro export real da KSB1 (via `scripts/sap/extrair_ksb1.py`) foi gerado direto em `data/raw/`, que não estava no `.gitignore` até então — risco de vazar dados financeiros reais da Pirelli no primeiro commit. A usuária confirmou: repositório é privado, mas "não pode vazar dados" de jeito nenhum.

**Como aplicar:** sempre checar `git status` antes de `git add`/commit e confirmar que nenhum arquivo de `data/` está sendo staged (regra também adicionada ao `CLAUDE.md` → Segurança de dados).

---

## 2026-08-10 — Cópia dos exports da KSB1 para a área de rede da Pirelli (Fitted Units)

**Decisão:** além de salvar em `data/raw/`, o `scripts/sap/extrair_ksb1.py` também copia o export bruto da KSB1 para `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\<ano>\00.Extração Base KSB1\<MM - Mês>\`, mas só quando a BU é Fitted Units.

**Motivo:** pedido direto da usuária — ela já usa essa área de rede para guardar os resultados e quer o export automático lá também. Circuito Panamericano ainda não tem um caminho de rede definido.

**Como funciona:** pastas de mês (formato `01 - Jan` a `12 - Dec`, abreviação em inglês) já foram criadas para 2026. O caminho é fixo no script (é um recurso corporativo compartilhado, não um caminho pessoal — não fere a regra de não hardcodar caminhos do usuário).

---

## 2026-08-10 — Atalho "ATUALIZAR KSB1.bat" na área de rede

**Decisão:** criado `ATUALIZAR KSB1.bat` direto em `...\00.Extração Base KSB1\` na rede. Ao dar duplo clique, ele entra em `C:\Users\silveju001\Projetos Claude` (caminho fixo, local do computador da Juliana) e abre `scripts\sap\atualizar_ksb1_gui.py` via `pythonw` (sem console preto).

**Motivo:** pedido da usuária, pra não precisar abrir o Claude Code/terminal toda vez que for atualizar a KSB1 — só clicar no atalho a partir da pasta de rede que ela já acessa.

**Limitação:** esse `.bat` só funciona no computador da Juliana (caminho local hardcoded), e o `.bat` em si não é versionado no Git (vive na rede, fora do repositório). Se ela trocar de máquina ou o projeto mudar de pasta, o atalho precisa ser recriado com o novo caminho.

**Atualização 2026-08-10 (mesma sessão):** o `.bat` da rede estava desatualizado — era uma versão antiga que chamava `python` (bloqueante, console preto) direto no `extrair_ksb1.py`, sem passar pela GUI nova. Corrigido para usar `cd /d` fixo + `start "" pythonw scripts\sap\atualizar_ksb1_gui.py` + `exit /b 0`, igual à versão local em `scripts/sap/ATUALIZAR KSB1.bat` (que usa `%~dp0`, só funciona por ser copiada, não porque roda direto no repo). Testado com sucesso pela usuária. Importante: a `atualizar_ksb1_gui.py` fica só local (`C:\Users\silveju001\Projetos Claude\scripts\sap\`) — não precisa copiar o `.py` para a rede, só o `.bat` de disparo.

---

## 2026-08-10 — "DIEGO" ao lado da variante /DESPFITTED: deixar como está

**Contexto:** na tela de seleção da KSB1, o campo "Variante de exibição" mostra `/DESPFITTED` e, ao lado, o texto "DIEGO" — a descrição/título salvo dessa variante ALV no próprio SAP (provavelmente definida por quem a criou). Não é algo que o script `atualizar_ksb1_gui.py` define; ele só digita `/DESPFITTED` no campo, e o SAP exibe a descrição automaticamente.

**Decisão:** não mexer. Perguntei à usuária se ela queria renomear a descrição da variante direto no SAP (o que mudaria a exibição para todo mundo que usa essa mesma variante, não só ela) ou deixar como está; ela escolheu deixar como está.

**Motivo:** é só texto informativo, não afeta o resultado da extração nem aparece nos arquivos gerados; renomear seria uma mudança em objeto compartilhado do SAP sem ganho real.

---

## 2026-08-10 — Nome do sub-projeto: "Fitted Units Despesas"

**Decisão:** todo o trabalho feito nesta sessão (extração da KSB1 via GUI + check de agrupamentos gestoriais) fica identificado como o sub-projeto **"Fitted Units Despesas"**, dentro do domínio maior Fitted Units (que também inclui faturamento, EBIT e P&L, ainda não automatizados).

**Motivo:** pedido explícito da usuária, para deixar claro que essa frente pertence à Fitted Units especificamente na parte de despesas.

**Como foi aplicado:** só a nível de documentação/memória (`memory/PROJECT_MAP.md`) — não houve renomeação de pastas ou arquivos físicos, para não quebrar os caminhos já usados pelos scripts (`REDE_BASE`, caminho local do projeto, etc.).

---

## 2026-08-11 — Automatizar a Base Intermediária direto da KSB1, sem passar pelo BASE_KSB1 acumulado

**Decisão:** o passo 3 do processo (`scripts/sap/gerar_base_intermediaria.py`) soma o `Valor/MR` do extrato bruto do mês (Gestoriais ou Sem Agrupamento, mesma regra do check de agrupamentos) agrupado por `(Centro de Custo, Conta Fiscal)` e cola direto na aba `Intermediária` do arquivo `Base Intermediária Fitted <Mês> <Ciclo> <Ano>.xlsx`. **Não abre nem cresce** o arquivo acumulado `BASE_KSB1` (~45 mil linhas, com Tabelas Dinâmicas nativas `Pivot_Inter.`/`Pivot_Detalhes`).

**Motivo:** a Juliana levantou a preocupação de que o BASE_KSB1 é onde a classificação Fixo/Variável é feita corretamente contra a base de contas, e perguntou se dava pra pular esse arquivo. Confirmado com ela que o BASE_KSB1 só tem 2 usos: alimentar a Base Intermediária e fazer o "check do agrupamento" — e o segundo já é feito hoje direto dos extratos brutos por `check_agrupamentos_ksb1.py`, sem depender do BASE_KSB1. A classificação Fixo/Variável de cada `(Conta Fiscal, Centro de Custo)` já está gravada na própria aba `Intermediária` (coluna H), não precisa ser recalculada. Evita abrir/gravar um arquivo de 45 mil linhas e evita automatizar refresh de Tabela Dinâmica nativa via COM do Excel (frágil e sem necessidade real).

**Trava de qualidade:** o script exige que o "Check de Agrupamentos" do mês já exista (gera se não existir) antes de somar qualquer valor — é essa checagem que garante o vínculo/classificação de cada conta, papel que antes dependia do BASE_KSB1.

**Validação:** rodado contra junho/2026 Actual real (extratos antigos em `Bases SAP/`, fora do fluxo automatizado, usado só como teste de leitura — nada foi sobrescrito). De 440 combinações `(Centro, Conta)`, 409 bateram exatamente com o valor já colado manualmente na coluna June. As 31 restantes: 16 eram linhas que a `Intermediária` tinha zeradas mas deveriam ter valor (a automação teria acertado — não é erro da lógica), 14 eram combinações com **linha duplicada** na `Intermediária` (problema estrutural conhecido da planilha, corretamente reportado como pendência em vez de preenchido às cegas) e 1 batia exatamente com uma reclassificação que a própria usuária já tinha anotado manualmente na planilha ("REVISAR TODOS OS MESES, TROCAR DE 8297 PARA CC FIXO 8295").

**Como funciona:** nunca sobrescreve o arquivo de trabalho — sempre salva uma cópia nova versionada (`... - gerado.xlsx`, mesmo padrão `nome_com_versao`) com os valores preenchidos e uma aba extra `Pendências` (combinações sem linha correspondente ou com chave ambígua). Botão "Atualizar KSB1 Pivot" na GUI (`scripts/sap/atualizar_ksb1_gui.py`), que ganhou também um seletor de Ciclo (Actual/Flash).

**Link externo conhecido como lixo:** o BASE_KSB1 tem um link externo para `RHFitted <Mês> Actual <Ano>_.xlsx` que fica desatualizado (em junho/2026 ainda apontava fevereiro). A usuária confirmou que não usa esse link pra nada — ignorar, não tentar corrigir.

---

## 2026-08-11 (retomada à tarde) — Reversão: voltar a alimentar via BASE_KSB1 + Pivot nativo, em vez do atalho direto do extrato

**Decisão:** revertida a decisão acima. O passo 3 agora vai replicar o processo manual real, e não mais pular o `BASE_KSB1`:
1. Copiar o `KSB1 <mês anterior> Actual <ano>.xlsx` (já acumulado) → nova cópia versionada `KSB1 <mês> <ciclo> <ano>.xlsx`.
2. Colar as linhas novas do extrato bruto do mês (mesma regra Gestoriais vs. Sem Agrupamento) no fim da aba `BASE_KSB1` (colunas A-R, mapeamento 1:1 confirmado) e replicar as fórmulas das colunas S-AI pra essas linhas novas.
3. Dar refresh nas Pivot Tables nativas (`Pivot_Inter.`, `Pivot_Detalhes`) via automação COM do Excel (`win32com`, já em `requirements.txt`).
4. Ler o `Pivot_Inter.` já atualizado (chave Centro custo + Classe de custo, igual à `Intermediária`) e colar o valor do mês na aba `Intermediária` — **esse último passo (5) ainda não foi fechado com a usuária**, ficou pendente porque ela interrompeu pra explicar a lógica das linhas coloridas (ver `memory/BRIEFING.md`).

**Motivo:** pedido explícito da usuária ("vamos voltar atrás") — não foi pra resolver o problema das linhas duplicadas especificamente (ela confirmou isso quando perguntado), e sim porque prefere que a automação **fique fiel ao processo manual atual** (BASE_KSB1 → Pivot_Inter. → Intermediária), mesmo topando com mais trabalho técnico (automação de refresh de Tabela Dinâmica via COM, que a decisão anterior evitava de propósito por ser mais frágil).

**Investigação feita antes de reverter o código:** estrutura real do `BASE_KSB1` e das Pivot Tables inspecionada ao vivo no arquivo `KSB1 June Actual 2026.xlsx` (só leitura) — detalhes completos em `memory/learnings/2026-08-11_estrutura_real_base_ksb1_e_pivot.md`. Achados chave: BASE_KSB1 não é Tabela do Excel; colunas 1-18 batem 1:1 com o extrato bruto; colunas 19-35 são fórmulas replicáveis por padrão de linha; o range de origem das duas pivots já cobre até a última linha do Excel (não precisa reajustar range, só refresh); `Pivot_Inter.` já sai agrupado pela mesma chave da `Intermediária`.

**Ainda não implementado em código** — só a investigação e o plano foram fechados nesta sessão. `scripts/sap/gerar_base_intermediaria.py` (o atalho antigo) continua no repositório, mas será substituído/reescrito pra seguir o novo plano.

**Status:** implementação pausada — usuária encerrou a sessão por cansaço, retoma amanhã. Ver `memory/BRIEFING.md` pra pendências detalhadas, incluindo um tema novo (linhas coloridas na `Intermediária`, só no Flash) que surgiu no meio da explicação do passo 5 e ainda não foi totalmente ensinado.

---

## 2026-08-13 — ZLFIB: exportar pra arquivo em vez de ler grid via COM

**Decisão:** `scripts/sap/analisar_zlfib_duplicidade.py` passou a exportar a grade de resultado da ZLFIB pra um arquivo `.xlsx` temporário via menu nativo do SAP (Lista > Exportar > Planilha eletrônica) e ler esse arquivo com `openpyxl`, em vez de ler célula a célula via `grid.GetCellValue()` (COM/GUI Scripting).

**Motivo:** `GetCellValue` em loop se mostrou não confiável em grades grandes. Comparando duas consultas onde uma deveria ser subconjunto da outra (SJP com filtro de Direção=Entrada vs. sem filtro), a leitura por COM devolveu MENOS notas únicas na consulta filtrada (38) do que na maior (27 — que deveria ser o total, mas ficou menor ainda, inconsistência dupla) — matematicamente impossível se a leitura fosse confiável. Exportando a mesma consulta pra arquivo, o número real de SJP filtrado é 565 notas únicas, não 38. Detalhe completo em `memory/errors/2026-08-13_zlfib_getcellvalue_dados_incorretos.md`.

**Lição geral pra outros scripts SAP GUI Scripting deste projeto:** `GetCellValue` em loop não é confiável pra grades com mais de ~1.000 linhas (provavelmente por virtualização da grid ALV — só as linhas renderizadas na tela ficam com valor correto e acessível via COM). Preferir sempre exportação nativa (arquivo) quando o volume não for pequeno, mesmo que isso exija negociar um popup de "Salvar como" (mesmo padrão já usado em `atualizar_ksb1_gui.py` pra KSB1).

---

## 2026-08-13 — ZLFIB: excluir Tipo NF 'R8' (transferência de material) da análise de duplicidade

**Decisão:** `analisar_zlfib_duplicidade.py` exclui notas com `NFTYPE == 'R8'` antes de procurar duplicidade (constante `NFTYPE_EXCLUIDOS`).

**Motivo:** rodando SOR+GOI (jan-jul/2026), o script encontrou 2 grupos "duplicados" por Chave de Acesso — mas com inspeção, eram pares de notas com a mesma chave de acesso e mesma Nota Fiscal, tipo 'R8', sempre com o parceiro FIAT AUTOMOVEIS S/A (uma montadora cliente, não um fornecedor comum) e valores diferentes entre as duas linhas do par. A usuária confirmou que isso é transferência de material, não duplicidade de pagamento a fornecedor — "pode ignorar estes". Pode ser relacionado ao conceito de "operação A24"/campo "OPERA" que ela mencionou antes (não confirmado 100%, mas o padrão bate: transferência, tipo específico de NF).

**Como aplicar:** se aparecer um novo caso de "duplicidade" envolvendo Tipo NF diferente de R8 mas ainda parecendo transferência entre plantas/cliente, perguntar à usuária antes de assumir que é outro falso positivo do mesmo tipo — a lista `NFTYPE_EXCLUIDOS` só tem R8 confirmado até agora.

---

## 2026-08-13 — Checagem mensal de duplicidade ZLFIB: automação com envio automático de e-mail

**Decisão:** nova automação (`scripts/sap/verificacao_mensal_zlfib.py` + `watcher_mensal_zlfib.bat`) que roda de hora em hora (via Agendador de Tarefas do Windows, ainda não registrado — ver `memory/BRIEFING.md`), checando no primeiro dia útil de cada mês se o SAP já está logado. Assim que achar, roda a checagem de duplicidade ZLFIB do mês anterior (4 filiais, Direção=Entrada, exclui R8) e, se achar duplicidade real, **envia automaticamente** (não como rascunho) um e-mail pra `juliana.silveira@pirelli.com` com o Excel de duplicidade em anexo. Se não achar duplicidade, não notifica. Se passar das 18h no primeiro dia útil sem achar o SAP logado, manda um e-mail de aviso (sem anexo) uma única vez por mês.

**Motivo:** pedido explícito da usuária, com as decisões de envio automático (vs. rascunho) e do aviso de indisponibilidade do SAP confirmadas via pergunta direta antes de implementar — envio de e-mail é ação irreversível (regra do `CLAUDE.md` → Autonomia), então confirmei antes de automatizar o envio recorrente sem supervisão.

**Por que polling de hora em hora, e não um gatilho direto no login do SAP:** não existe um jeito de "escutar" o evento de login do SAP GUI de fora via Scripting — a automação também não pode abrir/logar no SAP sozinha (não tem a senha da usuária). A alternativa viável é checar periodicamente se já existe uma sessão do SAP GUI aberta e logada (heurística validada ao vivo: `session.Info.User` vem preenchido só depois do login). A usuária pediu explicitamente de hora em hora (sugeri 15 em 15 min inicialmente).

**Limitação aceita:** "primeiro dia útil do mês" considera só segunda-sexta, sem calendário de feriados nacionais/municipais — se o dia 1 útil "de calendário" cair num feriado, a rotina roda mesmo assim nesse dia.

**Atualização:** tarefa registrada com sucesso ainda nesta mesma sessão (ver entrada seguinte sobre a reorganização de pastas — o `schtasks` funcionou usando `MSYS2_ARG_CONV_EXCL="*"` antes do comando, pra evitar o mangling de argumentos do Git Bash).

---

## 2026-08-13 — Reorganização de `scripts/sap/` em pastas por sub-projeto

**Decisão:** todos os scripts de Fitted Units (antes soltos numa pasta só, `scripts/sap/`) foram reorganizados em `scripts/sap/fitted_units/`, com uma subpasta por sub-projeto (`fitted_units_despesas/`, `fitted_recuperacao/`) e uma pasta `_shared/` (com `ksb1_core.py` — conexão/navegação SAP compartilhada — e `ferramentas/` — scripts de diagnóstico genéricos: `inspecionar_tela.py`, `test_conexao_sap.py`, `diagnosticar_popup.py`).

**Motivo:** pedido explícito da usuária — ela notou que, conforme mais sub-projetos forem criados (Circuito Panamericano em breve), tudo ficaria misturado numa pasta só sem separação. Pediu especificamente que os projetos de Fitted Units (Despesas e Recuperação) ficassem organizados como sub-projetos dentro de uma pasta "Fitted Units".

**Como foi feito:** `git mv` (preserva histórico) pra cada arquivo; extraído o código compartilhado (`connect_session`, navegação da KSB1, `nome_com_versao`, constantes `BU`/`REDE_BASE`/`MESES_*`) de `atualizar_ksb1_gui.py` pra `_shared/ksb1_core.py`, com `atualizar_ksb1_gui.py` reimportando esses nomes (scripts que já importavam dele na mesma pasta, como `check_agrupamentos_ksb1.py`, continuam funcionando sem mudança). Scripts de outra pasta que precisavam desse código compartilhado (`extrair_ksb1_periodo.py`, `analisar_zlfib_duplicidade.py`) passaram a importar direto de `ksb1_core.py` via `sys.path.insert()` (mesmo padrão já usado em `verificacao_mensal_zlfib.py`).

**Pontos externos corrigidos na mesma leva** (todos que tinham caminho hardcoded pro local antigo dos arquivos):
- Tarefa agendada do Windows `Verificacao_ZLFIB_Duplicidade_Mensal`: `schtasks /change /tr ...` pro novo caminho de `watcher_mensal_zlfib.bat`.
- Atalho `ATUALIZAR KSB1.lnk` na área de rede: regenerado rodando `criar_atalho_ksb1.ps1` de novo (o `.lnk` em si não é versionado, é local na rede — regenerar é a forma certa de "mover" um atalho).
- `atualizar_ksb1_launcher.vbs` e `criar_atalho_ksb1.ps1`: caminhos internos atualizados pro novo local de `atualizar_ksb1_gui.py`.

**Validação:** todos os `.py` movidos/editados passaram em `python -m py_compile` e depois em `import` real (não só sintaxe); o `watcher_mensal_zlfib.bat` foi rodado manualmente do novo local e funcionou (criou `logs/zlfib_mensal.log` vazio, como esperado por não ser o 1º dia útil); o `.lnk` regenerado foi conferido via PowerShell (`TargetPath`/`Arguments` corretos).

**Pendência:** nenhuma automação ficou quebrada, mas ainda falta atualizar o `.bat` legado `ATUALIZAR KSB1.bat` (dentro de `fitted_units_despesas/`, referência relativa `%~dp0` — deveria continuar funcionando sem mudança, mas não foi testado ao vivo) e confirmar se a usuária quer que `Circuito Panamericano`/`Original Equipment` sigam o mesmo padrão de pastas quando esses sub-projetos começarem de verdade.

---

## 2026-08-14 — Regra de processo (vale para todos os projetos/sub-projetos): validar em pasta apartada antes de ir para a rede

**Decisão:** todo script novo (ou nova versão de script existente) que vai gerar/atualizar um arquivo "oficial" na pasta de rede primeiro roda **numa pasta apartada** (local, em `data/processed/...`, fora da rede) até a usuária confirmar que o resultado bate/faz sentido. Só depois desse "OK" o script passa a escrever de fato no caminho de rede oficial, e só nesse momento qualquer GUI/atalho/botão relacionado é atualizado pra apontar pra ele.

**Motivo:** pedido explícito da usuária (2026-08-14, confirmando a validação de julho/2026 Flash do passo 3/BASE_KSB1) — quer sempre esse fluxo de "testar isolado → validar → promover pra produção", não só para este caso específico, mas como padrão para qualquer projeto novo daqui pra frente.

**Como aplicar:**
1. Script novo escreve output em `data/processed/<sub-projeto>/<algo>_teste/`, nunca direto na pasta de rede do processo real.
2. Reportar o resultado da validação pra usuária (bateu/não bateu, com números).
3. Só depois da aprovação dela: (a) trocar o destino do script pra pasta de rede oficial, (b) se houver GUI/atalho ligado ao processo antigo, atualizar pra chamar o script novo.
4. Exemplo em andamento: `gerar_ksb1_mensal.py` (passo 3, Fitted Units Despesas) — hoje escreve em `data/processed/fitted_units_despesas/base_ksb1_teste/`; quando validado com julho/2026 Flash, passa a escrever em `<REDE_BASE>/<ano>/<MM - Mês>/<MM>_<Mês3>_<Ciclo>/` e o botão "Atualizar KSB1 Pivot" da GUI (`atualizar_ksb1_gui.py`) é atualizado pra chamá-lo (hoje esse botão ainda chama o script antigo/revertido `gerar_base_intermediaria.py`).

---

## 2026-08-19 — Passo 3 (BASE_KSB1 + Pivot) validado com julho/2026 Actual: bateu, diferença 100% explicada por regra de negócio conhecida

**Decisão/achado:** o teste às cegas de `gerar_ksb1_mensal.py` gerado em 2026-08-14 (`KSB1 July Flash 2026 - TESTE VALIDAÇÃO.xlsx`) foi comparado — a pedido explícito da usuária, contra o **Actual** de julho/2026 real (`Base Intermediária Fitted July Actual 2026.xlsx`), e não contra o Flash como estava planejado antes. Motivo da correção: os arquivos brutos de extração usados no teste (`00.Extração Base KSB1/07 - Jul/...v4`) foram puxados do SAP em 10-11/08/2026, depois do fechamento Actual (~05/08) — já deveriam conter as contas PIS/COFINS ("PC"), então comparar contra o Flash real (que ainda não tem essas contas integradas) seria inválido por natureza.

**Resultado:** 587 combinações (Conta Fiscal, Centro de Custo) comparadas — **571 bateram exatamente**. As 16 restantes (R$ 112.275,58 de diferença) são 100% explicadas por uma regra de negócio confirmada pela usuária nesta sessão, não são erro da automação: **unidades com status "encerrada" continuam aparecendo na KSB1/BASE_KSB1 (lançamento retroativo/custo residual), mas não são coladas na `Intermediária` — não entram no EBIT.** O motivo: existe uma provisão em "Não Recorrente" pra cobrir custos de encerramento dessas unidades, e o residual é estornado contra essa provisão em vez de ir pro EBIT normal. 14 das 16 diferenças eram do centro de custo 8269 (Sorocaba), 1 do 8292 (Sorocaba) e 1 do 8247 (também Sorocaba, confirmado pela usuária — já estava correto em `centros_de_custo_por_unidade`). Regra completa registrada em `ontology/fitted_units.json` → `regra_unidades_encerradas_no_ebit`.

**Impacto pra próxima etapa (passo 4, ainda não escrito):** quando o script que lê `Pivot_Inter.` e cola o valor do mês nas linhas brancas da `Intermediária` for implementado, ele precisa **excluir qualquer combinação cujo Centro de Custo pertença a uma unidade "encerrada"** antes de colar — senão vai tentar preencher uma linha que não existe mais na planilha manual.

**Ainda pendente:** promover `gerar_ksb1_mensal.py` da pasta de teste local pra pasta de rede oficial (regra do processo confirmada em 2026-08-14 — só faz isso depois da validação, que agora está feita) e trocar o botão da GUI pra chamar esse script. Não foi feito ainda nesta sessão — perguntar à usuária se quer seguir pra isso agora.

---

## 2026-08-21 — Ciclo (Actual/Flash) passa a ser marcado já na extração da KSB1, e o Passo 3 busca por Ciclo em vez de "arquivo mais recente"

**Decisão (opção 1, recomendada, escolhida pela usuária das 3 propostas em 2026-08-19):** o Passo 1 (extração, `atualizar_ksb1_gui.py` → `extrair_um`) grava o Ciclo no nome do arquivo bruto (`nome_arquivo_ksb1` em `ksb1_core.py`, ex: `KSB1 - Fitted Units 07.2026 - Gestoriais - Actual.XLSX`). O Passo 2 (Check, `check_agrupamentos_ksb1.py` → `gerar_check`) e o Passo 3 (`gerar_ksb1_mensal.py` → `decidir_fonte_e_ler_linhas`) agora recebem o Ciclo como parâmetro e buscam pelo arquivo daquele Ciclo específico (`ksb1_core.encontrar_arquivo_ksb1`), em vez do arquivo com data de modificação mais recente na pasta do mês.

**Motivo:** achado da sessão de 2026-08-19 — se a usuária extrai a KSB1 2x no mês (Flash dia 1, Actual dia ~5) e depois regera/reroda o Flash com a extração do Actual já disponível, o script antigo pegaria por engano os dados do Actual (mais recente por `mtime`) com o nome do Flash. Mesmo tipo de inconsistência que motivou a comparação contra Actual (não Flash) na validação de julho.

**Compatibilidade com meses já extraídos (jan-jul/2026, sem Ciclo no nome do arquivo):** `encontrar_arquivo_ksb1` primeiro procura o arquivo com o Ciclo explícito no nome; se não achar, cai para o arquivo mais recente com o prefixo antigo (sem Ciclo) — mas nunca escolhe um arquivo que pertença claramente a OUTRO Ciclo (nome novo, com sufixo `- Flash`/`- Actual` diferente do pedido), pra não repetir o mesmo bug com dados mistos antigo/novo no mesmo mês. Testado com 4 cenários (só arquivo antigo, Flash+Actual novos coexistindo, antigo+novo Actual coexistindo, mês inexistente) — todos corretos.

**Também ajustado:** nome de saída do "Check de agrupamentos" passou a incluir o Ciclo (`Check de agrupamentos - MM.AAAA - Ciclo.xlsx`), já que agora cada Check é específico de um Ciclo, não mais do mês só.

**Não alterado:** `extrair_ksb1.py` (script standalone antigo na raiz de `fitted_units_despesas/`) — confirmado que não é chamado por nenhum `.bat`/`.vbs`/atalho em uso (o fluxo real é todo via `atualizar_ksb1_gui.py`); ficou como código morto, fora do escopo desta mudança.

**Ainda pendente (não fechado nesta sessão):** trocar o botão "Atualizar KSB1 Pivot" da GUI pra chamar `gerar_ksb1_mensal.py` (hoje ainda chama o script antigo `gerar_base_intermediaria.py`) e apontar a pasta de saída do Passo 3 pra rede oficial (hoje ainda escreve em `data/processed/fitted_units_despesas/base_ksb1_teste/`). Isso é o item #2 da lista de pendências antes de "colocar o cockpit em produção" — perguntar à usuária se quer seguir pra isso agora.

---

## 2026-08-21 (continuação) — Extração ao vivo de Jan/Fev/Mar/Jun 2026 (Actual): lacuna de dados preenchida, teste completo Jan-Jul confirmado

**Achado durante o teste ao vivo do Ciclo:** rodando `decidir_fonte_e_ler_linhas` (Passo 3) pro cenário Actual em todos os meses de jan-jul/2026, Jan/Fev/Mar/Jun retornaram erro "arquivo não encontrado" — não por causa da mudança do Ciclo, mas porque a pasta `00.Extração Base KSB1/<mês>` desses 4 meses estava **completamente vazia** (nenhum arquivo `KSB1 - Fitted Units...`, em nenhum padrão de nome, em lugar nenhum da árvore de 2026). Confirmado que os arquivos finais acumulados (`KSB1 <Mês> Actual 2026.xlsx`) existem pra Fev/Jun (Jan só tem versão Flash; Mar nem a pasta Actual existe) — ou seja, esses meses foram fechados normalmente, só que a extração bruta que alimenta o Passo 3 nunca foi salva nessa pasta (o fluxo `00.Extração Base KSB1` só passou a ser usado a partir de abril/2026).

**Decisão da usuária:** extrair agora, ao vivo, via SAP (`extrair_um` de `atualizar_ksb1_gui.py`, mesma função de produção), as bases brutas Gestoriais + Sem Agrupamento de Jan/Fev/Mar/Jun 2026, Ciclo Actual.

**Executado com sucesso** — 8 arquivos novos gerados (todos já com o Ciclo no nome, ex: `KSB1 - Fitted Units 01.2026 - Gestoriais - Actual.XLSX`), salvos em `00.Extração Base KSB1/<mês>/`. Único obstáculo: o popup nativo "Segurança SAPGUI" apareceu uma vez por pasta nova (Jan, Fev, Mar, Jun — primeira vez que cada uma recebe um arquivo via scripting), pedindo confirmação manual da usuária (não pode ser fechado por script, é proteção por design do SAP GUI). Usuária confirmou com "Memorizar minha decisão" em cada uma — igual ao que já tinha acontecido com Abr/Mai/Jul em sessões anteriores (ver 2026-08-13).

**Resultado final do teste (Jan-Jul/2026, Ciclo Actual, `decidir_fonte_e_ler_linhas`):** todos os 7 meses retornam valor, sem erro:
| Mês | Linhas | Soma Valor/MR |
|---|---|---|
| Jan | 7.448 | 5.370.934,68 |
| Fev | 7.994 | 3.855.960,16 |
| Mar | 9.938 | 4.952.336,15 |
| Abr | 8.130 | 4.339.578,33 |
| Mai | 4.811 | 5.493.635,51 |
| Jun | 6.655 | 5.717.965,18 |
| Jul | 6.063 | 6.767.317,49 |

**Limitação do teste:** confirma disponibilidade de dado e ausência de erro/ambiguidade — não é reconciliação de valor contra a Base Intermediária real de cada mês (isso só foi feito formalmente pra Julho, ver validação de 2026-08-19). Jan/Fev/Mar/Jun ainda não tiveram esse nível de conferência.

---

## 2026-08-21 (continuação) — Popup "Segurança SAPGUI": decisão de NÃO editar saprules.xml, manter clique manual mensal

**Decisão:** a usuária optou por continuar clicando "Permitir + Memorizar minha decisão" manualmente no popup "Segurança SAPGUI" a cada mês (1x por pasta nova), em vez de adicionar uma regra curinga em `saprules.xml`.

**Motivo:** a tentativa de editar o arquivo (via Bash e via Edit) foi bloqueada duas vezes pelo classificador de segurança do modo Auto do Claude Code (arquivo fora da pasta do projeto). Dei 3 opções pra usuária (mudar modo de permissão, adicionar regra de permissão, ou editar manualmente ela mesma) e ela preferiu simplesmente manter o clique manual — baixo custo (1x/mês) e zero risco de corromper um arquivo de configuração do SAP com mais de 1 ano de regras acumuladas.

**Estado do arquivo:** `saprules.xml` não foi alterado. Existe um backup (`saprules.xml.backup_2026-08-21`, mesma pasta `AppData\Roaming\SAP\Common\`) que pode ser removido a qualquer momento — não tem mais uso previsto.

**Se o assunto voltar no futuro:** a regra proposta (não aplicada) era um `<rule>` com diretório em curinga `.../Resultados Fitted/2026/00.Extração Base KSB1/*`, mesmo formato/contexto (KSB1, SAPLSLVC_FULLSCREEN/0200) das regras já existentes — texto completo já formulado, ver histórico desta sessão se precisar retomar.

---

## 2026-08-21 (continuação) — Cockpit em produção: botão "Atualizar Pivot KSB1" trocado pra gerar_ksb1_mensal.py + rede oficial; Ano vira seletor, antes do Mês

**Decisão:** promovido o Passo 3 pra produção de fato, fechando a pendência que vinha desde 2026-08-14/2026-08-19.

**Mudanças em `atualizar_ksb1_gui.py`:**
1. Botão da aba ③ renomeado de "Atualizar KSB1 Pivot" pra **"Atualizar Pivot KSB1"** (só o nome do botão — usuária pediu explicitamente não mexer no resto da aba por enquanto, um segundo botão pra "arquivo colorido" fica pra depois, ainda sem escopo definido).
2. `ao_clicar_pivot` trocado: chamava `gerar_base_intermediaria.atualizar_base_intermediaria` (script antigo/revertido, pré-BASE_KSB1); agora chama `gerar_ksb1_mensal.gerar_ksb1_mensal` (o validado com julho/2026 Actual em 2026-08-19), com `pasta_saida` apontando pra pasta de rede oficial do ciclo: `<REDE_BASE>/<ano>/<MM - Mês>/<MM>_<Mês3>_<Ciclo>/` (mesmo padrão que `localizar_ksb1_actual_anterior` já usa pra achar o Actual do mês anterior). Antes escrevia em `data/processed/fitted_units_despesas/base_ksb1_teste/` (pasta local de teste).
3. Painel superior: **Ano agora vem antes do Mês** (pedido explícito da usuária) e virou um Combobox **selecionável** (lista de anos gerada dinamicamente, ano atual -2 a +1, mas ainda editável — não travado em "readonly" como Mês/Ciclo, pra não bloquear um ano fora da lista se precisar). Antes era um campo de texto livre (`ttk.Entry`).

**Validado antes de aplicar:** `py_compile` + import real + checagem de que o rótulo do botão e a ordem Ano/Mês ficaram corretos, mais teste isolado da montagem do caminho de saída (`REDE_BASE/<ano>/<MM - Mês>/<MM>_<Mês3>_<Ciclo>/`) contra 3 casos reais (Jul/2026 Actual, Jan/2026 Flash, Ago/2026 Actual) — todos os 3 caminhos batem com pastas que já existem de verdade na rede.

**Atenção operacional pra próxima vez que a usuária rodar isso de verdade:** a pasta de destino de meses já fechados (ex: julho) já tem o arquivo oficial (`KSB1 July Actual 2026.xlsx`) dentro. `nome_com_versao` NUNCA sobrescreve — se rodar de novo pra um mês/ciclo que já tem arquivo, gera um `..._v2.xlsx` do lado, não substitui o oficial. A usuária precisa saber disso e decidir manualmente se promove o `_v2` a oficial (renomear) ou não.

**Ainda não fechado (fora do escopo desta mudança):** segundo botão da aba ③ pra "atualizar arquivo colorido" — usuária mencionou mas decidiu adiar a definição do escopo exato ("vamos fazer assim, altera o nome do primeiro botão apenas e vamos passar pra produção"). Retomar quando ela quiser detalhar o que esse botão deve fazer.

**Nota:** a janela do cockpit já estava aberta (lançada mais cedo nesta sessão) quando essas mudanças foram feitas no código — como é uma aplicação Tkinter carregada em memória, a usuária precisa **fechar e abrir de novo** a janela pra ver essas mudanças (não há hot-reload).

---

## 2026-08-21 (continuação) — Confirmação pedida pela usuária: Janeiro e Julho batem contra a Base Intermediária REAL (não só "sem erro")

**Contexto:** a usuária notou, corretamente, que o teste anterior (Jan-Jul/2026, Ciclo Actual) só confirmava que os arquivos eram encontrados e retornavam um total — não que esse total batia com o valor real já fechado. Pediu confirmação explícita pra Janeiro e Julho.

**Método:** comparação linha a linha por chave (Conta Fiscal, Centro de Custo) entre o KSB1 bruto (fonte decidida por `decidir_fonte_e_ler_linhas`) e a coluna do mês na aba `Intermediária` do arquivo real já fechado (`Base Intermediária Fitted <Mês> Actual 2026.xlsx`) — mesmo método já usado e validado pra Julho em 2026-08-19.

**Resultado:**
- **Julho:** diferença de R$ 112.275,58 (16 combinações) — **reconfirma exatamente** o resultado de 2026-08-19, 100% em centros de custo da Sorocaba (unidade encerrada).
- **Janeiro:** diferença de R$ 391.655,32 (34 combinações) — **também 100% em centros de custo da Sorocaba** (8231, 8247, 8269, 8292 — lista completa em `ontology/fitted_units.json` → `centros_de_custo_por_unidade.grupos.SOROCABA.centros`). Zero diferença fora da Sorocaba, zero diferença sem explicação.

**Conclusão:** confirma que a mudança do Ciclo não introduziu nenhum problema de valor — os dois meses batem exatamente com o valor real fechado, uma vez descontado o resíduo retroativo já conhecido e documentado (`regra_unidades_encerradas_no_ebit`). Fev/Mar/Jun ainda não passaram por essa mesma checagem linha a linha (só confirmado que retornam valor, sem erro) — fazer se a usuária pedir o mesmo nível de confiança pra esses meses.

---

## 2026-08-21 (continuação) — Bug real achado e corrigido: pastas de Março/Abril usam mês por extenso, quebrava Passo 3 pra Abr/Mai

**Achado (durante a checagem Fev-Jul pedida pela usuária):** as pastas de Ciclo de março e abril na rede usam o **mês por extenso** (`03_March_Actual`, `04_April_Actual`) em vez da abreviação de 3 letras que todos os outros meses de 2026 usam (`03_Mar_Actual`, `04_Apr_Actual`) — inconsistência real na estrutura de pastas, não relacionada à mudança do Ciclo. Isso quebrava duas coisas que constroem esse caminho a partir do padrão: `localizar_ksb1_actual_anterior` (Passo 3 buscando o Actual do mês anterior — travaria rodando Abril ou Maio) e a saída de produção que acabei de ligar no botão da GUI (travaria rodando Março ou Abril).

**Correção:** nova função `resolver_pasta_ciclo(pasta_mes, mes, ciclo)` em `ksb1_core.py` — tenta o nome padrão (abreviação); se não existir, cai para qualquer pasta existente que bata com `<MM>_*_<Ciclo>` (tolera a exceção sem precisar de uma lista hardcoded de meses problemáticos); se nada existir (mês/ciclo novo, pasta nunca criada), devolve o caminho padrão mesmo assim, pra quem chama decidir se cria (saída nova) ou reporta erro (entrada esperada). Aplicada em `localizar_ksb1_actual_anterior` (`gerar_ksb1_mensal.py`) e na construção de `pasta_saida` no botão "Atualizar Pivot KSB1" (`atualizar_ksb1_gui.py`).

**Validado:** meses 1-8/2026 resolvem certo agora (antes, mes=4 e mes=5 davam `FileNotFoundError` em `localizar_ksb1_actual_anterior`).

---

## 2026-08-21 (continuação) — Confirmação completa Fev-Jul/2026 (Actual): todos os meses batem 100% com o valor real fechado

**Pedido da usuária:** estender a mesma checagem linha a linha (feita antes só pra Jan/Jul) para Fev-Jul/2026, Ciclo Actual.

**Resultado (mesmo método: KSB1 bruto agrupado por (Conta Fiscal, Centro de Custo) vs coluna do mês na `Intermediária` real):**

| Mês | Diferença total | Combinações c/ diferença | Explicação |
|---|---|---|---|
| Fev | R$ 286.477,34 | 31 | 100% unidade encerrada |
| Mar | R$ 299.116,79 | 36 | 100% unidade encerrada |
| Abr | R$ 133.298,22 | 19 | 100% unidade encerrada |
| Mai | R$ 116.537,77 | 17 | 100% unidade encerrada |
| Jun | R$ 186.692,70 | 18 (16 unid. encerrada + 2 que se cancelam) | ver abaixo |
| Jul | R$ 112.275,58 | 16 | 100% unidade encerrada (reconfirma 2026-08-19) |

**Único ponto que não é "unidade encerrada" (Junho):** par `N151420000` nos centros `8295` (-R$354,60) e `8297` (+R$354,60), líquido R$ 0,00 — **não é problema novo**, é a mesma inconsistência já documentada em 2026-08-11 (linha 752 da planilha manual: "REVISAR TODOS OS MESES, TROCAR DE 8297 PARA CC FIXO 8295"). Ajuste manual conhecido, efeito líquido zero.

**Conclusão:** Jan-Jul/2026 (Ciclo Actual) batem 100% com o valor real fechado — toda diferença é explicada por regra de negócio já documentada (unidades encerradas + a inconsistência conhecida do centro 8295/8297). Zero diferença sem explicação em qualquer um dos 7 meses. A mudança do Ciclo e a correção da lacuna de dados (Jan/Fev/Mar/Jun extraídos nesta sessão) não introduziram nenhum problema de valor.

---

## 2026-08-21 (continuação) — Confirmado: coluna S ("Nº doc.de referência") do extrato bruto é ignorada de propósito no Passo 3

**Pergunta da usuária:** ao explicar o Passo 3 passo a passo (pedido "explica pro meu estagiário"), ela perguntou se a coluna S do arquivo extraído pelo Passo 1 (`Nº doc.de referência`) deveria ser ignorada ao colar na aba `BASE_KSB1`, já que o arquivo acumulado ("arquivo gigante") termina em `Soc.parc.negócios` (coluna R), sem essa coluna.

**Confirmado (checando ao vivo, não de memória):** sim, já é assim por construção, não é bug nem pendência. `N_COLS_BRUTO = 18` em `gerar_ksb1_mensal.py` limita a leitura/colagem às colunas A-R (18 colunas) — para exatamente em `Soc.parc.negócios`, nunca chega na coluna S. Verificado nos dois arquivos brutos (Gestoriais e Sem Agrupamento) de julho/2026: ambos têm 19 colunas, mesma ordem, coluna S sempre `Nº doc.de referência` por último — consistente nos dois tipos de arquivo.

---

## 2026-08-21 (continuação) — Atalho de rede do cockpit renomeado: "ATUALIZAR KSB1" -> "Fechamento Custo Fitted Units"

**Decisão:** a pedido da usuária, o atalho `.lnk` na pasta `00.Extração Base KSB1` (que abre o cockpit) foi renomeado de "ATUALIZAR KSB1.lnk" pra **"Fechamento Custo Fitted Units.lnk"**. `criar_atalho_ksb1.ps1` atualizado com o novo nome (arquivo + descrição), pra ficar consistente se precisar regenerar o atalho no futuro.

---

## 2026-08-21 (continuação) — Passo 4 (Base Intermediária) implementado e validado; achado e corrigido um bug real de corrupção de dados no COM/Excel

**Contexto:** implementado o Passo 4 (`gerar_base_intermediaria.py`, reescrito do zero — a versão antiga, pré-BASE_KSB1, foi substituída) seguindo a lógica detalhada pela usuária: copia a Base Intermediária Actual do mês anterior, lê o `Pivot_Inter.` do BASE_KSB1 do mês atual (já gerado pelo Passo 3), apaga toda a área de dados a partir da primeira linha sem cor (linhas 2-67 são reservadas pra reclassificações/provisões manuais, só existem de verdade no Flash, nunca tocadas) e cola tudo de novo — "full rebuild" a cada rodada em vez de atualização incremental célula a célula, decisão explícita da usuária "pra evitar erros". Unidades encerradas (Sorocaba, Itatiaia, Camaçari, Santo André, Juiz de Fora — lista confirmada pela usuária) continuam aparecendo como linha (rótulos preenchidos), mas com o valor do mês em branco; o que foi excluído vai pra um arquivo separado por mês (`Histórico Unidades Encerradas - <Mês> <Ciclo> <Ano>.xlsx`, colunas A-AA, só o mês atual — usuária vai enviar por e-mail pra contabilidade). Arrasta fórmula da coluna U (Total Ano) e Y:AJ (Gestorial II até Conta Geral) da primeira linha de dado pra todas as linhas coladas. Só Ciclo Actual implementado — Flash fica pra quando a usuária detalhar a lógica das linhas coloridas manuais.

**Achado crítico durante os testes — bug real de corrupção silenciosa no COM/Excel:** escrever um array grande (testado com 601 linhas × 15 colunas) numa única chamada `Range.Value = <lista de listas>` via pywin32/COM corrompe **aleatoriamente algumas células em erro `#N/A`** — sem relação com o conteúdo (linhas com valores completamente normais viravam erro), sem relação com fórmulas (`HasFormula=False`, era um valor literal de erro). O bug só aparece quando o arquivo é **salvo e reaberto** (célula lida na mesma sessão viva do Excel via COM parecia limpa — só descobri o problema real depois de perceber que minha checagem via COM comparava contra a string `"#N/A"`, mas o COM devolve erro como inteiro negativo tipo `-2146826246`, não como texto — checagem estava mascarando o bug o tempo todo). Testado sistematicamente: 166 erros colando tudo de uma vez, 25 erros em blocos de 50 linhas, **zero erros colando linha por linha**. Corrigido em `gerar_base_intermediaria.py` (Passo 4) e também **preventivamente em `gerar_ksb1_mensal.py`** (Passo 3), que tinha o mesmo padrão de colagem em bloco único (colava até ~6.063 linhas de uma vez em julho) — ambos os scripts agora colam linha por linha e verificam, logo depois de colar, se alguma célula virou erro (aborta sem salvar se achar, com mensagem clara).

**Confirmado que NENHUM arquivo de produção foi afetado por esse bug até agora:** conferido nos 8 meses de 2026 (jan-ago) — nenhum arquivo `KSB1 <Mês> Actual 2026_v2.xlsx` (ou versão posterior) criado pela automação existe em nenhuma pasta de rede real (o único achado, em abril, é de 08-12/05/2026, três meses antes da automação existir — falso alarme, não relacionado). Ou seja, `gerar_ksb1_mensal.py` nunca rodou de fato contra a rede oficial até agora (só em pastas de teste, sempre descartadas depois de validadas) — o bug foi pego antes de causar qualquer dano real.

**Achado secundário (não relacionado ao bug acima, mas também corrigido):** o `KSB1 July Actual 2026.xlsx` (arquivo real, usado como base pros próximos meses) tinha um link externo quebrado pro `RHFitted February Actual 2026_.xlsx` (já confirmado lixo em 2026-08-11). Usuária removeu manualmente pela interface do Excel (Dados → Editar Links → Quebrar Link) — **backup do arquivo original criado antes** (`KSB1 July Actual 2026.backup_2026-08-21.xlsx`, mesma pasta). Esse link quebrado **não era a causa do bug do `#N/A`** (testado e descartado) — é uma limpeza à parte, mas boa de ter feito já que julho serve de base daqui pra frente.

**Validação final (July, Ciclo Actual):** comparação linha a linha (587 combinações Conta Fiscal + Centro de Custo) entre o arquivo gerado pelo Passo 4 e o arquivo real fechado — **zero diferença, soma idêntica (R$ 6.655.041,91 nos dois)**. 60 linhas de unidade encerrada identificadas e corretamente zeradas (mais granulares que os 16 combos da checagem em 2 chaves feita antes — o Pivot_Inter agrupa por 8 dimensões, não só Conta Fiscal + Centro de Custo, então uma combinação "grossa" pode virar várias linhas "finas" no Pivot_Inter).

**Pendente:** Ciclo Flash (lógica das linhas coloridas ainda não detalhada pela usuária); segundo botão da aba ③ da GUI ("arquivo colorido", também ainda sem escopo); wiring do Passo 4 na GUI (`atualizar_ksb1_gui.py` ainda não chama `gerar_base_intermediaria.py` — só foi testado via linha de comando/pasta local até agora).

---

## 2026-08-21 (continuação) — Passo 4 ligado na GUI: segundo botão na aba ③, "Finalização da Base Intermediária"

**Decisão:** o Passo 4 (`gerar_base_intermediaria.py`) ganhou botão na GUI — a pedido explícito da usuária, **não** virou uma aba/passo novo, ficou como um segundo botão dentro da mesma aba ③ "Base Intermediária" (junto do Passo 3), rotulado **"Finalização da Base Intermediária"**.

**Mudanças em `atualizar_ksb1_gui.py`:**
1. Estrutura de dados `PASSOS` trocou `"botao": "..."` (string única) por `"botoes": [...]` (lista) em todas as 3 abas, pra suportar mais de um botão por aba.
2. `fazer_aba` agora cria um botão por item da lista, empilhados verticalmente.
3. `_todos_botoes` (liga/desliga botões durante uma operação) ajustado pra percorrer listas em vez de widgets únicos.
4. Novo handler `ao_clicar_finalizar_intermediaria`: lê Mês/Ano/Ciclo do painel compartilhado, resolve a pasta de rede oficial (`resolver_pasta_ciclo`, mesma tolerância a nome de mês por extenso), chama `atualizar_base_intermediaria`, mostra os dois caminhos gerados (Base Intermediária + arquivo de histórico de unidades encerradas, quando existir) na mensagem de conclusão.

**Ainda não testado ao vivo pela GUI** (só via linha de comando até agora, contra pasta de teste local) — a usuária vai testar agora.

---

## 2026-08-21 (continuação) — Passo 4: refresh da Pivot + quadro "(+) gain" (comparação Flash x Actual) implementados e validados

**Achado 1 — faltava refresh da Pivot Table da aba "Pivot":** a aba "Pivot" da Base Intermediária tem Tabelas Dinâmicas próprias (linhas 4-11, agrupadas por Var./MO-DG&Var), que não se atualizavam sozinhas só porque os dados da `Intermediária` mudaram. Corrigido com `wb.RefreshAll()` logo depois de zerar as unidades encerradas.

**Achado 2 — quadro "(+) gain" (linhas 14-22 da aba Pivot):** compara Despesas/Mão de Obra do próprio ciclo do arquivo (linha 15/16, já é fórmula automática — soma da própria Pivot Table) contra o ciclo oposto (linha 18/19 — hoje um valor fixo, colado à mão todo mês, buscando no arquivo do outro ciclo). Confirmado com a usuária: linha 15/16 da aba Pivot do arquivo **Flash do mesmo mês** é a fonte certa. Implementado em `atualizar_comparacao_flash` — só roda quando o ciclo sendo gerado é Actual (não teria sentido comparar Flash com Flash); não é fatal se o Flash do mês não existir (avisa e segue).

**Validação (Julho):** valores trazidos bateram exatamente com o real (Despesas R$ 4.334.644,06, Mão de Obra R$ 2.426.107,40). Intermediária continua 100% batendo (zero diferença, soma R$ 6.655.041,91).

**Achado à parte, não é bug:** a coluna de Junho (mês anterior) no mesmo quadro ficou diferente entre o arquivo real de Julho e o meu teste (R$ 3.419.475,14 vs R$ 3.470.779,65, herdado do arquivo de Junho original) — usuária confirmou que foi uma **correção manual feita depois**, direto no arquivo de Julho, sem propagar de volta pro arquivo de Junho. Não afeta a lógica (o script nunca toca em meses passados dessa linha, só escreve a coluna do mês atual) — só registrado aqui pra não confundir numa validação futura.

**Passo 4 (Ciclo Actual) considerado completo e validado** — cobre: full rebuild da Intermediária a partir do Pivot_Inter., exclusão de unidades encerradas com arquivo histórico separado, refresh da Pivot, e o quadro de comparação Flash x Actual.

---

## 2026-08-21 (continuação) — Quadro amarelo "Month/Flash/Actual/delta": fórmulas de Custos (H26/I26) corrigidas pra sempre apontar pro mês certo

**Achado:** H26 (Custos, Flash) e I26 (Custos, Actual) do quadro amarelo abaixo do "(+) gain" ficavam **travadas na coluna do último mês que alguém editou a mão** (I26 = `=I11/1000`, sempre coluna I/julho, mesmo processando outro mês) — H26 nem era fórmula, era um valor fixo colado.

**Confirmado com a usuária:** H26 = soma das linhas 18+19 (Flash) da coluna do mês atual, dividido por 1000 (keur). I26 = linha 11 (Grand Total, equivalente a somar 15+16) da coluna do mês atual, dividido por 1000. Faturamento (linha 25) fica de fora — usuária vai automatizar em outro momento.

**Implementado** em `atualizar_comparacao_flash` (mesma função que já traz Despesas/Mão de Obra do Flash): reescreve H26 e I26 como fórmulas com a coluna do mês atual calculada dinamicamente (`_letra_coluna_mes`), em vez de valor fixo/coluna travada.

**Validado (Julho):** H26 e I26 bateram exatamente com o arquivo real (R$ 6.760,75 e R$ 6.655,04, em milhares), agora como fórmula dinâmica em vez de valor/coluna fixos.

**Passo 4 (Ciclo Actual) considerado completo** — cobre: rebuild da Intermediária, exclusão de unidades encerradas + histórico separado, refresh da Pivot, comparação Flash x Actual (linhas 18/19 e quadro Custos H26/I26). Faturamento (linha 25) e Ciclo Flash continuam pendentes, ambos por decisão da usuária de tratar depois.

---

## 2026-08-21 (continuação) — Câmbio (L25) passa a ser puxado do Flash

**Confirmado com a usuária:** o câmbio (célula L25 da aba Pivot, usado no cálculo de "keur"/milhares de euros) só é alterado no arquivo Flash — o Actual deve sempre puxar exatamente o mesmo valor de lá, célula por célula (não depende de coluna de mês, é fixo).

**Implementado** em `atualizar_comparacao_flash` (mesma função que já traz Despesas/Mão de Obra e ajusta H26/I26) — reaproveita a mesma instância já aberta do arquivo Flash.

**Validado (Julho):** L25 = 5,849, igual ao real.

**Nota à parte (não implementada agora):** usuária mostrou um quadro de Faturamento por Centro de Montagem (Forecast/Flash/Delta — São José dos Pinhais, Ibirité, Sorocaba, Goiana) que será automatizado depois, junto com o resto do Faturamento (linha 25 do quadro amarelo) — confirmado que por enquanto continua manual.

---

## 2026-08-21 (continuação) — Ciclo Flash implementado (básico): preenchimento de provisões + rebuild igual ao Actual

**Lógica confirmada com a usuária:** diferente do Actual (linhas coloridas sempre em branco), no Flash as linhas coloridas — geralmente só as amarelas, que são provisões — são preenchidas manualmente todo mês, porque a contabilidade ainda não lançou (e estorna o mês seguinte, então nunca há nada acumulado pra herdar; como sempre partimos do Actual do mês anterior, que já vem com essas linhas em branco, cada mês começa do zero).

**Fonte das provisões:** pasta `Provisões e Reclassificações` dentro do Flash do mês (`\...\Resultados Fitted\2026\<mês>\<mês>_Flash\Provisões e Reclassificações\`), arquivo `Fast Provisão_<Mês>[_vN].xlsx` de **versão mais alta** (se não tiver `_v2` etc., usa o arquivo base). Aba "Ficha de Solicitação", dados a partir da linha 13: coluna C (Conta Contábil) → coluna C da Intermediária; coluna D (Centro de Custo) → coluna E da Intermediária; coluna H (Valor) → coluna do mês que está fechando. Mapeamento sequencial direto: linha 13→linha 2 da Intermediária, linha 14→linha 3, etc.

**Fórmulas arrastadas (colunas A, B, D, F, G):** cada linha nova recebe a mesma fórmula VLOOKUP (buscando na base de contas/centros externa) que já existe na última linha colorida do arquivo copiado (o "molde") — implementado com substituição de referência de linha via regex (não dá pra usar AutoFill normal porque as linhas ficam ACIMA da linha-molde, não abaixo).

**Limitação aceita por enquanto:** se as provisões não couberem nas linhas coloridas já existentes, o script para com erro claro em vez de inserir linhas novas automaticamente (usuária mencionou que às vezes precisa inserir linha inteira via Ctrl+ — não implementado ainda, fica pra quando/se acontecer de verdade).

**Implementado:** `localizar_fast_provisao` (acha a versão mais alta) + `preencher_provisoes_flash` (lê a Ficha de Solicitação, preenche C/E/valor, arrasta fórmulas). Chamado antes do resto do Passo 4 (que segue idêntico ao Actual: rebuild da área branca a partir do Pivot_Inter, unidades encerradas, refresh da Pivot). **O quadro de comparação Flash x Actual (linhas 18/19/H26/I26/L25) fica de fora de propósito quando o ciclo é Flash** — usuária vai detalhar depois como deve funcionar nesse caso.

**Validação (Julho Flash):**
- Área branca (dados do Pivot_Inter): zero diferença, soma idêntica (R$ 5.008.431,48).
- 28 provisões encontradas e preenchidas (linhas 2-29), batendo com a contagem do arquivo real.
- 2 achados nas linhas coloridas, ambos **confirmados pela usuária como esperados, não bugs**: (1) coluna D do arquivo real estava em branco porque a fórmula nunca foi arrastada lá de verdade (**o arquivo real estava incompleto — o script está certo**); (2) pequenas diferenças de centro/valor em 5 linhas porque a usuária foi revisando o Fast Provisão de v4 pra v5 depois que o arquivo real foi montado — o script sempre usa a versão mais alta (v5), então reflete a informação mais atual, não a desatualizada.

**Ciclo Flash (básico) considerado completo.** Pendente: o quadro de comparação Flash x Actual pro caso do próprio Flash (a usuária vai detalhar), e a inserção automática de linhas coloridas quando as provisões excederem a capacidade existente.

---

## 2026-08-21 (continuação) — Cockpit reestruturado: novo Passo 3 "Provisões", antigo Passo 3 vira Passo 4

**Decisão (a pedido da usuária):** o cockpit ganha um passo novo dedicado a Provisões (só Ciclo Flash), inserido ANTES do antigo Passo 3 (que passa a ser Passo 4):
- **Passo 3 (novo) · Provisões** — 2 botões: **"Lançar Provisões"** (cria a Base Intermediária Flash do mês, copiando do Actual do mês anterior, e preenche as linhas coloridas pela primeira vez) e **"Atualizar Provisões"** (relê o Fast Provisão — útil se ela corrigir e salvar uma versão nova depois — e atualiza um arquivo já existente, sem criar cópia nova).
- **Passo 4 (era Passo 3) · Base Intermediária** — mesmos 2 botões de antes ("Atualizar Pivot KSB1", "Finalização da Base Intermediária"), só renomeado.

**Trava de segurança implementada:** antes de ler o Fast Provisão (nos dois botões de Provisões), o script confere se existe um arquivo de lock do Excel (`~$Fast Provisão_...xlsx`, mesmo padrão que já vimos na pasta) — se o arquivo estiver aberto, para com erro claro em vez de ler um estado potencialmente incompleto. Confirmado com a usuária: é o arquivo Fast Provisão que precisa estar fechado, não a Base Intermediária.

**Refatoração no `gerar_base_intermediaria.py`:**
- Novas funções: `arquivo_esta_aberto` (checa o lock), `lancar_provisoes` (cria + preenche), `atualizar_provisoes` (relê + limpa + repreenche num arquivo já existente), `limpar_provisoes` (apaga rótulos + 12 meses das linhas coloridas antes de repreencher, mesma filosofia de full-rebuild já usada na área branca), `localizar_base_intermediaria_flash_existente` (acha o arquivo já criado pelo Passo 3, versão mais recente).
- `atualizar_base_intermediaria` (Finalização): pra Flash, não cria mais cópia nova nem chama `preencher_provisoes_flash` — usa o arquivo que o Passo 3 já criou. Pra Actual, comportamento idêntico a antes.

**Validado de ponta a ponta (Julho Flash):** rodei Lançar Provisões isolado, depois Finalização em cima do mesmo arquivo — bateu exatamente igual ao teste anterior (área branca zero diferença, soma R$ 5.008.431,48; provisões com as mesmas 5 diferenças já explicadas por versão do Fast Provisão). A trava de arquivo aberto também foi testada ao vivo (bloqueou de verdade com o Fast Provisão aberto, e passou depois que a usuária fechou).

**GUI (`atualizar_ksb1_gui.py`):** `PASSOS` reestruturado com o novo passo inserido no índice 2; wiring dos botões ajustado (índices deslocados pro antigo Passo 3). Não testado ao vivo pela interface ainda, só via linha de comando.

---

## 2026-08-21 (continuação) — Cores do cockpit: vermelho removido, botões/abas em amarelo claro

**Decisão:** a pedido da usuária ("tem que ficar gostoso de ver"), o vermelho Pirelli (`#DA291C`, usado nos botões, aba selecionada, seleção de combobox e na linha de trim do cabeçalho) foi substituído por um amarelo claro (`#FFE9A8`) em todo o cockpit. Constante renomeada de `VERMELHO_PIRELLI` pra `AMARELO_CLARO`. Cor de botão desabilitado também ajustada pra combinar (`#ecdfb0`, antes era um marrom-avermelhado escuro).

---

## 2026-08-21 (continuação) — Tamanho padrão da janela do cockpit fixado em 1317x800

**Decisão:** a pedido da usuária, a janela do cockpit sempre abre no tamanho `1317x800` (medido a partir de um print que ela mandou, 1317x835 incluindo a barra de título do Windows — descontei ~35px pra chegar no tamanho da área da janela em si). `minsize` ajustado proporcionalmente pra `1000x650` (antes era `780x560`, do tempo da janela pequena 860x640).

---

## 2026-08-22 — Quadro de comparação (linhas 18/19 da aba Pivot) implementado pro caso Flash: fonte passa a ser o Forecast (R<mês>), não outro Flash

**Contexto:** pendência #1 da lista de fim de sessão de 2026-08-21 — o quadro `atualizar_comparacao_flash` (implementado em 2026-08-21) só rodava quando o ciclo gerado era Actual, comparando contra o Flash do mesmo mês. Não fazia sentido comparar Flash contra Flash, então esse quadro ficava vazio pro Ciclo Flash.

**Decisão/lógica confirmada pela usuária:** pro Ciclo Flash, o cenário de referência (linhas 18/19) passa a ser o **Forecast mais recente**. Na Pirelli, Forecast = "R" + número do mês em que foi feito (REFRESH) — ex: R7 = refresh feito em julho. Não existem R1 (usa Budget/MP do ano anterior), R8 (time de HQ da Itália de férias) e R12 (time trabalhando no budget do ano seguinte) — confirmado que R2-R7 e R9-R11 cobrem o restante do ano.

**Regra de busca (implementada em `localizar_forecast_para_comparacao`, `gerar_base_intermediaria.py`):** procura primeiro o Forecast do próprio mês que está fechando (pasta `<MM>_<Mês3>_Forecast`, arquivo `<MM>_P&L Fitted Units_Forecast_<MêsInglês>_<AA>_.xlsx` — versão com `_` no final, recomendada pela usuária por não ter fórmula e não correr risco de perder a informação original se alguém mexer por engano). Se não existir, cai pro Forecast do mês anterior (ex: fechamento de Agosto usa R7) e **abre um popup de aviso na GUI** dizendo qual Forecast foi usado. Janeiro (R1) não tem fallback dentro do Forecast — usa Budget/MP, fonte ainda não mapeada (fica pendente, quadro não preenchido nesse caso, com aviso).

**Estrutura da fonte, confirmada inspecionando ao vivo o arquivo real de julho/2026 (R7)** — aba "Resumo Resultado Ano", cabeçalho de mês na linha 4 (coluna D=Janeiro...O=Dezembro), valores em '000 BRL com custo negativo:
- Linha 19 = Variable Cost (total), Linha 20 = Labour (dentro do Variable)
- Linha 30 = Fixed Cost (total), Linha 31 = Labour (dentro do Fixed)
- **Linha 38 = Total Costs** — confirmado bater exatamente com Linha19+Linha30 nos 12 meses testados. A usuária pediu que esse check rode em runtime sempre (não só na implementação) e avise se não bater — implementado em `ler_forecast_despesas_mao_de_obra`.
- Lê-se sempre a coluna do **mês que está fechando** (não do mês do próprio R) — ex: fechando Agosto com fallback pro R7, lê a coluna Agosto dentro do arquivo R7 (R7 projeta jul-dez).
- Fórmula: Mão de Obra = -(L20+L31)×1000; Despesas = -(L38)×1000 − Mão de Obra (inverte sinal e converte de '000 BRL negativo pra BRL absoluto positivo, mesmo padrão das linhas 18/19 da Intermediária).

**Validação forte:** os valores calculados pela fórmula acima pro R7/coluna Julho bateram **exatamente** com o que a usuária já tinha colado manualmente no arquivo real `Base Intermediária Fitted July Flash 2026.xlsx` (Despesas R$ 3.940.062,77 / Mão de Obra R$ 2.380.392,91) — confirma a lógica sem dúvida. Rodado também de ponta a ponta (Lançar Provisões → Finalização) contra pasta de teste local, mesmo resultado.

**Achado extra (confirmado com a usuária antes de implementar):** o arquivo Flash é copiado do Actual do mês anterior, que carrega rótulos de texto errados (herdados do Actual). Corrigido em `atualizar_comparacao_forecast`: linha 15/coluna A "Actual"→"Flash" (próprio ciclo do arquivo), linha 18/coluna A "Flash"→"Forecast" (cenário de referência), cabeçalho do quadro amarelo (linha 24, colunas H/I) "Flash"/"Actual"→"Forecast"/"Flash". Confirmado batendo com o arquivo real de julho.

**Reaproveitado sem mudança:** a escrita das fórmulas de Custos H26/I26 (que sempre apontam pra coluna do mês atual) já era genérica — passou a rodar também pro caso Flash sem precisar de nenhuma alteração.

**Ainda pendente:** fonte de Budget/MP pra Janeiro (sem R1) — retomar quando ficar mais perto do fechamento de Janeiro. Faturamento (linha 25) continua manual, fora do escopo desta mudança.

---

## 2026-08-22 (continuação) — Cockpit: operações rodam em thread separada + indicadores visuais (pneuzinho girando, cursor)

**Contexto:** testando o quadro de comparação Forecast pela GUI de verdade (primeira vez que a GUI era testada ao vivo, não só via linha de comando), a usuária notou que a janela "travava" durante operações longas (Excel via COM) — era o comportamento já conhecido/nunca resolvido (arquitetura 100% síncrona numa thread só).

**Decisão:** as 6 operações da GUI (`atualizar_ksb1_gui.py`) passam a rodar numa `threading.Thread` separada, via um helper novo `rodar_em_thread(descricao, func, ao_concluir)` — mantém a janela respondendo. Requereu refatorar `rodar()` (Passo 1/Extração) pra não chamar `messagebox` direto de dentro da função (só a thread principal do Tk pode) — nova exceção `ErroComTitulo` preserva o título específico de cada erro. Log agora é thread-safe via `queue.Queue` (fila) + drenagem em loop (`root.after`).

**Indicadores visuais pedidos pela usuária, nesta ordem de refinamento:**
1. Barra de progresso "na parte de baixo" → depois "não precisa de %" (indeterminada tá bom) → "um ícone rodando" (spinner) → cursor "ocupado" durante processamento → cursor de mãozinha em qualquer botão → **pneuzinho Pirelli com calota** (não um bloco/spinner genérico) → barra maior (32px) pra ficar mais visível.
2. **Bug de posicionamento encontrado e corrigido durante o próprio teste:** a primeira versão (status/barra dentro do corpo, embaixo do log) não aparecia pra usuária — o log tem `expand=True` e a janela (800px) não sobrava espaço quando a descrição da aba era longa, empurrando o indicador pra fora da área visível. Corrigido movendo pro cabeçalho (área fixa, nunca encolhe): texto de status no canto superior direito do cabeçalho escuro, barra de progresso logo abaixo do trim amarelo.
3. **Pneuzinho final:** `Canvas` (não `ttk.Progressbar`) com um ícone de pneu desenhado via PIL em runtime (`_gerar_frames_pneu` em `atualizar_ksb1_gui.py`) — banda preta com sulcos, calota prateada com 6 raios, miolo escuro com detalhe amarelo claro (cor do cockpit) — 12 frames pré-rotacionados, anima girando e deslizando de um lado a outro da barra (`root.after(40, ...)`). Desenhado em runtime (não arquivo/base64 separado) pra manter o script autossuficiente, mesma filosofia do logo Pirelli já embutido.

**Nova dependência:** Pillow adicionado a `requirements.txt`.

**Validação:** testado ao vivo pela GUI de verdade (não só linha de comando, primeira vez), contra pasta de teste local (REDE_BASE do módulo da GUI trocado via script auxiliar fora do repo, sem tocar rede) — janela permaneceu responsiva durante "Atualizar Pivot KSB1" real (Julho/Flash), indicadores visíveis e aprovados pela usuária ("ficou maravilhoso").

---

## 2026-08-22 (continuação) — Fallback de Agosto (R8→R7) confirmado end-to-end pela GUI real

**Contexto:** último item pendente de teste do quadro de comparação Forecast (ver decisão "Quadro de comparação... fonte vira Forecast" acima) — o fallback R8→R7 só tinha sido testado via função isolada, não pela GUI de ponta a ponta.

**Problema:** Agosto/2026 ainda não tem nenhum dado real na rede (pasta `2026/08 - Aug` totalmente vazia — nem KSB1 extraído nem Fast Provisão) porque o fechamento real de agosto ainda não aconteceu. Pra testar o mecanismo sem esperar o mês fechar de verdade nem escrever nada na rede, foram "emprestados" o KSB1 e a Base Intermediária REAIS de julho (só leitura da rede), copiados pra pasta de teste local e rotulados como se fossem de agosto (`gerar_base_intermediaria.localizar_base_ksb1_do_mes` interceptado por monkeypatch só pra esse caso específico `(8, 2026, "Flash")`). **A busca do Forecast (R8/R7) em si não foi mockada** — continuou 100% real, direto na rede, que é exatamente o mecanismo sendo validado.

**Resultado:** rodando "Finalização da Base Intermediária" pela GUI real com Mês=August/Ciclo=Flash, o popup apareceu corretamente: *"Fechamento de August/2026: não existe Forecast R8 — usei o Forecast de July (R7) como comparação."* Confirma que a busca (`localizar_forecast_para_comparacao`), o fallback e a exibição do aviso (`ao_clicar_finalizar_intermediaria` → `messagebox.showwarning`) funcionam corretamente de ponta a ponta pela interface real, não só via chamada direta de função.

**Nota:** os valores em R$ gerados nesse teste (Despesas/Mão de Obra) não têm significado real — são de julho, só com rótulo de agosto. Arquivos de teste (`_fontes_fake_agosto/`, Base Intermediária "de agosto" em `gui_teste_rede/`) removidos depois da validação — nada disso é dado oficial nem foi escrito na rede.

**Item de teste concluído** — o quadro de comparação Forecast (Actual e Flash) está validado em todas as frentes: lógica de valores (bate com o real de julho), rótulos de texto, GUI ao vivo, e agora o fallback R8→R7.

---

## 2026-08-22 (continuação) — Janeiro/Budget-MP implementado e validado: fecha o último item pendente do quadro de comparação Forecast

**Contexto:** último item pendente desde 2026-08-21 — Janeiro (mes=1) não tem Forecast R1 (a Pirelli não produz), a usuária confirmou que usa o Budget/MP do ano.

**Confirmado com a usuária:**
1. **Propósito:** é a mesma fonte de comparação do quadro (linhas 18/19), só que pra Janeiro em vez de usar Forecast.
2. **Budget e MP são a mesma coisa** (nomes alternativos).
3. **Caminho:** `\\FSS024-01BR.group.pirelli.com\GFU_DAC\Management Plan\MP <ano>\P&L Fitted Units_Budget<AA>_.xlsx` (AA = 2 últimos dígitos do ano) — área de rede totalmente separada da dos outros ciclos.
4. **Ano da pasta = MESMO ano que está sendo fechado, não o anterior** — corrige uma suposição errada que eu tinha registrado antes ("Budget/MP do ano anterior"). A usuária confirmou explicitamente: "nós construímos o MP27, no final de 2026, mas é referente 2027" — ou seja, o MP é PREPARADO no fim do ano anterior mas ARQUIVADO sob o ano que ele cobre.
5. **Estrutura idêntica ao Forecast** — confirmado inspecionando ao vivo o arquivo real de 2026: mesma aba "Resumo Resultado Ano", mesmas linhas (19/20/30/31/38), mesmo cabeçalho de mês. Único diferencial: linha 5 (rótulo do cenário) diz "Budget" em todas as 12 colunas (o arquivo inteiro é projeção, nenhum mês é Actual quando foi criado, ao contrário do Forecast que vai "ganhando" meses Actual conforme o ano avança).

**Implementado:** `localizar_arquivo_budget(ano)` (novo) + `localizar_forecast_para_comparacao` reescrita pra tratar mes=1 como caso especial (vai direto pro Budget, sem tentar Forecast/fallback) — devolve `(caminho, aviso)`, simplificado de 3 pra 2 elementos já que a coluna a ler é sempre `mes` independente da fonte. `atualizar_comparacao_forecast` ajustada: rótulo do quadro vira "Budget" (em vez de "Forecast") quando mes=1.

**Validação forte (GUI real, sem tocar rede):** Janeiro/2026 já tinha fechamento real (mês passado) — copiada a Base Intermediária Flash REAL de Janeiro pra pasta de teste local (única mudança: destino de escrita), sem precisar emprestar dado de outro mês como no teste de Agosto. Rodando "Finalização" pela GUI real: popup apareceu certinho ("Fechamento de January/2026: não existe Forecast R1 — usei o Budget/MP 2026 como comparação"), rótulo "Budget" escrito corretamente na célula, e o valor calculado (Despesas R$ 2.983.872,25) bateu **exatamente** com o que já estava na célula do arquivo original antes do teste (rotulada "MP'26" — a usuária já tinha feito essa mesma conta à mão, usando a mesma fonte, quando Janeiro fechou de verdade).

**Achado no meio do teste (não é bug, só curiosidade registrada):** a pasta `01_Jan_Forecast` na rede tem arquivos com "R1" no nome (`Base_Book_Fitted Units_Forecast R1'26.xlsx`, etc.) — mas NENHUM segue o padrão `P&L Fitted Units_Forecast_...` que a automação procura, então não interfere em nada. Parece ser resíduo de outro processo/relatório, não o Forecast formal que estamos automatizando.

**Falso alarme no meio da sessão:** um teste inicial "não funcionou" (nenhum popup, arquivo de teste intocado) — investigado e confirmado que foi a usuária esquecer de trocar o Ciclo/Mês da rodada anterior (estava em Julho/Actual) antes de clicar Finalização, não um bug. Repetindo com Janeiro/Flash selecionado corretamente, funcionou de primeira.

**Todos os itens do quadro de comparação Forecast agora estão fechados:** Actual (Flash como referência), Flash normal (Forecast R<mês>), fallback R8/R12→mês anterior, e Janeiro (Budget/MP). Só Faturamento (linha 25) e inserção automática de linha colorida continuam pendentes, sem relação com este quadro.

---

## 2026-08-22 (continuação) — Bug real corrigido + inserção automática de linha amarela quando provisões excedem capacidade

**Contexto:** último item pendente da lista — hoje o script para com erro claro quando as provisões do mês não cabem nas linhas amarelas já existentes, pedindo inserção manual (Ctrl+).

**Bug real encontrado ANTES de implementar a inserção (não relacionado ao pedido original, achado investigando as cores):** o cálculo de "capacidade" e a limpeza de conteúdo (`limpar_provisoes`, usada por "Atualizar Provisões") usavam `encontrar_primeira_linha_sem_cor` — que encontra o fim de TODA a área colorida (amarelo+verde+roxo), não só do amarelo. Duas consequências:
1. **Capacidade contava errado:** em julho, a "capacidade" calculada era 66 linhas (2-67, incluindo verde+roxo), quando a capacidade REAL de provisões (só amarelo) é 46 (2-47).
2. **`limpar_provisoes` apagava verde+roxo também** — incluindo a fórmula "molde" da última linha roxa, que era apagada ANTES do código tentar copiá-la pra novas linhas (bug duplo: a cópia sempre pegaria uma fórmula vazia).

**Por que não causou dano real ainda:** em todos os testes até agora (julho/2026), as linhas verde/roxo estavam vazias (nenhuma reclassificação de verdade esse ano) — então "apagar" conteúdo vazio não perdeu nada visível. Mas o bug era real e ia aparecer no primeiro mês com reclassificação de verdade, ou se as provisões um dia passassem de 46.

**Correção:** nova função `encontrar_ultima_linha_amarela(ws)` — detecta o fim do bloco amarelo via `Interior.Color` (não `Pattern`, que é igual pras 3 cores, nem `ColorIndex`, que não discrimina bem cores customizadas — confirmado inspecionando ao vivo que verde e roxo têm o mesmo ColorIndex, só o Color numérico difere: amarelo=65535, verde=10213316, roxo=14336204 no arquivo real de julho). `limpar_provisoes` e o cálculo de `capacidade` em `preencher_provisoes_flash` passaram a usar essa função em vez de `encontrar_primeira_linha_sem_cor`. A fórmula "molde" da roxa (`ultima_linha_colorida`, ainda calculada do jeito antigo — é o comportamento CORRETO, roxa é o molde de propósito) agora é capturada ANTES de qualquer limpeza/inserção, nunca depois.

**Inserção automática implementada:** `inserir_linhas_amarelas_novas(ws, n_linhas, ultima_linha_amarela, log)` — insere `n_linhas` linhas via `Rows(...).Insert(Shift=xlShiftDown, CopyOrigin=xlFormatFromLeftOrAbove)` (nativo do Excel, não copia/cola por área de transferência), sempre logo antes da primeira linha verde, copiando o formato da última amarela. **Nunca mexe no conteúdo das linhas verdes/roxas** — só a posição delas muda (efeito colateral inevitável de inserir linha acima, mas o Excel ajusta sozinho).

**Confirmado com a usuária antes de implementar:**
1. Linha nova é sempre amarela (nunca verde/roxa) — confirmado.
2. Frequência imprevisível ("pode acontecer todos os meses... não tem como prever, mas pode acontecer") — confirma que vale a pena automatizar, não é caso raro descartável.
3. Inserção sempre antes das linhas verdes, empurrando verde+roxo+branco pra baixo — confirmado, bate com a implementação.
4. Roxa mantém só 1 linha, de propósito, com as fórmulas de molde — confirmado, não é bug a ausência de dado real na roxa.

**Validação forte (cópia local do arquivo real de julho, sem tocar rede):**
- **Teste de regressão** (sem forçar inserção, 28 provisões reais dentro da capacidade de 46): `limpar_provisoes` + `preencher_provisoes_flash` rodados — verde/roxo permaneceram exatamente como antes (mesma cor, mesmo conteúdo vazio, fórmula molde da roxa intacta).
- **Teste forçando inserção** (3 linhas amarelas novas): antes `ultima_amarela=47, ultima_colorida=67`; depois `ultima_amarela=50, ultima_colorida=70` (deslocamento de exatamente 3, correto). Verde (era linha 48) passou pra 51, ainda verde, ainda vazia. Roxa (era 67) passou pra 70, ainda roxa — e a fórmula molde se AUTO-AJUSTOU sozinha (`=VLOOKUP(C67,...)` virou `=VLOOKUP(C70,...)`, o próprio Excel corrigiu a referência interna, comportamento nativo de inserção de linha de verdade — não aconteceria assim numa cópia manual de célula, é uma confirmação a mais de que o `Rows().Insert()` nativo foi a escolha certa em vez de manipular linhas na mão). Área branca também deslocou corretamente (era linha 68, passou pra 71).

**Pendente, tema à parte (a usuária vai detalhar depois):** como as linhas VERDES (reclassificações) devem ser preenchidas/atualizadas automaticamente — hoje a automação nunca escreve nelas, só desloca a posição quando insere amarela nova acima.

---

## 2026-08-22 (continuação) — Linhas verdes (reclassificações) deprecadas — decisão da usuária

**Decisão:** o bloco verde (reclassificações) na aba Intermediária nunca mais será usado. Motivo confirmado pela usuária: reclassificações já são feitas diretamente no SAP antes do fechamento ("já vai estar tudo dentro do SAP"), então o mecanismo de linha verde na Excel ficou obsoleto — ela já reclassifica antes do fechamento, não durante.

**Decisão de implementação:** NÃO automatizar a remoção física das linhas verdes do template. Motivo: a automação já ignora 100% esse bloco (nunca escreve nem apaga nada nele, só desloca a posição quando insere linha amarela nova acima) — apagar as linhas de vez seria uma operação estrutural (mesmo tipo de risco/teste da inserção de linha) só por um ganho cosmético (espaço vazio no meio da planilha). Proposto à usuária como alternativa: se ela quiser, apaga manualmente uma vez no arquivo-modelo atual (o Actual do mês corrente, que toda Flash futura copia) — depois disso os próximos meses já nascem sem a área verde, sem precisar de automação nenhuma. Ela aceitou deixar como está (ignorado).

**Roxa continua sendo mantida** — confirmado que é o "molde" de fórmula usado pela inserção automática de linha amarela (ver decisão anterior), não é a mesma situação da verde.

---

## 2026-08-22 (continuação) — Revisão completa do processo de fechamento (pedida pela usuária): 1 bug real corrigido, tudo testado de ponta a ponta

**Pedido da usuária:** revisar todo o processo (Passo 1-4, Actual e Flash), achar fragilidades/erros, testar, "preciso que isso esteja funcionando no fechamento."

**Bug real encontrado (revisão de código):** 4 funções de leitura buscavam arquivo pelo **nome exato**, sem considerar versões `_v2`/`_v3` geradas por `nome_com_versao` — inconsistente com o resto do sistema, que sempre versiona na escrita mas nem sempre considera versão na leitura:
1. `localizar_ksb1_actual_anterior` (gerar_ksb1_mensal.py) — lê o KSB1 Actual do mês anterior.
2. `localizar_base_ksb1_do_mes` (gerar_base_intermediaria.py) — lê o BASE_KSB1 do próprio mês/ciclo (usado pela Finalização).
3. `localizar_base_intermediaria_mes_anterior` (gerar_base_intermediaria.py) — lê a Base Intermediária Actual do mês anterior.
4. `localizar_base_intermediaria_flash_do_mes` (gerar_base_intermediaria.py) — lê a Base Intermediária Flash do mesmo mês (fonte de comparação do Actual).

**Risco real:** se qualquer um desses arquivos fosse regenerado por qualquer motivo (ex: corrigir um erro, rodar de novo por engano), o resultado ficava "invisível" pros passos seguintes — que continuavam lendo a versão antiga (`_v1`) em silêncio, sem erro nem aviso, porque o nome exato dela ainda existia. Isso poderia levar a um fechamento usando dado desatualizado sem ninguém perceber.

**Correção:** nova função `encontrar_arquivo_mais_recente(pasta, nome_base)` em `ksb1_core.py` (glob pelo padrão `nome_v*.xlsx` + pega o de `mtime` mais recente — mesmo método já usado com sucesso em `localizar_base_intermediaria_flash_existente`), aplicada nas 4 funções acima. Testada isoladamente (pasta vazia → None; só v1 → v1; v1+v2 → escolhe v2). Os arquivos de Forecast/Budget/Fast Provisão (que vêm de OUTROS processos, não da nossa automação, e não seguem esse padrão de versão) continuam com busca por nome exato/versão-com-`_`, que é o comportamento certo pra eles.

**Testado de ponta a ponta (Julho/2026, Actual E Flash, sem tocar rede — pasta de teste local, arquivos de origem lidos direto da rede real):**
- Passo 4a (Atualizar Pivot KSB1) Actual: 6.063 linhas coladas, sem erro.
- Passo 4b (Finalização) Actual: Intermediária com soma R$ 6.655.041,91 (bate exatamente com o valor real já validado antes), quadro de comparação (Flash) trazendo Despesas R$ 4.334.644,06 / Mão de Obra R$ 2.426.107,40 (idem).
- Passo 4a (Atualizar Pivot KSB1) Flash: 6.063 linhas coladas, sem erro.
- Passo 3 (Lançar Provisões): 28 provisões, linhas 2-29, sem erro.
- Passo 3 (Atualizar Provisões, testando a correção de cor de novo): log confirma "Linhas amarelas (2-47) limpas... verdes/roxas não são tocadas" — a correção do bug de 2026-08-22 anterior está ativa e funcionando.
- Passo 4b (Finalização) Flash: Intermediária com soma R$ 6.760.751,46 (bate exatamente), quadro de comparação (Forecast R7) trazendo Despesas R$ 3.940.062,77 / Mão de Obra R$ 2.380.392,91 (idem).

**Conclusão da revisão:** nenhum outro bug de correção encontrado na parte que dá pra testar sem SAP ao vivo (Passos 3 e 4, Actual e Flash). O Passo 1 (extração via SAP GUI Scripting) não pôde ser testado de ponta a ponta nesta revisão por depender de sessão SAP ao vivo — revisão de código não achou problema óbvio ali, mas não é o mesmo nível de confiança dos passos testados de verdade.

**Risco conhecido, não resolvido nesta revisão (já documentado antes, 2026-08-14):** se o Excel travar/hangar de verdade durante uma automação (não um erro, um hang de verdade), não existe timeout/vigia — o processo ficaria preso indefinidamente. Baixa probabilidade, mas não é zero. Caso queira, dá pra endereçar depois com um watchdog externo.

---

## 2026-08-24 — Watchdog de travamento (Excel isolado, 12 min): implementado, testado ao vivo e confirmado

**Decisão (design confirmado com a usuária antes de implementar, via pergunta direta):** se uma operação que usa Excel (Passos 3/4) ficar mais de **12 minutos** sem terminar, a GUI avisa e oferece **forçar o encerramento só da instância isolada do Excel** dessa operação (processo próprio, identificado por PID, não afeta outro Excel que a usuária tenha aberto). No Passo 1 (SAP) e Passo 2 (Check, sem COM), o aviso é só informativo, sem oferecer matar nada — decisão explícita de não automatizar o encerramento do SAP GUI, porque isso fecharia TODAS as sessões abertas dele, não só a desta automação.

**Como funciona:**
- `ksb1_core.abrir_excel_isolado(log, pid_callback)` centraliza a abertura do Excel isolado (`DispatchEx`, `Visible=False`, já existia repetida em 4 lugares) e captura o PID do processo via `win32process.GetWindowThreadProcessId(excel.Hwnd)` — funciona mesmo com a janela invisível.
- Os 4 pontos que abrem Excel (`gerar_ksb1_mensal.py` e os 3 em `gerar_base_intermediaria.py`) ganharam parâmetro opcional `pid_callback`, repassado até essa função.
- Na GUI, `rodar_em_thread` mede o tempo decorrido (`time.monotonic()`); passado o limite, mostra o aviso (dedup por intervalo de 12 min, não fica repetindo). Se a usuária confirmar "forçar", mata o processo via `psutil.Process(pid).terminate()` — só depois de checar que o processo ainda se chama `EXCEL.EXE` (proteção contra PID reciclado) — e libera a janela na hora.

**Bug real encontrado e corrigido durante o próprio teste desta implementação:** a primeira versão usava `messagebox.askyesno` pro aviso Sim/Não — o Tk usa botões fixos em **inglês** ("Yes"/"No") nesse widget, mesmo com todo o resto do texto em português, quebrando a regra do projeto (REGRAS_RAPIDAS #11, pedido explícito da usuária 2026-08-10: tudo em português). Corrigido com um diálogo customizado (`_perguntar_sim_nao`, `tk.Toplevel` com botões "Sim"/"Não" de verdade), no mesmo estilo visual do resto da GUI.

**Testado ao vivo, ponta a ponta, sem tocar rede/dados reais:** script de teste isolado (fora do repo) substituiu `gerar_ksb1_mensal.gerar_ksb1_mensal` por uma versão falsa que abre uma instância Excel isolada REAL (mesma chamada de produção) e trava de propósito, com `TIMEOUT_AVISO_SEGUNDOS` reduzido pra 15s só no teste. Automação de clique (via travessia da árvore de widgets do Tk) clicou "Atualizar Pivot KSB1" e depois "Sim" no aviso, sem precisar de interação manual. Resultado confirmado por monitoramento direto de processos (`tasklist`, antes/durante/depois):
- Um segundo `EXCEL.EXE` (instância isolada do teste) apareceu e, ~15s depois, foi encerrado sozinho.
- O Excel real da usuária, que estava aberto no fundo o tempo todo (arquivo de trabalho dela), **nunca foi tocado** (mesmo PID do início ao fim).
- A janela do cockpit voltou ao estado normal depois do "Sim" (botões e cursor recuperados), sem erro, sem travar.
- O aviso disparou no tempo esperado (~15s) e o diálogo em português foi encontrado e clicado normalmente pela automação, confirmando que o fix do idioma renderiza certo.

---

## 2026-08-25 — Popup "Segurança SAPGUI" no Passo 1: eliminado via pasta de staging fixa (ideia da usuária), não via edição de saprules.xml

**Contexto:** o popup nativo "Segurança SAPGUI" aparecia 1x por mês (toda vez que a extração passou a escrever numa subpasta de Ciclo nova, ver mudança de 2026-08-24). Primeira tentativa foi generalizar a regra em `saprules.xml` (arquivo de segurança do SAP GUI, fora do projeto) com um curinga — **bloqueada duas vezes pelo classificador de segurança do Auto mode do Claude Code** (edição de arquivo fora da pasta do projeto), mesmo com autorização explícita da usuária na conversa e SAP fechado. Anteriormente (2026-08-21) essa mesma tentativa já tinha sido bloqueada, e na época a usuária optou por manter clique manual.

**Decisão (ideia da própria usuária, melhor que a tentativa de editar saprules.xml):** em vez de mudar a permissão do SAP, mudar o **destino da exportação nativa do SAP** pra uma pasta fixa que nunca muda de mês/Ciclo/ano: `REDE_BASE/<ano>/00.Extração Base KSB1/Temporario/`. O SAP só pede autorização de escrita quando grava numa pasta ainda não conhecida em `saprules.xml` — como essa pasta passa a ser sempre a mesma, só pede 1x (por ano, na prática quase nunca).

**Implementado em `extrair_um` (`atualizar_ksb1_gui.py`):**
- SAP exporta sempre pra `pasta_staging` (a pasta `Temporario` fixa), não mais direto pra `pasta_rede` (a subpasta real do Ciclo).
- Depois que o arquivo aparece na staging, `shutil.move()` (operação de arquivo comum, **não passa pelo SAP GUI Scripting**, por isso não é vigiada pela segurança do SAP) realoca pro destino real (`pasta_rede`), com o nome já versionado (`nome_com_versao`, calculado contra o destino real, não a staging).
- Proteção contra sobra: se já existir um arquivo com o mesmo nome na staging (de uma tentativa anterior que falhou antes de mover), apaga antes de exportar de novo — evita o SAP perguntar "sobrescrever?" (diálogo que o script não trata).

**Testado isolado (pasta temporária local, não é rede nem repo):** 4 cenários — mover Gestoriais, mover Sem Agrupamento, sobra de staging de tentativa anterior sendo limpa antes do novo export, staging ficando vazia ao final, e versionamento (`_v2`) funcionando contra o destino real. Todos passaram.

**Pasta `Temporario` já existe na rede (confirmada vazia)** — a usuária indicou o caminho exato, script não precisou criá-la.

**Problema real encontrado testando contra o SAP de verdade (não previsto):** o SAP abre o arquivo recém-exportado automaticamente no Excel (visível, não é instância isolada) — isso trava o arquivo (`WinError 32`) na hora do `shutil.move`. Resolvido em rodadas sucessivas, todas em `ksb1_core.py`:

1. **Retentativa no move** (30x, 1s cada) em `extrair_um`, em vez de derrubar a extração no primeiro `OSError`.
2. **`fechar_excel_se_aberto(caminho_arquivo, log)`** — fecha só a aba específica que está travando o arquivo, via COM. Primeira versão usava `win32com.client.GetObject(Class='Excel.Application')`, que só enxerga UMA instância (ambíguo se a usuária tiver o Excel dela mesma aberto ao mesmo tempo, e falha se essa instância estiver ocupada com um popup modal). **Reescrita pra buscar direto na Running Object Table (via `pythoncom`)** pelo caminho completo do arquivo — cada pasta de trabalho aberta se registra na ROT com o próprio caminho como nome, então dá pra achar a certa direto, sem precisar adivinhar qual `Application` pegar. Testado isolado com 2 instâncias do Excel simultâneas (uma simulando "a planilha real da usuária", outra "a que o SAP abre") — fecha só a certa, confirmado que a da usuária nunca é tocada.
3. **Limpeza da sobra na pasta `Temporario`** (antes de exportar de novo) também virou resiliente a arquivo travado (retentativa 30x com `fechar_excel_se_aberto`, erro claro com `ErroComTitulo` em vez de travar a extração inteira se não conseguir).
4. **`limpar_excel_orfao(log)`** — a pedido da usuária: depois que o arquivo é movido com sucesso, o Excel que o SAP abriu fica "órfão" (o arquivo que ele tinha aberto sumiu do lugar original) e mostra um aviso "Sorry, we couldn't find..." + deixa uma janela vazia pra trás. **Bug real encontrado testando ao vivo:** a versão moderna do Office desenha esse aviso numa janela de classe `NUIDialog` (não o `#32770` clássico do Windows), e o botão "OK" não é um controle `Button` de verdade — é só um desenho dentro de um único `NetUIHWND` (confirmado inspecionando a janela ao vivo via `win32gui.EnumChildWindows`, só achou 1 filho sem nome/classe útil). Como não dá pra "clicar" num botão que não existe como controle, a correção fecha a **janela inteira** via `WM_CLOSE` (mesmo efeito prático do X) quando o título é exatamente "Microsoft Excel" — testado ao vivo contra os avisos reais (2 de uma vez, Gestoriais + Sem Agrupamento) e fechou os dois. Depois, fecha (`Quit()`) qualquer instância do Excel com `Workbooks.Count == 0` — nunca mexe numa instância com algo aberto.

**CONFIRMADO AO VIVO pela usuária, ponta a ponta (Agosto/2026, Ciclo Flash), sem clicar em nada:** os 2 avisos do Excel fecharam sozinhos, a janela vazia fechou sozinha, e os 2 arquivos (Gestoriais + Sem Agrupamento) caíram certinho em `08_Aug_Flash`, com nome limpo (sem versionamento espúrio, prova de que não ficou nenhuma sobra de tentativa anterior atrapalhando).

**Cuidado descoberto durante o debug, registrado pra próximas sessões:** o `Get-ChildItem` do PowerShell, usado por mim (Claude) pra checar essa pasta de rede específica, retornou resultado desatualizado ("pasta vazia") pelo menos 2 vezes durante essa sessão, quando na real já tinha arquivo lá — me levou a diagnosticar erroneamente que o `move` estava falhando sempre. `[System.IO.Directory]::GetFiles(...)` via .NET direto deu resultado correto nas duas vezes em que comparei os dois. Preferir `.GetFiles()` .NET (ou pedir print do Explorer da usuária) em vez de confiar cegamente num único `Get-ChildItem` ao checar esse caminho de novo.

**Lição para o projeto:** o classificador de segurança do Auto mode bloqueia consistentemente edição de arquivos fora da pasta do projeto, mesmo com autorização explícita da usuária na hora — para esse tipo de restrição (política do SAP GUI, arquivos de sistema fora do repo), a solução tem que vir de dentro do código do projeto (como essa), não de editar configuração externa.

**Conclusão:** item fechado — staging fixo elimina o popup de Segurança SAPGUI (1x por ano em vez de 1x por mês), e o travamento do Excel + avisos residuais são tratados automaticamente pelo próprio código. Validado ao vivo de ponta a ponta, sem qualquer risco ao Excel real da usuária ou a dados de produção.

---
## 2026-08-26 — Passo 5 (Rateio de Custos): APROVADO pela usuária

**Decisão:** depois de testar o botão "Abertura de Custos por Unidade" (arquivo real gerado na rede, Julho/2026 Actual), o botão "Atualizar Rateio", encontrar e corrigir um bug real (Check sobrando R$0,10 mil por arredondamento em cadeia), validar 7 meses em Actual e 7 meses em Flash (todos sem erro de fórmula, check por unidade ✓, Check final 0,00), a usuária confirmou explicitamente: **"passo 5 esta aprovado"**.

**Motivo:** validação extensa (Jan-Jul/2026, Actual e Flash) não encontrou nenhum erro real de lógica do script — todo diff contra os arquivos antigos teve explicação de negócio (lançamento manual da usuária em Abril, regra de resíduo de encerrada em Fevereiro) ou é irrelevante (< R$5 mil, ordem de grandeza de late-posting no SAP). O bug real que apareceu (Check 0,10) foi encontrado pela própria usuária testando ao vivo e corrigido no mesmo dia.

**Exceção pontual adicionada (não é reintrodução do mapeamento próprio por conta, banido em 2026-08-25):** conta `4255200` "Recuperação PIS/COFINS Depreciação" é forçada pra categoria "Depreciation" (`CONTAS_FORCADAS_DEPRECIATION` em `gerar_rateio_custos.py`), mesmo a Base Intermediária trazendo Conta Geral (AJ) = "Others". Justificativa: a usuária pediu "como ela está alocada nos meses anteriores?" — conferido no arquivo antigo que essa conta específica sempre foi classificada em "Depreciation", com valor constante mês a mês, nas 3 unidades (SJP, IBI, GOI) e nos 7 meses conferidos (Jan-Jul/2026). Correção category-neutra (não muda nenhum Total Costs).

**Ciclo Flash validado:** mesma lógica do Actual (confirmado pela usuária: "ciclo flash é a mesma forma de rodar do que o actual"). Testado Jan-Jul/2026 Flash: sem erro de fórmula, check ✓, Check final 0,00 nos 7 meses. Meses fechados (Jan-Jun) batem com os mesmos valores do Actual (esperado - mês fechado não muda entre Ciclos); Julho (mês ainda em Flash na época) mostrou valor de forecast genuinamente diferente do Actual, também batendo exato com o arquivo antigo Flash de referência.

**Pendências que NÃO bloqueiam a aprovação (aceitas pela usuária):**
- Resíduo de R$4.070,00 em Maio/IBI segue sem explicação — usuária disse explicitamente "não precisa explicar".
- RES (Resende) segue sem custo real testado — usuária confirmou "resende ainda nao teve custo, mas tera a partir do proximo mes", ciente que a lógica de RES nunca foi exercitada com valor real ainda.

**Rateio de Agosto/2026 (com Resende) confirmado alocado:** `ontology/rateio_gerencia.json`, entrada `vigente_desde: "2026-08"` — SJP 21% / IBI 49% / GOI 27% / RES 3%, já configurado desde 2026-08-25, confirmado pra usuária nesta sessão.

---

## 2026-08-27 — Passo 7 (P&L): mecanismo de automação dos links externos confirmado (ChangeLink) + erro real de PY encontrado e corrigido a partir de Agosto

**Contexto:** início do escopo do Passo 7 (P&L), aba "Resumo Resultado Ano" do Actual. A usuária perguntou se era mais seguro manter os links externos (em vez de converter pra valor) e ir trocando o destino todo mês, já que o arquivo não é pesado.

**Decisão confirmada, com evidência do arquivo real:** sim, manter os links — e a investigação mostrou que é mais simples do que parecia. O ponto que mais consumia tempo manual (trocar a fórmula das colunas D/F/G/H na aba "Resumo Resultado Mês" todo mês) **não é link externo**, é só referência de coluna dentro do próprio arquivo (`='Resumo Resultado Ano'!J8`, desliza 1 coluna por mês) — automatizável por reescrita direta da fórmula, sem precisar da API de Link do Excel.

**Os links externos de verdade (aba "Resumo Resultado Ano") e a frequência de troca, mapeados no arquivo real:**
- **Mensalização** (bloco Actual, colunas D:O — os 12 meses vêm de 1 arquivo só, aba "TOTAL") → aponta pro `MENS FITTED ACTUAL <Mês>.xls` do mês fechado, **já é a saída do Passo 6**. Troca todo mês.
- **Forecast** (coluna AF) → aponta pro P&L Forecast do mês fechado, mesmo fallback R8→R7 já usado na Mensalização. Troca todo mês.
- **Flash** (coluna E, aba "Resumo Resultado Mês") → aponta pro P&L Flash do mesmo mês. Troca todo mês.
- **PY** (coluna BJ, "Actual <ano-1>") → só troca em Janeiro.
- **MP** (coluna AU) → só troca em Janeiro.

Mecanismo validado via COM: `wb.LinkSources(1)` lista os links; `wb.ChangeLink(nome_antigo, nome_novo, 1)` troca o destino mantendo a fórmula de link viva (não converte pra valor); `excel.CalculateFullRebuild()` recalcula depois. Testado com `Workbooks.Open(..., ReadOnly=True)` + `wb.Close(SaveChanges=False)` — zero risco ao arquivo de produção.

**Erro real encontrado e quantificado (ver `memory/errors/2026-08-27_pnl_link_py_apontava_2024.md` para detalhe completo):** o link de PY (coluna BJ, "Actual 2025") está apontando pra `2024\12_December_Actual\...December-24.xlsx` em vez de `2025\...December-25.xlsx`, confirmado em **todos os 7 arquivos Actual de 2026 já fechados** (Jan-Jul) — o rollover de Janeiro não aconteceu. Testado o fix (`ChangeLink`) contra o arquivo real de Julho (só leitura, nada salvo): EBIT PY corrigido de 18.317,80 mil para 22.513,66 mil (+4.195,86 mil), ROS de 19,59% para 23,44% (+3,85 p.p.) — ou seja, todo P&L enviado pra consolidação em 2026 subestimou o EBIT do ano anterior. Link de MP conferido e está correto (não precisa de fix).

**Decisão da usuária:** **não corrigir retroativamente Jan-Jul/2026** (já fechados/enviados). A correção do link de PY entra a partir do arquivo de **Agosto/2026** (quando for criado a partir da cópia de Julho) — como PY só muda 1x/ano, deve durar até Jan/2027.

**Ainda pendente antes de implementar o Passo 7 de verdade:** a usuária ainda vai explicar a lógica do arquivo/aba Flash (só o Actual foi coberto até aqui). Nada foi implementado em script ainda — só análise e teste isolado (leitura) contra os arquivos reais.

---

## 2026-08-27 (continuação) — Segundo erro real achado nos links do P&L: Flash/Forecast apontam pra pasta sem o nível "MM - Mês" (Mai/Jun/Jul); EO_CONSUMER vs EO_FITTED

**Achado (retomada da sessão, antes de codar o Passo 7):** os links de Flash (coluna E, "Resumo Resultado Mês") e Forecast (coluna AF, "Resumo Resultado Ano") dos arquivos Actual de **Maio, Junho e Julho/2026** apontam pra `...\Resultados Fitted\2026\<MM>_<Mês3>_Flash\...` — falta o nível de pasta `<MM> - <Mês3>` que existe de verdade na rede (`...\2026\07 - Jul\07_Jul_Flash\...`). Confirmado com `Test-Path` (PowerShell nativo, não só `ls`) que o caminho curto não existe nos 3 meses. O valor mostrado na célula (ex: Flash de Julho = 451.071) bate com o valor real hoje só por coincidência — é cache congelado de antes do link quebrar, não atualiza mais se a fonte mudar. Detalhe completo em `memory/errors/2026-08-27_pnl_link_flash_forecast_pasta_faltando.md`.

**Causa confirmada pela usuária:** a pasta de rede foi reorganizada (nível "MM - Mês" adicionado) e os links dentro dos arquivos de P&L (trocados manualmente todo mês) nunca foram re-apontados pra incluir esse nível — mesmo padrão de "correção não propagada" do bug do PY.

**Achado relacionado, também confirmado:** o link de Mensalização usa `EO_CONSUMER\BU FITTED\...` em Maio/Junho mas `EO_FITTED\BU FITTED\...` em Julho (os dois compartilhamentos existem de fato). **Usuária confirmou: `EO_FITTED` é o correto/atual** — `EO_CONSUMER` não deve ser usado daqui pra frente.

---

## 2026-08-28 — Sem tabela de controle de acesso ao cockpit (estagiário)

**Contexto:** o estagiário da usuária tentou abrir o cockpit e o launcher (`atualizar_ksb1_launcher.vbs`) falhou porque tinha o caminho absoluto da máquina da Juliana hardcoded (`C:\Users\silveju001\...`) — corrigido pra calcular o caminho a partir de onde o próprio `.vbs` está salvo (commit `93605c0`), funciona em qualquer máquina desde que a pasta do projeto seja copiada inteira.

**Decisão:** não implementar uma tabela/allowlist de usuários autorizados a rodar o cockpit.

**Motivo:** o estagiário já tem acesso próprio às pastas de rede da Pirelli (SAP, financeiro) que os scripts leem/escrevem — essa é a barreira de segurança real. Só ele vai usar o cockpit além da Juliana; uma tabela de controle seria só um guardrail de cortesia sem valor de segurança de fato (não impede alguém de editar o script), e ela preferiu não adicionar essa camada.

**Como aplicar:** não construir controle de acesso local no cockpit a menos que a usuária peça explicitamente de novo (ex: se mais pessoas passarem a ter cópias do projeto).

**Decisão da usuária:** mesma política do bug do PY — **não corrigir retroativamente** Mai/Jun/Jul (já fechados/enviados). A automação do Passo 7 (`gerar_pnl.py`, ainda a implementar) sempre monta o caminho de destino do `ChangeLink` com o nível "MM - Mês" incluído (nunca replica o caminho curto) e sempre usa `EO_FITTED` pro link de Mensalização.

---

## 2026-09-01 — Ordem de clique confusa entre Passo 3 (Provisões) e Passo 4 (Base Intermediária): só reforçado o texto, não reestruturado

**Contexto:** a usuária notou que, no Ciclo Flash, a ordem real de clique não segue a numeração das abas — o botão "Atualizar Pivot KSB1" (aba ④) roda ANTES do Passo 3 (Provisões), e só depois disso é que "Finalização da Base Intermediária" (também aba ④) pode rodar, porque reaproveita o arquivo que o Passo 3 criou. Perguntou se não seria melhor inverter a ordem das abas 3 e 4.

**Opções levantadas:** (1) não mexer, (2) só reforçar o texto de aviso nas duas abas, (3) quebrar a aba ④ em duas abas separadas (uma pro "Atualizar Pivot KSB1", outra pra "Finalização"), com Provisões no meio — o que exigiria renumerar Rateio/Mensalização/P&L também.

**Decisão da usuária:** opção 2 — só reforçar o texto. Simplesmente inverter as abas 3↔4 não resolveria (o entrelaçamento continuaria, só que ao contrário), e quebrar em 8 abas era mudança grande demais pra esse problema.

**Aplicado em `atualizar_ksb1_gui.py`:** descrição da aba ③ (Provisões) e da aba ④ (Base Intermediária) agora abrem com um parágrafo "ORDEM DE CLIQUE NO FLASH" explicando a sequência real (① Atualizar Pivot KSB1 → ② Lançar Provisões → ③ Finalização da Base Intermediária) e avisando que rodar fora de ordem no Flash dá erro ou sobrescreve provisões já lançadas. No Actual, sem essa quebra — os dois botões da aba ④ rodam direto.

**Pendente:** rede não foi re-sincronizada (`robocopy /MIR` da cópia de rede pro estagiário) — fazer se/quando outras mudanças também estiverem prontas, pra não gerar vários syncs pequenos.

---

## 2026-09-01 (continuação) — Decisão anterior ("só reforçar texto") substituída: Passo 3 e 4 fundidos numa aba só

**Contexto:** logo depois de aplicar a decisão anterior (só reforçar o texto de aviso nas
abas ③ Provisões e ④ Base Intermediária — ver entrada acima), a usuária reconsiderou e
propôs uma solução melhor: já que o problema é a ordem de clique não seguir a numeração das
abas, por que não fundir as duas abas numa só e simplesmente numerar os botões na ordem
certa?

**Decisão da usuária:** fundir. Aplicada em `atualizar_ksb1_gui.py`:
- Aba ③ "Provisões" + aba ④ "Base Intermediária" → uma aba só, nova ③ "Base Intermediária",
  com 4 botões nomeados na ordem de clique real: "①  Atualizar Pivot KSB1", "②  Lançar
  Provisões", "②  Atualizar Provisões" (mesma numeração ② nos dois — são alternativas pro
  mesmo passo, não sequenciais entre si), "③  Finalização da Base Intermediária".
- Abas seguintes renumeradas: Rateio de Custos (era ⑤/Passo 5) → ④/Passo 4; Mensalização
  (era ⑥/Passo 6) → ⑤/Passo 5; P&L (era ⑦/Passo 7) → ⑥/Passo 6. Cockpit: 7 abas → 6 abas.
- Wiring de botões (`botoes[indice][i].config(command=...)`) reajustado pros novos índices.
- **Bug real achado durante o reajuste:** o aviso de Janeiro ("confirmar rateio da
  Gerência") checava `indice == 4` (índice antigo da aba Rateio de Custos) — corrigido pra
  `indice == 3` (novo índice). Sem essa correção, o aviso pararia de aparecer no Ciclo/mês
  certo, silenciosamente (sem erro, só o aviso sumindo).
- Textos corrigidos: aviso de Janeiro ("Vá na aba '④ Rateio de Custos'"), comentário e
  mensagem sobre "Atualizar Faturamento (Passo 5)" e "O Custo do Passo 5 foi atualizado"
  (ambos eram "Passo 6" antes da renumeração).

**Validado:** `py_compile` sem erro. Não testado visualmente ainda na janela (pendente:
confirmar que a aba ③ fundida renderiza os 4 botões e que cada um dispara a função certa).

**Memória atualizada:** `reference_cockpit_fechamento_fitted.md` (7 abas → 6 abas, descrição
da nova aba ③).

**Cópia de rede:** re-sincronizada nesta mesma sessão (`robocopy /MIR`) — ver rodapé desta
entrada ou a mensagem da sessão pra confirmação do resultado.

---

## 2026-09-01 (continuação) — Layout final do cockpit: texto lado a lado virou tooltip no hover; botão "Atualizar Provisões" ganhou cor cinza

**Contexto:** depois de fundir Passo 3+4 numa aba só (entrada anterior), a primeira tentativa
de resolver "botões empurrados pra baixo" foi colocar os 3 itens numerados lado a lado
(3 colunas). A usuária testou e achou "um horror" — pediu pra inverter a prioridade: botões
sempre no topo da aba, e o texto explicativo só aparece ao passar o mouse em cima do botão.

**Decisão e implementação:**
- Nova classe `_Tooltip` (Toplevel sem borda, aparece no `<Enter>`, some no `<Leave>`/clique)
  aplicada a TODOS os botões do cockpit (não só na aba ③) — cada `passo["botoes"][i]` agora
  tem um `passo["tooltips"][i]` correspondente. As chaves `"descricao"`/`"itens"` (texto
  sempre visível) foram removidas do dicionário `PASSOS` e do `fazer_aba` — cada aba mostra
  só `titulo` + botões, nada mais, o que deixou os botões colados no topo em todas as 6 abas
  (não só na ③).
- **Pedido adicional:** o botão "② Atualizar Provisões" (ação secundária/alternativa a
  "Lançar Provisões", só usada depois de uma correção no Fast Provisão) ganhou uma cor
  cinza-clara própria (`Secundario.TButton`, `#E7E7EA`), pra se diferenciar visualmente dos
  demais botões amarelos (`Pirelli.TButton`). Suporte genérico adicionado (`passo["estilos"]`,
  paralelo a `"botoes"`), reutilizável se outro botão precisar de estilo diferente no futuro.

**Validado:** `py_compile` sem erro a cada mudança; usuária confirmou visualmente ("ficou
ótimo") depois de reabrir o cockpit 3x ao longo da sessão (fusão de abas → tooltip → cor do
botão). Wiring de botões (`botoes[indice][i].config(command=...)`) reconferido depois de cada
mudança de estrutura — nenhuma função de negócio (`ao_clicar_*`) foi alterada, só a
apresentação (título/tooltip/estilo) e a estrutura de abas.
