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
