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
