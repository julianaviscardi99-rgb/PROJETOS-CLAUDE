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
