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
