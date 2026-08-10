# Sandbox do Claude Code intercepta chamadas diretas a cmd.exe

O ambiente de execução de comandos (Bash tool, Git Bash/MSYS) não consegue rodar `cmd.exe /c algum_script.bat` diretamente para testar `.bat` — a chamada é interceptada e abre um shell interativo em vez de executar o comando, sem erro explícito.

**Contorno:** para comandos nativos do Windows que dependem de path com espaços ou sintaxe estilo `/flag` (ex: `schtasks`), usar `MSYS_NO_PATHCONV=1` antes do comando, para o Git Bash não tentar converter `/tn`, `/tr` etc. em caminhos Unix. Isso funcionou para `schtasks /create ...`.

**Não resolvido:** não foi possível validar a execução do `Auto_Backup_GitHub.bat` diretamente neste ambiente — a lógica foi validada rodando os comandos `git add/commit/push` equivalentes manualmente. Testar o `.bat` com duplo-clique fora deste chat.
