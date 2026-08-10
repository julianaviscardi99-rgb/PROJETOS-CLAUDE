# Erro: pasta de rede da Pirelli inacessível a partir do Bash tool (mesmo com rede ok)

**O que aconteceu:** ao tentar copiar o `.bat` corrigido para `\\FSS024-01BR.group.pirelli.com\GFU_DAC\...`, `Path.exists()` retornou `False` mesmo no share raiz (`\\FSS024-01BR.group.pirelli.com\GFU_DAC`). A usuária confirmou que a rede estava ok do lado dela. `ping` no servidor respondeu normalmente (13ms), mas `net use` não mostrou nenhum drive mapeado.

**Causa provável:** o processo do Bash tool (Git Bash/MSYS) parece não herdar a sessão SMB autenticada da usuária da mesma forma que uma sessão interativa (Explorer) — servidor alcançável na rede, mas sem sessão/autenticação SMB estabelecida para esse processo.

**Contorno:** pedir para a usuária copiar manualmente o arquivo via Explorer, ou abrir a pasta de rede no Explorer primeiro (o que costuma forçar a autenticação SMB) antes de tentar de novo pelo Bash.

**Não resolvido:** não confirmado se abrir a pasta no Explorer primeiro destrava o acesso pelo Bash na mesma sessão. Testar da próxima vez.

**Atualização (mesma sessão, mais tarde):** numa tentativa seguinte, o Bash conseguiu ler e escrever normalmente em `\\FSS024-01BR.group.pirelli.com\GFU_DAC\...` sem nenhuma ação extra da usuária. Ou seja, o problema é intermitente (provavelmente a sessão SMB expira/precisa ser renovada de tempos em tempos), não uma limitação permanente do Bash tool. Recomendação: sempre tentar direto pelo Bash primeiro; só pedir para a usuária copiar manualmente via Explorer se a tentativa falhar.
