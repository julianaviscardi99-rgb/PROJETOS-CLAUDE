# Aprendizado — KSB1: campos alternativos + popup de segurança de arquivo por pasta nova

**Contexto:** explorando a KSB1 pro novo sub-projeto "Energia Elétrica Fitted".

## 1. Centro de custo/Classe de custo vs. seus grupos são alternativos

Na tela de seleção da KSB1, `Centro de custo` (KOSTL) e `Grupo de centros de custo` (KSTGR) são alternativos (rótulo "ou" entre eles no layout) — o mesmo vale para `Classe de custo` (KSTAR) e `Grupo de classes de custo` (KOAGR). Se os dois lados de qualquer um desses pares ficarem preenchidos ao mesmo tempo (ex: alguém navegou manualmente e digitou um centro de custo específico, depois um script preenche o grupo por cima sem limpar o campo específico), o SAP recusa a execução com o erro genérico **"Selecionar uma das alternativas indicadas"** — sem apontar qual campo é o culpado.

**Como evitar:** todo script que usa a KSB1 deve sempre usar os campos de **grupo** (`KSTGR`=`BU['kstgr']`, `KOAGR`="" ou "gestoriais") e **nunca tocar** em `KOSTL-LOW/HIGH` nem `KSTAR-LOW/HIGH`. Se esse erro aparecer, checar primeiro se algum desses 4 campos está preenchido inesperadamente (pode ter sobrado de navegação manual anterior no mesmo GUI, já que `/nKSB1` não limpa a tela — só assume valores novos por cima dos antigos).

## 2. Popup "Segurança SAPGUI" ao exportar pra uma pasta nova

Ao usar o menu nativo de exportação (Lista > Exportar > Planilha eletrônica) pra uma pasta que o SAP GUI ainda não "conhece", ele abre um popup modal (fora da árvore normal de `session.Children`, então **não é detectável só checando `wnd[1]`/`wnd[2]` via COM** — um script que espera o arquivo aparecer fica preso indefinidamente sem erro nenhum, parece só lento) pedindo autorização (Permitir/Rejeitar + "Memorizar minha decisão"). Isso trava qualquer script que dependa da exportação até alguém clicar manualmente no SAP.

**Como evitar:** nunca usar pastas efêmeras/temporárias pra exportação (ex: a pasta de scratchpad do Claude Code muda a cada sessão — cada uma dispara o popup de novo). Sempre exportar pra uma pasta estável já conhecida: `data/processed/<algo>` (local, gitignored) ou a pasta de rede já usada nos outros scripts. Isso também é a regra de qualidade do próprio `CLAUDE.md` ("Todo output salvar em `data/processed/`") — seguir ela evita esse problema de brinde.

**Se travar mesmo assim:** checar `tasklist` — o processo Python trava com status "Not Responding" no `tasklist /v` enquanto espera; não adianta esperar mais, precisa a usuária ir no SAP e clicar Permitir/Rejeitar manualmente.
