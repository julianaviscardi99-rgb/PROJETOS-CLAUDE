# Erro: popup "Definir área contab.custos" ao sair e voltar a entrar no SAP

**O que aconteceu:** a Juliana identificou que, toda vez que sai e volta a entrar no SAP GUI (nova sessão/login do dia), a primeira transação aberta dispara um popup modal "Definir área contab.custos" pedindo para preencher o campo "Área contab.custos" e confirmar na seta verde antes de liberar a tela principal. Os scripts (`atualizar_ksb1_gui.py`, `extrair_ksb1.py`) não tratavam esse popup — se ele aparecesse no meio da automação, os campos da tela de seleção da KSB1 (`wnd[0]/usr/...`) ficariam bloqueados/inacessíveis por trás do popup (`wnd[1]`), quebrando o script.

**Causa raiz:** popup de sessão nova do SAP, independente da tela de seleção da KSB1 (que já tinha o campo `ctxtP_KOKRS` tratado — ver `memory/errors/2026-08-10_ksb1_kokrs_vazio.md`). Esse popup é anterior a qualquer tela de transação.

**Correção:** adicionada a função `tratar_popup_area_contabil` em `scripts/sap/atualizar_ksb1_gui.py` e `scripts/sap/extrair_ksb1.py`. Ela verifica se existe `wnd[1]` logo após abrir a KSB1; se existir, preenche o primeiro campo de texto/combo dentro de `usr` com `"0580"` e pressiona `wnd[1]/tbar[0]/btn[0]` (seta verde/confirmar).

**Atenção — não testado ao vivo:** o ID técnico exato do campo dentro do popup não foi confirmado contra o SAP real (não temos acesso à sessão SAP da usuária para inspecionar `wnd[1]` diretamente). A correção usa uma busca genérica pelo primeiro campo editável (`GuiCTextField`/`GuiTextField`) dentro de `usr`, e assume que o botão de confirmar é o primeiro botão da toolbar do popup (`btn[0]`), com base no padrão visual da tela (seta verde é sempre a primeira ação). **A usuária precisa rodar o script depois de sair/voltar a entrar no SAP para confirmar que o popup é fechado corretamente.** Se falhar, o próximo passo é capturar o ID exato do campo (ex: com o SAP Scripting Recorder ou inspecionando `wnd[1].FindById(...).Id` na hora que o popup aparece) e trocar a busca genérica por esse ID fixo.

**Lição:** popups de sessão do SAP (login/nova sessão) podem interferir em scripts de automação mesmo fora da tela da transação-alvo — vale sempre checar `wnd[1]` logo depois de abrir qualquer transação, não só tratar os campos da tela esperada.

---

## Atualização 2026-08-11 — primeira tentativa falhou, causou erro em cascata

A usuária testou e a busca genérica pelo campo (primeiro `GuiCTextField`/`GuiTextField` dentro de `usr`) não encontrou o campo certo, mas o script ainda assim clicou no botão "confirmar" (`btn[0]`) sem ter preenchido nada. Isso fez o próprio SAP abrir um segundo popup "Preencher todos os campos obrigatórios" empilhado por cima, e a extração seguinte falhou com `The virtual key is not enabled` (efeito colateral de tentar mandar `SendVKey` com um popup modal ainda bloqueando a tela).

**Correção aplicada:** `tratar_popup_area_contabil` agora **não clica em "confirmar" se não encontrar o campo** — em vez disso, levanta um erro claro pedindo para rodar `scripts/sap/diagnosticar_popup.py` (script novo, lista Id/Type/Text de tudo dentro do `wnd[1]` aberto) para descobrir o Id técnico exato do campo e do botão, em vez de continuar adivinhando.

**Status:** aguardando a usuária rodar `diagnosticar_popup.py` com o popup aberto e mandar a saída, para eu trocar a busca genérica pelo Id fixo correto.

---

## Atualização 2026-08-11 — causa raiz confirmada e corrigida

Saída do `diagnosticar_popup.py` (com o popup aberto):
```
/app/con[0]/ses[0]/wnd[1]                                          [GuiModalWindow]     'Definir área contab.custos'
  /app/con[0]/ses[0]/wnd[1]/tbar[0]                                 [GuiToolbar]
    /app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]                        [GuiButton]
    /app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[5]                        [GuiButton]
    /app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[12]                       [GuiButton]
  /app/con[0]/ses[0]/wnd[1]/usr                                     [GuiUserArea]
    /app/con[0]/ses[0]/wnd[1]/usr/sub:SAPLSPO4:0300                 [GuiSimpleContainer]
      .../txtSVALD-KEYTEXT[0,0]                                     [GuiTextField]   'Área contab.custos' (label, não editável)
      .../ctxtSVALD-VALUE[0,21]                                     [GuiCTextField]  '' (campo editável — este é o certo)
```

**Causa raiz confirmada:** o popup é o diálogo genérico padrão do SAP para "definir parâmetro" (programa `SAPLSPO4`). O campo editável (`ctxtSVALD-VALUE`) fica dentro de um subscreen (`usr/sub:SAPLSPO4:0300/...`), não como filho direto de `usr` — por isso a busca original (`usr.Children` direto) nunca encontrava nada, mas o código antigo pressionava "confirmar" mesmo assim.

**Correção final:** troquei a busca por filhos diretos por uma busca **recursiva** (`_buscar_campo_editavel`, em ambos os scripts) que desce por todos os níveis de `usr` até achar o primeiro `GuiCTextField` (o tipo do campo editável — o label é `GuiTextField` e é ignorado). Isso deve funcionar independente do nome/índice exato do subscreen (`SAPLSPO4:0300` pode variar). O botão "confirmar" continua sendo `tbar[0]/btn[0]` (primeiro botão da toolbar do popup — bateu com o esperado, é a seta verde).

**Ainda não testado ao vivo após essa correção** — usuária vai confirmar na próxima vez que sair/voltar a entrar no SAP.
