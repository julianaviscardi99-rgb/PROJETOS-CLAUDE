# Erro: KSB1 falha com "Área de contabilidade de custos não existe"

**O que aconteceu:** os scripts de extração da KSB1 (`extrair_ksb1.py`, `atualizar_ksb1_gui.py`) nunca preenchiam o campo "Área contab. custos" (`ctxtP_KOKRS`). Funcionou nos primeiros testes porque o campo já estava preenchido manualmente de uma navegação anterior. Quando a usuária abriu a KSB1 do zero (`/nKSB1`), o campo veio vazio, o SAP acusou "Área de contabilidade de custos não existe" ao executar, e a exportação falhou com `The control could not be found by id` (porque a tela de resultado nunca abriu — o menu "Lista > Exportar" não existe na tela de seleção).

**Causa raiz:** campo obrigatório (`ctxtP_KOKRS`) não estava sendo preenchido pelo script.

**Correção:** os dois scripts agora preenchem `ctxtP_KOKRS = "0580"` explicitamente antes de executar. A usuária confirmou: para Fitted Units, a área contábil de custos é **sempre** `0580` (fixo, não muda).

**Lição:** ao automatizar uma tela do SAP, nunca assumir que um campo "já está preenchido" só porque estava assim durante o teste — mapear e preencher **todos** os campos obrigatórios explicitamente, mesmo que pareçam ter valor default.
