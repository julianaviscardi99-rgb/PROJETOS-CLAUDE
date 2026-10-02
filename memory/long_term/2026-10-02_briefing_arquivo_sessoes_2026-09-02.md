# Arquivo — sessões de 2026-09-02 (movidas do BRIEFING em 2026-10-02)

## EM ANDAMENTO 2026-09-02 (nova janela) — Usuária voltou reportando "vários problemas" usando o cockpit em produção; revisão do Passo 3 em andamento

**Contexto:** a usuária confirmou que Passos ① e ② do cockpit e os botões ①/② do Passo 3
("Atualizar Pivot KSB1" e "Lançar/Atualizar Provisões") estão **funcionando perfeitamente,
não mexer**. Os problemas reais estão no botão ③ "Finalização da Base Intermediária".

**1. Bug real confirmado e CORRIGIDO — botão ③ sobrescrevia o arquivo Flash em vez de
versionar.** Pedido explícito da usuária: toda rodada de Finalização deve gerar `_v2`, `_v3`
etc., nunca sobrepor (mesma regra já vale pro resto do projeto, `REGRAS_RAPIDAS.md` #2/#12).
- Causa: no branch `ciclo == "Flash"` de `atualizar_base_intermediaria`
  (`gerar_base_intermediaria.py`), `caminho_saida` apontava direto pro arquivo já criado por
  "Lançar/Atualizar Provisões" e `wb.Save()` gravava por cima dele. O branch `Actual` já fazia
  certo (sempre `nome_com_versao` + `shutil.copy2` antes de editar).
- **Correção aplicada:** branch Flash agora localiza o arquivo mais recente
  (`localizar_base_intermediaria_flash_existente`), copia pra um nome novo via
  `nome_com_versao` (mesmo padrão do resto do projeto) e só então abre/edita a cópia. `py_compile`
  OK. Passos seguintes (Rateio de Custos, `gerar_rateio_custos.py`) já usam
  `encontrar_arquivo_mais_recente` pra achar a Base Intermediária, então vão pegar a versão nova
  sozinhos, sem precisar de mais nenhuma mudança.
- **NÃO sincronizado na rede ainda** (`_Cockpit_KSB1\scripts\`) e **NÃO testado ao vivo** —
  próximo passo.

**2. Suspeita de bug NO quadro de comparação Forecast (Custos H26/I26) — investigado, mas os
dados batem, aguardando esclarecimento da usuária.** Ela relatou que o quadro trouxe "a
informação do forecast R7, mês de julho" em vez de agosto (mostrou print: Custos Forecast =
5.925, Flash = 5.137). Fui direto no arquivo real de rede
(`...\2026\07 - Jul\07_Jul_Forecast\07_P&L Fitted Units_Forecast_July_26_.xlsx`, aba "Resumo
Resultado Ano") e conferi:
- Cabeçalho confirma coluna J=Julho, K=Agosto.
- Total Costs em **K (Agosto) = -5.925,41** — bate EXATAMENTE com o "5.925" do quadro dela.
- Total Costs em J (Julho) = -6.320,46 — não bate.
- Confirmei também que `08_Aug_Forecast` (pasta de rede) está vazia — não existe R8 de
  verdade, então o fallback pro R7 está correto por design (já documentado em
  `DECISOES.md`, 2026-08-22).
- **Conclusão até agora: não achei o bug — os números automatizados parecem corretos
  (coluna de Agosto, não Julho).** Perguntei pra ela se conferiu comparando célula a célula
  ou só "pareceu" errado — pode ser outra célula (câmbio, Faturamento manual) que ela
  confundiu, ou pode haver algo que eu não vi ainda. **Resposta dela ainda pendente.**

**Outra coisa feita nesta sessão (fora do cockpit):** configurada a statusLine do Claude Code
pra mostrar "Modelo | Pasta | Ctx XX%" (`~/.claude/statusline-command.sh` +
`~/.claude/settings.json`). Trocado de `jq` (não instalado nesta máquina) pra Python (já
instalado) depois que o primeiro teste falhou — testado com JSON de exemplo, funcionando.

**PRIMEIRA COISA A PERGUNTAR NA PRÓXIMA SESSÃO/MENSAGEM:**
1. A resposta dela sobre o quadro de comparação Forecast (item 2 acima) — ela confirmou que
   comparou célula a célula, ou foi outra célula que ela viu errada?
2. Sincronizar a correção do botão ③ (item 1) pra cópia de rede do cockpit e testar ao vivo.
3. Ela mencionou "vários problemas" no plural — só cobrimos os 2 acima (botão ③ Finalização);
   perguntar se tem mais algum problema noutro passo (④ Rateio, ⑤ Mensalização, ⑥ P&L) que
   ainda não foi reportado.

---
## EM ANDAMENTO 2026-09-02 — 9º bug real do fechamento de Agosto/2026 Flash: a PivotTable não considerava as provisões. CORRIGIDO no código, mas a usuária ainda NÃO rodou com o código novo carregado

**PRIMEIRA COISA A PERGUNTAR NA PRÓXIMA SESSÃO:** ela fechou e reabriu o cockpit e rodou
"② Atualizar Provisões" + "③ Finalização" de novo? O Grand Total de Agosto da aba Pivot
passou a mostrar **R$ 5.137.087,24** (em vez dos R$ 3.473.552,82 errados)? Se sim, seguir pro
Passo ④ Rateio de Custos (que continua pendente de rodar pelo cockpit, ver entrada de 09-01).

**Sintoma relatado pela usuária:** na aba "Pivot" da `Base Intermediária Fitted August Flash
2026.xlsx`, o Grand Total de Agosto mostrava R$ 3.473.552,82, mas a soma real da coluna P
(August) da aba "Intermediária" era R$ 5.137.087,24 — diferença de R$ 1.663.534,42.

**Causa raiz (achada abrindo o arquivo real via COM, sem alterar nada):**
- A diferença batia EXATAMENTE com a soma da coluna P das linhas coloridas (provisões).
- `PivotCache.RecordCount` = 912: o cache tinha todas as linhas, não era problema de fonte.
- Os campos que a Pivot usa pra agrupar linhas — **"Var." (col 27/AA) e "MO/DG & Var" (col
  29/AC)** — estavam **em branco** nas 29 linhas de provisão com valor, e o item `(blank)` do
  campo "Var." está **desmarcado** (`Visible=False`) no filtro da PivotTable. Linha com "Var."
  em branco simplesmente some do Grand Total, sem erro nenhum.
- Por quê: `preencher_provisoes_flash` só arrastava a fórmula "molde" (linha roxa 67) pras
  colunas `COL_FORMULA_MODELO = [1,2,4,6,7]` (A,B,D,F,G) — nunca pra Y:AJ (25-36, que inclui
  Var. e MO/DG & Var) nem pro Total Ano (U/21). Em "① Lançar Provisões" isso passava batido
  (não limpa nada antes, as fórmulas herdadas do mês anterior seguiam lá); mas em
  "② Atualizar Provisões", `limpar_provisoes` apaga A:AJ das amarelas antes e o
  preenchimento nunca repunha Y:AJ/Total Ano nas linhas ATIVAS (só limpa as sobrando).

**Correção aplicada:** nova constante `COL_FORMULA_MOLDE_EXTRA = [COL_TOTAL_ANO] +
list(range(COL_FORMULA_INICIO, COL_FORMULA_FIM + 1))`, somada a `COL_FORMULA_MODELO` na
captura/aplicação da fórmula molde em `preencher_provisoes_flash`. `py_compile` OK. Cópia de
rede do cockpit ressincronizada às 09:39 e conferida idêntica (`diff` sem diferença). Detalhe
completo em `memory/errors/2026-09-02_pivot_nao_considera_provisoes_var_em_branco.md`.

**Por que ela disse "ainda está com erro" depois do fix:** ela rodou "② Atualizar Provisões"
às 09:43, DEPOIS do fix estar na rede (09:39) — mas o cockpit já estava ABERTO desde antes,
com o módulo antigo carregado na memória do Python. Conferido no arquivo: a linha 2 tem
fórmula só em A,B,D,F,G (comportamento antigo), nada em U/21 nem Y:AJ. **O fix está certo, só
não foi carregado.** Regra que já vale pra qualquer correção de script: fechar e reabrir o
cockpit antes de testar.

**Próximo passo (combinado, ainda não feito):** fechar/reabrir o cockpit → "② Atualizar
Provisões" → "③ Finalização da Base Intermediária" (essa é a que dá `wb.RefreshAll()` na
Pivot; "Atualizar Provisões" sozinho NÃO atualiza a PivotTable). Alternativa oferecida: eu
rodar os dois passos por linha de comando (código novo carrega sozinho), mas mexe no arquivo
real da rede — precisa do OK dela.

**Layout das linhas coloridas da Intermediária (conferido neste arquivo, útil pra referência):**
amarelas (provisão) = 2-47 (29 com valor em Agosto), verdes = 48-51, roxas = 52-67, sendo a
**67 a linha "molde"** de fórmula; dados brancos (KSB1) começam na 68 e vão até 912.

---
---

