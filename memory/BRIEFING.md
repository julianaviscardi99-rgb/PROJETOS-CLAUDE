# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## CONCLUÍDO 2026-09-04 — Fechamento de Agosto/2026 ACTUAL destravado: R$ 79.787,48 que a Pivot escondia em silêncio, corrigido na raiz + 2 travas novas no Passo ①

**Sintoma trazido pela usuária:** dois números que deveriam bater no fechamento de Agosto/2026
**Actual** — extração da KSB1 = **R$ 5.671.131,15**, mas o que chegava na Base Intermediária =
**R$ 5.750.918,63**. Diferença de **R$ 79.787,48**, com o custo do mês INFLADO.

**Primeira hipótese (ERRADA, registrar pra não repetir):** achei que era cache da PivotTable
dessincronizado (`recordCount` com 1 registro a mais). Ela rodou o refresh de novo e o número não
mudou — hipótese descartada. **A pista boa veio de reconciliar a Pivot contra os dados brutos pelas
8 chaves de `rowFields`**, em vez de teorizar: 73 combinações com diferença, mas 72 eram pares que
se cancelavam (só troca de Variabilidade F↔V, líquido zero). Sobrou **uma** diferença líquida.

**CAUSA RAIZ (cadeia completa):**
1. Duas linhas lançadas em **31/08/2026**, CC **8296 (Ibirité)**, conta **`M240600000`
   "Rech cost reco:FI-Gr"**, ambas CRÉDITO: "Repasse Man. Bancais Ibirité - Materiais"
   (**-29.595,22**) e "- Mdo" (**-50.192,26**) = **-79.787,48**.
2. A conta **não estava cadastrada** na `Base_Contas_Contábeis_Fitted_22.xlsx` (aba `Contas`) →
   a coluna T (Gestorial) do BASE_KSB1, que é `=VLOOKUP(D;[1]Contas!$A:$J;10;0)`, resolvia `#N/A`.
3. **O filtro do campo "Gestorial" da `Pivot_Inter.` tem os itens `#N/A` e `(vazio)` DESMARCADOS**
   (`h="1"` em `pivotTable1.xml`) → as 2 linhas sumiam do Grand Total **sem erro nenhum**. Como são
   crédito, a Pivot ficava R$ 79.787,48 mais ALTA que a realidade, e a Base Intermediária herdava.

**É o mesmo padrão do 9º bug de 2026-09-02** (item `(blank)` do campo "Var." desmarcado escondendo
as provisões). Vale desconfiar sempre desse filtro quando um total não bater.

**AS 3 CORREÇÕES:**
1. **Conta cadastrada** — `M240600000` na linha 564 da aba `Contas` como **4263000 / Outras
   Despesas** (escolha da usuária, mesmo tratamento da conta irmã `M230600000` "Rec. de Custos
   Terceiros"), com **backup datado** (`Base_Contas_Contábeis_Fitted_22.backup_2026-09-04.xlsx`).
   Conferido antes: a aba não é ListObject, os VLOOKUPs usam coluna inteira, e a gestorial 4263000
   já existia nas abas `Classificação Despesas`/`Classificação Despesas Fixo`.
2. **Links externos passam a ser resolvidos AO VIVO** (`abrir_fontes_dos_links` /
   `fechar_fontes_dos_links` em `gerar_ksb1_mensal.py`), revertendo o `UpdateLinks=0` de 2026-08-11.
   **A abordagem foi ideia da usuária** ("e se você abrisse o arquivo base de contas pra ele ler
   sempre?") — testei a hipótese isolada, **sem nenhum `UpdateLink`**, e a coluna T saiu de `#N/A`
   pra `4263000` só por a base de contas estar aberta. Abre a origem SOMENTE LEITURA, descobre o
   caminho pelo próprio `wb.LinkSources` (sem hardcode), e `UpdateLink` virou plano B.
   **Ordem que importa: só fechar as fontes DEPOIS do `Save`** — fonte fechada antes, o Excel volta
   a gravar o valor em cache.
3. **Trava nova** (`conferir_pivot_contra_base`) no fim do Passo ①: compara o Grand Total do mês na
   `Pivot_Inter.` com um `SUMIF` direto do BASE_KSB1 e, se divergir, aborta **listando as contas com
   Gestorial em erro** (via `SpecialCells(xlCellTypeFormulas, xlErrors)`, sem varrer 67 mil linhas).
   Roda **depois** do `Save`, de propósito — não joga fora os 10+ minutos de colagem.
   Pega a classe inteira do problema: qualquer item escondido em filtro de qualquer campo da Pivot.

**RESULTADO CONFIRMADO NOS ARQUIVOS REAIS (a usuária rodou e deu certo):**
| Arquivo | Valor |
|---|---|
| `KSB1 August Actual 2026_v3.xlsx` → Pivot_Inter. Agosto | **5.671.131,15** ✓ |
| `Base Intermediária ... Actual 2026_v3.xlsx` → coluna P | 5.022.726,06 (= 5.671.131,15 − 648.405,09 de unidades encerradas) ✓ |
| Linha 491 da Intermediária | `M240600000` → Gestorial **4263000**, CC 8296, MF 0491, **-79.787,48** ✓ |

**ARMADILHA EM QUE EU CAÍ (lição geral, vale pra qualquer escrita em xlsx via COM):** a primeira
tentativa de cadastrar a conta **falhou em silêncio** — a `Base_Contas_Contábeis_Fitted_22.xlsx`
também tem `<fileSharing readOnlyRecommended="1"/>`, o Excel abriu em modo leitura com
`DisplayAlerts=False`, o `Save()` virou no-op e o script ainda imprimiu "salvo". O
`IgnoreReadOnlyRecommended=True` **não** resolveu. Só descobri porque fui reconferir o arquivo em
disco. Solução num arquivo COMPARTILHADO: remover a flag do XML → gravar via COM → **restaurar a
flag** no fim. **Depois de escrever xlsx via COM, sempre conferir no disco com openpyxl** — "salvo"
sem erro não prova nada neste ambiente.

**Outras conferências feitas (todas OK):**
- **Flash de Agosto NÃO foi afetado** — a conta `M240600000` não aparece na extração do Ciclo Flash
  (0 linhas), só na do Actual. O P&L Flash que já foi enviado está correto nesse ponto.
- Varredura do BASE_KSB1 inteiro (Jan-Ago, 67 mil linhas): essas 2 linhas de Agosto são as **únicas**
  com Gestorial `#N/A`/vazio no ano todo — não houve vazamento silencioso em meses anteriores.
- **O "Check de Agrupamentos" NÃO pega esse caso** (deu "OK - valores batem" às 15:35): ele compara
  Gestoriais × Sem Agrupamento (os dois extratos do SAP, que ambos contêm a conta) e **não verifica
  se cada conta do mês existe no de-para em Excel**. Ponto cego real, ainda em aberto.

**Detalhe completo:** `memory/errors/2026-09-04_pivot_inter_ksb1_cache_dessincronizado.md` e
`memory/DECISOES.md` (entrada de 2026-09-04).

**PENDÊNCIAS / PRÓXIMOS PASSOS:**
1. **Seguir o fechamento de Agosto/2026 Actual** a partir do Passo ④ (Rateio de Custos) — Passos ①
   e ③ concluídos e conferidos hoje. Passos ⑤ Mensalização e ⑥ P&L ainda não rodados no Actual.
2. **Considerar estender o "Check de Agrupamentos" (Passo ②)** para conferir que toda conta do mês
   existe na aba `Contas` do de-para — pegaria esse tipo de problema um passo antes, já no ②, em vez
   de só no fim do ①. Ofereci hoje, ela preferiu só a trava; retomar quando fizer sentido.
3. Quando aparecer conta nova, alguém precisa cadastrá-la no de-para — vale confirmar com ela **quem
   é o dono** da `Base_Contas_Contábeis_Fitted_22.xlsx` (é arquivo corporativo compartilhado, hoje
   editei com backup + restauração da flag de somente leitura).
4. Continua valendo do dia 09-02: investigar por que o Excel às vezes não fecha sozinho
   (`excel.Quit()`), e consolidar os aprendizados recorrentes em `memory/learnings/`.

---
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

## Histórico anterior (2026-09-01 para trás)

Movido para `memory/long_term/2026-09-04_briefing_arquivo_ate_2026-09-01.md` em 2026-09-04
(regra: no máximo 2 sessões inline no BRIEFING). Snapshots automáticos completos também ficam
em `memory/long_term/*_briefing_snapshot.md`.
