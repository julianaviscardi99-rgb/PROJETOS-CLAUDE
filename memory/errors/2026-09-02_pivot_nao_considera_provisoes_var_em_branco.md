# 2026-09-02 — Pivot da aba "Pivot" não considerava o valor das provisões no cálculo

**Relatado pela usuária:** depois de clicar "② Atualizar Provisões" e depois "③ Finalização
da Base Intermediária" (Agosto/2026 Flash), o Grand Total da PivotTable (aba "Pivot",
`Base Intermediária Fitted August Flash 2026.xlsx`) mostrava R$ 3.473.552,82 pra Agosto —
mas a soma real da coluna de Agosto (P) da aba "Intermediária" era R$ 5.137.087,24. Diferença
de R$ 1.663.534,42.

**Diagnóstico (abrindo o arquivo real via COM, sem alterar nada):**
- A diferença exata batia com a soma da coluna P **só das linhas COLORIDAS (2-67, provisões/
  reclassificações)** — a Pivot estava contando SÓ as linhas brancas (KSB1), como se as
  provisões não existissem.
- `PivotCache.RecordCount` = 912 (o cache tinha TODAS as linhas, inclusive as coloridas) — não
  era problema de range/fonte de dados.
- A causa real: os campos de linha da Pivot ("Var." e "MO/DG & Var", colunas AA/AC da
  Intermediária) estavam **em branco** em todas as 49 linhas coloridas com valor lançado. E o
  item de filtro `(blank)` do campo "Var." estava **desmarcado** (`Visible=False`) na
  PivotTable — então qualquer linha com "Var." em branco simplesmente some do Grand Total,
  sem erro nenhum.
- Por que "Var."/"MO/DG & Var" ficavam em branco: `preencher_provisoes_flash`
  (`gerar_base_intermediaria.py`) só arrasta a fórmula "molde" (da última linha colorida, a
  roxa) pras colunas `COL_FORMULA_MODELO = [1,2,4,6,7]` (A,B,D,F,G) de cada linha de provisão
  ativa — nunca pras colunas Y:AJ (25-36, Gestorial II até Conta Geral, que incluem "Var." e
  "MO/DG & Var") nem pra Total Ano (coluna U/21).
- Isso nunca dava problema em "① Lançar Provisões" (não limpa nada antes, então as fórmulas
  Y:AJ/Total Ano herdadas do arquivo do mês anterior continuavam lá, recalculando sozinhas com
  base no novo Centro de Custo da linha). Mas em "② Atualizar Provisões", `limpar_provisoes`
  apaga A:AJ (inclusive Y:AJ e Total Ano) de TODAS as linhas amarelas antes de
  `preencher_provisoes_flash` preencher de novo — e como esse preenchimento nunca repunha
  Y:AJ/Total Ano nas linhas ativas (só nas "sobrando sem provisão", que ficam limpas de
  propósito), as linhas com provisão real ficavam com "Var." em branco depois de qualquer
  "Atualizar Provisões".

**Corrigido em `preencher_provisoes_flash`:** nova constante
`COL_FORMULA_MOLDE_EXTRA = [COL_TOTAL_ANO] + list(range(COL_FORMULA_INICIO, COL_FORMULA_FIM + 1))`
somada a `COL_FORMULA_MODELO` na captura/aplicação da fórmula "molde" — agora TODAS as colunas
de fórmula (A,B,D,F,G + Total Ano + Y:AJ) são arrastadas pra cada linha de provisão ativa,
igual já acontecia só com A,B,D,F,G. `py_compile` OK. Cópia de rede do cockpit
(`_Cockpit_KSB1\scripts\...`) ressincronizada e conferida idêntica (`diff` sem diferença).

**Ainda NÃO aplicado ao arquivo de Agosto já existente** (`Base Intermediária Fitted August
Flash 2026.xlsx` na rede) — o código só corrige o preenchimento DAÍ PRA FRENTE. Pra corrigir
o arquivo atual, a usuária precisa fechar o arquivo (se estiver aberto no Excel) e rodar de
novo, nesta ordem: "② Atualizar Provisões" (agora com o fix, recompõe Var./Total Ano nas
linhas coloridas) e depois "③ Finalização da Base Intermediária" (recalcula tudo E dá
`wb.RefreshAll()` na Pivot — sem isso a Pivot não atualiza sozinha, mesmo com os valores
certos na Intermediária).

**Lição pra generalizar:** sempre que uma correção adicionar fórmula "molde" só numa
sub-lista de colunas (ex: só A,B,D,F,G), checar se existe uma sub-lista IRMÃ (Y:AJ, Total
Ano) que também depende da mesma linha "molde" e que ficou de fora — o sintoma pode não dar
erro nenhum, só um número errado numa tabela dinâmica que filtra "(blank)" silenciosamente.
