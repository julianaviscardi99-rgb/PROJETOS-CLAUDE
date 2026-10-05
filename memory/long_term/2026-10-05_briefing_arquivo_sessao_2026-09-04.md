# Arquivo — sessão 2026-09-04 (movida do BRIEFING em 2026-10-05)

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

