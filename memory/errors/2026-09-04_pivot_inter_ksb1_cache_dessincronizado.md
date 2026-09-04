# 2026-09-04 — Agosto/2026 Actual: Pivot_Inter. descarta em silêncio linhas com Gestorial #N/A (conta nova fora do de-para)

> NOTA: a primeira hipótese registrada neste arquivo (cache da PivotTable dessincronizado,
> `recordCount` com 1 registro a mais) estava **ERRADA** e foi descartada — a usuária rodou o
> refresh de novo (`_v2`, 16:26) e o número não mudou. A causa raiz real está abaixo.

## Sintoma
Fechamento de Agosto/2026 **Actual**. Dois números que deveriam bater:
- Extração bruta da KSB1 (Gestoriais), Agosto: **R$ 5.671.131,15**
- Valor que chega na Base Intermediária (via `Pivot_Inter.`): **R$ 5.750.918,63**
- Diferença: **R$ 79.787,48** (a Pivot mostra MAIS custo que a realidade)

## Causa raiz (cadeia completa, confirmada nos arquivos reais)
1. Duas linhas lançadas em **31/08/2026**, centro de custo **8296 (Ibirité)**, conta contábil
   **`M240600000` — "Rech cost reco:FI-Gr"** (repasse/recuperação de custo), ambas CRÉDITO:
   - "Repasse Man. Bancais Ibirité - **Materiais**": **-29.595,22**
   - "Repasse Man. Bancais Ibirité - **Mdo**": **-50.192,26**
   - Total: **-79.787,48** (linhas 51232 e 51233 do `BASE_KSB1`)
2. A conta `M240600000` **NÃO está cadastrada** na `Base_Contas_Contábeis_Fitted_22.xlsx`
   (aba `Contas`, 562 contas) — a que mais se aproxima é a irmã `M230600000` "Rec. de Custos
   Terceiros" → Gestorial 4263000 "Outras Despesas".
3. A coluna T (Gestorial) do `BASE_KSB1` é `=VLOOKUP(D<n>;[1]Contas!$A:$J;10;0)` — sem a conta no
   de-para, resolve **`#N/A`** nessas 2 linhas. (Aqui o range é coluna inteira `$A:$J`, então NÃO é
   o mesmo problema de range travado da coluna AH/Resende de 2026-09-01 — a conta simplesmente
   não existe na tabela.)
4. **O filtro do campo "Gestorial" da `Pivot_Inter.` tem exatamente 2 itens desmarcados**
   (`pivotTable1.xml`, campo índice 19): `<e v="#N/A"/>` e `<m/>` (vazio), ambos com `h="1"`.
   Ou seja, **toda linha cujo Gestorial seja #N/A ou vazio some do Grand Total sem erro nenhum**.
5. Como as linhas que somem são CRÉDITO, a Pivot fica R$ 79.787,48 **mais alta** que o real, e a
   Base Intermediária herda esse custo inflado.

**É exatamente o mesmo padrão do 9º bug de 2026-09-02** (item `(blank)` do campo "Var." desmarcado
fazia as provisões sumirem) — só que agora no campo "Gestorial" e com o item `#N/A`.

## Conferências feitas
- Reconciliação Pivot × BASE_KSB1 pelas 8 chaves de `rowFields`: 73 combinações com diferença, mas
  72 delas são **pares que se cancelam** (mesma chave, só trocando Variabilidade F↔V — ruído de
  reclassificação, líquido zero). A **única** diferença líquida é a linha de `#N/A`: R$ 79.787,48.
- Varredura do `BASE_KSB1` inteiro (Jan-Ago, 67.076 linhas): essas 2 linhas de Agosto são as
  **únicas** com Gestorial #N/A/vazio no ano todo — não houve vazamento silencioso em meses
  anteriores.
- **Flash de Agosto NÃO foi afetado**: a conta `M240600000` não aparece na extração do Ciclo Flash
  (0 linhas) — só na do Actual. O P&L Flash que já foi enviado está OK nesse ponto.
- A aba `Pivot` da Base Intermediária está consistente com a aba `Intermediária` (mesmo valor nos
  dois) — o erro entra antes, pela `Pivot_Inter.` do KSB1.
- O **"Check de Agrupamentos" não pega esse caso** (rodou hoje 15:35 e deu "OK - valores batem"):
  ele compara Gestoriais × Sem Agrupamento (os dois extratos do SAP, que ambos contêm a conta) —
  ele **não verifica se cada conta do mês existe no de-para em Excel** (`Contas`), que é onde a
  falha está. Ponto cego real do Check.

## O QUE JÁ FOI FEITO (2026-09-04, mesma sessão)
1. **Conta cadastrada.** A usuária escolheu tratar a `M240600000` igual à conta irmã `M230600000`
   → **Gestorial 4263000 / Outras Despesas**. Linha 564 da aba `Contas` da
   `Base_Contas_Contábeis_Fitted_22.xlsx`, gravada via COM (preserva formatação/links, o que o
   openpyxl destruiria): `A=M240600000 | B=Rech cost reco:FI-Gr | C=4263000 | D=Outras Despesas |
   E=Others | F=Others | H=F | I=DG | J=4263000 | K=1`. **Backup datado criado ANTES**:
   `Base_Contas_Contábeis_Fitted_22.backup_2026-09-04.xlsx`, mesma pasta.
   - Conferido antes de gravar: a aba `Contas` **não** é Tabela/ListObject (append simples serve),
     os VLOOKUPs usam coluna inteira (`$A:$J`), e a gestorial 4263000 **já existe** nas abas
     `Classificação Despesas` e `Classificação Despesas Fixo` — nenhum outro cadastro necessário.
   - Conferido também que a Base Intermediária puxa do mesmo de-para (`[2]Contas`) nas colunas
     Y (Gestorial II/col J), Z (MO/DG/col I), AH (Var./col H) e AJ (Conta Geral/col F) — todas
     resolvidas com essa mesma linha.
2. **Trava implementada em `gerar_ksb1_mensal.py`** (nova função `conferir_pivot_contra_base`,
   chamada no fim de `colar_linhas_e_atualizar_pivots`): compara o Grand Total do mês na
   `Pivot_Inter.` com `SUMIF` direto do BASE_KSB1 (1 chamada COM, Excel calcula nativamente) e
   levanta erro se divergir. Quando dispara, usa `SpecialCells(xlCellTypeFormulas, xlErrors)` na
   coluna T pra **listar as contas culpadas com valor e nº de linhas** — sem varrer as 67 mil linhas.
   - **Roda DEPOIS do `Save`, de propósito:** os dados da BASE_KSB1 estão corretos e a colagem custa
     10+ minutos; o arquivo é preservado e o erro sobe pra usuária resolver o de-para antes de
     seguir pro Passo ③.
   - **Testado contra o arquivo real** (`KSB1 August Actual 2026_v2.xlsx`, cópia local, só leitura):
     disparou certinho, reportando base 5.671.131,15 × pivot 5.750.918,63, diferença 79.787,48 e
     `M240600000 (-79.787,48 em 2 linha(s))`.
   - O teste também pegou um detalhe real: a seta "→" na mensagem quebrava em console cp1252
     (`UnicodeEncodeError`) — trocada por ">". Arquivo inteiro conferido: codifica em cp1252.
   - `py_compile` OK. **Sincronizado na cópia de rede do cockpit** (conferido idêntico com `diff`).

3. **Links externos passam a ser atualizados a cada geração** (decisão dela, 2026-09-04): nova
   função `atualizar_links_externos(wb)` em `gerar_ksb1_mensal.py`, chamada logo depois de pôr o
   cálculo em manual (assim não dispara recálculo — o `CalculateFullRebuild` seguinte resolve tudo
   de uma vez). Reverte a decisão de 2026-08-11 (`UpdateLinks=0`), que era pra **validação** contra
   mês fechado mas em **produção** fazia conta nova nunca resolver. Falha de link não derruba a
   geração (só loga aviso) — o BASE_KSB1 já teve link quebrado pra RHFitted (lixo conhecido).
   - Risco conferido antes de mudar: a base de contas não era modificada desde **16/07** e o KSB1
     de Julho foi salvo em **21/08** — ou seja, o cache já estava atualizado pra todo o resto;
     atualizar os links muda **só** as 2 linhas da conta nova, nenhum mês histórico se mexe.
   - **Teste ponta a ponta na cópia local** (`KSB1 August Actual 2026_v2.xlsx`, nada escrito na
     rede): antes `T=#N/A`; depois de `atualizar_links_externos` + recálculo, `T=4263000` e
     `U=Outras Despesas`; depois do `RefreshAll`, a conferência **PASSOU**
     ("Pivot_Inter. e BASE_KSB1 batem no mês 8 (5.671.131,15)"). `py_compile` OK, cp1252 OK,
     sincronizado na rede.

## ARMADILHA EM QUE EU CAÍ (registrar pra não repetir)
A primeira tentativa de cadastrar a conta **falhou em silêncio**: o script escreveu, imprimiu
"salvo." e o arquivo em disco continuou sem a linha. Causa: a
`Base_Contas_Contábeis_Fitted_22.xlsx` também tem `<fileSharing readOnlyRecommended="1"/>` — o
Excel abre em modo leitura **silenciosamente** com `DisplayAlerts=False`, e o `Save()` vira no-op
sem erro. `IgnoreReadOnlyRecommended=True` **não resolveu** (limitação já documentada no próprio
projeto, em `remover_flag_somente_leitura_recomendada`) e `wb.ReadOnly` continuou `True`.
- **Solução usada (arquivo compartilhado, não uma cópia nossa):** remover a flag do XML → gravar
  via COM → **restaurar a flag** no fim, pro arquivo terminar com o mesmo comportamento pra quem
  abrir. Conferido depois: linha presente em disco **e** `fileSharing` de volta.
- **Lição geral:** depois de escrever em xlsx via COM, **conferir no disco** (openpyxl) que o dado
  persistiu — `wb.Save()` sem erro e `print("salvo")` não provam nada neste ambiente. Checar
  `wb.ReadOnly` logo depois do `Open()` também (o `gerar_ksb1_mensal.py` já faz isso; meu script
  pontual não fazia).

## Correção necessária (pendente de decisão da usuária)
1. **Cadastrar `M240600000` na `Base_Contas_Contábeis_Fitted_22.xlsx` (aba `Contas`)** com o
   agrupamento gestorial correto — **decisão de negócio dela** (arquivo corporativo compartilhado,
   não mexer sem OK). Sugestão de partida: 4263000 "Outras Despesas", igual à irmã `M230600000`;
   mas como uma das linhas é de **Mdo**, ela pode querer que o repasse abata mão de obra
   (muda a linha do P&L: Labour vs Others, e o MO/DG).
2. **Atenção ao link externo:** `gerar_ksb1_mensal.py` abre o BASE_KSB1 com `UpdateLinks=0`
   (deliberado, ver docstring/DECISOES 2026-08-11) — então, mesmo cadastrando a conta, o VLOOKUP
   **não** vai reresolver sozinho: o arquivo do mês herda o cache do link do arquivo do mês
   anterior. Precisa atualizar os links do `KSB1 August Actual 2026_v2.xlsx` (Dados → Editar Links
   → Atualizar valores) antes de refazer a Pivot — ou revisar essa decisão de `UpdateLinks=0`
   para produção.

## Guarda proposta no código (ainda não implementada)
- **Invariante barato e forte:** depois do `RefreshAll`, comparar o **Grand Total do mês na
  `Pivot_Inter.`** contra a **soma direta do `BASE_KSB1` para aquele mês** (que o script já tem em
  mãos — é o total da extração, gravado em `AK1`). Se divergir → abortar com mensagem clara.
  Isso pegaria de uma vez qualquer item escondido no filtro de qualquer campo da Pivot, não só
  este caso.
- Complementar: estender o "Check de Agrupamentos" para conferir que **toda conta do mês existe na
  aba `Contas`** do de-para — pegaria o problema um passo antes, já no Passo ②.
