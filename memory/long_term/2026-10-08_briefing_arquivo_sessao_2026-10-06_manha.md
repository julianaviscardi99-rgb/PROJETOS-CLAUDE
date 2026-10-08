# Arquivo do BRIEFING — sessão 2026-10-06 (manhã)

---
## EM ANDAMENTO 2026-10-06 (manhã) — MP27 V2: revisão da IBIRITÉ (só leitura feita, NADA gravado) + decisão sobre saída da JLR

**Estado do arquivo:** `...\Forecast\MP\MP 2027\MENS FITTED MP27_v2.xls` — última gravação 05/10 18:27, **hoje nada foi gravado na V2**. SJP e GO seguem prontas (ver seção 05/10). Backup de 05/10 em `data/processed/backups_mp27/2026-10-05/`.

### FEITO 06/10 11:28 — aba nova "Impacto JLR" gravada na V2 (só leitura das outras abas; backup `v2_BACKUP_antes_aba_impacto_jlr_1114.xls`, script `criar_aba_impacto_jlr.py`)
- Seções: (1) dados da JLR no R9 e MP'26 (lidos dos arquivos-fonte, em azul; **frete e material JLR = ESTIMATIVA** 21,84×0,9075 e 6,99/pç, células amarelas — usuária aprovou os ~R$ 21 mil no R9); (2) e (3) volume/mix/preço por cliente (Fiat, Iveco, CNH, **JLR em linha própria**) vs R9 e vs MP'26; (4) conferência com IBI!T16/W16 e volume + EBIT sem JLR. Status OK nos dois.
- **Resultados (R$ mil):** margem direta JLR = **522,5 (R9)** e **526,8 (MP'26)**. Vs R9 a variante preço+mix da IBI (−1.457,1) é **quase toda JLR (−1.472,1)**; volume da JLR só −0,5 (1.059 pç de 2,9 mi) — a saída aparece como MIX. EBIT IBI: variante vs R9 −1.272,7 → **−750,2 sem a JLR**; vs MP'26 +506,3 → **+1.033,2 sem a JLR**.
- Aba NÃO alimenta a TOTAL (não foi pedido). Rótulos da aba nova em português.

### (concluído) pedido 06/10 ~11:15: medir o impacto da SAÍDA DA JLR (Ibirité) e expurgá-lo da variante volume/preço/mix
- A usuária corrigiu: **não é Sorocaba, é a JLR de Ibirité** ("quero medir o quanto a saída da JLR impacta"). Fonte indicada: `...\Forecast\Fcst\Fcst 2026\R9 2026\MENS FITTED FORECAST SETEMBRO_FINAL.xls` (aba IBI). Também existe `...\MP\MP 2026\MENS FITTED MP2026_v3.xls` (aba IBI, mesmas linhas de JLR).
- **Dados já localizados (dump em scratchpad, só leitura):** no R9 e no MP'26 a aba IBI tem Volume JLR (linha 55), Preço JLR (62), Faturamento JLR (69), JLR/Rodas insumos (R9: 76/77/83/84/88-89; MP'26: 73/77/81), Frete (R9 86 / MP'26 79). Na V2 atual a JLR **já foi retirada** da IBI pela usuária (volume = Fiat+Iveco+CNH; linhas 53-55).
- **Plano:** criar na V2 uma ABA NOVA "Impacto JLR" (não mexe nas abas existentes), com volume/receita/materiais+rodas/frete/margem da JLR no R9 e no MP'26 (valores lidos dos arquivos-fonte, origem anotada), efeito volume/preço/mix da saída e a variante ex-JLR da IBI; conferir que Σ = variante já existente. Ainda NÃO criado. Pendente confirmar com ela: frete JLR Jan-Jun estimado (≈1.059 un × 21,84 × 0,9075 ≈ R$ 21 mil) e se a aba nova deve alimentar a TOTAL.
- Lembrete: a usuária edita a V2 em paralelo (salvou 10:40); reler e conferir `LastWriteTime`/lock antes de gravar. Rótulos ficam em INGLÊS.

### GRAVADO 06/10 10:59 na V2 — REVISÃO ESTRUTURAL (RES, TOTAL, linha 50) + aba SOR EXCLUÍDA
- Backup: `data/processed/backups_mp27/2026-10-06/v2_BACKUP_antes_revisao_estrutural_1040.xls` (+ `revisao_estrutural.py`, `audita.py`). Testado em cópia, reaberto e reauditado sem erros.
- **A usuária salvou a V2 às 10:40** (IBI no layout da SJP: T=var R9, V=MP'26, W=var MP'26; JLR fora do volume; **rótulos de volta em INGLÊS de propósito — decisão dela: NÃO traduzir**).
- **Variantes:** SJP, GO, IBI, RES e TOTAL com T43 = Q43−R9, W43 = Q43−MP'26 e conferência (linha 51) = OK.
- **Correções:** linha 50 (∆ vs MP'26) da IBI e RES usava `E44−E48` (ROS) → `E43−E48`. **RES** tinha MC `E10−E18` e EBIT `E28−E31` com custo já negativo (EBIT inflado: 5.812,7 → **2.586,4**), S26 com sinal trocado, `#DIV/0!` em W (V8=0 → `IFERROR`), Z5 `#REF!`. SJP!M32 padronizada.
- **TOTAL reconstruída** no layout da SJP (T/V/W/Y/Z), convenção negativa, soma só de SJP+IBI+GO+RES; T/W = soma das variantes das unidades; Outras Receitas no nível TOTAL (`T41=Q41−S41`, que capta a saída da SOR de −188,2). EBIT TOTAL Q43 = **12.151,1**; R9 TOTAL = soma das unidades + 188,203 (SOR) ✔; MP'26 TOTAL = soma ✔. **SOR excluída** (só tinha R$188 mil de Outras Receitas em Agosto, vindos do R9).
- **Sobrou (ocultas, não mexi):** `Rateio Fixo!D5 =IBI!Q10-IBI!#REF!/1000` (quebrou quando a linha da JLR saiu da IBI) e 14 `#REF!` na MDO que leem `'Rateio Fixo'!#REF!`.
- Pendências antigas seguem (hardcodes IBI J52/H13, MO Set-Dez lê `Labour Cost R09`, linhas 19/32 mistas Jan-Ago × Set-Dez, JLR vs MP'26).

### GRAVADO 06/10 09:16 na V2 — IBI: custo negativo (padrão SJP `*-1`) + português + IFERROR + conferência
- Backup: `data/processed/backups_mp27/2026-10-06/v2_BACKUP_antes_IBI_sinal_portugues.xls` (+ script `editar_ibi_sinal_portugues.py`). Testado antes numa cópia local; reaberto sem erros.
- **Sinal:** constantes Jan-Ago e colunas S/T/AB das linhas de custo viraram negativas; links Set-Dez `=(link/1000)*-1` / `=link*-1`; linhas 21/22 `=M79/1000*-1`, `=M86/1000*-1` (bloco de apoio 79-89 segue POSITIVO). MC `=E10+E18`, EBIT `=E28+E31+E41`; variantes W/X/AC no padrão da SJP (`S16+S26`, `Q-S`). Linhas 26/29/44 e preços/insumos com IFERROR (o `IF(>0)` antigo zeraria custo negativo).
- **Resultado:** EBIT 6.808,70, W43 −591,4 e AC43 +1.187,6 IDÊNTICOS; linha 50 = conferência (OK). 31 rótulos em português. W41/AC41 criadas (Outras Receitas na variante, valor 0).
- **ATENÇÃO TOTAL:** a aba TOTAL soma `SJP+IBI+GO+RES` célula a célula, mas as fórmulas próprias dela (MC `=Q10-Q18`, EBIT `=Q28-Q31`) assumem custo POSITIVO — já estava misturado antes (SJP/GO negativos, IBI positivo) e agora precisa ser refeita com a convenção negativa (+ 22 erros). `Confronto` e `Rateio Fixo` também leem a IBI (sinal agora consistente nos dois lados). Aba RES: convenção ainda não conferida.
- **Não feito (fora do pedido):** hardcodes J52/H13, MO Set-Dez lendo `Labour Cost R09`, linha 72 (83 vs 84), colunas ocultas T-Z, JLR (zerar + "Efeito saída JLR").

### IBI — achados (lido na V2, só leitura)
- **OK:** variantes fecham (W43 = Q49 = −591,4; AC43 = Q43−AB43 = +1.187,6); EBIT 6.808,70 = MC − Fixo = soma dos meses; R9 (7.400,14) e MP'26 (5.621,12) = soma dos meses; **zero `#REF!`** (os "10 erros" de 05/10 eram `####`).
- **CONVENÇÃO DE SINAL DIFERENTE da SJP/GO:** na IBI o custo é POSITIVO (MC = Vendas − Variável; EBIT = MC − Fixo; variantes S−Q, ganho +). Por isso `W8 = (Q8-S8)*(S16-S26)` está CERTO (corrigi o erro que eu tinha registrado). **Não aplicar o padrão da SJP sem converter o sinal.**
- **A fazer (quando ela mandar):** (1) português (Pieces, NET SALES, Variable Cost, Labour, Handling, Rents, Other Fixed, FEB/APR/OCT/DEC, "Variabile / Pc", "Condominio"); (2) IFERROR em E44, Q26, Q29, Q44, linhas de preço/insumos; (3) linha de conferência (colunas W e AC; row 50 está livre); (4) hardcodes: `J52 =245303-4`, `H13 =(-7802.603325/1000)+0.5058`, frete Set-Dez com 311,55 e 21,84 embutidos; (5) Jan-Ago digitados × Set-Dez fórmula/link nas linhas 19-23, 32-37, 66-69 — MO Set-Dez (linhas 19 e 32) ainda lê `Labour Cost R09 2026` (mesmo link que ela quebrou na SJP); (6) linha 72: Jan-Fev usam linha 83, Mar-Dez usam linha 84; (7) colunas ocultas T,U,V,X,Y,Z com restos (U28/V28 = margem/pç rotulada como variação); (8) Outras Receitas (linha 41) fora do EBIT/variante.
- **Preço JLR de abril (H62) digitado 0**, mas faturamento/volume = 1.335,78; média Q62 (1.390,59) e Q63 misturam JLR.

### DECISÃO DA USUÁRIA (06/10) — JLR
"Em Ibirité o cliente JLR sai; para 2027 não tem volume previsto." Escolheu: **zerar a JLR nos 12 meses (volume, preço, materiais, rodas, frete) E criar na variante uma linha "Efeito saída JLR"**, com volume/preço/custos calculados SEM a JLR (volume e preço da variante ex-JLR) para a variante continuar fechando com o EBIT.
- **Dados achados:** a JLR na V2 é IDÊNTICA à do R9 (Jan 90, Abr 270, Mai 540, Jun 159 = 1.059 un; faturamento 1.472.631,9; rodas JLR 921.763,3 nas linhas 83/84; frete: termo `M55*21,84*0,9075` só existe Set-Dez, Jan-Ago é total digitado — o frete JLR de Jan-Jun precisa ser ESTIMADO ≈ 1.059×21,84×0,9075 ≈ R$ 21 mil, confirmar com ela). Margem JLR no R9 ≈ 1.472,6 − 921,8 − ~21 ≈ **R$ 530 mil** (aprox.).
- **Ainda falta:** valores da JLR no **MP'26** (arquivo `\...\Forecast\MP\MP 2026\MENS FITTED MP2026_v3.xls`, aba IBI, mesmas linhas 55/69/83/84/86) para o efeito vs MP'26 (colunas AB/AC).
- **Plano de fórmulas (gravar só com ok dela):** bloco auxiliar "JLR na referência" (volume, vendas, materiais/rodas, frete) sob as colunas S (R9) e AB (MP'26); bases ex-JLR (S8x = S8−vol JLR etc.); `W8=(Q8−S8x)*(S16x−S26x)`, `W16=(Q16−S16x)*Q8`, linhas 21/22 usam S21/S22 menos JLR; nova linha (ex.: 42) "Efeito saída JLR" = −margem JLR na referência; `W43` soma a nova linha; conferir W43 = Q49 e AC43 = Q43−AB43 antes de salvar. Zerar JLR: linha 55 (Jan-Jun→0), preços 62/88/89, materiais 76/77, frete.
- **Atenção:** zerar a JLR muda EBIT da IBI (−~R$ 530 mil de margem) e os totais; ela disse que "valores serão revistos depois", mas ESCOLHEU zerar agora.

### PRÓXIMOS PASSOS
1. Ler MP'26 (JLR) → montar o bloco auxiliar → mostrar a ela as fórmulas e o impacto → **gravar na V2 com backup** (pedir que feche a V2; PowerShell, não Bash, para a rede; `win32com` via arquivo .py).
2. Aplicar português + IFERROR + conferência na IBI (padrão adaptado à convenção de sinal da IBI).
3. **RESENDE** (ainda não vista), depois **TOTAL** (20 `#REF!`).
4. Pendências antigas: ver seção 05/10 e 02/10.

