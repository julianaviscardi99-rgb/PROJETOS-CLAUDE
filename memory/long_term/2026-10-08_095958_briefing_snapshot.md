# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## SESSÃO 2026-10-07 — Por que o resultado do MP27 V2 caiu + R9 REVISADO com efetivo Jan-Set (nada gravado nos arquivos da rede)

### Por que o EBIT caiu (V2 salva 06/10 21:01 vs backup das 11:14; só leitura)
- **EBIT TOTAL 12.151 → 2.916 (−9.235, R$ mil):** receita **−6.794** (volumes novos em todas as unidades, −618 mil pç = −11%: SJP 771→711, IBI 2.923→2.703, GO 1.124→1.004, RES 575→357; RES preço 7,3→11,1 amorteceu), custo variável só **+765** (MO variável = R9, provisória, não acompanha volume; Handling IBI fixo 350,3/mês = +791), custo fixo **−3.206** (IBI Outros Fixos 4.523→6.614 vindo do Detalhe V3; RES Labour 0→−1.944 — de manhã a RES estava SEM mão de obra, então o número das 11:14 estava inflado; SJP/GO aliviaram 644/251).
- **EBIT por unidade (antes → agora):** SJP 2.110→1.825, IBI 6.127→193 (ROS 0,5%), GO 1.327→487, RES 2.586→412. **vs R9 (TOTAL): −10.085**, sendo IBI −7.207 (MC −4.691 por volume + saída JLR ≈ −520; fixo −2.516), SJP −2.246, GO −947, RES +503.
- IBI no Detalhe V3 (despesas 22.308 vs R9 21.147 vs MP'26 20.477): Aluguéis +972, Outras Despesas +588, Alimentação +454 ("Desjejum + Café" 401 novo, Refeição +438), Manut. Terceiros +297, Transp. Funcionários +226, Energia +162; Materiais Diretos −1.119. ~860 do "aumento" é o R9 estar baixo (linha manual "Delta MP'26" −500 + créditos negativos).

### R9 REVISADO (pedido dela: R9 sem efetivo de Jul/Ago)
- **Script:** `scripts/sap/fitted_units/mp2027/gerar_r9_revisado_com_efetivo.py` (args: KSB1, R9, MP27, pasta, último mês efetivo). **Saída:** `data/processed/mp2027/MP27_vs_R9_Revisado_Efetivo_Jan_Set.xlsx` (Resumo, abas por unidade, Fora do Detalhe, Notas) — **já aberto no Excel dela**.
- **Fonte do efetivo:** `...\Resultados Fitted\2026\09 - Sep\09_Sep_Flash\KSB1 September Flash 2026.xlsx` (BASE_KSB1, Jan-Set, **Flash**, não Actual fechado). Out-Dez = R9 do Detalhe. Chave CM+Gestorial, só gestoriais do Detalhe. Validado: Jan-Jun KSB1 ≈ R9 (SJP −1,7; GOI +3,7; IBI −36).
- **Resultado (despesas, R$ mil):** R9 orig 35.340 → **R9 rev 34.139** (−1.201; IBI −877: efetivo Jul-Set 840 abaixo do previsto) ; MP27 37.059 → **Δ vs R9 rev +2.920** (vs orig +1.719). IBI: rev 20.270 vs MP27 22.308 (+2.038). Efetivo IBI com créditos negativos: Outras Despesas −1.059, Manut. Terceiros −372 (parecem reclassificação, **não confirmado**).
- Só cobre as DESPESAS do Detalhe; o EBIT do R9 no MENS (receita/volume/MO) não foi revisado.

### PENDENTE
1. Abrir mês a mês os créditos de Outras Despesas/Manut. Terceiros da IBI (reclass? para qual voz foi o custo?) — próximo passo sugerido.
2. Ela confirmar: volumes novos são os definitivos? Outros Fixos da IBI no Detalhe V3 está certo? Handling IBI deve variar com volume? MO variável continua = R9 (aguarda RH)? Aluguéis IBI 4.099 vs R9 3.127; "Desjejum + Café" e transporte de funcionários novos = decisão ou duplicidade?
3. Quando sair o Actual fechado de Set, rodar o script de novo. Itens da sessão 06/10 abaixo (MDO `#REF!`, `Sheet2`/`#REF!` do Detalhe V3, refazer comparativo de vozes com V3 atual) seguem abertos.
- **Ambiente:** verificador de segurança do Bash/PowerShell falhou várias vezes seguidas (reenviar o mesmo comando funcionou). Sessão 05/10 arquivada em `memory/long_term/2026-10-07_briefing_arquivo_sessao_2026-10-05.md`.

---
## SESSÃO 2026-10-06 (tarde/noite) — Comparativo de despesas MP26 x R9 x MP27 + auditoria de erros de fórmula (NADA gravado nos arquivos da rede)

### Comparativo por voz (feito, em `data/processed/mp2027/Comparativo_Vozes_MP26_R9_MP27.xlsx`)
- Script: `scripts/sap/fitted_units/mp2027/gerar_comparativo_vozes_mp26_mp27.py` (args opcionais: MP26, MP27, R9, saída). Lê a aba `DataBase_Detail` dos **Detalhes de Despesas** ("Total C/ Curva" Jan-Dez; MP26 e R9 = CQ:DB, MP27 = BC:BN), R$ mil, custo positivo.
- **Arquivos-fonte (GFU_DAC\Management Plan):** MP26 `MP 2026\Detalhe_Despesas_Fitted Units_MP2026_v3.xlsx`; MP27 `MP 2027\Detalhe_Despesas_Fitted Units_Budget'27_V3.xlsx`; R9 `EO_FITTED\...\Fcst 2026\R9 2026\Detalhe_Despesas_Fitted Units_Forecast Setembro_final.xlsx`. (Engano meu no início: usei o MENS FITTED; "o arquivo é o DETALHE DE DESPESAS".)
- Abas: Resumo, Vozes (total), Vozes por unidade, Categorias, **Serviço por CM** e **Serviço (total)** (campo "Detalhe Serviço/Produto", pedido dela), Maiores linhas, Notas. Ranking pela variação MP27 vs R9. Totais fechavam com o `Resumo Custos` de cada arquivo.
- **DEFASADO:** gerado com o V3 das 16:37 (MP27 = 42.486). O V3 foi salvo de novo às 19:17 e agora fecha em **37.886** (SJP 7.757, IBI 22.999, GOI 5.288, RES 1.609, GER 233). MP26 = 34.722,6 e R9 = 35.340,1. **Refazer** o comparativo com o V3 atual (rodar o script de novo) antes de usar os números.
- Achados (com o V3 antigo): maiores vozes vs R9 = Prestados-Serviços, Materiais Diretos (R9 subestimado), Aluguéis, Transporte de Mats. Vários; IFRS16-Aluguéis da IBI (−1.507,8 no MP26) some no MP27 = provável reclassificação; "HSE HQ" 2.161 só no MP27; Transporte de Funcionários com nomes diferentes por versão (reclassificação); R9 tem linhas "Delta MP'26" (ajuste manual) e Gerência/Xerox negativos.

### Auditoria de erros de fórmula
- **Detalhe de Despesas V3 (varrido, nada alterado):** `DataBase_Detail!L192` e `L312` com `#REF!` (chave L; ninguém usa a coluna L, o Resumo usa M); chave L aponta para a coluna R de outra linha em L6/96/311/313/315/316/383; `E304` digitado no meio de fórmulas; aba `Sheet2` (só no V3) com `#DIV/0!` (E17) e `#N/A` (C31:E31) e 27 fórmulas lendo o arquivo antigo "Forecast September". Resumo Custos fecha com a DataBase_Detail em todos os meses. V2 tem os mesmos 2 `#REF!` (L195, L327).
- **MENS FITTED MP27_v2.xls (varrido às 19:44, pedido final dela — era este o arquivo):** abas de uso (SJP, IBI, GO, RES, TOTAL, Impacto JLR, Rateio Fixo) **sem erro**; só a aba oculta **MDO** tem 21 `#REF!` (linhas 63/64/67, cols H-L,N,O, vêm de `'Rateio Fixo'!#REF!`); nenhuma aba lê a MDO. `Rateio Fixo!D5` já foi corrigido por ela. Linha de conferência (51) = OK em todas.
- **Mudança de números no MENS (não é erro):** as despesas das 4 unidades agora vêm do Detalhe V3 (72 células/aba, `Resumo Custos!AF..`). EBIT Q43 hoje: SJP 1.818, IBI **114 (ROS 0,3%)**, GO 477, RES 409, TOTAL 2.818 (de manhã: SJP 2.110, IBI 6.808, GO 1.327, TOTAL 12.151). Confirmar com ela se a queda da IBI é esperada.
- **Vínculo nos dois sentidos:** MENS lê o Detalhe V3 e o Detalhe V3 lê o MENS (Frete e Materiais Diretos SJP/IBI/GOI, 72 células, `=[MENS]SJP!E94*-1` etc.; GO usa `*-1000`). Sem circularidade numérica, mas quebra se um dos arquivos mudar de nome/lugar.
- `Rateio Fixo` coluna "Rateio MP'27" é digitada (21/49/27/3%) e a média calculada por volume+faturamento dá 18,6/55,2/19,6/6,6% — perguntar se é intencional.

### PENDENTE desta sessão (aguardando ela)
1. Decidir a MDO: corrigir a referência para o Rateio Fixo certo ou excluir a aba (ela disse que está desatualizada). Ela precisa **fechar o MENS** antes de eu gravar.
2. Se quiser, corrigir no Detalhe V3: `#REF!` e chaves da coluna L, E304, Sheet2 (rascunho?) — precisa fechar o V3.
3. **Refazer o comparativo** com o V3 atual; opcional: consolidar nomes parecidos de "Detalhe Serviço/Produto" (ex.: Transporte de Funcionários) — ela indica quais são o mesmo serviço.
4. Itens abertos da seção abaixo (IBI/RES/TOTAL já foram tratados em 06/10; ver blocos GRAVADO).

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

---
## Histórico anterior (2026-09-01 para trás)

Movido para `memory/long_term/2026-09-04_briefing_arquivo_ate_2026-09-01.md` em 2026-09-04
(regra: no máximo 2 sessões inline no BRIEFING). Snapshots automáticos completos também ficam
em `memory/long_term/*_briefing_snapshot.md`.
