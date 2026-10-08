# 2026-10-07 — R9 revisado: efetivo Jan-Set (KSB1 Flash) + Out-Dez pela premissa do R9

**Pedido (Juliana):** o Detalhe de Despesas do R9 não tem o efetivo de Jul/Ago; explicar a variação grande do MP27 vs R9.
Fonte do efetivo: `...\Resultados Fitted\2026\09 - Sep\09_Sep_Flash\KSB1 September Flash 2026.xlsx` (aba `BASE_KSB1`, 74.665 linhas, Jan-Set, Flash = ainda não é o Actual fechado).

**Script:** `scripts/sap/fitted_units/mp2027/gerar_r9_revisado_com_efetivo.py` (args: KSB1, R9, MP27, pasta, último mês efetivo).
Saída: `data/processed/mp2027/MP27_vs_R9_Revisado_Efetivo_Jan_Set.xlsx` (Resumo, uma aba por unidade, Fora do Detalhe, Notas).

**Método validado:** chave (CM + Gestorial), só Gestoriais que existem no Detalhe. KSB1 Jan-Jun bate com o R9 (SJP -1,7; GOI +3,7; IBI -36 de 9.298; R$ mil).
KSB1 traz TODA a despesa (mão de obra etc.) — o que não está no Detalhe fica na aba "Fora do Detalhe".
Colunas da BASE_KSB1 (0-based): valor 16, mês 18, gestorial 19, descrição 20, CM 21.

**Resultado (despesas, R$ mil):** R9 original 35.340 → R9 revisado 34.139 (-1.201); MP27 37.059.
Δ MP27 vs R9 revisado +2.920 (vs R9 original +1.719). IBI: R9 rev 20.270 vs MP27 22.308 (+2.038).
O efetivo de Jul-Set da IBI veio 840 abaixo do previsto no R9; créditos negativos em Outras Despesas (-1.059) e Manutenção Terceiros (-372) no efetivo.

**Ambiente:** o verificador de segurança do Bash/PowerShell falhou várias vezes seguidas (sem veredito); funcionou ao reenviar o mesmo comando. Python via PowerShell lê a rede; Bash não.
