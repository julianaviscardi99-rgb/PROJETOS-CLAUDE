# Passo 3 (gerar_ksb1_mensal.py) quebrado desde 26/08 — chamada com assinatura antiga

**Contexto:** durante o teste de fechamento pedido pela usuária em 2026-08-28 (reproduzir os
7 passos para Julho/2026 Actual), tentei validar o Passo 3 e encontrei um `TypeError` real.

**Causa raiz:** em 2026-08-26, `check_agrupamentos_ksb1.eh_conta_ignorada` ganhou os
parâmetros `mes`/`ano` (para a vigência da conta N410400000, só ignorada a partir de
Agosto/2026 — ver `ontology/fitted_units.json`). `gerar_ksb1_mensal.py`
(`decidir_fonte_e_ler_linhas`, linha 99) importa e chama essa função, mas ninguém atualizou
a chamada lá — continuou `eh_conta_ignorada(c)`, só 1 argumento. Como a função não tem
default para `mes`/`ano`, isso quebra com `TypeError: eh_conta_ignorada() missing 2 required
positional arguments: 'mes' and 'ano'` — e a linha roda incondicionalmente (antes de decidir
Gestoriais vs. Sem Agrupamento), então **qualquer** chamada a `decidir_fonte_e_ler_linhas`
quebraria, para qualquer mês/Ciclo, desde 26/08.

**Por que não foi pego antes:** nenhuma sessão rodou o Passo 3 de ponta a ponta contra dado
real depois de 26/08 — as sessões seguintes focaram no Passo 6 (Mensalização) e Passo 7
(P&L), que já leem a Base Intermediária pronta, não passam por `decidir_fonte_e_ler_linhas`.

**Correção:** `eh_conta_ignorada(c)` → `eh_conta_ignorada(c, mes, ano)` (linha 99).

**Validação pós-fix:** rodei a lógica corrigida contra os arquivos brutos reais de
Julho/2026 Actual (localizados em `.../07 - Jul/07_Jul_Actual/Bases SAP/`, formato
pré-automação — ver aviso relacionado em
`memory/learnings/2026-08-28_check_julho_actual_n410400000.md`), agrupei por
(Conta Fiscal, Centro de Custo) e comparei contra a Base Intermediária real de Julho/Actual:
mesmíssimo padrão de diferença já documentado (16 combinações, 100% unidades encerradas —
Sorocaba, R$ 112.275,58) — confirma que a correção não muda nenhum valor de Julho (a
N410400000 nunca foi ignorável nesse mês de qualquer forma) e que a lógica volta a funcionar.

**Lição:** quando uma função compartilhada muda de assinatura, grep por todos os call sites
antes de considerar a mudança fechada — `check_agrupamentos_ksb1.py` foi corrigido/testado
sozinho em 26/08, mas `gerar_ksb1_mensal.py` (outro script que importa a mesma função) não
foi re-testado na mesma sessão.
