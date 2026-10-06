# 2026-10-05 — GO/IBI: efeito volume da variante com sinal trocado

**Sintoma:** conferência da variante vs MP'26 na GO dava −727 (W43 = −1.284 vs Q43−V43 = −558).
**Causa:** `W8 = +(Q8-V8)*(V16-V26)` (e `T8` idem). Linha 26 é negativa → `V16-V26` soma o custo em vez de subtrair. Correto (igual SJP): `(V16+V26)`.
**Por que passou batido:** T8 = 0 (volume igual ao R9), e eu disse na 1ª revisão que a variante fecha; não conferi o sinal da fórmula, só o resultado.
**Correção:** feita pela usuária na GO (18:07/18:19). **IBI NÃO tem esse erro** (corrigido em 06/10): lá o custo é positivo (MC = Vendas − Variável), então `(S16-S26)` é o certo. Antes de aplicar o padrão da SJP numa aba, confirmar a convenção de sinal dela.
**Prevenção:** conferir a *fórmula* da variante contra a SJP, não só o valor; testar com volume diferente do R9.
