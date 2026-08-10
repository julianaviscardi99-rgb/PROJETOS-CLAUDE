# REGRAS_RAPIDAS — Top 10 (leitura de 60 segundos)

> Lido no início de cada sessão, antes de qualquer outra coisa.

1. Ações que MODIFICAM dados externos (banco, email, sistema X) → sempre confirmar antes.
2. Nunca sobrescrever arquivos originais — salvar sempre em pasta separada com data.
3. Nunca usar caminhos absolutos hardcoded nos scripts Python.
4. Ao errar algo → registrar imediatamente em `memory/errors/`.
5. Ao aprender algo novo → registrar imediatamente em `memory/learnings/`.
6. Ao tomar decisão de processo → registrar em `memory/DECISOES.md` com o motivo.
7. Ao explicar conceito de negócio → registrar em `ontology/`.
8. `BRIEFING.md`: manter máximo 2 sessões inline — sessões mais antigas vão para `long_term/`.
9. Commit sempre seguido de push imediato no mesmo momento.
10. Antes de qualquer script novo: verificar se o conhecimento já existe nas ontologias.
11. Comandos, status e perguntas para a usuária sempre em português (pedido explícito, 2026-08-10).
12. Ao salvar qualquer arquivo em uma pasta que já tem um arquivo com o mesmo nome, nunca sobrescrever: salvar como "_v2"; se "_v2" já existir, "_v3", e assim por diante (pedido explícito, 2026-08-10).

---

## Gate pré-execução

Antes de qualquer operação que modifica dados (externos ou locais), responder:

1. Essa ação é reversível? Se não for, o usuário já confirmou explicitamente?
2. Essa ação sobrescreve algum arquivo original? Se sim, parar — salvar em `data/processed/` com data em vez disso.
3. O script usa algum caminho absoluto do usuário? Se sim, corrigir antes de rodar.
4. Existe uma ontologia (`ontology/*.json`) que já descreve esse conceito de negócio? Se sim, ler antes de assumir a regra.
5. Se a ação falhar ou o resultado for inesperado, o erro será registrado em `memory/errors/` antes de tentar de novo?
