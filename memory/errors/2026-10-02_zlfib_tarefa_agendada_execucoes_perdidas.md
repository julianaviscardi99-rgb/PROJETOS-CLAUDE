# 2026-10-02 — ZLFIB mensal não rodou no 1º dia útil (01/10/2026)

**Sintoma:** só 1 execução real em 01/10 (09:00, SAP ainda não logado). Demais horas (09:36, 13:34, 13:36, 14:36, 17:07, 17:36) = evento 153 do Agendador "missed its schedule". Sem e-mail de aviso das 18h.
**Causa:** tarefa `Verificacao_ZLFIB_Duplicidade_Mensal` com `StartWhenAvailable=False`; PC bloqueado/suspenso nesses horários (hipótese, o log não diz o motivo).
**Correção:** `StartWhenAvailable=True` aplicado e conferido em 2026-10-02 (gatilhos e .bat inalterados).
**Recuperação:** checagem de Setembro/2026 rodada manualmente em 02/10 — 0 duplicidades (arquivo ..._v8.xlsx); estado gravado em data/processed/zlfib_mensal_estado.json.
**Lição:** tarefa que depende de PC ligado/logado precisa de StartWhenAvailable; conferir o histórico do Agendador (eventos 153) quando uma rotina "não rodou".
