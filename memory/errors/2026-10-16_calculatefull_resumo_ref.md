# 2026-10-16 — CalculateFull com o Resultado aberto read-only gerou #REF! no Resumo

- O que aconteceu: ao criar a v2 do Resumo R10 abri o Resultado v2 (read-only) + Resumo v2, `ChangeLink` e `xl.CalculateFull()`. O CalculateFull recalculou também o Resultado, cujas SUMIFS dependem de ~25 arquivos fechados → viraram erro em memória e o Resumo gravou #REF! (Resumo FY C6 etc.).
- Correção: refiz a cópia; no Resultado aberto `sheet.EnableCalculation = False` em todas as abas e usei `xl.Calculate()` (não Full). Resumo FY voltou a 92.886 / 4.070.
- Regra: com Resultado TLM de fonte, NUNCA CalculateFull sem abrir todas as fontes de link; para só repuxar valores, desligar o cálculo do Resultado e usar Calculate.
