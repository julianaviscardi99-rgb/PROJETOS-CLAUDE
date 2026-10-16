# Erro: contagem de erros célula a célula via COM travou o Resultado MP27 por 1h40

**Data:** 08/10/2026

**Tarefa:** remontar o `Resultado TLM - MP27 vs R09.xlsm` (`scripts/tlm/mp27_resultado.py`) para ler o faturamento novo.

**Erro:** o script ficou 1h40 rodando depois do `CalculateFull`, com Excel e Python a 100 % de CPU, sem gravar. Tive de
encerrar os dois processos (pelo PID) e rodar de novo.

**Causa:** a conferência final `sum(1 for c in ws.UsedRange.Cells if ... c.Value ...)` faz uma ida e volta Python × Excel
por célula. Nas abas de WH e no P&L do Resultado o UsedRange tem centenas de milhares de células.

**Solução:** `contar_erros(ws)` lê `ws.UsedRange.Value` de uma vez (uma chamada COM, devolve tupla de tuplas) e conta os
códigos de erro (inteiros < -2146826000) em Python. Aplicado em `mp27_resultado.py`, `mp27_resumo.py` e `mp27_tarifa.py`.

**Evitar:** nunca iterar `Range.Cells` via win32com para ler valores em massa — ler `Range.Value` (ou `Value2`) inteiro e
percorrer em Python. Vale para qualquer script que confira planilha grande pelo Excel.
