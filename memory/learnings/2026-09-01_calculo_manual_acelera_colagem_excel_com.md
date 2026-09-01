# 2026-09-01 — Calculo manual (Excel COM) resolve lentidão real no Passo 3 do cockpit

**Sintoma:** "Atualizar Pivot KSB1" (botão ① do Passo 3, `gerar_ksb1_mensal.py`) levou mais de
12 minutos rodando pra Agosto/2026 Flash — usuária achou que podia estar travado.

**Causa raiz confirmada:** a colagem no BASE_KSB1 é linha a linha, de propósito (proteção
contra um bug real de corrupção `#N/A` do COM/pywin32 achado em 2026-08-21 — colar em bloco
corrompe células aleatoriamente). Mas com o Excel em modo de cálculo automático (padrão),
CADA uma das milhares de escritas linha a linha disparava recálculo do arquivo inteiro
(fórmulas + Pivot Tables) — confirmado no monitor de processo: a instância do Excel travada
tinha ~2510s de CPU acumulado em ~15-20min de execução (CPU pegada o tempo todo).

**Correção aplicada:** `excel.Calculation = -4135` (xlCalculationManual) +
`excel.ScreenUpdating = False` logo antes da colagem/AutoFill, restaurado pra `-4105`
(xlCalculationAutomatic) + `ScreenUpdating = True` logo depois do `CalculateFullRebuild()`
explícito que já existia no fim da rotina. Não muda a granularidade da escrita (continua
linha a linha, mesma proteção contra o bug de corrupção) — só evita recalcular o arquivo
inteiro milhares de vezes em vez de uma.

Aplicada em dois lugares com o mesmo padrão de colagem lenta:
- `gerar_ksb1_mensal.py` (`colar_linhas_e_atualizar_pivots`, botão "① Atualizar Pivot KSB1")
- `gerar_base_intermediaria.py` (`atualizar_base_intermediaria`, botão "③ Finalização")

**Resultado:** confirmado ao vivo pela usuária (Agosto/2026 Flash, mesma operação que levou
12+ min antes) — rodou rápido, sem travar, sem erro. **Não medimos o tempo exato**, mas o
padrão de CPU do novo processo Excel era bem mais baixo (2,3s de CPU nos primeiros 5min, vs
2510s do processo anterior) — consistente com o diagnóstico (o gargalo era recálculo
repetido, não a escrita linha a linha em si).

**Se o mesmo sintoma aparecer de novo em outro script do cockpit** (ex: futura automação com
colagem em massa via COM/Excel): checar primeiro se `Application.Calculation` está em modo
automático durante um loop de escrita — é um ganho grande e de baixo risco, desde que exista
um `CalculateFullRebuild()` explícito garantindo os valores certos antes de ler/salvar.

**Nota operacional:** o fix só entra em vigor numa janela NOVA do cockpit — Python já
carregado numa janela aberta continua com o código antigo em memória até fechar e reabrir.
