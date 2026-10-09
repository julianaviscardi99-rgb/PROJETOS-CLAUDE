# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

---
## SESSÃO 2026-10-09 (janela "Circuito Panamericano: faturamento e B.U") — Onboarding do Circuito + classificação de custos

### Feito
- **Domínio explicado e registrado** em `ontology/circuito_panamericano.json`: site/estrutura, nomes (CP / Elias Fausto), empresa SAP **2281**, gerente **Camila Maluf**, faturamento = **intercompany** (R&D + demais empresas do grupo, ex.: Pirelli Pneus) e **terceiros** (eventos, pistas, antena Alfasite, valor fixo mensal). Faturamento chega por e-mail com **recibos (sem NF)**; Juliana cria arquivo e abre chamado na Contabilidade. Só **custo fixo**.
- **Foco do projeto:** acelerar a classificação dos custos da KSB1 (horas manuais por mês). Arquivos em `...\Reporting ACT_FCST_MP Cons. e Ind\Circuito Panamericano\<ano>\<MM - Mes>\<02. Flash | 03. Actual>` (classificação manual em `KSB1 - Jan-Dez_.xlsx`: aba por mês + Parâmetros + Check; Jan–Jul na coluna N verde, Ago+ na coluna "Classificação").
- **Script** `scripts/sap/circuito_panamericano/classificar_ksb1.py` (+ `regras_manuais.json`: Optimus = I.T, não predial): sugere classificação (histórico + Parâmetros); teste: 94–99% de acerto na confiança alta, 61–90% de cobertura. Gera xlsx novo em `data/processed/`. Juliana viu e **não gostou do formato** (quer manter o dela) — **NÃO está ligado ao cockpit ainda**.
- **Cockpit do Circuito (janela separada; não importa nada da Fitted)** — `scripts/sap/circuito_panamericano/cockpit_circuito_gui.py` + `ksb1_cp.py`. Layout = **mesmo formato da Fitted** (painel Ano/Mês/Ciclo, abas escuras, pneu girando, barra de progresso, LOG, spinner), título "COCKPIT FECHAMENTO CIRCUITO", foto da pista no cabeçalho (pedido dela). **Passo 1 = Extração KSB1** pronto e **validado ao vivo** (Set/Actual): área 2281, mês inteiro, variante `/DESPFITTED`, sem grupo de centros de custo/agrupamento; salva em `<ano>\<MM - Mes>\<02. Flash|03. Actual>\` com nome no **padrão da Fitted**: `KSB1 - Circuito Panamericano MM.AAAA - <Ciclo>.XLSX` (`_v2/_v3` se já existir). Comparado ao arquivo manual dela: 172x19, soma Valor/MR idêntica (R$ 1.487.805,12). Passos ②–⑤ **em branco** (ela direciona um a um).
- **Popup "Segurança SAPGUI"**: o 1º teste travou (pasta temporária nova do CP). Solução: exportar sempre pela pasta `Temporario` de **2026** da Fitted (`GFU_DAC\...\Resultados Fitted\2026\00.Extração Base KSB1\Temporario`), já autorizada — e **a pedido dela vale também em 2027+** (nos dois cockpits; `atualizar_ksb1_gui.py` da Fitted ajustado em uma linha).
- **Cópias de rede atualizadas (autorizado):** cockpit da Fitted (`_Cockpit_KSB1`: `scripts\sap\fitted_units` + `ontology` sem o json do circuito + requirements) e cockpit do CP (`Circuito Panamericano\Extração SAP\_Cockpit_CP` + atalho `Circuito Panamericano.lnk`). **Nenhuma sincroniza sozinha** — re-copiar (robocopy) após mudanças.
- **Limpeza (autorizada):** arquivo do teste renomeado para o padrão novo em `2026\09 - Sep\03. Actual`; `_v2` duplicado e pasta vazia `Extração SAP\Temporario` apagados. Original dela `Base KSB1 - 09.2026_Actual.xlsx` intacto.

### Pendente / próximos passos (segunda-feira 12/10)
1. **Passo 2 do cockpit do CP** — a Juliana define. Provável: classificação dos custos da KSB1 na mensalização (`classificar_ksb1.py` como base), **mantendo o formato/arquivo dela** (não criar layout novo).
2. Faturamento por recibos (e-mail da Camila → arquivo → chamado na Contabilidade) e mensalização: ainda a explicar/automatizar. Pendências de negócio na ontologia (`faturamento_tipos`, `fluxo_faturamento`).
3. **Regras dela:** não alterar arquivos da área na rede sem confirmar (só leitura / saída em `data/processed/` ou arquivo novo versionado); apagar/renomear só com confirmação.
4. Opcional: aprender com as decisões dela nos "pendentes" do classificador; revisar a aba Parâmetros (Optimus ainda como "Predial" lá — ela decide se atualiza).
5. Dica: o gravador de script do SAP (Alt+F12) mostra os cliques dela para novos passos de SAP.

---
## SESSÃO 2026-10-09 — Revisão da MENS v2/v3 "teste daniel" + Bridge de EBIT MP'27 linkado

### Feito
- **Revisão da `MP 2027\MENS FITTED MP27_v2_teste daniel.xlsx`** (só leitura). Erros achados: IBI linha 77 (materiais CNH usava o volume da Iveco, `E54`); MO com +7% colado em SJP/IBI/GO mas não na RES. Ela corrigiu os dois na **`MENS FITTED MP27_v3_teste daniel.xlsx`**.
- **Revisão 100% da v3:** 0 erros de célula, fórmulas consistentes nos 12 meses, TOTAL = soma das unidades, conferências OK. Volumes agora vêm de `BU FITTED\PO's\Volume\MP'27.xlsx`. **EBIT MP'27 = −828** (SJP 1.492, IBI −1.667, GO 289, RES −941).
- **Explicações dela (premissas, não são erros):** RES 7,3 = preço de lista (a unidade identifica os ganhos); frete da IBI subiu pela nova regra da ANTT (−1,5707 → −1,9763/pç); o frete da GO estava superestimado (agora −58×1,03/mês); ROS esperado ajustado para 13,8%; DU da GO em outubro = férias coletivas da Jeep; MO = estimativa até o RH; rateio fixo só depois de quadrar o faturamento; vínculos de volume mantidos de propósito; ela vai apagar a célula solta `GO!AC17`.
- **Bridge criado:** `MP 2027\Bridge EBIT MP27 - linkado v3 teste daniel_v2.xlsx` (a v1 sem a Jaguar −550 ficou na pasta). Abas: Bridge (vs R9 e vs MP'26: Volume, Mix, Preço, Jaguar MD, Jaguar rateio fixo, ANTT, Demais var, Demais fixos, Outros), Premissas (amarelo = digitar), Detalhe custos. Vínculo externo real e relativo (mesma pasta), testado no Excel: 0 erros, conferências OK. Script: `scripts/sap/fitted_units/mp2027/gerar_bridge_ebit_linkado.py`. Learning: `memory/learnings/2026-10-09_bridge_ebit_mp27_linkado.md`.
- **Números (R$ mil):** vs R9: base 12.813 → −828 (Δ −13.641); Jaguar −550 (informado por ela; margem direta −522,5 + −27,5 na linha de rateio); ANTT −1.073,8. vs MP'26: base 9.816 (Δ −10.644); Jaguar −526,8.

### PENDENTE
1. Ela inserir o rateio de custo fixo de Ibirité da Jaguar (`Premissas!D14/E14`; hoje D14 = −27,5 = −550 informado − margem direta).
2. **R9 da aba TOTAL 188,2 maior que a soma das unidades** (`TOTAL!L47`, agosto, digitado) — corrigir na fonte ou explicar.
3. MP'26 da RES sem detalhe (só EBIT 337,5 na coluna V) → bridge vs MP'26 da RES fica distorcido (Preço = receita inteira, Outros −337,5).
4. Confirmar se o frete anterior da IBI (−1,5707, `Premissas!D6`) vale para as bases R9/MP'26.
5. Itens das sessões 07/10 e 08/10 seguem abertos (Faturamento Actual I25, IFRS16/MAPPING WRONG, linhas de agosto na KSB1, créditos da IBI, MDO `#REF!`) — ver arquivo `long_term/2026-10-09_briefing_arquivo_sessoes_2026-10-06_e_07.md`.

### NOTA DE AMBIENTE (09/10 tarde) — notebook lento/travando (diagnóstico, nada alterado)
- Hardware OK (SSD saudável, 211 GB livres, RAM 66%). Ela usa **Wi-Fi** (sinal 82%, estável: 9 quedas em 7 dias).
- **Placa Ethernet com link oscilando:** 223 eventos "link disconnected" em 7 dias (22 hoje, picos 12h e 14-15h, coincidindo com o travamento do Notes às 14:47). Mesmo usando Wi-Fi, cada oscilação faz o Windows/VPN Check Point reavaliar a rede → apps de rede (Outlook/Teams/SAP) congelam. Sugestão: tirar qualquer cabo/dock da porta de rede ou pedir à TI para desativar o adaptador Ethernet.
- **Teams** é o maior consumidor (webview2 ~900 MB, mais de 2.000 s de CPU); o serviço de câmera do Windows (Camera Frame Server) caiu 13 vezes hoje. Notebook estava ligado havia 2 dias → ela foi reiniciar.
- Throttling de CPU pelo firmware: um episódio só (07/10 17:23).
- `%TEMP%\Outlook Logging` = **14,5 GB** (17.687 .etl desde ago/2025; logging NÃO está ligado nas opções, é log automático do Office). Instaladores antigos do Claude (.msix) ~2,5 GB. Limpeza **aguardando OK dela** (é exclusão).

### Feito (09/10 fim da tarde, Haiku/Sonnet 5.5) — atalhos, relink do P&L Budget27 e pastas do Circuito
- **Locais de rede no "This PC"** (pastas em `%APPDATA%\Microsoft\Windows\Network Shortcuts\<nome>` com `desktop.ini` + `target.lnk`, mesmo formato do "Arquivo de Mensalização"): **CUSTOS MP** → `GFU_DAC\Management Plan\MP 2027` e **GFU_DAC_CUSTOS** → `GFU_DAC\Custos Fitted Units\Resultados Fitted\2026`. Os `.lnk` de teste (área de trabalho e `Management Plan`) foram apagados com OK dela.
- **P&L `MP 2027\P&L Fitted Units_Budget27.xlsx` relinkado (4 links):** Mensalização → `EO_FITTED\...\MP 2027\MENS FITTED MP27_v3_teste daniel.xlsx` (aba TOTAL); Forecast R10/25 → `2026\09 - Sep\09_Sep_Forecast\09_P&L ...Forecast_September_26_.xlsx` (R9); MP25 → `MP 2026\P&L Fitted Units_Budget26_.xlsx`; Dec-24 → Dec-25 (aba Resultado YTD). Removido o `*-1` das 132 fórmulas de custo (linhas 20-24 e 31-36, D:O) porque o MP27 já mantém custos negativos. EBIT jan esperado = −49 (conferido pela TOTAL). Backup: `data/processed/PL_Fitted_Units_Budget27_backup_antes_relink_2026-10-09.xlsx`. Decisão em `DECISOES.md`. Futuras trocas: ela me chama aqui (sem botão, "poucas trocas").
- **Pastas do Circuito Panamericano** (`CONTROLLING\Reporting\Reporting ACT_FCST_MP Cons. e Ind\Circuito Panamericano`): 2026 → criadas `10 - Oct`, `11 - Nov`, `12 - Dec`; 2027 → 12 meses (`01 - Jan`…`12 - Dec`), todos com `01. Forecast / 02. Flash / 03. Actual` (padrão de setembro; sem `Pre Flash`). `2027\FATURAMENTO` criada só com a estrutura (12 meses vazios) + cópia exata de `Faturamento CP - YTD-YTG_.xlsx` como `Faturamento CP - YTD-YTG_2027.xlsx`.

- **Cockpit novo do Circuito Panamericano** (separado da Fitted, nada da Fitted editado): `scripts/sap/circuito_panamericano/cockpit_circuito_gui.py` (+ launcher `.vbs`, `assets/`), cabeçalho com foto da pista do site oficial no lugar do fundo preto, título "CIRCUITO PANAMERICANO". Hoje só a casca, sem etapas. Cópia de rede: `Circuito Panamericano\Extração SAP\_Cockpit_CP\` + atalho `Circuito Panamericano.lnk` (ícone de pneu). Rede não sincroniza sozinha. Próximo passo natural: primeira etapa = classificação de custos da KSB1 (ver memória do foco do Circuito). Decisão em `DECISOES.md`.

- **Handoff (09/10 ~18:40):** a Juliana tem outra janela ("Circuito Panamericano: faturamento e B.U", sessão `projetos-claude-4c`) que já ia montar o cockpit do CP com o Passo 1 = Extração KSB1 (empresa 2281, layouts/filtros e destino ainda em aberto). Mandei a ela, por mensagem entre sessões, o aviso de que a casca já existe (para construir em cima, não duplicar) e que a Fitted não deve ser tocada. Sessão longa (45 ações): contexto salvo.

### PENDENTE (adicional)
0. Cockpit CP: definir as etapas (extração SAP / classificação KSB1 / mensalização) e se quer outra foto/recorte.
1. Abrir o P&L Budget27 no Excel, clicar **Atualizar** nos links e conferir EBIT jan = −49 e demais colunas.
2. Linha 2 do P&L (anos 2026/2025/2025/2024) segue com os anos antigos; linha 42 (Slow Moving) vazia na TOTAL do MP27; se sair MENS nova (v4, sem "teste daniel"), repontar o link 1.
3. `Faturamento CP - YTD-YTG_2027.xlsx` é cópia exata de 2026 (dados 2026, 3 links externos antigos, ano 2026): precisa ser zerado e re-linkado; falta ela dizer quais abas/valores manter.
4. Pastas extras de 2026 (`MP'26`, `Risk Assessment`) não foram copiadas para 2027; `Pre Flash` (existe só em agosto) não foi criada.

---
## SESSÃO 2026-10-08 — Fechamento Fitted Set/2026: quadro amarelo (Flash x Actual), ganho de 84K e KSB1 Jul-Set

### Feito
- **Quadro amarelo (aba Pivot da Base Intermediária Actual):** explicada a lógica (Custos H26/I26 por fórmula; Faturamento manual). Achado: H25 (Faturamento Flash) estava 6.465 e o Flash real é **6.547** (I25 do Flash) → delta de resultado inflado (84 em vez de ~2,3).
- **Código:** `gerar_base_intermediaria.py` → `atualizar_comparacao_flash` agora copia I25 do Flash para H25 do Actual. Testado só em cópia local (H25 = 6.546,87; Resultado Flash 1.305,5; delta 2,3). **Faturamento Actual (I25) segue manual** — fonte ainda não definida. Decisão em `DECISOES.md` (2026-10-08).
- **Ganho Actual vs Flash (R$ 83.893,31, todo em Despesas):** script `comparar_flash_actual.py`; arquivo em `data/processed/fitted_units_despesas/ganho_flash_actual/`. Ponte: créditos PIS/COFINS `_PC` −59.316; deprec. prédios alugados IFRS16 −45.897; conservação +14.966; demais +9.755; veículos −2.002; armazém −1.398; manutenção MG 0 (troca de conta). Learning em `memory/learnings/2026-10-08_ganho_actual_vs_flash_setembro.md`.
- **KSB1 Actual Set só Jul-Ago-Set:** script `filtrar_ksb1_por_meses.py`; criada na rede (`09_Sep_Actual\KSB1 September Actual 2026 - meses 07-09.xlsx`, 29.842 linhas, soma 17.817.647,29; conferido contra o original). Original intacto. Agosto tem 16.038 linhas (o dobro dos outros) — a investigar se duplicado.

### PENDENTE
1. Definir a fonte do Faturamento Actual (I25) para automatizar; hoje igual ao antigo 6.465 (conferir se é o real).
2. Confirmar com a contabilidade depreciação IFRS16 e "MAPPING WRONG" (hipóteses, não confirmadas).
3. Investigar as 16.038 linhas de agosto na KSB1.
4. Nada commitado nesta sessão (scripts novos + memória). Cópia local do Flash/Actual de teste fica só no scratchpad.
5. Itens do MP27/R9 da sessão 07/10 abaixo seguem abertos.

---
## Histórico anterior (2026-09-01 para trás)

Movido para `memory/long_term/2026-09-04_briefing_arquivo_ate_2026-09-01.md` em 2026-09-04
(regra: no máximo 2 sessões inline no BRIEFING). Snapshots automáticos completos também ficam
em `memory/long_term/*_briefing_snapshot.md`.
