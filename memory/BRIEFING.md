# BRIEFING — Documento Vivo da Sessão
> Atualizado por Claude em tempo real. Lido no início de cada sessão.
> Manter apenas as últimas 2 sessões inline — sessões mais antigas vão para long_term/.

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
