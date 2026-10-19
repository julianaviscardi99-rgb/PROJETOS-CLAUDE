# Auditoria Labour Cost FITTED MP27 (v1 x v2 x R9) — 2026-10-19

- Script: `scripts/fitted/auditoria_labour_mp27.py` (args: v1, v2, r9, MENS MP27 v3, P&L R9, saida). Saída local: `data/processed/Auditoria Labour FITTED MP27 v1 v2 R9 - 20261019_v2.xlsx` (7 abas; "Achados" é a primeira). Não copiado para a rede (aguarda OK dela).
- **Conciliações que fecham:** MENS MP27 v3 = Labour v2 (link externo `[3]Sheet1`, diferença 0,0; v1 daria até 40 mil/mês); Labour R9 = P&L R9 em set–dez (Var e Fixo). Jan–ago do R9 não fecha (+268 mil, PLR jan/26 em CCs 8231/8269/8292 sem unidade = ENC).
- **Detalhe de Despesas não tem MO** (só DG): conciliar MO com P&L/MENS, não com o Detalhe.
- **Provisão de férias do horista (v2)**: IBI/SJP/GOI ≈5,5% de (salário+HE) = metade do estatutário (11,1%); RES 11,4% (correto). Encargo s/ férias (4201170P) ficou em ~89% da provisão (esperado ~35–40%). R9 realizado jan–ago ≈12–13%; R9 forecast set–dez 6–9% → origem provável no forecast. Exposição ≈ R$ 0,95 mi/ano (hipótese; confirmar com RH).
- MO "variável" do MP27 tem elasticidade 0,01–0,12 ao volume (peças da MENS): praticamente fixa.
- SJP CC 8289 horista fixo +35% vs run-rate set–dez/26 (+R$ 460 mil/ano). RES CC 8349 muda de Fixo (R9) para Variável (MP27), R$ 873 mil.
- Reajuste no arquivo: mensalistas +4,8% em jun/27 (GER/GOI/IBI/RES), SJP +3,1% em out/27; em jun as provisões têm retroativo (5 meses × Δ × %). Premissa "~7%" do deck não aparece.
- Rateio GER no MP27 (aba Rateio Fixo da MENS) é digitado (22/52/22,5/3,5) e difere da média calculada na mesma aba (20,6/55,6/18,0/5,8).
- Falso alarme meu: rateio da Sheet1 soma 100% (0,035 e 0,225 aparecem como 0,04/0,23 arredondados).
- Gotcha técnico: `python -I` ignora PYTHONIOENCODING; usar `sys.stdout.reconfigure(encoding='utf-8')` para imprimir "Δ".
