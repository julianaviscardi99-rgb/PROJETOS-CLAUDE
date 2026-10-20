# 2026-10-20 — TLM MP27: causa do #VALUE! (HC R09/R07) e como o Actual lança aprendizes

## #VALUE! no Resultado MP27 (N124 / N13) — CAUSA CONFIRMADA
- Teste com recálculo completo (instância separada, somente leitura):
  - só o Resultado aberto: N13/N50/N124 = #REF!
  - + `Base Total MP27.xlsx`: custos OK (N50 −50.860), mas N13 e N124 = #VALUE!
  - + `2026\1. FCST\R09\HC R09 TLM.xlsx` e `R07\HC R07.xlsx`: **tudo OK, PBIT 1.996**
- Ou seja: precisa de 3 arquivos abertos (Base Total MP27 + HC R09 TLM + HC R07). O HC alimenta o faturamento
  (ex. CPFab F148 = custo do trabalho / HC da linha 19); SUMIF em pasta fechada devolve #VALUE!.
- Só acontece quando o Excel recalcula tudo ao abrir (às vezes abre com os valores gravados e mostra 1.996).
- 16 células de comparação (BAQA/Fitted/FSCheck) leem o Resultado R09 e seguem com erro mesmo com tudo aberto — antigas, sem efeito no P&L.

## Aprendizes no Actual 2026 (Base_KSB1)
- Classe **N14211S001 "COM DESP Mão de obra"** → gestorial **4236020 Prestados - Serviços** → Desc.Resultado
  **"Prestados - Serviços"**, Fixo (provisão em N14211S003 → 4248000, também "Prestados - Serviços").
- Fornecedores: Associação de Ensino Social e Rede Nacional de Aprendizagem (texto "Aprendiz TLM Matriz/SBC").
- ~R$ 60 mil/mês (Jan–Set 2026 ≈ R$ 608 mil líquido), coerente com os R$ 695.940/ano da Iana no MP27.
- As classes da Iana (N14215M001 → 4248000 Mão de obra temporária; N14219M000 → 4236020) são classes MFG (M),
  não existem na base 2026 da TLM (que usa S) — por isso ficaram sem Desc.Resultado.

## Volume oficial x Ecolabel / Outras Atividades / m²
- Os Volumes Movimentados MP27 (16/10 e 20/10 1148) NÃO têm Ecolabel, m² nem Outras Atividades — só toneladas.
  Essas linhas do Faturamento seguem como simulação digitada; o volume novo não as afeta.
