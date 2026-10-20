# 2026-10-20 — TLM MP27: volume oficial (Jefferson), conciliação de custos e erro de -390 herdado do Flash

## Feito
- Gap de R$ 560 mil na MO: era número errado dela (53.950.513,03). Pyplan 16/10 18:03 = arquivo da rede = 53.390.513,03 (diferença zero); ela confirmou. Nenhuma versão recebida soma 53,95.
- Volume oficial do Jefferson: 20/10 11:00 e 11:48 (Revenda). A aba virou `03. Fluxos_MP27_Ofic` e os rótulos ganharam sufixo `_Ofic`; os arquivos na pasta Volume foram renomeados `... 1100.xlsm` e `... 1148.xlsm`. `mp27_faturamento_link_volume.py` agora aceita aba e sufixo (args 4 e 5).
- `Faturamento MP27 - base R09 - 20261020_v2.xlsx` (a partir da v4_L, ligado ao 1148). O `mp27_resultado.py` escolhe o faturamento pelo ÚLTIMO NOME (`_ultimo`), por isso o nome precisa ordenar depois (`_v2` > `.xlsx`). Tons: BA Car 29,085 / CP Car 134,080 / FS 81,818 mil t; fat. líq. 86.972 igual (tarifa a custo: tons não mexem na receita).
- Base Total, Resultado e Resumo regerados (13:26–13:42). Versões anteriores em `_Versoes_anteriores\20261019_*`.
- Conciliação de custos: originais Labour 53.390.513,03 + áreas MP27 Inf 29.351.433,93 + Iana 2.948.480,30 = 85.690.427; base única 85.684.627 (R$ 5.800 do RH de Cabreúva/BAMP fora do perímetro); Base Total 85.684.628 (arredondamento).
- Evidências e rascunhos de e-mail (Outlook, resposta ao Jefferson 16/10 07:27): sobreposição CC x classe (5 combos; ajuste máximo ~R$ 28,6 mil) e IFRS16 (blocos +X/-X/-X/+Y que fecham perto do R09, provavelmente intencionais). Scripts `mp27_evidencia_sobreposicao.py` e `mp27_evidencia_ifrs16.py`; saídas em `data/processed/tlm/mp27/`. Ela enviou os dois e aguarda a resposta.

## ERRO ENCONTRADO (a corrigir)
- `BAADM!J105` do Resultado: `=IF(J$6="Flash_KSB1", ... , -SUMIFS(...Intermediaria...)/1000)-390`. O `-390` é ajuste digitado do Flash Set herdado do template. No MP27 não há Outras Despesas em BAADM na Base Total => R$ 390 mil de custo falso; PBIT 1.605 -> ~1.995 mil.
- Como achei: conciliando Base Total x P&L por Desc.Resultado e Variabilidade; só "Outras Despesas" fixas divergia (exatamente 390,0). Setembro J105 = -373,4 vs ~16,6 nos demais meses.
- A varredura por "ajuste no fim da fórmula" também apontou `BALog!D24` (fórmula termina em constante, valor -238,1 — verificar) e CPFab linhas 150/151 (números 4/7/27, parecem HC).
- Os 5 erros #VALUE! do P&L (O5, P5, BY5, O6, BY6) são fórmulas de diferença (`=J5-X5`) sobre os textos "MP27"/"R09" das linhas 5/6: cosméticos, sem efeito nos valores.

## Fora do Resultado (por desenho)
- 19 linhas de aprendiz da Iana (classes N14215M001/N14219M000, R$ 695.940): estão na Base Total, sem Desc.Resultado => sem linha no P&L. O Labour do Pyplan não tem aprendizes (só DIRETOR/GERENTE/HORISTA/MENSALISTA/SENIOR), então não é duplicidade. Classes não existem na base 2026.
- Previdência V 4247100 (R$ 12,4 mil): decisão dela de manter "como veio".

## Lição
- Ao regenerar o Resultado a partir da cópia do Flash, varrer fórmulas por ajustes digitados (`)-390`) e números soltos antes de entregar; a conciliação base x P&L por linha é o que pega.
