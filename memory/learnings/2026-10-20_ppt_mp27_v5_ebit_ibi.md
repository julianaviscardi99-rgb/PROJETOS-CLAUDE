# 2026-10-20 — PPT MP27 Fitted v5: novo EBIT de Ibirité

- Pedido: atualizar valores/comentários com o novo EBIT de IBI **sem mudar layout**; **não mexer nos slides 2, 3 e 4** (feitos por ela).
- Fontes (20/10 manhã): Base Book (GFU_DAC, 08:26), MENS v3 teste daniel (08:20; preço Fiat 15,29 = 12,97+0,6+1,55+0,17), Rentabilidade IBI v3 (EO_FITTED, 08:24; ela digitou realocação +300 Fiat / −150 Iveco / −150 CNH).
- Números: IBI EBIT 6,565 (ROS 13,9%; era 6,08); Brasil 11,455 (ROS 13,7%, −11,9% vs R9); reajuste Fiat +16,2% = 6,14 mi (cobre 102% de 3º turno + ANTT); sem reajuste Brasil ~5,3.
- Script: `scripts/fitted/atualizar_ppt_mp27_ebit_ibi.py` — edição cirúrgica (troca de texto por run, barras do bridge reposicionadas com o mesmo algoritmo da rev2, pizza via replace_data). **Rótulos da pizza são texto fixo no XML do gráfico (c:dLbl/c:tx) — atualizar à parte.**
- Na tabela de clientes, Fixed Cost inclui a realocação de 0,3 (senão a tabela não fecha no EBIT).
- Slides 2–4 NÃO atualizados a pedido; divergem do Base Book: slide 3 IBI 6,8 / Δ −0,6 (Base Book 6,6 / −0,8), 4 unidades Δ −1,3/−10,2% (é −1,4/−10,6%), 86% (é 84%), ROS 13,8% (13,7%); slide 4 NS 83,2/+8,6% (83,7/+9,2%); slide 2 é a versão antiga (11,0).
- Saída: rede `MP 2027\Fitted_Units_Meeting_Budget'27_v5.pptx` + data/processed.
