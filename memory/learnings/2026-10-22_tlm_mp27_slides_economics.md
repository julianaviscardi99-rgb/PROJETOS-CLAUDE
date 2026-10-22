# 2026-10-22 - TLM MP27: slides Economics (total + por unidade)

- Cache do `Resultado TLM - MP27 vs R09_v2.xlsm` salvo tem #REF!/#VALUE! nos custos (salvo sem fontes abertas). Para ler os numeros: recalcular em instancia separada do Excel (DispatchEx, ReadOnly, UpdateLinks=0) com Base Total MP27 + HC R09 + HC R07 abertos (script `scripts/tlm/mp27_extrair_resumo_slide.py`). Abrir tambem o Resultado R09 faz as colunas R09 virarem #REF! - ler o R09 direto do cache do proprio arquivo.
- R09: net sales do cache = #VALUE! (HC). Total R09 = 93.023,973 (Resumo FY G6); fabricas = residuo. PBIT R09 3.437,1 fecha com soma das unidades.
- Estrutura: P&L por aba (BATotal, CPLog, FSLog, CBTotal, SBLog, GVLog, P&L Outras Atividades, P&L Fab). PBIT = net + var + fix; D&A (linhas 42+106) de todo o TLM esta em BA.
- Gerador: `scripts/tlm/gerar_ppt_tlm_mp27.py` (template = deck do CP; bridge nativo com base zero, aceita PBIT negativo).
