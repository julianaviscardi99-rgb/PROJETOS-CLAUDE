# Erro — leitura da grade da ZLFIB via COM (GetCellValue) retornava dados incorretos

**Contexto:** `scripts/sap/analisar_zlfib_duplicidade.py` lia a grid ALV da ZLFIB linha a linha via `grid.GetCellValue(r, coluna)` (COM/GUI Scripting), em vez de exportar pra arquivo como já era feito na extração da KSB1.

**Sintoma:** rodando SJP (Filial 0031, jan-jul/2026, filtro Direção=Entrada), a leitura por COM relatou corretamente 1.348 linhas de item, mas ao agrupar por `Nr Documento` (DOCNUM) encontrou só **38 notas fiscais únicas**. Comparando com o mesmo período sem o filtro de Direção (3.311 linhas), a leitura por COM encontrou **27 notas únicas** — matematicamente impossível, já que "Entrada" deveria ser um subconjunto de "todas as direções" e não pode ter mais notas únicas que o total.

**Causa raiz confirmada:** exportei a mesma consulta (1.348 linhas, Entrada) usando o menu real do SAP (Lista → Exportar → Planilha eletrônica, mesmo mecanismo de `session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]")` já usado em `atualizar_ksb1_gui.py`) e comparei: a contagem de linhas bateu (1.348), mas o número de `DOCNUM` únicos no arquivo exportado é **565**, não 38. Ou seja, `GetCellValue` estava devolvendo valores errados/repetidos pra coluna `DOCNUM` em grades grandes — provavelmente por causa da virtualização da grid ALV (só as linhas atualmente renderizadas na tela ficam com valor correto e acessível de forma confiável via COM; linhas fora da área visível podem devolver valor de cache/vizinho).

**Impacto:** todos os resultados gerados por esse script antes da correção (SJP, IBI, SOR, GOI — todos relatando "0 duplicidades") **não são confiáveis** e precisam ser refeitos com o método corrigido.

**Correção aplicada:** reescrito `analisar_zlfib_duplicidade.py` pra exportar a grade pra arquivo (`.xlsx`) via menu nativo do SAP, igual à extração da KSB1, e ler o arquivo exportado com `openpyxl` em vez de `GetCellValue` linha a linha. Ver `memory/DECISOES.md` pra decisão completa.

**Licao pra generalizar:** `GetCellValue` em loop (linha a linha) numa `GuiShell`/ALV Grid via SAP GUI Scripting **não é confiável para grades grandes** (testado: falhou já em ~1.300 linhas). Preferir sempre exportação nativa (arquivo/planilha) quando o volume de linhas não for pequeno (dezenas), mesmo que a exportação nativa exija negociar um popup de "Salvar como".
