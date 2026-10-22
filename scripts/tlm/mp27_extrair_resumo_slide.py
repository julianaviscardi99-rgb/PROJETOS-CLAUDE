"""Recalcula (somente leitura) o Resultado TLM MP27 v2 numa instância separada do Excel
e extrai o P&L por unidade (colunas A, N=MP27, AB=R09, AP=MP26) para JSON.
Uso: python mp27_extrair_resumo_slide.py <resultado.xlsm> <saida.json> [fonte1 fonte2 ...]"""
import sys, json, os
import win32com.client as w

ABAS = ["P&L - TLM", "BATotal", "SBLog", "CPLog", "CBTotal", "GVLog", "FSLog",
        "P&L Outras Atividades", "SBTransp", "CPEng", "BAMP", "Hevea", "BAQA", "Campneus", "CPCheck",
        "P&L Fab", "CPFab", "FSFab", "P&L Log"]
resultado, saida, fontes = sys.argv[1], sys.argv[2], sys.argv[3:]
xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
try:
    abertos = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in fontes]
    wb = xl.Workbooks.Open(resultado, UpdateLinks=0, ReadOnly=True)
    xl.CalculateFull()
    out = {}
    for nome in ABAS:
        ws = wb.Worksheets(nome)
        rng = ws.Range("A1:AP150").Value
        out[nome] = [[r[0], r[13], r[27], r[41]] for r in rng]  # A, N, AB, AP
    json.dump(out, open(saida, "w", encoding="utf-8"), ensure_ascii=False, default=str)
    wb.Close(False)
    for a in abertos: a.Close(False)
finally:
    xl.Quit()
