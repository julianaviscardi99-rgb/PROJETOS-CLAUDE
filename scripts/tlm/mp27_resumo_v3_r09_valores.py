"""Resumo MP27 _v3: parte do _v2 (link no Resultado v2) e, onde o R09 chega como erro (#VALUE!/#REF!, vindo do HC do R09
no Resultado R09), grava o valor que o Resumo antigo (20/10) tinha para essa celula. So toca celulas de ERRO cuja
formula le outro arquivo (ou INDIRECT); as demais (deltas) recalculam. Nunca sobrescreve."""
import os, shutil, sys
import win32com.client as w

REDE = "\\\\SRV054-01BR.group.pirelli.com\\DAC_TLM"
PASTA = os.path.join(REDE, "MP27", "1\u00aa Vers\u00e3o MP27 vs R09", "TLM_Automation")
R26 = os.path.join(REDE, "2026", "1. FCST")
V2X = os.path.join(PASTA, "Resumo Resultado TLM - MP27 vs R09_v2.xlsx")
V3X = os.path.join(PASTA, "Resumo Resultado TLM - MP27 vs R09_v3.xlsx")
ANT = os.path.join(PASTA, "Resumo Resultado TLM - MP27 vs R09.xlsx")
RES = os.path.join(PASTA, "Resultado TLM - MP27 vs R09_v2.xlsm")
FONTES = [os.path.join(PASTA, "Base Total MP27.xlsx"), os.path.join(R26, "R09", "HC R09 TLM.xlsx"),
          os.path.join(R26, "R07", "HC R07.xlsx"), os.path.join(REDE, "2026", "Base_KSB1.xlsx"),
          os.path.join(REDE, "2026", "3. Actual", "Act_09.Setembro", "Headcount TLM - 2026.09.xlsx")]
ABAS = ["Resumo FY", "Resumo Month", "Resumo YTD", "Resumo YTG", "P&L - TLM", "Resumo de Resultado"]
if os.path.exists(V3X):
    sys.exit("saida ja existe")
shutil.copy2(V2X, V3X)
xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False


def erro(x):
    return isinstance(x, int) and x < -2146826000


try:
    ab = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in FONTES]
    res = xl.Workbooks.Open(RES, UpdateLinks=0, ReadOnly=True)
    old = xl.Workbooks.Open(ANT, UpdateLinks=0, ReadOnly=True)
    wb = xl.Workbooks.Open(V3X, UpdateLinks=0)
    xl.CalculateFull()
    trocadas = {}
    for n in ABAS:
        ws, wo = wb.Worksheets(n), old.Worksheets(n)
        ur = ws.UsedRange
        vals = ur.Value
        ov = wo.Range(ur.Address).Value
        r0, c0 = ur.Row, ur.Column
        cnt = 0
        for i, lin in enumerate(vals):
            for j, x in enumerate(lin):
                if erro(x) and not erro(ov[i][j]) and ov[i][j] is not None:
                    c = ws.Cells(r0 + i, c0 + j)
                    f = str(c.Formula)
                    if "[" in f or "INDIRECT" in f.upper():
                        c.Value = ov[i][j]
                        cnt += 1
        trocadas[n] = cnt
    xl.CalculateFull()
    restam = {}
    for n in ABAS:
        v = wb.Worksheets(n).UsedRange.Value
        restam[n] = sum(1 for l in v for x in l if erro(x))
    fy = wb.Worksheets("Resumo FY")
    print("trocadas", trocadas)
    print("erros restantes", restam)
    for k in ("C6", "G6", "K6", "C27", "G27", "K27", "I27"):
        print(k, fy.Range(k).Value)
    wb.Save()
    wb.Close(False)
    old.Close(False)
    res.Close(False)
    for a in ab:
        a.Close(False)
finally:
    xl.Quit()
