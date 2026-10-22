"""Teste (somente leitura): com as fontes do R09 abertas (Base Total Act/Flash, KSB1, 3 HC), o Resultado R09 calcula
HC/faturamento sem #VALUE!, e o Resultado MP27 v2 passa a enxergar isso?"""
import os
import win32com.client as w

REDE = "\\\\SRV054-01BR.group.pirelli.com\\DAC_TLM"
PASTA = os.path.join(REDE, "MP27", "1ª Versão MP27 vs R09", "TLM_Automation")
A26 = os.path.join(REDE, "2026")
FONTES = [os.path.join(A26, "Base Total_Act_Flash_Fct.xlsx"), os.path.join(A26, "Base_KSB1.xlsx"),
          os.path.join(A26, "1. FCST", "R09", "HC R09 TLM.xlsx"), os.path.join(A26, "1. FCST", "R07", "HC R07.xlsx"),
          os.path.join(A26, "3. Actual", "Act_09.Setembro", "Headcount TLM - 2026.09.xlsx"),
          os.path.join(PASTA, "Base Total MP27.xlsx")]
R09 = os.path.join(A26, "1. FCST", "R09", "Resultado TLM - R09 '26.xlsm")
V2 = os.path.join(PASTA, "Resultado TLM - MP27 vs R09_v2.xlsm")
xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False


def ver(wb, tag):
    ws = wb.Worksheets("P&L - TLM")
    print(tag, {a: ws.Range(a).Value for a in ("N13", "AB13", "AB19", "T13", "AB21", "AB135", "AB142")})


try:
    ab = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in FONTES]
    r = xl.Workbooks.Open(R09, UpdateLinks=0, ReadOnly=True)
    xl.CalculateFull()
    ver(r, "R09 recalculado:")
    v = xl.Workbooks.Open(V2, UpdateLinks=0, ReadOnly=True)
    xl.CalculateFull()
    ver(v, "v2 (links em cache):")
    for s in v.LinkSources(1):
        if "Resultado TLM - R09" in str(s):
            v.UpdateLink(s, 1)
    xl.CalculateFull()
    ver(v, "v2 apos atualizar vinculo R09:")
    v.Close(False)
    r.Close(False)
    for a in ab:
        a.Close(False)
finally:
    xl.Quit()
