"""Cria `Resumo Resultado TLM - MP27 vs R09_v2.xlsx` a partir do Resumo atual, trocando o link do Resultado para o _v2.
O Resultado v2 e as fontes ficam abertos (somente leitura, instancia separada do Excel) para o Resumo gravar valores
calculados; o v2 nao e salvo. Nunca sobrescreve: se a saida existir, aborta."""
import os, shutil, sys
import win32com.client as w

REDE = "\\\\SRV054-01BR.group.pirelli.com\\DAC_TLM"
PASTA = os.path.join(REDE, "MP27", "1ª Versão MP27 vs R09", "TLM_Automation")
R26 = os.path.join(REDE, "2026", "1. FCST")
ORIG = os.path.join(PASTA, "Resumo Resultado TLM - MP27 vs R09.xlsx")
SAIDA = os.path.join(PASTA, "Resumo Resultado TLM - MP27 vs R09_v2.xlsx")
V2 = os.path.join(PASTA, "Resultado TLM - MP27 vs R09_v2.xlsm")
V1 = "Resultado TLM - MP27 vs R09.xlsm"
FONTES = [os.path.join(PASTA, "Base Total MP27.xlsx"), os.path.join(R26, "R09", "HC R09 TLM.xlsx"),
          os.path.join(R26, "R07", "HC R07.xlsx"),
          os.path.join(REDE, "2026", "Base_KSB1.xlsx"),
          os.path.join(REDE, "2026", "3. Actual", "Act_09.Setembro", "Headcount TLM - 2026.09.xlsx")]
if os.path.exists(SAIDA):
    sys.exit("saida ja existe: " + SAIDA)
shutil.copy2(ORIG, SAIDA)
xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
try:
    ab = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in FONTES]
    v2 = xl.Workbooks.Open(V2, UpdateLinks=0, ReadOnly=True)
    xl.CalculateFull()
    wb = xl.Workbooks.Open(SAIDA, UpdateLinks=0)
    antigo = [s for s in wb.LinkSources(1) if os.path.basename(str(s)).lower() == V1.lower()]
    wb.ChangeLink(antigo[0], V2, 1)
    for nome in ("Resumo YTD", "Resumo YTG", "Resumo Month", "Resumo FY"):
        wb.Worksheets(nome).UsedRange.Replace("[%s]" % V1, "[%s]" % os.path.basename(V2), 2)
    xl.CalculateFull()
    fy = wb.Worksheets("Resumo FY")
    for k in ("C6", "C11", "C19", "C27", "G6", "G27", "I27", "K27"):
        print(k, fy.Range(k).Value)
    erros = {}
    for n in ("Resumo FY", "Resumo Month", "Resumo YTD", "Resumo YTG", "P&L - TLM", "Resumo de Resultado"):
        v = wb.Worksheets(n).UsedRange.Value
        erros[n] = sum(1 for l in v for x in l if isinstance(x, int) and x < -2146826000)
    print("erros", erros)
    print("links", [os.path.basename(str(s)) for s in wb.LinkSources(1)])
    wb.Save()
    wb.Close(False)
    v2.Close(False)
    for a in ab:
        a.Close(False)
finally:
    xl.Quit()
