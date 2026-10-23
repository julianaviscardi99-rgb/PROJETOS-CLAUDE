"""Gera a copia SEM LINK de um Resumo MP27 (uso: python mp27_resumo_sem_link.py <resumo.xlsx> <resultado.xlsm>).

Derivado de mp27_gerar_sem_link.py (parte do Resumo), com os arquivos por argumento.

Nunca sobrescreve: se o arquivo de saida existir, aborta. Os originais nao sao alterados (abertos somente leitura;
a copia sem link e gravada com SaveAs).
- Resultado: abre com as fontes (para os valores estarem calculados), rompe todos os vinculos externos (viram valores)
  e grava `_sem link.xlsm`.
- Resumo: celulas cuja formula le outro arquivo ou usa INDIRECT viram valor; demais formulas ficam; vinculos rompidos;
  grava `_sem link.xlsx`.
Depois reabre as duas copias sozinhas (sem nenhuma fonte), confere 0 vinculos e compara valores.
"""
import os
import sys
import win32com.client as w

REDE = "\\\\SRV054-01BR.group.pirelli.com\\DAC_TLM"
PASTA = os.path.join(REDE, "MP27", "1ª Versão MP27 vs R09", "TLM_Automation")
A26 = os.path.join(REDE, "2026")
FONTES = [os.path.join(A26, "Base Total_Act_Flash_Fct.xlsx"), os.path.join(A26, "Base_KSB1.xlsx"),
          os.path.join(A26, "1. FCST", "R09", "HC R09 TLM.xlsx"), os.path.join(A26, "1. FCST", "R07", "HC R07.xlsx"),
          os.path.join(A26, "3. Actual", "Act_09.Setembro", "Headcount TLM - 2026.09.xlsx"),
          os.path.join(PASTA, "Base Total MP27.xlsx"), os.path.join(A26, "1. FCST", "R09", "Resultado TLM - R09 '26.xlsm")]
RSM, RES = sys.argv[1], sys.argv[2]
RSM_OUT = RSM[:-5] + "_sem link.xlsx"
ABAS_RSM = ["Resumo FY", "Resumo Month", "Resumo YTD", "Resumo YTG", "Resumo Month old", "P&L - TLM", "Resumo de Resultado",
            "Gui", "Planilha1", "Provisões"]
if os.path.exists(RSM_OUT):
    sys.exit("saida ja existe: " + RSM_OUT)


def novo_excel():
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    return xl


def erro(x):
    return isinstance(x, int) and x < -2146826000


def snapshot(wb, abas, rng_fn=None):
    out = {}
    for n in abas:
        v = wb.Worksheets(n).UsedRange.Value
        out[n] = v if isinstance(v, tuple) else ((v,),)
    return out


def comparar(a, b):
    dif, tot, err = 0, 0, 0
    for n in a:
        for la, lb in zip(a[n], b[n]):
            for x, y in zip(la, lb):
                tot += 1
                if erro(y):
                    err += 1
                if x != y and not (isinstance(x, float) and isinstance(y, float) and abs(x - y) < 1e-6):
                    dif += 1
    return dif, tot, err


def quebrar(wb):
    n = 0
    for tipo in (1,):
        for s in list(wb.LinkSources(tipo) or []):
            wb.BreakLink(s, tipo)
            n += 1
    return n


# ---------------------------------------------------------------- Resumo
xl = novo_excel()
try:
    ab = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in FONTES]
    res = xl.Workbooks.Open(RES, UpdateLinks=0, ReadOnly=True)
    wb = xl.Workbooks.Open(RSM, UpdateLinks=0, ReadOnly=True)
    xl.CalculateFull()
    abas = [n for n in ABAS_RSM if n in [s.Name for s in wb.Worksheets]]
    antes = snapshot(wb, abas)
    trocadas = 0
    for n in abas:
        ws = wb.Worksheets(n)
        ur = ws.UsedRange
        fm = ur.Formula
        if not isinstance(fm, tuple):
            continue
        r0, c0 = ur.Row, ur.Column
        for i, lin in enumerate(fm):
            for j, f in enumerate(lin):
                if isinstance(f, str) and f.startswith("=") and ("[" in f or "INDIRECT" in f.upper()):
                    c = ws.Cells(r0 + i, c0 + j)
                    c.Value = c.Value
                    trocadas += 1
    quebrar(wb)
    print("Resumo: formulas externas/INDIRECT viradas valor:", trocadas)
    xl.CalculateFull()
    depois = snapshot(wb, abas)
    print("Resumo: dif (antes x depois) =", comparar(antes, depois))
    wb.SaveAs(RSM_OUT, FileFormat=51)
    wb.Close(False)
    res.Close(False)
    for a in ab:
        a.Close(False)
finally:
    xl.Quit()

xl = novo_excel()
try:
    wb = xl.Workbooks.Open(RSM_OUT, UpdateLinks=0, ReadOnly=True)
    print("Resumo sem link, aberto sozinho: vinculos =", wb.LinkSources(1))
    xl.CalculateFull()
    fy = wb.Worksheets("Resumo FY")
    print("  FY C6 %s | G6 %s | C27 %s | G27 %s | K27 %s" % tuple(fy.Range(a).Value for a in ("C6", "G6", "C27", "G27", "K27")))
    print("  erros por aba:", {n: sum(1 for l in snapshot(wb, [n])[n] for x in l if erro(x)) for n in
                                ("Resumo FY", "Resumo Month", "Resumo YTD", "Resumo YTG", "P&L - TLM", "Resumo de Resultado")})
    wb.Close(False)
finally:
    xl.Quit()
