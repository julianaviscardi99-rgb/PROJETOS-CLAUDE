# -*- coding: utf-8 -*-
"""mp27_recalcular_resultado.py -- gera `Resultado TLM - MP27 vs R09_vN.xlsm` recalculado com a Base Total MP27 atual.

USO: python scripts/tlm/mp27_recalcular_resultado.py <resultado_origem.xlsm>
     (a origem NAO e tocada; a copia ganha o proximo _vN livre e e aberta/gravada em instancia separada do Excel
     com as fontes abertas -- receita que evita #VALUE! no HC do R09, ver learnings 2026-10-22)
"""
import os
import re
import shutil
import sys

import win32com.client as w

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mp27_draft as dr                                                # noqa: E402

A26 = os.path.join(dr.REDE, "2026")
FONTES = [os.path.join(A26, "Base Total_Act_Flash_Fct.xlsx"), os.path.join(A26, "Base_KSB1.xlsx"),
          os.path.join(A26, "1. FCST", "R09", "HC R09 TLM.xlsx"), os.path.join(A26, "1. FCST", "R07", "HC R07.xlsx"),
          os.path.join(A26, "3. Actual", "Act_09.Setembro", "Headcount TLM - 2026.09.xlsx"),
          os.path.join(dr.PASTA_SAIDA, "Base Total MP27.xlsx"),
          os.path.join(A26, "1. FCST", "R09", "Resultado TLM - R09 '26.xlsm")]
CHECKS = {"Fat. liquido": "N13", "PBIT": "N142", "R9 fat. liquido": "AB13", "R9 PBIT": "AB142", "HC R9": "AB19"}


def proximo(origem):
    base = re.sub(r"_v\d+(?=\.xlsm$)", "", origem)[:-5]
    n = 2
    while os.path.exists("%s_v%d.xlsm" % (base, n)):
        n += 1
    return "%s_v%d.xlsm" % (base, n)


def main():
    origem = sys.argv[1]
    destino = proximo(origem)
    shutil.copy2(origem, destino)
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        abertos = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in FONTES]
        wb = xl.Workbooks.Open(destino, UpdateLinks=0)
        ws = wb.Worksheets("P&L - TLM")
        antes = {k: ws.Range(c).Value for k, c in CHECKS.items()}
        xl.CalculateFull()
        depois = {k: ws.Range(c).Value for k, c in CHECKS.items()}
        wb.Save()
        wb.Close(False)
        for a in abertos:
            a.Close(False)
    finally:
        xl.Quit()
    print("Saida:", destino)
    for k in CHECKS:
        print("  %-18s antes %s  depois %s" % (k, antes[k], depois[k]))


if __name__ == "__main__":
    main()
