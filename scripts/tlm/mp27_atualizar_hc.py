# -*- coding: utf-8 -*-
"""mp27_atualizar_hc.py -- copia o Resultado para o proximo _vN e troca o HC do MP27 pela ultima base do Jefferson.

USO: python scripts/tlm/mp27_atualizar_hc.py <resultado_origem.xlsm>
     le o ultimo `* - Base Geral HC*MP27*.xlsx` de TLM_Automation (versao MP27_vN mais nova), regrava a aba 'HC MP27',
     religa a linha 19 das abas de WH, recalcula (fontes abertas) e grava. A origem nao e tocada.
"""
import os
import sys

import win32com.client as w

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mp27_recalcular_resultado as rc                                 # noqa: E402
import mp27_resultado as res                                           # noqa: E402

CHECKS = {"Fat. liquido": "N13", "PBIT": "N142", "HC MP27": "N19", "R9 PBIT": "AB142", "HC R9": "AB19"}


def main():
    origem = sys.argv[1]
    destino = rc.proximo(origem)
    arq, hc = res.hc_mp()
    import shutil
    shutil.copy2(origem, destino)
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        abertos = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in rc.FONTES]
        wb = xl.Workbooks.Open(destino, UpdateLinks=0)
        ws = wb.Worksheets("P&L - TLM")
        antes = {k: ws.Range(c).Value for k, c in CHECKS.items()}
        res.gravar_hc(wb, arq, hc)
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
        print("  %-14s antes %s  depois %s" % (k, antes[k], depois[k]))


if __name__ == "__main__":
    main()
