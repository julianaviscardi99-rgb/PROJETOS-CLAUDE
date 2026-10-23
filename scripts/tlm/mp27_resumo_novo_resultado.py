# -*- coding: utf-8 -*-
"""mp27_resumo_novo_resultado.py -- cria o proximo `Resumo Resultado TLM - MP27 vs R09_vN.xlsx` apontando para um Resultado novo.

USO: python scripts/tlm/mp27_resumo_novo_resultado.py <resumo_origem.xlsx> <resultado_novo.xlsm>
     Copia o Resumo (origem intacta), troca o vinculo do Resultado e os textos "[arquivo]" dos INDIRECT das abas
     Resumo FY/Month/YTD/YTG, recalcula com o Resultado novo e as fontes abertas (somente leitura, instancia separada)
     e grava a copia. Celulas ja gravadas como valor (R09 do _v3) nao sao tocadas. Nunca sobrescreve.
"""
import os
import re
import shutil
import sys

import win32com.client as w

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mp27_recalcular_resultado as rc                                 # noqa: E402

ABAS_ERRO = ["Resumo FY", "Resumo Month", "Resumo YTD", "Resumo YTG", "P&L - TLM", "Resumo de Resultado"]


def erro(x):
    return isinstance(x, int) and x < -2146826000


def main():
    origem, resultado = sys.argv[1], sys.argv[2]
    base = re.sub(r"_v\d+(?=\.xlsx$)", "", origem)[:-5]
    n = 2
    while os.path.exists("%s_v%d.xlsx" % (base, n)):
        n += 1
    saida = "%s_v%d.xlsx" % (base, n)
    shutil.copy2(origem, saida)
    novo = os.path.basename(resultado)
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        ab = [xl.Workbooks.Open(f, UpdateLinks=0, ReadOnly=True) for f in rc.FONTES]
        res = xl.Workbooks.Open(resultado, UpdateLinks=0, ReadOnly=True)
        xl.CalculateFull()
        wb = xl.Workbooks.Open(saida, UpdateLinks=0)
        antigos = [str(s) for s in wb.LinkSources(1)
                   if os.path.basename(str(s)).lower().startswith("resultado tlm - mp27 vs r09")
                   and os.path.basename(str(s)).lower() != novo.lower()]
        for a in antigos:
            wb.ChangeLink(a, resultado, 1)
            for nome in ("Resumo YTD", "Resumo YTG", "Resumo Month", "Resumo FY"):
                wb.Worksheets(nome).UsedRange.Replace("[%s]" % os.path.basename(a), "[%s]" % novo, 2)
        xl.CalculateFull()
        fy = wb.Worksheets("Resumo FY")
        print("vinculos trocados:", [os.path.basename(a) for a in antigos], "->", novo)
        for k in ("C6", "C11", "C19", "C27", "G6", "G27", "I27", "K27"):
            print(" ", k, fy.Range(k).Value)
        erros = {}
        for nome in ABAS_ERRO:
            erros[nome] = sum(1 for l in wb.Worksheets(nome).UsedRange.Value for x in l if erro(x))
        print("erros", erros)
        print("links", [os.path.basename(str(s)) for s in wb.LinkSources(1)])
        wb.Save()
        wb.Close(False)
        # 2a passada: o cache gravado logo apos o ChangeLink sai com #REF!; reabrir ja com o vinculo novo e recalcular resolve
        wb = xl.Workbooks.Open(saida, UpdateLinks=0)
        xl.CalculateFull()
        erros = {nome: sum(1 for l in wb.Worksheets(nome).UsedRange.Value for x in l if erro(x)) for nome in ABAS_ERRO}
        print("erros (2a passada)", erros)
        for k in ("C6", "C27", "G6", "G27", "I27", "K27"):
            print(" ", k, wb.Worksheets("Resumo FY").Range(k).Value)
        wb.Save()
        wb.Close(False)
        res.Close(False)
        for a in ab:
            a.Close(False)
    finally:
        xl.Quit()
    print("Saida:", saida)


if __name__ == "__main__":
    main()
