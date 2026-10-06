#!/usr/bin/env python3
"""
Le uma tabela do SAP via SE16N (SOMENTE LEITURA) e salva em data/processed/.

Abre uma sessao NOVA do SAP GUI (nao mexe na tela atual) e a fecha ao terminar.
Uso: python ler_tabela_se16n.py T16FS [T16FC ...] [--max 5000]
"""
import argparse
import csv
import sys
import time
from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parents[5]
SAIDA = RAIZ / "data" / "processed" / "cadeia_aprovacao"


def nova_sessao():
    app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    sess = app.Children(0).Children(0)
    con = app.Children(0)
    antes = con.Children.Count
    sess.CreateSession()
    for _ in range(60):
        time.sleep(0.5)
        if con.Children.Count > antes:
            nova = con.Children(con.Children.Count - 1)
            try:
                nova.findById("wnd[0]/tbar[0]/okcd")
                return nova
            except Exception:
                pass
    raise RuntimeError("Nao consegui abrir uma sessao nova do SAP.")


def esperar(sess):
    while sess.Busy:
        time.sleep(0.2)


def ler_tabela(sess, tabela, maximo):
    sess.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16N"
    sess.findById("wnd[0]").sendVKey(0)
    esperar(sess)
    sess.findById("wnd[0]/usr/ctxtGD-TAB").text = tabela
    sess.findById("wnd[0]").sendVKey(0)
    esperar(sess)
    sess.findById("wnd[0]/usr/txtGD-MAX_LINES").text = str(maximo)
    sess.findById("wnd[0]").sendVKey(8)  # F8 executar
    esperar(sess)
    grid = sess.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    colunas = list(grid.ColumnOrder)
    linhas = []
    for r in range(grid.RowCount):
        linhas.append([grid.GetCellValue(r, c) for c in colunas])
    return colunas, linhas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("tabelas", nargs="+")
    ap.add_argument("--max", type=int, default=5000)
    args = ap.parse_args()

    SAIDA.mkdir(parents=True, exist_ok=True)
    sess = nova_sessao()
    try:
        for tab in args.tabelas:
            try:
                cols, linhas = ler_tabela(sess, tab.upper(), args.max)
            except Exception as e:
                print(f"{tab}: FALHOU ({e})")
                continue
            arq = SAIDA / f"{tab.upper()}.csv"
            with open(arq, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(cols)
                w.writerows(linhas)
            print(f"{tab}: {len(linhas)} linhas -> {arq}")
    finally:
        try:
            sess.findById("wnd[0]").close()
            sess.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
