#!/usr/bin/env python3
"""
Lista todos os campos (IDs tecnicos, tipo e texto) da tela atual do SAP GUI.

Uso: deixe a tela do SAP que voce quer mapear aberta (ex: tela de selecao da
KSB1) e rode este script. Ele nao clica em nada, so le o que ja esta na tela.
"""
import sys

import win32com.client


def connect_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)


def dump(element, depth=0):
    try:
        elem_id = element.Id
        elem_type = element.Type
    except Exception:
        return

    texto = ""
    try:
        texto = element.Text
    except Exception:
        pass

    prefixo = "  " * depth
    if texto.strip():
        print(f'{prefixo}[{elem_type}] {elem_id}  ->  "{texto}"')
    else:
        print(f"{prefixo}[{elem_type}] {elem_id}")

    try:
        children = element.Children
        for i in range(children.Count):
            dump(children.Item(i), depth + 1)
    except Exception:
        pass


def main():
    try:
        session = connect_session()
    except Exception as e:
        print(f"ERRO ao conectar: {e}")
        sys.exit(1)

    print(f"Transacao atual: {session.Info.Transaction}")
    print("=" * 80)
    wnd = session.FindById("wnd[0]")
    dump(wnd)


if __name__ == "__main__":
    main()
