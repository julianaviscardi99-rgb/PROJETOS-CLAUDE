#!/usr/bin/env python3
"""
Script de diagnostico: lista o Id, o Type e o Text/Name de cada elemento
dentro do popup aberto no momento (wnd[1]) no SAP GUI.

Usado para descobrir o Id tecnico exato de um campo/botao de um popup do SAP
(ex: "Definir área contab.custos"), em vez de adivinhar — a primeira
tentativa de fechar esse popup automaticamente adivinhou o campo errado e
gerou uma cascata de erros (ver
memory/errors/2026-08-11_popup_area_contabil_ao_reentrar_sap.md).

Uso:
1. No SAP, chegue ate a tela onde o popup aparece (ex: abra a transacao
   KSB1 logo apos entrar numa sessao nova do SAP).
2. Com o popup aberto E SEM CLICAR EM NADA nele, rode:
       python scripts/sap/diagnosticar_popup.py
3. Copie a saida e mande para o Claude.
"""
import win32com.client


def connect_session():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)


def listar(elemento, prefixo=""):
    try:
        tipo = elemento.Type
    except Exception:
        tipo = "?"
    texto = ""
    for attr in ("Text", "Name"):
        try:
            texto = getattr(elemento, attr)
            break
        except Exception:
            continue
    try:
        id_ = elemento.Id
    except Exception:
        id_ = "?"
    print(f"{prefixo}{id_}  [{tipo}]  '{texto}'")

    try:
        filhos = elemento.Children
    except Exception:
        return
    for filho in filhos:
        listar(filho, prefixo + "  ")


def main():
    try:
        session = connect_session()
    except Exception as e:
        print(f"ERRO ao conectar no SAP: {e}")
        return

    wnd1 = session.FindById("wnd[1]", False)
    if wnd1 is None:
        print("Não encontrei nenhum popup (wnd[1]) aberto agora.")
        print("Abra o popup no SAP primeiro (sem clicar em nada nele) e rode de novo.")
        return

    print("Elementos dentro do popup wnd[1]:\n")
    listar(wnd1)


if __name__ == "__main__":
    main()
