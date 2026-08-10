#!/usr/bin/env python3
"""
Testa se e possivel conectar via scripting a uma sessao do SAP GUI ja aberta.

Pre-requisito: abrir o SAP GUI, logar em um sistema (ex: G20) e deixar a
janela aberta antes de rodar este script.
"""
import sys

try:
    import win32com.client
except ImportError:
    print("ERRO: pywin32 nao esta instalado. Rode: pip install pywin32")
    sys.exit(1)


def main():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        print("FALHA: nao encontrei nenhuma instancia do SAP GUI rodando.")
        print("Abra o SAP GUI e faca login em um sistema antes de rodar este teste.")
        print(f"Detalhe: {e}")
        sys.exit(1)

    try:
        application = sap_gui_auto.GetScriptingEngine
    except Exception as e:
        print("FALHA: o SAP GUI esta aberto, mas o scripting nao respondeu.")
        print("Isso normalmente significa que o SCRIPTING ESTA BLOQUEADO NO SERVIDOR")
        print("(parametro sapgui/user_scripting controlado pela TI/Basis da Pirelli).")
        print(f"Detalhe: {e}")
        sys.exit(1)

    if application.Children.Count == 0:
        print("O SAP GUI esta aberto, mas nao encontrei nenhuma sessao logada (conexao).")
        print("Faca login em um sistema (ex: G20) e rode o teste novamente.")
        sys.exit(1)

    connection = application.Children(0)
    session = connection.Children(0)

    print("SUCESSO! Conectado via scripting.")
    print(f"Sistema: {session.Info.SystemName}")
    print(f"Mandante (client): {session.Info.Client}")
    print(f"Usuario: {session.Info.User}")
    print(f"Transacao atual: {session.Info.Transaction}")


if __name__ == "__main__":
    main()
