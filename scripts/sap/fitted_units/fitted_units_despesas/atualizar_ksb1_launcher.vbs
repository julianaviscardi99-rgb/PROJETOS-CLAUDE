Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\silveju001\Projetos Claude"
WshShell.Run "pythonw ""C:\Users\silveju001\Projetos Claude\scripts\sap\fitted_units\fitted_units_despesas\atualizar_ksb1_gui.py""", 0, False
