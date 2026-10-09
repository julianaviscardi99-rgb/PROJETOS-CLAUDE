Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Roda o cockpit do CP a partir de onde este .vbs estiver salvo (projeto ou rede).
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyFile = scriptDir & "\cockpit_circuito_gui.py"

If Not fso.FileExists(pyFile) Then
    MsgBox "Não encontrei 'cockpit_circuito_gui.py' em:" & vbCrLf & pyFile, vbCritical, "Circuito Panamericano"
    WScript.Quit
End If

WshShell.Run "pythonw """ & pyFile & """", 0, False
