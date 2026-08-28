Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Descobre os caminhos a partir de onde este .vbs está salvo, sem depender
' do usuário ou de onde a pasta do projeto foi clonada/copiada.
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pyFile = scriptDir & "\atualizar_ksb1_gui.py"

' scriptDir = ...\scripts\sap\fitted_units\fitted_units_despesas
' raiz do projeto = 4 níveis acima
projectRoot = scriptDir
For i = 1 To 4
    projectRoot = fso.GetParentFolderName(projectRoot)
Next

If Not fso.FileExists(pyFile) Then
    MsgBox "Não encontrei 'atualizar_ksb1_gui.py' em:" & vbCrLf & pyFile & vbCrLf & vbCrLf & _
        "Verifique se a pasta do projeto foi copiada inteira (com toda a estrutura de subpastas).", vbCritical, "Cockpit KSB1"
    WScript.Quit
End If

WshShell.CurrentDirectory = projectRoot
WshShell.Run "pythonw """ & pyFile & """", 0, False
