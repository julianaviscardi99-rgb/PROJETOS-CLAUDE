$raiz = '\\FSS024-01BR.group.pirelli.com\GFU_DAC\Custos Fitted Units\Resultados Fitted\2026'
$pasta = Get-ChildItem -Path $raiz -Directory | Where-Object { $_.Name -like '00.Extra*Base KSB1' } | Select-Object -First 1
if (-not $pasta) {
    Write-Error 'Pasta 00.Extracao Base KSB1 nao encontrada'
    exit 1
}
$shell = New-Object -ComObject WScript.Shell
$atalho = $shell.CreateShortcut((Join-Path $pasta.FullName 'ATUALIZAR KSB1.lnk'))
$atalho.TargetPath = 'C:\Windows\System32\wscript.exe'
$atalho.Arguments = '"C:\Users\silveju001\Projetos Claude\scripts\sap\fitted_units\fitted_units_despesas\atualizar_ksb1_launcher.vbs"'
$atalho.WorkingDirectory = 'C:\Users\silveju001\Projetos Claude'
$atalho.IconLocation = 'C:\Users\silveju001\Projetos Claude\scripts\sap\fitted_units\fitted_units_despesas\assets\pirelli_tire.ico'
$atalho.Description = 'Atualizar KSB1 - Fitted Units'
$atalho.Save()
Write-Output "Atalho criado em: $($pasta.FullName)"
