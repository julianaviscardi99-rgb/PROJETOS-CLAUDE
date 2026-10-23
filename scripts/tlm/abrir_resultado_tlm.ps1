# abrir_resultado_tlm.ps1 -- abre o Resultado TLM MP27 mais recente JA com as fontes abertas (evita #VALUE!/#N/A).
# Uso: duplo clique em "Abrir Resultado TLM MP27" (so o Resultado) ou "Abrir Resumo TLM MP27" (-Resumo: fontes + Resultado + Resumo).
# Parametros opcionais: -Arquivo <caminho> (outra versao do Resultado), -Resumo (abre tambem o Resumo mais recente).
# Nao altera nenhum arquivo: fontes em somente leitura, Resultado aberto sem atualizar links (usa o valor gravado e recalcula com as fontes).
param([string]$Arquivo = "", [switch]$Resumo)
$ErrorActionPreference = "Stop"
$rede = "\\SRV054-01BR.group.pirelli.com\DAC_TLM"
$pasta = (Get-ChildItem "$rede\MP27" -Directory | Where-Object { $_.Name -like "1*MP27 vs R09" } | Select-Object -First 1).FullName + "\TLM_Automation"
$a26 = "$rede\2026"
$fontes = @(
    "$a26\Base Total_Act_Flash_Fct.xlsx",
    "$a26\Base_KSB1.xlsx",
    "$a26\1. FCST\R09\HC R09 TLM.xlsx",
    "$a26\1. FCST\R07\HC R07.xlsx",
    "$a26\3. Actual\Act_09.Setembro\Headcount TLM - 2026.09.xlsx",
    "$pasta\Base Total MP27.xlsx",
    "$a26\1. FCST\R09\Resultado TLM - R09 '26.xlsm"
)
function UltimaVersao($filtro) {   # maior _vN do nome (sem _v = 1), ignora "sem link" e temporarios
    (Get-ChildItem $pasta -Filter $filtro |
        Where-Object { $_.Name -notlike "*sem link*" -and $_.Name -notlike "~*" } |
        Sort-Object { if ($_.BaseName -match "_v(\d+)$") { [int]$Matches[1] } else { 1 } } -Descending |
        Select-Object -First 1).FullName
}
if (-not $Arquivo) { $Arquivo = UltimaVersao "Resultado TLM - MP27 vs R09*.xlsm" }
$ArquivoResumo = $null
if ($Resumo) { $ArquivoResumo = UltimaVersao "Resumo Resultado TLM - MP27 vs R09*.xlsx"; Write-Host "Resumo: $ArquivoResumo" }
Write-Host "Resultado: $Arquivo"
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
catch { $xl = New-Object -ComObject Excel.Application }
$xl.Visible = $true
$xl.AskToUpdateLinks = $false
$abertos = @(); foreach ($b in $xl.Workbooks) { $abertos += $b.FullName.ToLower() }
$ordem = $fontes + $Arquivo
if ($ArquivoResumo) { $ordem += $ArquivoResumo }     # o Resumo por ultimo: precisa do Resultado e das fontes ja abertos
foreach ($f in $ordem) {
    if ($abertos -contains $f.ToLower()) { Write-Host "  ja aberto: $(Split-Path $f -Leaf)"; continue }
    Write-Host "  abrindo: $(Split-Path $f -Leaf)"
    if ($f -eq $Arquivo -or $f -eq $ArquivoResumo) { [void]$xl.Workbooks.Open($f, 0) } else { [void]$xl.Workbooks.Open($f, 0, $true) }
}
$alvo = if ($ArquivoResumo) { $ArquivoResumo } else { $Arquivo }
$xl.Workbooks | Where-Object { $_.FullName -eq $alvo } | ForEach-Object { $_.Activate() }
Write-Host "Pronto. Deixe as fontes e o Resultado abertos enquanto usa o Resumo."
