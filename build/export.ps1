<#
.SYNOPSIS
  Exporta el codigo VBA de un .xlsm de vuelta a src\ (para versionarlo en Git).

.DESCRIPTION
  Usalo cuando hagas cambios directamente en el editor de VBA y quieras
  pasarlos al repositorio. modSecrets no se exporta (tu clave no llega a Git).
  El proyecto VBA NO debe tener contrasena al exportar.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File build\export.ps1
  powershell -ExecutionPolicy Bypass -File build\export.ps1 -Workbook "C:\ruta\archivo.xlsm"
#>
param(
    [string]$Workbook = ""
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'src'
if (-not $Workbook) { $Workbook = Join-Path $root 'dist\Cotiva-Cotizador.xlsm' }
if (-not (Test-Path $Workbook)) { throw "No existe el archivo: $Workbook" }

$regPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$previousVbom = (Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $regPath -Name AccessVBOM -Value 1 -Type DWord

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3            # No ejecutar macros al abrir
    $wb = $excel.Workbooks.Open($Workbook, 0, $true)   # Solo lectura
    $encoding = [System.Text.Encoding]::GetEncoding(1252)

    foreach ($component in $wb.VBProject.VBComponents) {
        $name = $component.Name
        if ($component.Type -eq 1 -and $name -ne 'modSecrets') {
            $target = Join-Path $src "modules\$name.bas"
            $component.Export($target)
            Write-Host "  modulo  -> $target"
        }
        elseif ($component.Type -eq 2) {
            $target = Join-Path $src "classes\$name.cls"
            $component.Export($target)
            Write-Host "  clase   -> $target"
        }
        elseif ($component.Type -eq 3) {
            $target = Join-Path $src "forms\$name.vba"
            $codeModule = $component.CodeModule
            $code = ""
            if ($codeModule.CountOfLines -gt 0) { $code = $codeModule.Lines(1, $codeModule.CountOfLines) }
            [System.IO.File]::WriteAllText($target, $code + "`r`n", $encoding)
            Write-Host "  form    -> $target"
        }
    }
    Write-Host "Exportacion terminada." -ForegroundColor Green
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if ($null -eq $previousVbom) {
        Remove-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $regPath -Name AccessVBOM -Value $previousVbom -Type DWord
    }
}
