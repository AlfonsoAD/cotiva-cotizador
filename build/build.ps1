<#
.SYNOPSIS
  Genera dist\Cotiva-Cotizador.xlsm a partir del codigo fuente en src\.

.DESCRIPTION
  1. Crea un libro nuevo de Excel.
  2. Importa modulos (.bas) y clases (.cls).
  3. Crea el formulario frmQuote y le inserta su codigo (src\forms\frmQuote.vba).
  4. Ejecuta modBuild.BuildWorkbook: hojas, boton de inicio, logos y proteccion.
  5. Elimina modBuild y guarda el .xlsm.

  Requisitos:
  - Windows con Excel de escritorio (Office 2016 o posterior).
  - src\modules\modSecrets.bas (copia de modSecrets.example.bas con tu clave).

  El script activa temporalmente "Confiar en el acceso al modelo de objetos
  de proyectos de VBA" y lo restaura al terminar.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File build\build.ps1
#>
param(
    [string]$Output = "",
    [switch]$Hidden          # Excel invisible (si hay un error de VBA no lo veras)
)

$ErrorActionPreference = 'Stop'

$root   = Split-Path -Parent $PSScriptRoot
$src    = Join-Path $root 'src'
$dist   = Join-Path $root 'dist'
$assets = Join-Path $root 'assets'
if (-not $Output) { $Output = Join-Path $dist 'Cotiva-Cotizador.xlsm' }

# --- Validaciones previas ---
$secrets = Join-Path $src 'modules\modSecrets.bas'
if (-not (Test-Path $secrets)) {
    throw "Falta src\modules\modSecrets.bas. Copia modSecrets.example.bas como modSecrets.bas y pon tu clave."
}
New-Item -ItemType Directory -Force -Path $dist | Out-Null
if (Test-Path $Output) { Remove-Item $Output -Force }

# --- Acceso al modelo de objetos de VBA (temporal) ---
$regPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$previousVbom = (Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $regPath -Name AccessVBOM -Value 1 -Type DWord

$excel = $null
$wb = $null
try {
    Write-Host "Abriendo Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = -not $Hidden
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($Output, 52)                  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $vbp = $wb.VBProject

    Write-Host "Importando modulos..."
    Get-ChildItem (Join-Path $src 'modules') -Filter *.bas |
        Where-Object { $_.Name -ne 'modSecrets.example.bas' } |
        ForEach-Object { Write-Host "  + $($_.Name)"; [void]$vbp.VBComponents.Import($_.FullName) }

    Write-Host "Importando clases..."
    Get-ChildItem (Join-Path $src 'classes') -Filter *.cls |
        ForEach-Object { Write-Host "  + $($_.Name)"; [void]$vbp.VBComponents.Import($_.FullName) }

    Write-Host "Importando modulo de build..."
    [void]$vbp.VBComponents.Import((Join-Path $src 'build\modBuild.bas'))

    Write-Host "Creando formulario frmQuote..."
    $form = $vbp.VBComponents.Add(3)          # 3 = vbext_ct_MSForm
    $form.Name = 'frmQuote'
    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $formCode = [System.IO.File]::ReadAllText((Join-Path $src 'forms\frmQuote.vba'), $encoding)
    $module = $form.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($formCode)

    Write-Host "Armando el libro (hojas, logos, proteccion)..."
    [void]$excel.Run("'" + $wb.Name + "'!modBuild.BuildWorkbook", $assets)

    Write-Host "Quitando modulo de build..."
    $vbp.VBComponents.Remove($vbp.VBComponents.Item('modBuild'))

    $wb.Save()
    Write-Host ""
    Write-Host "Listo: $Output" -ForegroundColor Green
    Write-Host "Recuerda: contrasena del proyecto VBA y firma digital se ponen a mano (ver README)."
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    # Restaurar la configuracion de seguridad original
    if ($null -eq $previousVbom) {
        Remove-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $regPath -Name AccessVBOM -Value $previousVbom -Type DWord
    }
}
