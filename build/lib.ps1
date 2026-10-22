<#
.SYNOPSIS
  Armado comun del libro de Excel. Lo cargan build.ps1 y test.ps1 con dot-source:
  . (Join-Path $PSScriptRoot 'lib.ps1')
#>

$script:VbaSecurityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'

# Enciende "Confiar en el acceso al modelo de objetos de proyectos de VBA".
# Regresa el valor anterior ($null si no existia) para pasarlo a Restore-VbaAccess.
function Enable-VbaAccess {
    if (-not (Test-Path $script:VbaSecurityKey)) { New-Item -Path $script:VbaSecurityKey -Force | Out-Null }
    $previous = (Get-ItemProperty -Path $script:VbaSecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty -Path $script:VbaSecurityKey -Name AccessVBOM -Value 1 -Type DWord
    return $previous
}

function Restore-VbaAccess {
    param($Previous)
    if ($null -eq $Previous) {
        Remove-ItemProperty -Path $script:VbaSecurityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $script:VbaSecurityKey -Name AccessVBOM -Value $Previous -Type DWord
    }
}

function New-ExcelApp {
    param([bool]$Visible)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    return $excel
}

# PID del proceso de esta instancia de Excel, a partir de su ventana principal.
function Get-ExcelProcessId {
    param($Excel)
    if (-not ('CotivaNative.User32' -as [type])) {
        Add-Type -Namespace CotivaNative -Name User32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
    }
    [uint32]$procId = 0
    [void][CotivaNative.User32]::GetWindowThreadProcessId([System.IntPtr]$Excel.Hwnd, [ref]$procId)
    return [int]$procId
}

# Importa el codigo de src\ al libro: modulos, secretos, clases, modBuild y frmQuote.
# Con -IncludeTests agrega al final src\tests\*.bas (solo para el libro de pruebas).
function Import-CotizadorCode {
    param(
        $Workbook,
        [string]$Src,
        [string]$SecretsFile,
        [switch]$IncludeTests
    )
    $vbp = $Workbook.VBProject

    Write-Host "Importando modulos..."
    Get-ChildItem (Join-Path $Src 'modules') -Filter *.bas |
        Where-Object { $_.Name -ne 'modSecrets.bas' -and $_.Name -ne 'modSecrets.example.bas' } |
        ForEach-Object { Write-Host "  + $($_.Name)"; [void]$vbp.VBComponents.Import($_.FullName) }
    Write-Host "  + $(Split-Path -Leaf $SecretsFile)"
    [void]$vbp.VBComponents.Import($SecretsFile)

    Write-Host "Importando clases..."
    Get-ChildItem (Join-Path $Src 'classes') -Filter *.cls |
        ForEach-Object { Write-Host "  + $($_.Name)"; [void]$vbp.VBComponents.Import($_.FullName) }

    Write-Host "Importando modulo de build..."
    [void]$vbp.VBComponents.Import((Join-Path $Src 'build\modBuild.bas'))

    Write-Host "Creando formulario frmQuote..."
    $form = $vbp.VBComponents.Add(3)          # 3 = vbext_ct_MSForm
    $form.Name = 'frmQuote'
    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $formCode = [System.IO.File]::ReadAllText((Join-Path $Src 'forms\frmQuote.vba'), $encoding)
    $module = $form.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($formCode)

    if ($IncludeTests) {
        Write-Host "Importando pruebas..."
        Get-ChildItem (Join-Path $Src 'tests') -Filter *.bas |
            ForEach-Object { Write-Host "  + $($_.Name)"; [void]$vbp.VBComponents.Import($_.FullName) }
    }
}

# Cierra el libro sin guardar y sale de Excel. Tolera $null.
function Close-ExcelApp {
    param($Excel, $Workbook)
    if ($Workbook) { try { $Workbook.Close($false) } catch { } }
    if ($Excel) {
        try { $Excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }
}
