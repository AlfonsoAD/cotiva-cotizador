<#
.SYNOPSIS
  Compila y prueba el VBA dentro de Excel (oculto). Lo corre bflow check (paso "test").

.DESCRIPTION
  Arma en memoria un libro con el mismo codigo que build.ps1 mas src\tests\*.bas,
  llama por COM a cada prueba (modTestRunner.ListTests, BeginTest, EndTest y Report) y escribe dist\test-results.txt. El libro nunca
  se guarda en disco.

  Sale con 0 solo si hubo al menos una prueba y ninguna fallo; con 1 si falla una
  prueba, no hay pruebas, no hay Excel o se pasa el limite de tiempo.

  El proceso padre vigila el tiempo y lanza un hijo (-Worker) que maneja Excel:
  si Excel se queda esperando (por ejemplo, un dialogo de error de compilacion),
  el padre mata solo a ese hijo y a la instancia de Excel que el hijo registro.

  Requisitos: Windows con Excel de escritorio. Activa temporalmente "Confiar en el
  acceso al modelo de objetos de proyectos de VBA" y lo restaura al terminar.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File build\test.ps1
  powershell -ExecutionPolicy Bypass -File build\test.ps1 -Visible -TimeoutSeconds 600
#>
param(
    [int]$TimeoutSeconds = 300,
    [string]$Results = "",       # por omision dist\test-results.txt
    [switch]$Visible,            # para depurar: Excel visible
    [switch]$Worker,             # interno: lo pone el padre
    [string]$PidFile = ""        # interno
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lib.ps1')

$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
if (-not $Results) { $Results = Join-Path $dist 'test-results.txt' }

function Write-Results {
    param([string]$Text)
    $dir = Split-Path -Parent $Results
    if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($Results, $Text, (New-Object System.Text.UTF8Encoding($false)))
}

# ============================== Hijo: maneja Excel ==============================
if ($Worker) {
    $excel = $null
    $wb = $null
    $exitCode = 1
    try {
        $excel = New-ExcelApp -Visible $Visible.IsPresent
        if ($PidFile) { Set-Content -Path $PidFile -Value (Get-ExcelProcessId $excel) -Encoding ASCII }
        $wb = $excel.Workbooks.Add()          # No se guarda: queda como "Libro1" en memoria

        # Sin modSecrets.bas se usa el de ejemplo (ya declara VB_Name = "modSecrets")
        $secrets = Join-Path $src 'modules\modSecrets.bas'
        if (-not (Test-Path $secrets)) { $secrets = Join-Path $src 'modules\modSecrets.example.bas' }

        Import-CotizadorCode -Workbook $wb -Src $src -SecretsFile $secrets -IncludeTests

        Write-Host "Armando el libro y corriendo pruebas..."
        [void]$excel.Run("'" + $wb.Name + "'!modBuild.BuildWorkbook", (Join-Path $root 'assets'))
        # Cada prueba se llama por COM: un error sin manejar regresa aqui sin abrir el
        # dialogo de depuracion (desde Application.Run dentro de VBA si lo abriria).
        $book = "'" + $wb.Name + "'!"
        $names = @(([string]$excel.Run($book + "modTestRunner.ListTests")) -split '\|' | Where-Object { $_ })
        $passed = 0
        $failed = 0
        $failureLines = New-Object System.Collections.Generic.List[string]
        foreach ($name in $names) {
            [void]$excel.Run($book + "modTestRunner.BeginTest", $name)
            $failure = ""
            try {
                [void]$excel.Run($book + $name)
            } catch {
                $com = $_.Exception.InnerException
                if (-not $com) { $com = $_.Exception }
                $failure = "FALLA ${name}: error {0:X8}: {1}" -f [int]$com.HResult, $com.Message
            }
            $recorded = [string]$excel.Run($book + "modTestRunner.EndTest")
            if ($recorded -or $failure) { $failureLines.Add("--- FAIL: $name") }
            if ($recorded) { $failureLines.Add($recorded) }
            if ($failure) { $failureLines.Add($failure) }
            if ($recorded -or $failure) { $failed++ } else { $passed++ }
        }
        $report = [string]$excel.Run($book + "modTestRunner.Report", $passed, $failed, ($failureLines -join "`n"))
        $report = ($report -replace "`r`n", "`n" -replace "`r", "`n").TrimEnd()

        $lines = $report -split "`n"
        $ok = $lines[0] -match '^ok (\d+) / falla 0$' -and [int]$Matches[1] -gt 0
        $text = ($lines -join "`r`n") + "`r`n"
        Write-Results $text
        if ($ok) {
            Write-Host $lines[0] -ForegroundColor Green
            $exitCode = 0
        } else {
            Write-Host $lines[0] -ForegroundColor Red
            $lines | Select-Object -Skip 1 | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        }
    }
    catch {
        $msg = "error: " + $_.Exception.Message
        Write-Results ($msg + "`r`n")
        Write-Host $msg -ForegroundColor Red
        $exitCode = 1
    }
    finally {
        Close-ExcelApp -Excel $excel -Workbook $wb
    }
    exit $exitCode
}

# ============================== Padre: vigila el tiempo ==============================
if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) {
    $msg = "las pruebas de VBA necesitan Excel de escritorio; no esta instalado"
    Write-Results ($msg + "`r`n")
    Write-Host $msg -ForegroundColor Red
    exit 1
}

$exitCode = 1
$previousVbom = $null
$pidFilePath = Join-Path ([System.IO.Path]::GetTempPath()) ("cotiva-test-" + [guid]::NewGuid().ToString('N') + ".pid")
try {
    $previousVbom = Enable-VbaAccess
    if (Test-Path $Results) { Remove-Item $Results -Force }
    Set-Content -Path $pidFilePath -Value '' -Encoding ASCII

    $childArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'),
              '-Worker', '-PidFile', ('"' + $pidFilePath + '"'), '-Results', ('"' + $Results + '"'))
    if ($Visible) { $childArgs += '-Visible' }
    $child = Start-Process -FilePath 'powershell' -ArgumentList $childArgs -NoNewWindow -PassThru
    [void]$child.Handle                       # Necesario para que ExitCode quede disponible

    if ($child.WaitForExit($TimeoutSeconds * 1000)) {
        $exitCode = $child.ExitCode
        if (-not (Test-Path $Results)) {
            $msg = "el proceso de pruebas termino sin resultados"
            Write-Results ($msg + "`r`n")
            Write-Host $msg -ForegroundColor Red
            $exitCode = 1
        }
    } else {
        # Solo el hijo y la instancia de Excel cuyo PID dejo el hijo; nunca otros Excel
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
        $excelPid = 0
        if ([int]::TryParse(((Get-Content -Path $pidFilePath -ErrorAction SilentlyContinue) -join '').Trim(), [ref]$excelPid) -and $excelPid -gt 0) {
            $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($proc -and $proc.ProcessName -eq 'EXCEL') { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
        }
        $msg = "tiempo agotado ($TimeoutSeconds s): posible error de compilacion; corre build.ps1 para verlo"
        Write-Results ($msg + "`r`n")
        Write-Host $msg -ForegroundColor Red
        $exitCode = 1
    }
}
finally {
    Restore-VbaAccess -Previous $previousVbom
    Remove-Item $pidFilePath -Force -ErrorAction SilentlyContinue
}
exit $exitCode
