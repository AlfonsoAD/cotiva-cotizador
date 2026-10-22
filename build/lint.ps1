<#
.SYNOPSIS
  Revision estatica del codigo VBA de src\ (sin Excel). La usa bflow check.

.DESCRIPTION
  En todos los archivos de src\:
  - Option Explicit al inicio del codigo.
  - Codificacion Windows-1252: falla si el archivo quedo en UTF-8 (los acentos
    se verian mal al importarlo en el editor de VBA) o tiene bytes indefinidos.

  Solo en las lineas que agrega la rama respecto a -Base (el codigo previo
  tiene usos legitimos, como mostrar una hoja al usuario):
  - Sin Debug.Print ni Stop.
  - Sin Select, Activate, ActiveSheet, ActiveCell ni Selection.
  Una linea que de verdad lo necesite lleva el comentario  ' lint:allow

  Compilar en Excel y probar sigue siendo manual (build.ps1).

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File build\lint.ps1 -Base auto-repair-shop
#>
param(
    [string]$Base = ""
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'src'
$problems = New-Object System.Collections.Generic.List[string]

# Bytes que Windows-1252 no define
$undefined = @(0x81, 0x8D, 0x8F, 0x90, 0x9D)
$utf8 = New-Object System.Text.UTF8Encoding($false, $true)   # lanza si no es UTF-8 valido

$files = Get-ChildItem -Path $src -Recurse -File -Include *.bas, *.cls, *.vba |
    Where-Object { $_.Name -ne 'modSecrets.bas' }

foreach ($file in $files) {
    $rel = $file.FullName.Substring($root.Length + 1)
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    # --- Codificacion ---
    $hasHigh = $false
    foreach ($b in $bytes) {
        if ($b -ge 0x80) { $hasHigh = $true }
        if ($undefined -contains $b) {
            $problems.Add("${rel}: byte 0x{0:X2} no existe en Windows-1252" -f $b)
            break
        }
    }
    if ($hasHigh) {
        $isUtf8 = $true
        try { [void]$utf8.GetString($bytes) } catch { $isUtf8 = $false }
        if ($isUtf8) { $problems.Add("${rel}: esta en UTF-8; debe quedar en Windows-1252 (ver .gitattributes)") }
    }

    # --- Option Explicit ---
    $text = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
    if ($text -notmatch '(?im)^\s*Option\s+Explicit\s*$') {
        $problems.Add("${rel}: falta Option Explicit")
    }
}

# --- Reglas sobre las lineas nuevas de la rama ---
if ($Base) {
    $mergeBase = (& git -C $root merge-base $Base HEAD 2>$null)
    if (-not $mergeBase) {
        $problems.Add("no se encontro la rama base '$Base'")
    } else {
        $rules = @(
            @{ Pattern = '\bDebug\.Print\b';                         Message = 'Debug.Print olvidado' },
            @{ Pattern = '^\s*Stop\s*(''.*)?$';                      Message = 'Stop olvidado' },
            @{ Pattern = '\.(Select|Activate)\b';                    Message = 'Select/Activate: usa referencias explicitas' },
            @{ Pattern = '\b(ActiveSheet|ActiveCell|Selection)\b';   Message = 'ActiveSheet/ActiveCell/Selection: usa referencias explicitas' }
        )
        $diff = & git -C $root diff --unified=0 --no-color $mergeBase -- src
        $current = ''
        $lineNo = 0
        foreach ($line in $diff) {
            if ($line -match '^\+\+\+ b/(.+)$') { $current = $Matches[1]; continue }
            if ($line -match '^@@ -\d+(?:,\d+)? \+(\d+)') { $lineNo = [int]$Matches[1]; continue }
            if ($line.StartsWith('+') -and -not $line.StartsWith('+++')) {
                $code = $line.Substring(1)
                if ($code -notmatch "'\s*lint:allow") {
                    # Ignora lo que va despues de un comentario
                    $stmt = ($code -split "'", 2)[0]
                    foreach ($rule in $rules) {
                        if ($stmt -match $rule.Pattern) {
                            $problems.Add("${current}:${lineNo}: $($rule.Message)")
                        }
                    }
                }
                $lineNo++
            }
        }
    }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host $_ }
    Write-Host "lint: $($problems.Count) problema(s)"
    exit 1
}
Write-Host "lint: ok ($($files.Count) archivos)"
