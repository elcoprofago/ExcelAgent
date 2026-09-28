# Limpieza de los restos de desarrollo de ExcelAgent.
# Ver la lista SIN borrar nada:
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1 -WhatIf
# Ejecutar:
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1
# Cerrando ademas la instancia de Excel que quedo sin ventana:
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1 -CerrarExcel
[CmdletBinding(SupportsShouldProcess = $true)]
param([switch]$CerrarExcel)

$carpetas = @(
    'D:\REPOS\EXCEL\_scratch'
    'D:\REPOS\EXCEL\ExcelAgent\_partes'
    'D:\REPOS\EXCEL\ExcelAgent\excelagent'
    'D:\REPOS\EXCEL\ExcelAgent\__pycache__'
    'D:\REPOS\EXCEL\ExcelAgent\tests\__pycache__'
)

$archivos = @(
    'D:\REPOS\EXCEL\_probe.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_diag.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_e1.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_e2.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_len1.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_len3.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_os1.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_os2.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_os3.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_t.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_t3.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_vA.txt'
    'D:\REPOS\EXCEL\ExcelAgent\_vB.txt'
    'D:\REPOS\EXCEL\ExcelAgent\tests\diag.py'
)

foreach ($c in $carpetas) {
    if (Test-Path -LiteralPath $c) {
        Remove-Item -LiteralPath $c -Recurse -Force
        Write-Host ('carpeta eliminada: ' + $c)
    }
}

foreach ($a in $archivos) {
    if (Test-Path -LiteralPath $a) {
        Remove-Item -LiteralPath $a -Force
        Write-Host ('archivo eliminado: ' + $a)
    }
}

Write-Host 'Limpieza terminada. Se conservan .venv, tesseract y el proyecto.'

if ($CerrarExcel) {
    $procesos = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
    foreach ($p in $procesos) {
        if ([string]::IsNullOrEmpty($p.MainWindowTitle)) {
            Stop-Process -Id $p.Id -Force
            Write-Host ('cerrado EXCEL sin ventana: ' + $p.Id)
        }
        else {
            Write-Host ('NO se cierra EXCEL ' + $p.Id + ': tiene ventana -> ' + $p.MainWindowTitle)
        }
    }
}
