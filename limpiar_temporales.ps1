# Limpieza de los restos de desarrollo de ExcelAgent (los que dejo la version anterior en la notebook).
# Las rutas son relativas a este script: sirve igual en D:\REPOS\EXCEL\ExcelAgent que en F:\source\repos\EXCEL\ExcelAgent.
# No toca .venv, tesseract, logs, los respaldos ni el proyecto.
# Ver la lista SIN borrar nada:
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1 -WhatIf
# Ejecutar:
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1
# Listar ademas las instancias de Excel que quedaron sin ventana (NO las cierra):
#   powershell -ExecutionPolicy Bypass -File limpiar_temporales.ps1 -ExcelSinVentana
[CmdletBinding(SupportsShouldProcess = $true)]
param([switch]$ExcelSinVentana)

$proyecto = $PSScriptRoot
$raiz = Split-Path -Parent $proyecto

$carpetas = @(
    (Join-Path $raiz '_scratch')
    (Join-Path $proyecto '_partes')
    (Join-Path $proyecto 'excelagent')
    (Join-Path $proyecto '__pycache__')
    (Join-Path $proyecto 'tests\__pycache__')
    (Join-Path $proyecto 'tests\modulos\__pycache__')
    (Join-Path $env:TEMP 'ea_build')
)
$carpetas += @(Get-ChildItem -Path $env:TEMP -Directory -Filter 'contactos_*' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })

$archivos = @(
    (Join-Path $raiz '_probe.txt')
    (Join-Path $proyecto 'tests\diag.py')
)
foreach ($n in '_bat_nuevo', '_ml2', '_ml3', '_diag', '_e1', '_e2', '_len1', '_len3', '_os1', '_os2', '_os3', '_t', '_t3', '_vA', '_vB') {
    $archivos += (Join-Path $proyecto ($n + '.txt'))
}

foreach ($c in $carpetas) {
    if ((Test-Path -LiteralPath $c) -and $PSCmdlet.ShouldProcess($c, 'eliminar carpeta')) {
        Remove-Item -LiteralPath $c -Recurse -Force
        Write-Host ('carpeta eliminada: ' + $c)
    }
}

foreach ($a in $archivos) {
    if ((Test-Path -LiteralPath $a) -and $PSCmdlet.ShouldProcess($a, 'eliminar archivo')) {
        Remove-Item -LiteralPath $a -Force
        Write-Host ('archivo eliminado: ' + $a)
    }
}

Write-Host 'Limpieza terminada. Se conservan .venv, tesseract, logs, respaldos y el proyecto.'

if ($ExcelSinVentana) {
    # Antes esto era -CerrarExcel y hacia Stop-Process -Force: un Excel sin ventana puede tener un libro oculto sin
    # guardar (por ejemplo el que abrio ExcelAgent con 'visible' apagado), y matarlo lo pierde sin preguntar.
    # Ahora solo se listan; para cerrarlos, abrir ExcelAgent y usar Reconectar/Salir, o cerrarlos a mano a conciencia.
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ([string]::IsNullOrEmpty($p.MainWindowTitle)) {
            Write-Host ('EXCEL sin ventana: PID ' + $p.Id + ', iniciado ' + $p.StartTime + ' (no se cierra)')
        }
    }
}
