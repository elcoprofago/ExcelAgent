# Deja andando el entorno virtual ..\.venv con el Python instalado en ESTA PC.
# Uso: doble clic en preparar_entorno.bat, o
#   powershell -ExecutionPolicy Bypass -File preparar_entorno.ps1
#
# Caso tipico: la carpeta se copio desde otra PC (otro usuario, otro Python) y pyvenv.cfg apunta a un
# python.exe que aca no existe. Entonces:
#   - si el Python de esta PC es de la misma version menor (3.13 con 3.13): solo se corrige pyvenv.cfg;
#   - si es otra (3.12, 3.14...): las bibliotecas compiladas (pywin32, pillow) no sirven, y se arma un
#     entorno nuevo con requirements.txt.
# Nada se borra: el entorno anterior queda renombrado como .venv.respaldo-<fecha>, y si algo falla se lo
# vuelve a poner en su lugar.
# 'Continue' y no 'Stop': en PowerShell 5.1, con 'Stop' cualquier aviso que python escriba en stderr aborta el
# script. Los errores se controlan por codigo de salida ($LASTEXITCODE) y con -ErrorAction Stop donde importa.
$ErrorActionPreference = 'Continue'

$aqui = $PSScriptRoot
$raiz = Split-Path $aqui -Parent
$venv = Join-Path $raiz '.venv'
$cfg = Join-Path $venv 'pyvenv.cfg'
$requisitos = Join-Path $aqui 'requirements.txt'
# Lo que la aplicacion necesita importar para arrancar.
$prueba = 'import tkinter, win32com.client, openpyxl, pypdf, pytesseract, PIL'  # sin comillas dobles: PowerShell 5.1 las pierde al pasarlas a un .exe

function Salir($codigo) {
    if ($Host.Name -eq 'ConsoleHost' -and -not $env:EXCELAGENT_SIN_PAUSA) { Read-Host 'Enter para cerrar' | Out-Null }
    exit $codigo
}

# Version de un python.exe, preguntandole a el mismo (asi ademas se comprueba que arranca). $null si no anda.
function VersionDe($exe) {
    try {
        $v = & $exe -c "import sys; print('%d.%d.%d' % sys.version_info[:3])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $v -match '^\d+\.\d+\.\d+$') { return [version]$v }
    } catch {}
    return $null
}

# Todos los python.exe instalados: registro (PEP 514, lo que usa el lanzador py) y carpetas habituales.
function BuscarPythons {
    $rutas = @()
    foreach ($base in 'HKCU:\Software\Python\PythonCore', 'HKLM:\Software\Python\PythonCore',
                      'HKLM:\Software\WOW6432Node\Python\PythonCore') {
        if (-not (Test-Path $base)) { continue }
        foreach ($k in Get-ChildItem $base -ErrorAction SilentlyContinue) {
            $ip = Join-Path $k.PSPath 'InstallPath'
            if (-not (Test-Path $ip)) { continue }
            $p = Get-ItemProperty $ip -ErrorAction SilentlyContinue
            if ($p.ExecutablePath) { $rutas += $p.ExecutablePath }
            elseif ($p.'(default)') { $rutas += (Join-Path $p.'(default)' 'python.exe') }
        }
    }
    foreach ($patron in "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
                        "$env:ProgramFiles\Python3*\python.exe") {
        $rutas += (Get-ChildItem $patron -ErrorAction SilentlyContinue | ForEach-Object FullName)
    }
    $vistos = @{}
    foreach ($r in $rutas) {
        if (-not $r -or -not (Test-Path $r)) { continue }
        $r = (Resolve-Path $r).Path
        if ($vistos.ContainsKey($r.ToLower())) { continue }
        $vistos[$r.ToLower()] = $true
        $v = VersionDe $r
        if ($v -and $v -ge [version]'3.10') { [pscustomobject]@{ Ruta = $r; Version = $v } }
    }
}

# $callado: en la primera mirada un entorno roto es lo esperable; no hace falta mostrar el error crudo de python.
function ProbarEntorno([switch]$callado) {
    $py = Join-Path $venv 'Scripts\python.exe'
    if (-not (Test-Path $py)) { return $false }
    $r = & $py -c $prueba 2>&1
    if ($LASTEXITCODE -ne 0) { if (-not $callado) { Write-Host ($r | Out-String) }; return $false }
    # Que el codigo de la aplicacion compile con este Python (sintaxis nueva en uno viejo, por ejemplo).
    $r = & $py -m compileall -q -x '[\\/](tests|__pycache__)[\\/]' $aqui 2>&1
    if ($LASTEXITCODE -ne 0) { Write-Host ($r | Out-String); return $false }
    return $true
}

# --- 1. Python de esta PC ---
$pythons = @(BuscarPythons | Sort-Object Version -Descending)
if ($pythons.Count -eq 0) {
    Write-Host 'No se encontro ningun Python 3.10 o posterior instalado en esta PC.' -ForegroundColor Red
    Write-Host 'Instalarlo desde python.org (marcar "tcl/tk" en la instalacion) y volver a correr esto.'
    Salir 1
}
Write-Host 'Pythons encontrados:'
$pythons | ForEach-Object { Write-Host ('  {0}  {1}' -f $_.Version, $_.Ruta) }

# --- 2. El entorno actual, si ya anda, no se toca ---
if ((Test-Path $cfg) -and (ProbarEntorno -callado)) {
    Write-Host "El entorno $venv ya funciona con esta PC. No se cambio nada." -ForegroundColor Green
    Salir 0
}

$versionVenv = $null
if (Test-Path $cfg) {
    $linea = Get-Content $cfg | Where-Object { $_ -match '^\s*version(_info)?\s*=' } | Select-Object -First 1
    if ($linea -match '(\d+\.\d+)') { $versionVenv = [version]$Matches[1] }
    Write-Host "Entorno actual: Python $versionVenv, pyvenv.cfg apunta a:"
    Get-Content $cfg | Where-Object { $_ -match '^\s*(home|executable)\s*=' } | ForEach-Object { Write-Host "  $_" }
}

# --- 3. Misma version menor: alcanza con corregir pyvenv.cfg ---
if ($versionVenv) {
    $igual = $pythons | Where-Object { $_.Version.Major -eq $versionVenv.Major -and $_.Version.Minor -eq $versionVenv.Minor } |
        Select-Object -First 1
    if ($igual) {
        $copia = "$cfg.respaldo-$(Get-Date -Format yyyyMMdd-HHmmss)"
        Copy-Item $cfg $copia -ErrorAction Stop
        $home_ = Split-Path $igual.Ruta -Parent
        $texto = Get-Content $cfg | ForEach-Object {
            if ($_ -match '^\s*home\s*=') { "home = $home_" }
            elseif ($_ -match '^\s*executable\s*=') { "executable = $($igual.Ruta)" }
            else { $_ }
        }
        [IO.File]::WriteAllLines($cfg, [string[]]$texto)   # sin BOM: el lanzador del venv lo lee como texto plano
        if (ProbarEntorno) {
            Write-Host "Listo: pyvenv.cfg ahora apunta a $($igual.Ruta). Respaldo: $copia" -ForegroundColor Green
            Salir 0
        }
        Write-Host 'Corregir pyvenv.cfg no alcanzo; se vuelve al original y se arma un entorno nuevo.' -ForegroundColor Yellow
        Copy-Item $copia $cfg -Force
    }
}

# --- 4. Entorno nuevo con el Python mas reciente, guardando el anterior ---
$elegido = $pythons[0]
$respaldo = $null
if (Test-Path $venv) {
    $respaldo = "$venv.respaldo-$(Get-Date -Format yyyyMMdd-HHmmss)"
    Rename-Item $venv (Split-Path $respaldo -Leaf) -ErrorAction Stop
    Write-Host "Entorno anterior guardado en $respaldo"
}
$ok = $false
try {
    Write-Host "Creando el entorno con Python $($elegido.Version)..."
    & $elegido.Ruta -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw "python -m venv devolvio $LASTEXITCODE" }
    Write-Host 'Instalando bibliotecas (hace falta internet)...'
    & (Join-Path $venv 'Scripts\python.exe') -m pip install --disable-pip-version-check -r $requisitos
    if ($LASTEXITCODE -ne 0) { throw "pip install devolvio $LASTEXITCODE" }
    if (-not (ProbarEntorno)) { throw 'el entorno nuevo no pasa la prueba de importar las bibliotecas' }
    $ok = $true
} catch {
    Write-Host "Fallo: $_" -ForegroundColor Red
}

if ($ok) {
    Write-Host "Listo: entorno nuevo con Python $($elegido.Version) en $venv" -ForegroundColor Green
    if ($respaldo) { Write-Host "El anterior quedo en $respaldo; se puede borrar a mano cuando se confirme que todo anda." }
    Salir 0
}
# Rollback: se descarta lo que se armo recien y vuelve el anterior, tal como estaba.
if (Test-Path $venv) { Remove-Item $venv -Recurse -Force }
if ($respaldo) { Rename-Item $respaldo (Split-Path $venv -Leaf); Write-Host 'Se restauro el entorno anterior sin cambios.' }
Salir 1
