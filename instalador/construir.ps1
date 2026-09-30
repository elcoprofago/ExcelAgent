# Arma el instalador de ExcelAgent: ..\..\_instalador\ExcelAgent-Setup-<VERSION>.exe (fuera del repo).
#
#   .\instalador\construir.ps1                 usa el Python 3.13 de esta PC como base
#   .\instalador\construir.ps1 -Python <exe>   otra base (tiene que ser Python de 64 bits de python.org)
#
# Pasos, y por que:
#   1. copia el Python base a _instalador\etapa\python, sin su site-packages ni lo que no hace falta (documentacion,
#      pruebas de la biblioteca estandar, Scripts con rutas absolutas). Un Python de python.org es relocalizable: busca
#      su biblioteca al lado del .exe, asi que funciona en cualquier carpeta de cualquier PC;
#   2. le instala pip y requirements.txt adentro, con -E -s para que no se mezcle nada de esta PC;
#   3. copia el programa (los archivos del repo, sin los ignorados por git) y ..\tesseract;
#   4. verifica la etapa: que las bibliotecas importen, que el programa compile y que encuentre tesseract;
#   5. compila ExcelAgent.iss con la VERSION de version.py (una sola fuente: el instalador no tiene la suya).
param(
    [string]$Python
)
# Continue y no Stop: en PowerShell 5.1 lo que python o git escriben en stderr seria un error fatal.
# Cada paso se juzga por su codigo de salida.
$ErrorActionPreference = 'Continue'

function Falla($texto) {
    Write-Host "NO SE ARMO EL INSTALADOR: $texto" -ForegroundColor Red
    exit 1
}
function Paso($texto) { Write-Host "== $texto" -ForegroundColor Cyan }

$programa = Split-Path $PSScriptRoot -Parent                  # ...\EXCEL\ExcelAgent (el repo)
$salida = Join-Path (Split-Path $programa -Parent) '_instalador'
$etapa = Join-Path $salida 'etapa'
$iscc = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $iscc)) { Falla "no esta Inno Setup 6 ($iscc)" }

$m = Select-String -LiteralPath (Join-Path $programa 'version.py') -Pattern "^VERSION\s*=\s*'(\d+\.\d+\.\d+)'"
if (-not $m) { Falla 'no encuentro VERSION en version.py' }
$version = $m.Matches[0].Groups[1].Value

# --- Python base ---
if (-not $Python) {
    $Python = (& py -3.13 -c "import sys; print(sys.executable)" 2>$null)
    if ($LASTEXITCODE -ne 0 -or -not $Python) { $Python = "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe" }
}
if (-not (Test-Path $Python)) { Falla "no encuentro el Python base: $Python" }
$base = Split-Path $Python -Parent
$datos = & $Python -E -s -c "import sys, struct; print('%d.%d' % sys.version_info[:2], struct.calcsize('P') * 8, sys.prefix == sys.base_prefix)"
if ($LASTEXITCODE -ne 0) { Falla "el Python base no arranca: $Python" }
$pyver, $bits, $noEsVenv = $datos -split ' '
if ($bits -ne '64') { Falla "el Python base es de $bits bits; el instalador es de 64" }
if ($noEsVenv -ne 'True') { Falla 'el Python base es un entorno virtual; hace falta el Python instalado' }
$pyver_ = $pyver -replace '\.', ''
if (-not (Test-Path (Join-Path $base "python$pyver_.dll"))) { Falla "no es un Python de python.org (falta python$pyver_.dll)" }

$sucio = & git -C $programa status --porcelain
if ($sucio) {
    Write-Host 'AVISO: el repo tiene cambios sin commitear; el instalador lleva la carpeta tal como esta:' -ForegroundColor Yellow
    $sucio | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
}

Paso "ExcelAgent $version con Python $pyver ($base)"

# --- 1. etapa limpia: solo se borra la carpeta de etapa de este script ---
if (Test-Path $etapa) {
    if ((Split-Path $etapa -Leaf) -ne 'etapa' -or (Split-Path (Split-Path $etapa -Parent) -Leaf) -ne '_instalador') { Falla "ruta de etapa inesperada: $etapa" }
    Remove-Item $etapa -Recurse -Force
    if (Test-Path $etapa) { Falla "no pude vaciar $etapa (algun archivo abierto?)" }
}
New-Item -ItemType Directory $etapa | Out-Null

Paso 'Copiando el Python base'
$py = Join-Path $etapa 'python'
robocopy $base $py /E /XJ /NFL /NDL /NJH /NJS /NP `
    /XD "$base\Doc" "$base\Lib\test" "$base\Lib\idlelib" "$base\Lib\site-packages" "$base\Scripts" "$base\include" "$base\libs" __pycache__ `
    /XF *._pth | Out-Null
if ($LASTEXITCODE -ge 8) { Falla "robocopy del Python devolvio $LASTEXITCODE" }
$pyexe = Join-Path $py 'python.exe'

# --- 2. pip y bibliotecas ---
Paso 'Instalando pip y las bibliotecas'
& $pyexe -E -s -m ensurepip --default-pip
if ($LASTEXITCODE -ne 0) { Falla "ensurepip devolvio $LASTEXITCODE" }
& $pyexe -E -s -m pip install --disable-pip-version-check --no-warn-script-location -r (Join-Path $programa 'requirements.txt')
if ($LASTEXITCODE -ne 0) { Falla "pip install devolvio $LASTEXITCODE" }
# Los .exe de Scripts llevan la ruta absoluta de la etapa escrita adentro: en otra carpeta no sirven.
Remove-Item (Join-Path $py 'Scripts') -Recurse -Force -ErrorAction SilentlyContinue

# --- 3. programa y tesseract ---
Paso 'Copiando el programa'
$lista = & git -C $programa ls-files --cached --others --exclude-standard
if ($LASTEXITCODE -ne 0 -or -not $lista) { Falla 'git ls-files fallo' }
$dest = Join-Path $etapa 'ExcelAgent'
foreach ($r in $lista) {
    $origen = Join-Path $programa $r
    if (-not (Test-Path -LiteralPath $origen)) { continue }          # borrado en el arbol y todavia no commiteado
    $d = Join-Path $dest $r
    New-Item -ItemType Directory (Split-Path $d -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $origen $d -ErrorAction Stop
}

Paso 'Copiando tesseract'
$tess = Join-Path (Split-Path $programa -Parent) 'tesseract'
if (-not (Test-Path (Join-Path $tess 'tesseract.exe'))) { Falla "falta $tess\tesseract.exe" }
if (-not (Test-Path (Join-Path $tess 'tessdata\spa.traineddata'))) { Falla "falta el idioma espanol en $tess\tessdata" }
robocopy $tess (Join-Path $etapa 'tesseract') /E /XJ /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { Falla "robocopy de tesseract devolvio $LASTEXITCODE" }

Paso 'Icono'
$ico = Join-Path $etapa 'ExcelAgent.ico'
$dibujo = @'
import sys
from PIL import Image, ImageDraw
im = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
d.rounded_rectangle((8, 8, 248, 248), 48, fill=(16, 124, 65))
for i in range(3):
    y = 64 + i * 48
    d.rectangle((52, y, 204, y + 28), fill=(255, 255, 255))
    d.rectangle((110, y, 116, y + 28), fill=(16, 124, 65))
im.save(sys.argv[1], sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
'@
$dibujo | & $pyexe -E -s - $ico
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $ico)) { Falla 'no pude generar el icono' }

# --- 4. verificar la etapa ---
Paso 'Verificando la etapa'
& $pyexe -E -s -c 'import tkinter, win32com.client, pythoncom, pywintypes, openpyxl, pypdf, pytesseract, PIL'
if ($LASTEXITCODE -ne 0) { Falla 'las bibliotecas no importan en el Python de la etapa' }
& $pyexe -E -s -m compileall -q -x '[\\/]tests[\\/]' $dest
if ($LASTEXITCODE -ne 0) { Falla 'el programa no compila con el Python de la etapa' }
Push-Location $dest
$t = & $pyexe -E -s -c 'import os, excel; print(os.path.normcase(os.path.abspath(str(excel.ruta_tesseract()))))'
Pop-Location
if ($t -ne [IO.Path]::GetFullPath((Join-Path $etapa 'tesseract\tesseract.exe')).ToLower()) { Falla "el programa no encuentra el tesseract de la etapa (encontro: $t)" }
# El Python de la etapa no tiene que haber quedado apuntando a esta PC: ninguna ruta de sys.path fuera de la etapa.
$fuera = & $pyexe -E -s -c "import sys, os; e = os.path.normcase(sys.argv[1]); print([p for p in sys.path if p and not os.path.normcase(os.path.abspath(p)).startswith(e)])" $py
if ($fuera -ne '[]') { Falla "sys.path del Python de la etapa sale de la etapa: $fuera" }

# --- 5. compilar ---
Paso 'Compilando el instalador'
$nombre = "ExcelAgent-Setup-$version"
& $iscc /Q "/DMiVersion=$version" "/DEtapa=$etapa" "/O$salida" "/F$nombre" (Join-Path $PSScriptRoot 'ExcelAgent.iss')
if ($LASTEXITCODE -ne 0) { Falla "ISCC devolvio $LASTEXITCODE" }
$exe = Join-Path $salida "$nombre.exe"
if (-not (Test-Path $exe)) { Falla "ISCC termino bien pero no esta $exe" }
$mb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
Write-Host "Listo: $exe ($mb MB)" -ForegroundColor Green
