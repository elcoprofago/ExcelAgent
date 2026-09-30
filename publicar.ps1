# Publica una version nueva de ExcelAgent como release de GitHub (la que baja el boton Actualizar).
#
#   1. subir VERSION en version.py y commitear;
#   2. desde la rama main, con todo commiteado:   .\publicar.ps1 -Notas "Que trae esta version"
#
# Se niega (y no toca nada) si: hay cambios sin commitear, no se esta en main, main no coincide con origin/main
# despues del push, la etiqueta v<VERSION> ya existe, o VERSION no es mayor que la ultima release publicada.
# Es el control de "version.py y la etiqueta cambian juntas": el boton Actualizar, ademas, rechaza en tiempo de
# ejecucion una release cuya etiqueta no coincida con la VERSION de adentro.
#
# Tambien arma el instalador (instalador\construir.ps1) y lo adjunta: ExcelAgent-Setup-<VERSION>.exe. Se arma ANTES
# de subir nada, asi que si falla no se publica nada. Va en el mismo gh release create que crea la release, y al final
# se lo baja de GitHub y se compara su SHA-256 con el armado: se verifica lo publicado, no el codigo de salida.
param(
    [Parameter(Mandatory = $true)][string]$Notas
)
# Continue y no Stop: en PowerShell 5.1 lo que git escribe en stderr (el avance de un push) seria un error fatal.
# Cada paso se juzga por su codigo de salida.
Set-Location -LiteralPath $PSScriptRoot

function Falla($texto) {
    Write-Host "NO PUBLICADO: $texto" -ForegroundColor Red
    exit 1
}
# Para lo que falla cuando la etiqueta ya esta en GitHub: decir "no publicado" seria falso.
function FallaAMedias($texto) {
    Write-Host "PUBLICADO A MEDIAS: $texto" -ForegroundColor Red
    exit 1
}
function Correr-Git {
    $salida = & git.exe @args 2>&1
    if ($LASTEXITCODE -ne 0) { Falla ("git $($args -join ' ') -> " + ($salida -join ' ')) }
    return $salida
}

$gh = 'C:\Program Files\GitHub CLI\gh.exe'
if (-not (Test-Path -LiteralPath $gh)) { $gh = 'gh' }
# Solo para probar este script sin tocar GitHub: otro gh (uno falso).
if ($env:PUBLICAR_GH) { $gh = $env:PUBLICAR_GH }

$m = Select-String -LiteralPath 'version.py' -Pattern "^VERSION\s*=\s*'(\d+\.\d+\.\d+)'"
if (-not $m) { Falla 'no encuentro VERSION en version.py' }
$version = $m.Matches[0].Groups[1].Value
$etiqueta = "v$version"
$repo = (Select-String -LiteralPath 'version.py' -Pattern "^REPO\s*=\s*'([^']+)'").Matches[0].Groups[1].Value

if (Correr-Git status --porcelain) { Falla 'hay cambios sin commitear' }
$rama = (Correr-Git rev-parse --abbrev-ref HEAD) -join ''
if ($rama -ne 'main') { Falla "se publica desde main (estas en $rama)" }
if (Correr-Git tag --list $etiqueta) { Falla "la etiqueta $etiqueta ya existe: subi VERSION en version.py" }
if (Correr-Git ls-remote --tags origin "refs/tags/$etiqueta") { Falla "la etiqueta $etiqueta ya existe en GitHub" }

$ultima = & $gh release view --repo $repo --json tagName --jq .tagName 2>$null
if ($LASTEXITCODE -eq 0 -and $ultima) {
    if ([version]($ultima.TrimStart('v')) -ge [version]$version) {
        Falla "la ultima release es $ultima y version.py dice ${version}: tiene que ser mayor"
    }
}

# --- instalador, antes de subir nada ---
$inicio = Get-Date
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'instalador\construir.ps1')
if ($LASTEXITCODE -ne 0) { Falla "no se pudo armar el instalador (construir.ps1 devolvio $LASTEXITCODE)" }
$instalador = Join-Path (Split-Path $PSScriptRoot -Parent) "_instalador\ExcelAgent-Setup-$version.exe"
if (-not (Test-Path -LiteralPath $instalador)) { Falla "construir.ps1 termino bien pero no esta $instalador" }
if ((Get-Item -LiteralPath $instalador).LastWriteTime -lt $inicio) { Falla "$instalador es de una corrida anterior, no de esta" }
if (Correr-Git status --porcelain) { Falla 'armar el instalador dejo cambios en el repo' }
$hashLocal = (Get-FileHash -LiteralPath $instalador -Algorithm SHA256).Hash

Correr-Git push origin main | Out-Null
if (((Correr-Git rev-parse HEAD) -join '') -ne ((Correr-Git rev-parse origin/main) -join '')) { Falla 'main no quedo igual a origin/main' }
Correr-Git tag -a $etiqueta -m "ExcelAgent $version" | Out-Null
Correr-Git push origin $etiqueta | Out-Null

& $gh release create $etiqueta $instalador --repo $repo --title "ExcelAgent $version" --notes $Notas --verify-tag
if ($LASTEXITCODE -ne 0) { FallaAMedias "la etiqueta $etiqueta quedo subida pero no se creo la release; reintentar: gh release create $etiqueta `"$instalador`"" }

# verificar lo publicado, no el codigo de salida: lo mismo que va a consultar el boton Actualizar
$publicada = & $gh api "repos/$repo/releases/latest" --jq .tag_name
if ($publicada -ne $etiqueta) { FallaAMedias "GitHub dice que la ultima release es '$publicada', no $etiqueta" }

# el instalador publicado: bajarlo y compararlo (por SHA-256) con el que se armo
$archivo = Split-Path $instalador -Leaf
$reintento = "gh release upload $etiqueta `"$instalador`" --repo $repo --clobber"
$baja = Join-Path ([IO.Path]::GetTempPath()) ('excelagent-publicar-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $baja | Out-Null
& $gh release download $etiqueta --repo $repo --pattern $archivo --dir $baja
$bajado = Join-Path $baja $archivo
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $bajado)) { FallaAMedias "la release $etiqueta quedo publicada pero no pude bajar su instalador. Subirlo: $reintento" }
$hashPublicado = (Get-FileHash -LiteralPath $bajado -Algorithm SHA256).Hash
Remove-Item -LiteralPath $baja -Recurse -Force
if ($hashPublicado -ne $hashLocal) { FallaAMedias "el instalador publicado en $etiqueta no coincide con el armado. Volver a subirlo: $reintento" }
Write-Host "Instalador verificado en GitHub: $archivo (SHA-256 $hashLocal)" -ForegroundColor Green
Write-Host "Publicada $etiqueta en https://github.com/$repo/releases/tag/$etiqueta" -ForegroundColor Green
