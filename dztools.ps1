param(
    [string]$Branch = "develop"
)

# 1) Limpiar pantalla al inicio
Clear-Host

# 2) Función para mostrar barra de progreso simple
function Show-ProgressBar {
    param(
        [int]$Percent,
        [string]$Message = ""
    )

    $width = 20  # ancho de la barra
    if ($Percent -lt 0) { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }

    $filled = [math]::Round(($Percent / 100) * $width)
    $bar = "[" + ("=" * $filled).PadRight($width) + "]"

    $line = "{0} {1,3}%  {2}" -f $bar, $Percent, $Message
    Write-Host "`r$line" -NoNewline

    if ($Percent -ge 100) {
        Write-Host ""  # salto de línea al terminar
    }
}

# Configuración del repo
$Owner = "water0ff"
$Repo  = "dztools"

# 3) Carpeta PERSISTENTE (ya no se borra)
#    - Se reutiliza para que arranque más rápido si ya existe
$baseRuntimePath = Join-Path $env:LOCALAPPDATA "dztools-runtime"

if (-not (Test-Path $baseRuntimePath)) {
    New-Item -ItemType Directory -Path $baseRuntimePath | Out-Null
}

Show-ProgressBar -Percent 5 -Message "Revisando instalación previa..."

# 4) Ver si ya tenemos una versión descargada usable
$existingRoot = Get-ChildItem $baseRuntimePath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$Repo-*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

$projectRoot = $null
$mainPath    = $null

if ($existingRoot) {
    $projectRoot = $existingRoot.FullName
    $mainPath    = Join-Path $projectRoot "src\main.ps1"

    if (Test-Path $mainPath) {
        Show-ProgressBar -Percent 30 -Message "Usando versión ya descargada"
        goto ExecuteApp
    }
}

# Si llegamos aquí, no hay versión usable -> descargar
$zipUrl  = "https://github.com/$Owner/$Repo/archive/refs/heads/$Branch.zip"
$zipPath = Join-Path $baseRuntimePath "dztools.zip"

Show-ProgressBar -Percent 15 -Message "Descargando última versión..."

try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
}
catch {
    Show-ProgressBar -Percent 100 -Message "Error al descargar"
    Write-Host ""
    Write-Host "❌ No se pudo descargar el repositorio: $($_.Exception.Message)" -ForegroundColor Red
    return
}

Show-ProgressBar -Percent 50 -Message "Extrayendo archivos..."

# Opcional: limpiar versiones anteriores del mismo repo (sin borrar todo el runtime)
Get-ChildItem $baseRuntimePath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$Repo-*" } |
    ForEach-Object {
        try {
            Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop
        } catch {
            # si algo falla, lo ignoramos y seguimos
        }
    }

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $baseRuntimePath)
}
catch {
    Show-ProgressBar -Percent 100 -Message "Error al extraer"
    Write-Host ""
    Write-Host "❌ No se pudo extraer el ZIP: $($_.Exception.Message)" -ForegroundColor Red
    return
}

Show-ProgressBar -Percent 70 -Message "Localizando proyecto..."

$extractedFolder = Get-ChildItem $baseRuntimePath -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$Repo-*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $extractedFolder) {
    Show-ProgressBar -Percent 100 -Message "Error"
    Write-Host ""
    Write-Host "❌ No se encontró la carpeta extraída del repositorio." -ForegroundColor Red
    return
}

$projectRoot = $extractedFolder.FullName
$mainPath    = Join-Path $projectRoot "src\main.ps1"

Show-ProgressBar -Percent 90 -Message "Preparando aplicación..."

if (-not (Test-Path $mainPath)) {
    Show-ProgressBar -Percent 100 -Message "Error"
    Write-Host ""
    Write-Host "❌ No se encontró src\main.ps1 en el proyecto." -ForegroundColor Red
    return
}

Show-ProgressBar -Percent 100 -Message "Listo"

:ExecuteApp

Write-Host ""
Write-Host "=============================================" -ForegroundColor Gray
Write-Host "   Iniciando Daniel Tools desde GitHub" -ForegroundColor Green
Write-Host "   Rama: $Branch" -ForegroundColor DarkGray
Write-Host "   Carpeta: $projectRoot" -ForegroundColor DarkGray
Write-Host "=============================================" -ForegroundColor Gray
Write-Host ""

# 5) Ejecutar el main del proyecto
& $mainPath
