param(
    [string]$Branch = "develop",
    [switch]$KeepFiles  # Si quieres conservar los archivos para futuras ejecuciones
)

# Configuración
$Owner = "water0ff"
$Repo  = "dztools"

# Carpeta donde se va a descargar y ejecutar
# Opción A: carpeta temporal (no ensucia nada)
$baseRuntimePath = Join-Path $env:TEMP "dztools-runtime"

# Si quieres que se quede instalado por usuario, puedes cambiarlo a algo tipo:
# $baseRuntimePath = Join-Path $env:LOCALAPPDATA "dztools"

# Limpia runtime si no queremos conservar archivos
if (-not $KeepFiles -and (Test-Path $baseRuntimePath)) {
    try {
        Remove-Item $baseRuntimePath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Host "⚠ No se pudo limpiar el runtime anterior: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if (-not (Test-Path $baseRuntimePath)) {
    New-Item -ItemType Directory -Path $baseRuntimePath | Out-Null
}

# Descargar ZIP del repo
$zipUrl  = "https://github.com/$Owner/$Repo/archive/refs/heads/$Branch.zip"
$zipPath = Join-Path $baseRuntimePath "dztools.zip"

Write-Host "Descargando Daniel Tools desde $zipUrl ..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath

# Extraer ZIP
Write-Host "Extrayendo archivos..." -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $baseRuntimePath)

# Buscar carpeta extraída (dztools-<branch> o similar)
$extractedFolder = Get-ChildItem $baseRuntimePath -Directory |
    Where-Object { $_.Name -like "$Repo-*" } |
    Select-Object -First 1

if (-not $extractedFolder) {
    Write-Host "❌ No se encontró la carpeta extraída del repositorio." -ForegroundColor Red
    return
}

$projectRoot = $extractedFolder.FullName
$mainPath    = Join-Path $projectRoot "src\main.ps1"

if (-not (Test-Path $mainPath)) {
    Write-Host "❌ No se encontró src\main.ps1 en el proyecto." -ForegroundColor Red
    return
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Gray
Write-Host "   Iniciando Daniel Tools desde GitHub" -ForegroundColor Green
Write-Host "   Rama: $Branch" -ForegroundColor DarkGray
Write-Host "=============================================" -ForegroundColor Gray
Write-Host ""

# Ejecutar la herramienta
& $mainPath

# Limpieza final (solo si NO queremos conservar archivos)
if (-not $KeepFiles) {
    Write-Host "Limpiando archivos temporales..." -ForegroundColor DarkGray
    try {
        Remove-Item $baseRuntimePath -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Host "⚠ No se pudo eliminar la carpeta temporal: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
