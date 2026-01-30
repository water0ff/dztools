param(
    [string]$Branch = "release"
)
Write-Host "`n==============================================" -ForegroundColor Red
Write-Host "           ADVERTENCIA DE VERSIÓN BETA " -ForegroundColor Red
Write-Host "==============================================" -ForegroundColor Red
Write-Host "Esta aplicación se encuentra en fase de desarrollo BETA.`n" -ForegroundColor Yellow
Write-Host "Algunas funciones pueden realizar cambios irreversibles en: `n"
Write-Host " - Su equipo" -ForegroundColor Red
Write-Host " - Bases de datos" -ForegroundColor Red
Write-Host " - Configuraciones del sistema`n" -ForegroundColor Red
Write-Host "¿Acepta ejecutar esta aplicación bajo su propia responsabilidad? (Y/N)" -ForegroundColor Yellow
$response = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
$answer = $response.Character.ToString().ToUpper()
while ($answer -notin 'Y', 'N') {
    $response = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    $answer = $response.Character.ToString().ToUpper()
}
if ($answer -ne 'Y') {
    Write-Host "`nEjecución cancelada por el usuario.`n" -ForegroundColor Red
    return
}
Clear-Host
$baseRuntimePath = "C:\temp\dztools"
$releasePath = Join-Path $baseRuntimePath "release"
$Owner = "water0ff"
$Repo = "dztools"
if (-not (Test-Path $baseRuntimePath)) {
    New-Item -ItemType Directory -Path $baseRuntimePath | Out-Null
}
Write-Host "Preparando entorno..." -ForegroundColor Yellow

$zipPath = Join-Path $baseRuntimePath "dztools.zip"

Write-Host "Limpiando versión anterior..." -ForegroundColor Yellow
try {
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path $releasePath) { Remove-Item $releasePath -Recurse -Force -ErrorAction SilentlyContinue }
} catch {}
$zipUrl = "https://github.com/$Owner/$Repo/releases/latest/download/dztools-release.zip"

Write-Host "Descargando última versión..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Write-Host "  ✓ Descarga completada" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Error al descargar: $($_.Exception.Message)" -ForegroundColor Red
    return
}
Write-Host "Extrayendo archivos..." -ForegroundColor Yellow
try {
    if (-not (Test-Path $releasePath)) {
        New-Item -ItemType Directory -Path $releasePath | Out-Null
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $releasePath)
    Write-Host "  ✓ Extracción completada" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Error al extraer: $($_.Exception.Message)" -ForegroundColor Red
    return
}

Write-Host "Preparando aplicación..." -ForegroundColor Yellow
$projectRoot = $releasePath
$mainPath = Join-Path $projectRoot "main.ps1"
if (-not (Test-Path $mainPath)) {
    Write-Host "  ✗ No se encontró main.ps1 en la carpeta release." -ForegroundColor Red
    Write-Host "  Ruta esperada: $mainPath" -ForegroundColor DarkYellow
    return
}
Write-Host "  ✓ Listo" -ForegroundColor Green
Write-Host ""
Write-Host "=================================================" -ForegroundColor Gray
Write-Host "   Iniciando Gerardo Zermeño Tools desde GitHub" -ForegroundColor Green
Write-Host "   Canal: $Branch" -ForegroundColor DarkGray
Write-Host "   Carpeta: $projectRoot" -ForegroundColor DarkGray
Write-Host "=================================================" -ForegroundColor Gray
Write-Host ""

$exe = if ($PSVersionTable.PSVersion.Major -ge 6) { "pwsh" } else { "powershell" }
& $exe -NoProfile -ExecutionPolicy Bypass -File $mainPath