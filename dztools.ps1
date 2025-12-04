param(
    [string]$Branch = "release"   # solo informativo ahora
)
$baseRuntimePath = "C:\temp\dztools"
$releasePath = Join-Path $baseRuntimePath "release"
$Owner = "water0ff"
$Repo = "dztools"
Clear-Host
function Show-ProgressBar {
    param(
        [int]$Percent,
        [string]$Message = ""
    )
    $width = 20
    if ($Percent -lt 0) { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }
    $filled = [math]::Round(($Percent / 100) * $width)
    $bar = "[" + ("=" * $filled).PadRight($width) + "]"
    $line = "{0} {1,3}%  {2}" -f $bar, $Percent, $Message
    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $line = $line.PadRight($consoleWidth - 1)
    Write-Host "`r$line" -NoNewline
    if ($Percent -ge 100) {
        Write-Host ""
    }
}
if (-not (Test-Path $baseRuntimePath)) {
    New-Item -ItemType Directory -Path $baseRuntimePath | Out-Null
}
Show-ProgressBar -Percent 5 -Message "Preparando entorno..."
$zipPath = Join-Path $baseRuntimePath "dztools.zip"
Show-ProgressBar -Percent 10 -Message "Limpiando versión anterior..."

try {
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $releasePath) {
        Remove-Item $releasePath -Recurse -Force -ErrorAction SilentlyContinue
    }
} catch {
    # Si algo truena limpiando, no es fatal, seguimos intentando
}
$zipUrl = "https://github.com/$Owner/$Repo/releases/latest/download/dztools-release.zip"
Show-ProgressBar -Percent 20 -Message "Descargando última versión..."
try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
} catch {
    Show-ProgressBar -Percent 100 -Message "Error al descargar"
    Write-Host ""
    Write-Host "❌ No se pudo descargar el release: $($_.Exception.Message)" -ForegroundColor Red
    return
}
Show-ProgressBar -Percent 50 -Message "Extrayendo archivos..."
try {
    if (-not (Test-Path $releasePath)) {
        New-Item -ItemType Directory -Path $releasePath | Out-Null
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $releasePath)
} catch {
    Show-ProgressBar -Percent 100 -Message "Error al extraer"
    Write-Host ""
    Write-Host "❌ No se pudo extraer el ZIP: $($_.Exception.Message)" -ForegroundColor Red
    return
}
Show-ProgressBar -Percent 70 -Message "Preparando aplicación..."
$projectRoot = $releasePath
$mainPath = Join-Path $projectRoot "main.ps1"
if (-not (Test-Path $mainPath)) {
    Show-ProgressBar -Percent 100 -Message "Error"
    Write-Host ""
    Write-Host "❌ No se encontró main.ps1 en la carpeta release." -ForegroundColor Red
    Write-Host "Ruta esperada: $mainPath" -ForegroundColor DarkYellow
    return
}
Show-ProgressBar -Percent 100 -Message "Listo"
Write-Host ""
Write-Host "=============================================" -ForegroundColor Gray
Write-Host "   Iniciando Daniel Tools desde GitHub" -ForegroundColor Green
Write-Host "   Canal: $Branch" -ForegroundColor DarkGray
Write-Host "   Carpeta: $projectRoot" -ForegroundColor DarkGray
Write-Host "=============================================" -ForegroundColor Gray
Write-Host ""
powershell -ExecutionPolicy Bypass -NoProfile -File $mainPath