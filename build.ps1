#requires -Version 5.0

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("patch", "minor", "major")]
    [string]$Bump = "patch",
    
    [Parameter(Mandatory = $false)]
    [switch]$Test,
    
    [Parameter(Mandatory = $false)]
    [switch]$Release
)

Write-Host "=== Daniel Tools Build Script ===" -ForegroundColor Cyan
Write-Host "Directorio: $PSScriptRoot" -ForegroundColor Yellow
Write-Host "PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow

$ErrorActionPreference = "Stop"
$ProjectRoot = $PSScriptRoot
$SrcPath = Join-Path $ProjectRoot "src"
$ReleasePath = Join-Path $ProjectRoot "release"
$VersionFile = Join-Path $SrcPath "version.json"

function Get-CurrentVersion {
    if (Test-Path $VersionFile) {
        try {
            $content = Get-Content $VersionFile -Raw
            $data = $content | ConvertFrom-Json
            return [version]$data.Version
        }
        catch {
            Write-Host "Error leyendo version.json: $_" -ForegroundColor Yellow
            return [version]"0.1.0"
        }
    }
    return [version]"0.1.0"
}

if ($Test) {
    Write-Host "`nEJECUTANDO PRUEBAS..." -ForegroundColor Green
    
    if (-not (Get-Module -Name Pester -ListAvailable)) {
        Write-Host "Instalando Pester..." -ForegroundColor Yellow
        Install-Module -Name Pester -Force -Scope CurrentUser -AllowClobber
    }
    
    if (-not (Test-Path ".\tests")) {
        Write-Host "Creando carpeta tests básica..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path ".\tests" -Force | Out-Null
        
        @'
# Pruebas básicas
Describe "Pruebas iniciales" {
    It "Debe existir el archivo main.ps1" {
        Test-Path ".\src\main.ps1" | Should -Be $true
    }
    
    It "Deben existir los módulos" {
        $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Installers.psm1")
        foreach ($module in $modules) {
            Test-Path ".\src\modules\$module" | Should -Be $true
        }
    }
}
'@ | Out-File -FilePath ".\tests\Basic.Tests.ps1" -Encoding UTF8
    }
    
    try {
        $result = Invoke-Pester -Path ".\tests" -PassThru
        if ($result.FailedCount -gt 0) {
            Write-Host "Pruebas fallidas: $($result.FailedCount)" -ForegroundColor Red
            exit 1
        }
        Write-Host "✓ Todas las pruebas pasaron ($($result.PassedCount) pruebas)" -ForegroundColor Green
    }
    catch {
        Write-Host "Error ejecutando pruebas: $_" -ForegroundColor Red
        exit 1
    }
}

if ($Release) {
    Write-Host "`nCREANDO RELEASE..." -ForegroundColor Green
    
    $version = Get-CurrentVersion
    Write-Host "Versión: $version" -ForegroundColor Cyan
    
    if (Test-Path $ReleasePath) {
        Remove-Item $ReleasePath -Recurse -Force
    }
    
    New-Item -ItemType Directory -Path $ReleasePath -Force | Out-Null
    
    Write-Host "Copiando archivos..." -ForegroundColor Yellow
    Copy-Item -Path "$SrcPath\*" -Destination $ReleasePath -Recurse -Force
    
    $mainFile = Join-Path $ReleasePath "main.ps1"
    if (Test-Path $mainFile) {
        (Get-Content $mainFile) -replace '\$global:version = ".*?"', "`$global:version = `"$version`"" |
            Out-File $mainFile -Encoding UTF8
    }
    
    @'
@echo off
echo ==============================================
echo        Daniel Tools - Suite de Utilidades
echo ==============================================
echo.

if not exist "C:\Temp" (
    mkdir C:\Temp
    echo Carpeta C:\Temp creada
)

powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0main.ps1"

if %errorlevel% neq 0 (
    echo.
    echo Error al ejecutar la herramienta
    pause
)
'@ | Out-File -FilePath "$ReleasePath\run.bat" -Encoding ASCII
    
    Write-Host "✓ Release creado en: $ReleasePath" -ForegroundColor Green
}

Write-Host "`n✅ Proceso completado" -ForegroundColor Green