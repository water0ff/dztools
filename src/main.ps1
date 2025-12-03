#requires -Version 5.0

try {
    Write-Host "Cargando ensamblados de Windows Forms..." -ForegroundColor Yellow
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "✓ System.Windows.Forms cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando System.Windows.Forms: $_" -ForegroundColor Red
    pause
    exit 1
}
try {
    Add-Type -AssemblyName System.Drawing
    Write-Host "✓ System.Drawing cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando System.Drawing: $_" -ForegroundColor Red
    pause
    exit 1
}
try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "✓ VisualStyles habilitado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error habilitando VisualStyles: $_" -ForegroundColor Red
}
if (Get-Command Set-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}
Write-Host "`nImportando módulos..." -ForegroundColor Yellow
$modulesPath = Join-Path $PSScriptRoot "modules"
$modules = @(
    "GUI.psm1",
    "Database.psm1",
    "Utilities.psm1",
    "Installers.psm1"
)
foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Write-Host "  ✓ $module" -ForegroundColor Green
        } catch {
            Write-Host "  ✗ Error importando `$module: `$_" -ForegroundColor Red

        }
    } else {
        Write-Host "  ✗ $module no encontrado" -ForegroundColor Red
    }
}
$global:version = "0.1.0"
$global:defaultInstructions = @"
----- CAMBIOS -----
- Versión modularizada
- Mejor estructura de código
- Compatibilidad mejorada
"@
function Initialize-Environment {
    Write-Host "`n=============================================" -ForegroundColor DarkCyan
    Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
    Write-Host "              Versión: v$($global:version)               " -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor DarkCyan
    if (!(Test-Path -Path "C:\Temp")) {
        try {
            New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
            Write-Host "Carpeta 'C:\Temp' creada." -ForegroundColor Green
        } catch {
            Write-Host "Error creando C:\Temp: $_" -ForegroundColor Yellow
        }
    }

    return $true
}
function New-MainForm {
    Write-Host "Creando formulario principal..." -ForegroundColor Yellow
    try {
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $formPrincipal = New-Object System.Windows.Forms.Form
        $formPrincipal.Size = New-Object System.Drawing.Size(1000, 600)  # Aumentado de 720x400
        $formPrincipal.StartPosition = "CenterScreen"
        $formPrincipal.BackColor = [System.Drawing.Color]::White
        $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $formPrincipal.MaximizeBox = $false
        $formPrincipal.MinimizeBox = $false
        $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $formPrincipal.Text = "Daniel Tools v$version"
        Write-Host "`n=============================================" -ForegroundColor DarkCyan
        Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
        Write-Host "              Versión: v$($version)               " -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor DarkCyan
        Write-Host "`nTodos los derechos reservados para Daniel Tools." -ForegroundColor Cyan
        Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan
        $toolTip = New-Object System.Windows.Forms.ToolTip

        return $formPrincipal

    } catch {
        Write-Host "✗ Error creando formulario: $_" -ForegroundColor Red
        return $null
    }
}
function Start-Application {
    Write-Host "Iniciando aplicación..." -ForegroundColor Cyan

    if (-not (Initialize-Environment)) {
        Write-Host "Error inicializando entorno. Saliendo..." -ForegroundColor Red
        return
    }

    $mainForm = New-MainForm
    if ($mainForm -eq $null) {
        Write-Host "Error: No se pudo crear el formulario principal" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("No se pudo crear la interfaz gráfica. Verifique los logs.", "Error crítico")
        return
    }

    try {
        Write-Host "Mostrando formulario..." -ForegroundColor Yellow
        $mainForm.ShowDialog()
        Write-Host "Aplicación finalizada correctamente." -ForegroundColor Green
    } catch {
        Write-Host "Error mostrando formulario: $_" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error en la aplicación")
    }
}
try {
    Start-Application
} catch {
    Write-Host "Error fatal: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    pause
    exit 1
}