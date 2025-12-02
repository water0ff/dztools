#requires -Version 5.0

# Cargar ensamblados necesarios con manejo de errores
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

# Configurar política de ejecución
if (Get-Command Set-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}

# Importar módulos con verificación
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
        }
        catch {
                         Write-Host "  ✗ Error importando `$module: `$_" -ForegroundColor Red

        }
    } else {
        Write-Host "  ✗ $module no encontrado" -ForegroundColor Red
    }
}

# Variables globales
$global:version = "0.1.0"
$global:defaultInstructions = @"
----- CAMBIOS -----
- Versión modularizada
- Mejor estructura de código
- Compatibilidad mejorada
"@

# Configuración inicial simplificada
function Initialize-Environment {
    Write-Host "`n=============================================" -ForegroundColor DarkCyan
    Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
    Write-Host "              Versión: v$($global:version)               " -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor DarkCyan
    
    # Crear carpeta temporal si no existe
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

# Función para crear formulario principal SIMPLIFICADA
function New-MainForm {
    Write-Host "Creando formulario principal..." -ForegroundColor Yellow
    
    try {
        # Crear formulario usando la función del módulo GUI
        if (Get-Command New-FormBuilder -ErrorAction SilentlyContinue) {
            $form = New-FormBuilder -Title "Daniel Tools v$($global:version)" -Size (New-Object System.Drawing.Size(800, 500))
            Write-Host "✓ Formulario creado" -ForegroundColor Green
        } else {
            Write-Host "✗ Función New-FormBuilder no encontrada. Creando formulario manualmente..." -ForegroundColor Yellow
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "Daniel Tools v$($global:version)"
            $form.Size = New-Object System.Drawing.Size(800, 500)
            $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
            $form.BackColor = [System.Drawing.Color]::White
        }
        
        # Crear controles básicos
        $lblTitulo = New-Object System.Windows.Forms.Label
        $lblTitulo.Text = "Daniel Tools - Versión $($global:version)"
        $lblTitulo.Location = New-Object System.Drawing.Point(10, 10)
        $lblTitulo.Size = New-Object System.Drawing.Size(400, 30)
        $lblTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        
        $btnProbar = New-Object System.Windows.Forms.Button
        $btnProbar.Text = "Probar Conexión SQL"
        $btnProbar.Location = New-Object System.Drawing.Point(10, 50)
        $btnProbar.Size = New-Object System.Drawing.Size(150, 30)
        
        $btnSalir = New-Object System.Windows.Forms.Button
        $btnSalir.Text = "Salir"
        $btnSalir.Location = New-Object System.Drawing.Point(170, 50)
        $btnSalir.Size = New-Object System.Drawing.Size(150, 30)
        $btnSalir.BackColor = [System.Drawing.Color]::LightGray
        
        # Agregar controles
        $form.Controls.AddRange(@($lblTitulo, $btnProbar, $btnSalir))
        
        # Eventos
        $btnProbar.Add_Click({
            Write-Host "Botón Probar presionado" -ForegroundColor Cyan
            [System.Windows.Forms.MessageBox]::Show("Funcionalidad en desarrollo", "Información")
        })
        
	$btnSalir.Add_Click({
	    try {
	        Write-Host "Botón Salir presionado" -ForegroundColor Cyan
	
	        # $this es el botón; FindForm() regresa el formulario que lo contiene
	        $currentForm = $this.FindForm()
	
	        if ($null -ne $currentForm) {
	            Write-Host "  Cerrando formulario..." -ForegroundColor Yellow
	            $currentForm.Close()
	        }
	        else {
	            Write-Host "  No se encontró formulario, cerrando aplicación..." -ForegroundColor Yellow
	            [System.Windows.Forms.Application]::Exit()
	        }
	    }
	    catch {
	        Write-Host "Error en botón Salir: $($_.Exception.Message)" -ForegroundColor Red
	        [System.Windows.Forms.MessageBox]::Show(
	            "Ocurrió un error al intentar cerrar la aplicación:`n$($_.Exception.Message)",
	            "Error",
	            [System.Windows.Forms.MessageBoxButtons]::OK,
	            [System.Windows.Forms.MessageBoxIcon]::Error
	        )
	    }
	})

        
        return $form
        
    } catch {
        Write-Host "✗ Error creando formulario: $_" -ForegroundColor Red
        return $null
    }
}

# Función principal con mejor manejo de errores
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
    }
    catch {
        Write-Host "Error mostrando formulario: $_" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error en la aplicación")
    }
}

# Punto de entrada
try {
    Start-Application
}
catch {
    Write-Host "Error fatal: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    pause
    exit 1
}