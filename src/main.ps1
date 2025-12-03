#requires -Version 5.0
$global:version = "beta.25.12.03.1046"
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
            Write-Host "  ✗ Error importando módulo: $module" -ForegroundColor Red
            Write-Host "    Ruta    : $modulePath" -ForegroundColor DarkYellow
            Write-Host "    Mensaje : $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "    Detalle : $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkYellow
            Write-Host "    Stack   : $($_.ScriptStackTrace)" -ForegroundColor DarkGray
            throw
        }
    } else {
        Write-Host "  ✗ $module no encontrado" -ForegroundColor Red
    }
}

$global:defaultInstructions = @"
----- CAMBIOS -----
- Carga de INIS en la conexión a BDD.
- Se cambió la instalación de SSMS14 a SSMS21.
- Se deshabilitó la subida a mega.
- Restructura del proceso de Backups (choco).
- Se agregó subida a megaupload.
- Se agregó compresión con contraseña de respaldos
- Se agregó compresión con contraseña de respaldos
- Se agregó consola de cambios y tool tip para botones
- Reorganización de botones
- Query Browser para SQL en pestaña: Base de datos
- - Ahora se pueden agregar comentarios con "-" y entre "/* */"
- - Tabla en consola
- - Obtener columnas en consola
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
        $global:lastReportedPct = -1  # Añadir al inicio del script
        $tabControl = New-Object System.Windows.Forms.TabControl
        $tabControl.Size = New-Object System.Drawing.Size(990, 515)    # Original: 710x315
        $tabControl.Location = New-Object System.Drawing.Point(5, 5)   # Pequeño margen
        $tabAplicaciones = New-Object System.Windows.Forms.TabPage
        $tabAplicaciones.Text = "Aplicaciones"
        $tabProSql = New-Object System.Windows.Forms.TabPage
        $tabProSql.Text = "Base de datos"
        $tabProSql.AutoScroll = $true  # Habilitar scrollbar si el contenido excede el área
        $tabControl.TabPages.Add($tabAplicaciones)
        $tabControl.TabPages.Add($tabProSql)
        $lblServer = Create-Label -Text "Instancia SQL:" `
            -Location (New-Object System.Drawing.Point(10, 10)) `
            -Size (New-Object System.Drawing.Size(100, 10))
        $txtServer = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) `
            -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle "DropDown"
        $txtServer.Text = ".\NationalSoft"
        $lblUser = Create-Label -Text "Usuario:" `
            -Location (New-Object System.Drawing.Point(10, 50)) `
            -Size (New-Object System.Drawing.Size(100, 10))
        $txtUser = Create-TextBox -Location (New-Object System.Drawing.Point(10, 60)) `
            -Size (New-Object System.Drawing.Size(180, 20))
        $lblPassword = Create-Label -Text "Contraseña:" `
            -Location (New-Object System.Drawing.Point(10, 90)) `
            -Size (New-Object System.Drawing.Size(100, 10))
        $txtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(10, 100)) `
            -Size (New-Object System.Drawing.Size(180, 20)) -UseSystemPasswordChar $true
        $txtUser.Text = "sa"
        $lblbdd_cmb = Create-Label -Text "Base de datos" `
            -Location (New-Object System.Drawing.Point(10, 130)) `
            -Size (New-Object System.Drawing.Size(100, 10))
        $cmbDatabases = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 140)) `
            -Size (New-Object System.Drawing.Size(180, 20)) `
            -DropDownStyle DropDownList
        $cmbDatabases.Enabled = $false
        $lblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" `
            -Location (New-Object System.Drawing.Point(10, 400)) `
            -Size (New-Object System.Drawing.Size(180, 80))
        $btnExecute = Create-Button -Text "Ejecutar" -Location (New-Object System.Drawing.Point(220, 20))
        $btnExecute.Size = New-Object System.Drawing.Size(100, 30)
        $btnExecute.Enabled = $false
        $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 220)) `
            -Size (New-Object System.Drawing.Size(180, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))
        $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(10, 260)) `
            -Size (New-Object System.Drawing.Size(180, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) `
            -Enabled $false
        $btnBackup = Create-Button -Text "Backup BDD" -Location (New-Object System.Drawing.Point(10, 300)) `
            -Size (New-Object System.Drawing.Size(180, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(0, 192, 0)) `
            -ToolTip "Realizar backup de la base de datos seleccionada"
        $cmbQueries = New-Object System.Windows.Forms.ComboBox
        $cmbQueries.Location = New-Object System.Drawing.Point(330, 25)
        $cmbQueries.Size = New-Object System.Drawing.Size(350, 20)
        $cmbQueries.Enabled = $false
        $rtbQuery = New-Object System.Windows.Forms.RichTextBox
        $rtbQuery.Location = New-Object System.Drawing.Point(220, 60)
        $rtbQuery.Size = New-Object System.Drawing.Size(740, 140)     # Mayor ancho
        $rtbQuery.Multiline = $true
        $rtbQuery.ScrollBars = 'Vertical'
        $rtbQuery.WordWrap = $true
        $dgvResults = New-Object System.Windows.Forms.DataGridView
        $dgvResults.Location = New-Object System.Drawing.Point(220, 205)
        $dgvResults.Size = New-Object System.Drawing.Size(740, 280)
        $dgvResults.ReadOnly = $true
        $dgvResults.AllowUserToAddRows = $false
        $dgvResults.AllowUserToDeleteRows = $false
        $dgvResults.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditProgrammatically
        $dgvResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
        $dgvResults.MultiSelect = $true
        $dgvResults.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
        $dgvResults.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightBlue
        $dgvResults.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
        $script:originalForeColor = $dgvResults.DefaultCellStyle.ForeColor
        $script:originalHeaderBackColor = $dgvResults.ColumnHeadersDefaultCellStyle.BackColor
        $script:originalAutoSizeMode = $dgvResults.AutoSizeColumnsMode
        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $copyMenu = New-Object System.Windows.Forms.ToolStripMenuItem
        $copyMenu.Text = "Copiar selección"
        $contextMenu.Items.Add($copyMenu) | Out-Null
        $dgvResults.ContextMenuStrip = $contextMenu
        $panelGrid = New-Object System.Windows.Forms.Panel
        $panelGrid.Location = $dgvResults.Location
        $panelGrid.Size = $dgvResults.Size
        $panelGrid.AutoScroll = $true
        $dgvResults.Dock = [System.Windows.Forms.DockStyle]::Fill
        $panelGrid.Controls.Add($dgvResults)
        $btnConnectDb.Enabled = $True
        $btnBackup.Enabled = $false
        $btnDisconnectDb.Enabled = $false
        $btnExecute.Enabled = $false
        $rtbQuery.Enabled = $false
        $txtServer.Enabled = $true
        $txtUser.Enabled = $true
        $txtPassword.Enabled = $true
        $btnExecute.Enabled = $false
        $cmbQueries.Enabled = $false
        #Todavía no lo migramos                $btnReloadConnections,  # <-- Agregar este
        $tabProSql.Controls.AddRange(@(
                $btnConnectDb,
                $btnDisconnectDb,
                $cmbDatabases,
                $lblConnectionStatus,
                $btnExecute,
                $cmbQueries,
                $rtbQuery,
                $lblServer,
                $txtServer,
                $lblUser,
                $txtUser,
                $lblPassword,
                $txtPassword,
                $lblbdd_cmb,
                $panelGrid,
                $btnBackup
            ))
        # COLUMNA 1 | INSTALADORES EJECUTABLES | X:10
        $lblHostname = Create-Label -Text ([System.Net.Dns]::GetHostName()) -Location (New-Object System.Drawing.Point(10, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Hostname al portapapeles."
        $btnInstalarHerramientas = Create-Button -Text "Instalar Herramientas" -Location (New-Object System.Drawing.Point(10, 50)) `
            -ToolTip "Abrir el menú de instaladores de Chocolatey."
        $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 90)) `
            -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
        $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 130)) `
            -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
        $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 170)) `
            -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "De momento solo si es SQL 2014."
        $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 210)) `
            -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Busca SQL Management en tu equipo y te confirma la versión previo a ejecutarlo."
        $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 250)) `
            -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Herramienta de Star con funciones multiples para impresoras POS."
        $btnLectorDPicacls = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(10, 290)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
        $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Instalador LZMA" -Location (New-Object System.Drawing.Point(10, 330)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
        $btnConfigurarIPs = Create-Button -Text "Agregar IPs" -Location (New-Object System.Drawing.Point(10, 370)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
        $btnAddUser = Create-Button -Text "Agregar usuario de Windows" -Location (New-Object System.Drawing.Point(10, 410)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Crear nuevo usuario local en Windows"
        $btnForzarActualizacion = Create-Button -Text "Actualizar datos del sistema" -Location (New-Object System.Drawing.Point(10, 450)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Actualiza información de hardware del sistema"
        #COLUMNA 2 | FUNCIONES Y SERVICIOS DE WINDOWS | X: 250
        $lblPort = Create-Label -Text "Puerto: No disponible" -Location (New-Object System.Drawing.Point(250, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Puerto al portapapeles."
        $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(250, 50)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
        $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(250, 90)) `
            -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
        $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(250, 130)) `
            -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
        #COLUMNA 3 | FUNCIONES NATIONAL SOFT | X: 490
        $txt_IpAdress = Create-TextBox -Location (New-Object System.Drawing.Point(490, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) `
            -ScrollBars 'Vertical' -Multiline $true -ReadOnly $true
        $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(490, 50)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
        $btnCambiarOTM = Create-Button -Text "Cambiar OTM a SQL/DBF" -Location (New-Object System.Drawing.Point(490, 90)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Cambiar la configuración entre SQL y DBF para On The Minute."
        $btnCheckPermissions = Create-Button -Text "Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(490, 130)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
        $btnCreateAPK = Create-Button -Text "Creación de SRM APK" -Location (New-Object System.Drawing.Point(490, 170)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Generar archivo APK para Comandero Móvil"
        # Columna 4: | FIXES Y NOVEDADES |  X:730
        $txt_AdapterStatus = Create-TextBox -Location (New-Object System.Drawing.Point(730, 1)) -Size(New-Object System.Drawing.Size(220, 40)) `
            -BackColor([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) `
            -ScrollBars 'Vertical' -Multiline $true -ReadOnly  $true
        $toolTip.SetToolTip($txt_AdapterStatus, "Lista de adaptadores y su estado. Haga clic en 'Actualizar adaptadores' para refrescar.")
        $txt_InfoInstrucciones = Create-TextBox `
            -Location (New-Object System.Drawing.Point(730, 50)) `
            -Size     (New-Object System.Drawing.Size(220, 500)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 1, 36, 86)) `
            -ForeColor ([System.Drawing.Color]::White) `
            -Font      (New-Object System.Drawing.Font("Courier New", 10)) `
            -Multiline $true `
            -ReadOnly  $true
        $txt_InfoInstrucciones.WordWrap = $true
        $txt_InfoInstrucciones.Text = $global:defaultInstructions
        $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(350, 525)) `
            -Size (New-Object System.Drawing.Size(500, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
        $tabAplicaciones.Controls.AddRange(@(
                $btnInstalarHerramientas,
                $btnProfiler,
                $btnDatabase,
                $btnSQLManager,
                $btnSQLManagement,
                $btnClearPrintJobs,
                $btnClearAnyDesk,
                $btnShowPrinters,
                $btnPrinterTool,
                $btnAplicacionesNS,
                $btnCheckPermissions,
                $btnCambiarOTM,
                $btnConfigurarIPs,
                $btnAddUser,
                $btnForzarActualizacion,
                $LZMAbtnBuscarCarpeta,
                $btnLectorDPicacls,
                $lblHostname,
                $lblPort,
                $txt_IpAdress,
                $txt_AdapterStatus,
                $txt_InfoInstrucciones,
                $btnCreateAPK
            ))
        $formPrincipal.Controls.Add($tabControl)
        $formPrincipal.Controls.Add($btnExit)
        $ipsWithAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
        Where-Object { $_.OperationalStatus -eq 'Up' } |
        ForEach-Object {
            $interface = $_
            $interface.GetIPProperties().UnicastAddresses |
            Where-Object {
                $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '127.0.0.1'
            } |
            ForEach-Object {
                @{
                    AdapterName = $interface.Name
                    IPAddress   = $_.Address.ToString()
                }
            }
        }
        if ($ipsWithAdapters.Count -gt 0) {
            $ipsTextForClipboard = ($ipsWithAdapters | ForEach-Object {
                    $_.IPAddress
                }) -join ", "
            $ipsTextForLabel = $ipsWithAdapters | ForEach-Object {
                "- $($_.AdapterName) - IP: $($_.IPAddress)"
            } | Out-String
            $txt_IpAdress.Text = $ipsTextForLabel
        } else {
            $txt_IpAdress.Text = "No se encontraron direcciones IP"
        }
        $btnExit.Add_Click({
                $formPrincipal.Dispose()
                $formPrincipal.Close()
            })
        Load-IniConnectionsToComboBox
        return $formPrincipal

    } catch {
        Write-Host "✗ Error creando formulario: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Detalle: $($_.InvocationInfo.PositionMessage)" -ForegroundColor Yellow
        Write-Host "  Stack : $($_.ScriptStackTrace)" -ForegroundColor DarkYellow
        throw    # <-- re-lanza para que también lo veas fuera
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