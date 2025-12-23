#requires -Version 5.0
chcp 65001 > $null
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding
$global:version = "beta.25.12.03.1046"
try {
    Write-Host "Cargando ensamblados de WPF..." -ForegroundColor Yellow
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Write-Host "✓ WPF cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando WPF: $_" -ForegroundColor Red
    pause
    exit 1
}

if (Get-Command Set-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}

Write-Host "`nImportando módulos..." -ForegroundColor Yellow
$modulesPath = Join-Path $PSScriptRoot "modules"
$modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Queries.psm1", "Installers.psm1")

foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -ErrorAction Stop -DisableNameChecking
            Write-Host "  ✓ $module" -ForegroundColor Green
        } catch {
            Write-Host "  ✗ Error importando módulo: $module" -ForegroundColor Red
            Write-Host "    Ruta    : $modulePath" -ForegroundColor DarkYellow
            Write-Host "    Mensaje : $($_.Exception.Message)" -ForegroundColor Yellow
            throw
        }
    } else {
        Write-Host "  ✗ $module no encontrado" -ForegroundColor Red
    }
}
$global:defaultInstructions = @"
----- CAMBIOS -----
- Se agregó una mejor busqueda para SQL Server Management Studio.
- Migración completa a WPF
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
    if (!(Test-Path -Path "C:\Temp")) {
        try {
            New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
            Write-Host "Carpeta 'C:\Temp' creada." -ForegroundColor Green
        } catch {
            Write-Host "Error creando C:\Temp: $_" -ForegroundColor Yellow
        }
    }
    try {
        $debugEnabled = Initialize-DzToolsConfig
        Write-DzDebug "`t[DEBUG]`t[DEBUG]Configuración de debug cargada (debug=$debugEnabled)" -Color DarkGray
    } catch {
        Write-Host "Advertencia: No se pudo inicializar la configuración de debug. $_" -ForegroundColor Yellow
    }
    return $true
}
$global:isHighlightingQuery = $false
function New-MainForm {

    Write-Host "`tCreando formulario principal WPF..." -ForegroundColor Yellow
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Daniel Tools $global:version" Height="600" Width="1000"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid>
        <TabControl Name="tabControl" Margin="5">
            <TabItem Header="Aplicaciones" Name="tabAplicaciones">
                <Grid>
                    <Label Content="" Name="lblHostname" HorizontalAlignment="Left" VerticalAlignment="Top"
                           Width="220" Height="40" Margin="10,1,0,0" Background="Black" Foreground="White"
                           HorizontalContentAlignment="Center" VerticalContentAlignment="Center"
                           BorderBrush="Gray" BorderThickness="1" Cursor="Hand"/>
                    <TextBox Name="lblPort" HorizontalAlignment="Left" VerticalAlignment="Top"
                            Width="220" Height="40" Margin="250,1,0,0" Background="Black" Foreground="White"
                            Text="Puerto: No disponible"
                            IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Cursor="Hand"/>
                    <TextBox Name="txt_IpAdress" HorizontalAlignment="Left" VerticalAlignment="Top"
                             Width="220" Height="40" Margin="490,1,0,0" Background="Black" Foreground="White"
                             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Cursor="Hand"/>
                    <TextBox Name="txt_AdapterStatus" HorizontalAlignment="Left" VerticalAlignment="Top"
                             Width="220" Height="40" Margin="730,1,0,0" Background="Black" Foreground="White"
                             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Cursor="Hand"/>
                    <Button Content="Instalar Herramientas" Name="btnInstalarHerramientas" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,50,0,0"/>
                    <Button Content="Ejecutar ExpressProfiler" Name="btnProfiler" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,90,0,0" Background="#E0E0E0"/>
                    <Button Content="Ejecutar Database4" Name="btnDatabase" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Background="#E0E0E0"/>
                    <Button Content="Ejecutar Manager" Name="btnSQLManager" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,170,0,0" Background="#E0E0E0"/>
                    <Button Content="Ejecutar Management" Name="btnSQLManagement" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,210,0,0" Background="#E0E0E0"/>
                    <Button Content="Printer Tools" Name="btnPrinterTool" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,250,0,0" Background="#E0E0E0"/>
                    <Button Content="Lector DP - Permisos" Name="btnLectorDPicacls" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,290,0,0" Background="#96C8FF"/>
                    <Button Content="Buscar Instalador LZMA" Name="LZMAbtnBuscarCarpeta" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,330,0,0" Background="#96C8FF"/>
                    <Button Content="Agregar IPs" Name="btnConfigurarIPs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,370,0,0" Background="#96C8FF"/>
                    <Button Content="Agregar usuario de Windows" Name="btnAddUser" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,410,0,0" Background="#96C8FF"/>
                    <Button Content="Actualizar datos del sistema" Name="btnForzarActualizacion" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,450,0,0" Background="#96C8FF"/>
                    <Button Content="Clear AnyDesk" Name="btnClearAnyDesk" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,50,0,0" Background="#FF4C4C"/>
                    <Button Content="Mostrar Impresoras" Name="btnShowPrinters" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,90,0,0"/>
                    <Button Content="Limpia y Reinicia Cola de Impresión" Name="btnClearPrintJobs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,130,0,0"/>
                    <Button Content="Aplicaciones National Soft" Name="btnAplicacionesNS" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,50,0,0" Background="#FFC896"/>
                    <Button Content="Cambiar OTM a SQL/DBF" Name="btnCambiarOTM" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,90,0,0" Background="#FFC896"/>
                    <Button Content="Permisos C:\NationalSoft" Name="btnCheckPermissions" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,130,0,0" Background="#FFC896"/>
                    <Button Content="Creación de SRM APK" Name="btnCreateAPK" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,170,0,0" Background="#FFC896"/>
                    <TextBox Name="txt_InfoInstrucciones" HorizontalAlignment="Left" VerticalAlignment="Top"
                             Width="220" Height="400" Margin="730,50,0,0" Background="#012456" Foreground="White"
                             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                             FontFamily="Courier New" FontSize="10"/>
                </Grid>
            </TabItem>
            <TabItem Header="Base de datos" Name="tabProSql">
                <Grid>
                    <Label Content="Instancia SQL:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,10,0,0"/>
                    <ComboBox Name="txtServer" HorizontalAlignment="Left" VerticalAlignment="Top"
                              Width="180" Margin="10,30,0,0" IsEditable="True" Text=".\NationalSoft"/>
                    <Label Content="Usuario:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,60,0,0"/>
                    <TextBox Name="txtUser" HorizontalAlignment="Left" VerticalAlignment="Top"
                             Width="180" Margin="10,80,0,0" Text="sa"/>
                    <Label Content="Contraseña:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,110,0,0"/>
                    <PasswordBox Name="txtPassword" HorizontalAlignment="Left" VerticalAlignment="Top"
                                 Width="180" Margin="10,130,0,0"/>
                    <Label Content="Base de datos" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,160,0,0"/>
                    <ComboBox Name="cmbDatabases" HorizontalAlignment="Left" VerticalAlignment="Top"
                              Width="180" Margin="10,180,0,0" IsEnabled="False"/>
                    <Button Content="Conectar a BDD" Name="btnConnectDb" Width="180" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,220,0,0" Background="#96C8FF"/>
                    <Button Content="Desconectar de BDD" Name="btnDisconnectDb" Width="180" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,260,0,0" Background="#96C8FF" IsEnabled="False"/>
                    <Button Content="Backup BDD" Name="btnBackup" Width="180" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,300,0,0" Background="#00C000"/>
                    <Label Name="lblConnectionStatus" Content="Conectado a BDD: Ninguna" HorizontalAlignment="Left"
                           VerticalAlignment="Top" Width="180" Height="80" Margin="10,400,0,0"/>
                    <Button Content="Ejecutar" Name="btnExecute" Width="100" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="220,20,0,0" IsEnabled="False"/>
                    <ComboBox Name="cmbQueries" HorizontalAlignment="Left" VerticalAlignment="Top"
                              Width="350" Margin="330,25,0,0" IsEnabled="False"/>
                    <RichTextBox Name="rtbQuery" HorizontalAlignment="Left" VerticalAlignment="Top"
                                 Width="740" Height="140" Margin="220,60,0,0" VerticalScrollBarVisibility="Auto"
                                 IsEnabled="False"/>
                    <DataGrid Name="dgvResults" HorizontalAlignment="Left" VerticalAlignment="Top"
                              Width="740" Height="280" Margin="220,205,0,0" IsReadOnly="True"
                              AutoGenerateColumns="True" CanUserAddRows="False" CanUserDeleteRows="False"
                              SelectionMode="Extended"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Button Content="Salir" Name="btnExit" Width="500" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="350,0,0,10" Background="#A9A9A9"/>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $lblHostname = $window.FindName("lblHostname")
    $lblPort = $window.FindName("lblPort")
    $txt_IpAdress = $window.FindName("txt_IpAdress")
    $txt_AdapterStatus = $window.FindName("txt_AdapterStatus")
    $txt_InfoInstrucciones = $window.FindName("txt_InfoInstrucciones")
    $btnInstalarHerramientas = $window.FindName("btnInstalarHerramientas")
    $btnProfiler = $window.FindName("btnProfiler")
    $btnDatabase = $window.FindName("btnDatabase")
    $btnSQLManager = $window.FindName("btnSQLManager")
    $btnSQLManagement = $window.FindName("btnSQLManagement")
    $btnPrinterTool = $window.FindName("btnPrinterTool")
    $btnLectorDPicacls = $window.FindName("btnLectorDPicacls")
    $LZMAbtnBuscarCarpeta = $window.FindName("LZMAbtnBuscarCarpeta")
    $btnConfigurarIPs = $window.FindName("btnConfigurarIPs")
    $btnAddUser = $window.FindName("btnAddUser")
    $btnForzarActualizacion = $window.FindName("btnForzarActualizacion")
    $btnClearAnyDesk = $window.FindName("btnClearAnyDesk")
    $btnShowPrinters = $window.FindName("btnShowPrinters")
    $btnClearPrintJobs = $window.FindName("btnClearPrintJobs")
    $btnAplicacionesNS = $window.FindName("btnAplicacionesNS")
    $btnCambiarOTM = $window.FindName("btnCambiarOTM")
    $btnCheckPermissions = $window.FindName("btnCheckPermissions")
    $btnCreateAPK = $window.FindName("btnCreateAPK")
    $txtServer = $window.FindName("txtServer")
    $txtUser = $window.FindName("txtUser")
    $txtPassword = $window.FindName("txtPassword")
    $cmbDatabases = $window.FindName("cmbDatabases")
    $btnConnectDb = $window.FindName("btnConnectDb")
    $btnDisconnectDb = $window.FindName("btnDisconnectDb")
    $btnBackup = $window.FindName("btnBackup")
    $lblConnectionStatus = $window.FindName("lblConnectionStatus")
    $btnExecute = $window.FindName("btnExecute")
    $cmbQueries = $window.FindName("cmbQueries")
    $rtbQuery = $window.FindName("rtbQuery")
    $dgvResults = $window.FindName("dgvResults")
    $btnExit = $window.FindName("btnExit")
    $global:txtServer = $txtServer
    $global:txtUser = $txtUser
    $global:txtPassword = $txtPassword
    $global:cmbDatabases = $cmbDatabases
    $global:btnConnectDb = $btnConnectDb
    $global:btnDisconnectDb = $btnDisconnectDb
    $global:btnExecute = $btnExecute
    $global:btnBackup = $btnBackup
    $global:cmbQueries = $cmbQueries
    $global:rtbQuery = $rtbQuery
    $global:lblConnectionStatus = $lblConnectionStatus
    $global:dgvResults = $dgvResults
    $global:txt_AdapterStatus = $txt_AdapterStatus
    $lblHostname.Content = [System.Net.Dns]::GetHostName()
    $txt_InfoInstrucciones.Text = $global:defaultInstructions
    Write-Host "`n=============================================" -ForegroundColor DarkCyan
    Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
    Write-Host "              Versión: $($global:version)               " -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor DarkCyan
    Write-Host "`nTodos los derechos reservados para Daniel Tools." -ForegroundColor Cyan
    Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan
    Write-Host "O crea un issue en GitHub. https://github.com/water0ff/dztools/issues/new" -ForegroundColor Cyan
    $script:predefinedQueries = Get-PredefinedQueries
    Initialize-PredefinedQueries -ComboQueries $cmbQueries -RichTextBox $rtbQuery -Queries $script:predefinedQueries -Window $window
    $script:setInstructionText = {
        param([string]$Message)
        if ($null -ne $txt_InfoInstrucciones) {
            $txt_InfoInstrucciones.Dispatcher.Invoke([action] {
                    $txt_InfoInstrucciones.Text = $Message
                })
        }
    }.GetNewClosure()

    $ipsWithAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
    Where-Object { $_.OperationalStatus -eq 'Up' } |
    ForEach-Object {
        $interface = $_
        $interface.GetIPProperties().UnicastAddresses |
        Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '127.0.0.1' } |
        ForEach-Object {
            @{
                AdapterName = $interface.Name
                IPAddress   = $_.Address.ToString()
            }
        }
    }

    if ($ipsWithAdapters.Count -gt 0) {
        $ipsTextForLabel = $ipsWithAdapters | ForEach-Object { "- $($_.AdapterName) - IP: $($_.IPAddress)" } | Out-String
        $txt_IpAdress.Text = $ipsTextForLabel
        Write-DzDebug "`t[DEBUG]`t[DEBUG] txt_IpAdress poblado con: '$($txt_IpAdress.Text)'"
        Write-DzDebug "`t[DEBUG]`t[DEBUG] txt_IpAdress.Text.Length: $($txt_IpAdress.Text.Length)"
    } else {
        $txt_IpAdress.Text = "No se encontraron direcciones IP"
        Write-DzDebug "`t[DEBUG]`t[DEBUG] No se encontraron IPs" -Color Yellow
    }
    Refresh-AdapterStatus
    Load-IniConnectionsToComboBox -Combo $txtServer
    $changeColorOnHover = {
        param($sender, $e)
        $sender.Background = [System.Windows.Media.Brushes]::Orange
    }
    $restoreColorOnLeave = {
        param($sender, $e)
        $sender.Background = [System.Windows.Media.Brushes]::Black
    }
    $lblHostname.Add_MouseEnter($changeColorOnHover)
    $lblHostname.Add_MouseLeave($restoreColorOnLeave)
    $lblPort.Add_MouseEnter($changeColorOnHover)
    $lblPort.Add_MouseLeave($restoreColorOnLeave)
    $txt_IpAdress.Add_MouseEnter($changeColorOnHover)
    $txt_IpAdress.Add_MouseLeave($restoreColorOnLeave)
    $txt_AdapterStatus.Add_MouseEnter($changeColorOnHover)
    $txt_AdapterStatus.Add_MouseLeave($restoreColorOnLeave)
    $buttonsToUpdate = @(
        $LZMAbtnBuscarCarpeta, $btnInstalarHerramientas, $btnProfiler,
        $btnDatabase, $btnSQLManager, $btnSQLManagement, $btnPrinterTool,
        $btnLectorDPicacls, $btnConfigurarIPs, $btnAddUser, $btnForzarActualizacion,
        $btnClearAnyDesk, $btnShowPrinters, $btnClearPrintJobs, $btnAplicacionesNS,
        $btnCheckPermissions, $btnCambiarOTM, $btnCreateAPK
    )
    foreach ($button in $buttonsToUpdate) {
        $button.Add_MouseLeave({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke($global:defaultInstructions)
                }
            })
    }
    # Obtener puertos UNA VEZ y almacenar en variable estructurada
    $portsResult = Get-SqlPortWithDebug

    # Asegurarnos de que $portsResult sea un array (incluso si solo hay un elemento)
    $portsArray = @($portsResult)

    $global:sqlPortsData = @{
        Ports        = $portsArray
        Summary      = $null
        DetailedText = $null
        DisplayText  = $null
    }

    # Generar textos formateados
    if ($portsArray.Count -gt 0) {
        # Ordenar por instancia
        $sortedPorts = $portsArray | Sort-Object -Property Instance

        # Formatear para mostrar en un solo renglón
        $displayParts = @()
        foreach ($port in $sortedPorts) {
            $instanceName = if ($port.Instance -eq "MSSQLSERVER") { "Default" } else { $port.Instance }
            $displayParts += "$instanceName : $($port.Port)"
        }

        # Unir todo en una sola línea con separador
        $global:sqlPortsData.DisplayText = $displayParts -join " | "

        # Texto detallado para copiar (formato original)
        $global:sqlPortsData.DetailedText = $sortedPorts | ForEach-Object {
            $instanceName = if ($_.Instance -eq "MSSQLSERVER") { "Default" } else { $_.Instance }
            "- Instancia: $instanceName | Puerto: $($_.Port) | Tipo: $($_.Type)"
        } | Out-String

        $global:sqlPortsData.Summary = "Total de instancias con puerto encontradas: $($sortedPorts.Count)"

        # Configurar el TextBlock
        $lblPort.Text = $global:sqlPortsData.DisplayText
        $lblPort.Tag = $global:sqlPortsData.DetailedText.Trim()
        $lblPort.ToolTip = if ($sortedPorts.Count -eq 1) {
            "Haz clic para mostrar en consola y copiar al portapapeles"
        } else {
            "$($sortedPorts.Count) instancias encontradas. Haz clic para detalles"
        }

        # Mostrar resumen en consola CON COLORES
        Write-Host "`n=== RESUMEN DE BÚSQUEDA SQL ===" -ForegroundColor Cyan
        Write-Host $global:sqlPortsData.Summary -ForegroundColor White
        Write-Host "Puertos: " -ForegroundColor White -NoNewline

        foreach ($port in $sortedPorts) {
            $instanceName = if ($port.Instance -eq "MSSQLSERVER") { "Default" } else { $port.Instance }
            Write-Host "$instanceName : " -ForegroundColor White -NoNewline
            Write-Host "$($port.Port) " -ForegroundColor Magenta -NoNewline

            if ($port -ne $sortedPorts[-1]) {
                Write-Host "| " -ForegroundColor Gray -NoNewline
            }
        }
        Write-Host ""  # Nueva línea
        Write-Host "=== FIN DE BÚSQUEDA ===" -ForegroundColor Cyan

    } else {
        $global:sqlPortsData.DetailedText = "No se encontraron puertos SQL ni instalaciones de SQL Server"
        $global:sqlPortsData.Summary = "No se encontraron puertos SQL"
        $global:sqlPortsData.DisplayText = "No se encontraron puertos SQL"

        $lblPort.Text = "No se encontraron puertos SQL"
        $lblPort.Tag = $global:sqlPortsData.DetailedText
        $lblPort.ToolTip = "Haz clic para mostrar el resumen de búsqueda"
    }
    $lblPort.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Click en lblPort - Evento iniciado" -Color DarkGray

            try {
                $textToCopy = $global:sqlPortsData.DetailedText.Trim()
                Write-DzDebug "`t[DEBUG]`t[DEBUG] Contenido a copiar: '$textToCopy'" -Color DarkGray

                # Mostrar en consola
                Write-Host "`n=== INFORMACIÓN DE PUERTOS SQL ===" -ForegroundColor Cyan
                if ($global:sqlPortsData.Ports.Count -gt 0) {
                    Write-Host $global:sqlPortsData.Summary -ForegroundColor White
                    Write-Host ""
                    $textToCopy -split "`n" | ForEach-Object {
                        Write-Host $_ -ForegroundColor Green
                    }
                } else {
                    Write-Host $textToCopy -ForegroundColor Red
                }
                Write-Host "=====================================" -ForegroundColor Cyan

                # Intentar copiar al portapapeles con reintentos
                $retryCount = 0
                $maxRetries = 3
                $copied = $false
                while (-not $copied -and $retryCount -lt $maxRetries) {
                    try {
                        [System.Windows.Clipboard]::SetText($textToCopy)
                        $copied = $true
                    } catch {
                        $retryCount++
                        if ($retryCount -ge $maxRetries) {
                            Write-Host "`n[ERROR] No se pudo copiar al portapapeles: $($_.Exception.Message)" -ForegroundColor Red
                            [System.Windows.MessageBox]::Show("Error al copiar al portapapeles: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        } else {
                            Start-Sleep -Milliseconds 100
                        }
                    }
                }

                if ($copied) {
                    if ($global:sqlPortsData.Ports.Count -gt 0) {
                        Write-Host "`n[ÉXITO] Información de puertos SQL copiada al portapapeles:" -ForegroundColor Green
                        $textToCopy -split "`n" | ForEach-Object {
                            Write-Host "  $_" -ForegroundColor Gray
                        }
                    } else {
                        Write-Host "`n[INFORMACIÓN] $textToCopy (copiado al portapapeles)" -ForegroundColor Yellow
                    }
                }

            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
    $lblHostname.Add_MouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Click en lblHostname - Evento iniciado" -Color DarkGray
            try {
                $hostname = [System.Net.Dns]::GetHostName()

                if ([string]::IsNullOrWhiteSpace($hostname)) {
                    Write-Host "`n[AVISO] No se pudo obtener el hostname." -ForegroundColor Yellow
                    return
                }

                [System.Windows.Clipboard]::SetText($hostname)
                Write-Host "`nNombre del equipo copiado: $hostname" -ForegroundColor Green

            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })


    $txt_IpAdress.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Click en txt_IpAdress - Evento iniciado" -Color DarkGray
            try {
                # Usar $sender.Text en lugar de $txt_IpAdress.Text por si hay problema de scope
                $ipsText = $sender.Text

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Contenido (sender): '$ipsText'" -Color DarkGray
                Write-DzDebug "`t[DEBUG]`t[DEBUG] Contenido (variable): '$($txt_IpAdress.Text)'" -Color DarkGray
                Write-DzDebug "`t[DEBUG]`t[DEBUG] Length: $($ipsText.Length)" -Color DarkGray

                if ([string]::IsNullOrWhiteSpace($ipsText)) {
                    Write-Host "`n[AVISO] No hay IPs para copiar." -ForegroundColor Yellow
                    return
                }

                [System.Windows.Clipboard]::SetText($ipsText)
                Write-Host "`nIP's copiadas al portapapeles:`n$ipsText" -ForegroundColor Green

            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }.GetNewClosure())
    $txt_AdapterStatus.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Get-NetConnectionProfile |
            Where-Object { $_.NetworkCategory -ne 'Private' } |
            ForEach-Object {
                Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private
            }
            Write-Host "Todas las redes se han establecido como Privadas."
            Refresh-AdapterStatus
            $e.Handled = $true
        })
    $btnInstalarHerramientas.Add_Click({
            Write-Host ""
            Write-DzDebug ("`t[DEBUG] Click en 'Instalar Herramientas' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            if (-not (Check-Chocolatey)) {
                Write-Host "Chocolatey no está instalado. No se puede abrir el menú de instaladores." -ForegroundColor Red
                return
            }
            Show-ChocolateyInstallerMenu
        })

    $btnProfiler.Add_Click({
            Write-DzDebug "`t[DEBUG]BTN CLICK: Inicio ejecución (ExpressProfiler)" ([System.ConsoleColor]::DarkGray)

            try {
                $toolName = "ExpressProfiler"
                $url = "https://github.com/ststeiger/ExpressProfiler/releases/download/1.0/ExpressProfiler20.zip"
                $zip = "C:\Temp\ExpressProfiler22wAddinSigned.zip"
                $extract = "C:\Temp\ExpressProfiler2"
                $exe = "ExpressProfiler.exe"
                $exePath = Join-Path $extract $exe

                Write-DzDebug "`t[DEBUG]`t[DEBUG] toolName=$toolName" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] url=$url" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] zip=$zip" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] extract=$extract" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] exePath=$exePath" ([System.ConsoleColor]::DarkGray)

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Validando si existe exe en $exePath" ([System.ConsoleColor]::DarkGray)

                if (Test-Path $exePath) {
                    Write-DzDebug "`t[DEBUG]INFO: $toolName ya existe en: $exePath" ([System.ConsoleColor]::Cyan)

                    $rExist = Show-WpfMessageBox -Message "$toolName ya existe en:`n$exePath`n`nSí = Abrir local`nNo = Volver a descargar`nCancelar = Cancelar operación" `
                        -Title "Ya existe" `
                        -Buttons "YesNoCancel" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado existe (rExist) = $rExist" ([System.ConsoleColor]::Cyan)

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process $exePath
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Abriendo $toolName existente..." }
                        return
                    }

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Cancel) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }
                } else {
                    $r = Show-WpfMessageBox -Message "¿Deseas descargar $toolName?" `
                        -Title "Confirmar descarga" `
                        -Buttons "YesNo" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado confirmación (r) = $r" ([System.ConsoleColor]::Cyan)

                    if ($r -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }
                }

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Creando progress bar (Show-WpfProgressBar)..." ([System.ConsoleColor]::DarkGray)
                $pw = Show-WpfProgressBar -Title "Descargando $toolName" -Message "Iniciando..."
                if ($null -eq $pw -or $null -eq $pw.ProgressBar) {
                    Write-DzDebug "`t[DEBUG]ERROR: No se pudo crear la ventana de progreso." ([System.ConsoleColor]::Red)
                    return
                }

                if (Test-Path $zip) {
                    Write-DzDebug "`t[DEBUG]WARN: Borrando zip anterior: $zip" ([System.ConsoleColor]::Yellow)
                    Remove-Item $zip -Force -ErrorAction SilentlyContinue
                }

                try {
                    Write-DzDebug "`t[DEBUG]PROGRESS: 0% - Iniciando descarga..." ([System.ConsoleColor]::DarkGray)
                    Update-WpfProgressBar -Window $pw -Percent 0 -Message "Preparando descarga..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Preparando descarga..." }

                    $ok = Download-FileWithProgressWpfStream -Url $url -OutFile $zip -Window $pw -OnStatus {
                        param($p, $m)
                        Write-DzDebug "`t[DEBUG]PROGRESS: $p% - $m" ([System.ConsoleColor]::DarkGray)
                        if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = $m }
                    }

                    if (-not $ok) { throw "Descarga fallida." }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Extrayendo..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Extrayendo..." }

                    $exeRunning = Get-Process | Where-Object { $_.Path -and $_.Path -ieq $exePath }

                    if (-not $exeRunning) {
                        if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue }
                    } else {
                        Write-DzDebug "`t[DEBUG]WARN: $exe está en ejecución, se omitió limpieza previa" ([System.ConsoleColor]::Yellow)
                    }

                    Expand-Archive -Path $zip -DestinationPath $extract -Force

                    if (-not (Test-Path $exePath)) { throw "No se encontró $exePath" }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Ejecutando..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Ejecutando..." }

                    Start-Process $exePath
                    Write-DzDebug "`t[DEBUG]INFO: Listo: Ejecutado $toolName" ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Listo: Ejecutado $toolName" }

                } catch {
                    Write-DzDebug "`t[DEBUG]ERROR Download/Run: $($_.Exception.Message)" ([System.ConsoleColor]::Red)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error: $($_.Exception.Message)" }
                } finally {
                    Close-WpfProgressBar -Window $pw
                }

            } catch {
                Write-DzDebug "`t[DEBUG]ERROR TryCatch: $($_.Exception.Message)`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
            }
        })

    $btnPrinterTool.Add_Click({
            Write-DzDebug "`t[DEBUG]BTN CLICK: Inicio ejecución (POS Printer Test)" ([System.ConsoleColor]::DarkGray)

            try {
                $toolName = "POS Printer Test"
                $url = "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip"
                $zip = "C:\Temp\RPT-RPI-Printer-Tool-1.zip"
                $extract = "C:\Temp\RPT-RPI-Printer-Tool-1"
                $exe = "POS Printer Test.exe"
                $exePath = Join-Path $extract $exe

                Write-DzDebug "`t[DEBUG]`t[DEBUG] toolName=$toolName" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] url=$url" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] zip=$zip" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] extract=$extract" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] exePath=$exePath" ([System.ConsoleColor]::DarkGray)

                $shouldDownload = $true

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Validando si existe exe en $exePath" ([System.ConsoleColor]::DarkGray)
                if (Test-Path $exePath) {
                    Write-DzDebug "`t[DEBUG]INFO: $toolName ya existe en: $exePath" ([System.ConsoleColor]::Cyan)

                    $rExist = Show-WpfMessageBox -Message "$toolName ya existe en:`n$exePath`n`nSí = Abrir local`nNo = Volver a descargar`nCancelar = Cancelar operación" `
                        -Title "Ya existe" `
                        -Buttons "YesNoCancel" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado existe (rExist) = $rExist" ([System.ConsoleColor]::Cyan)

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process $exePath
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Abriendo $toolName existente..." }
                        return
                    }

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Cancel) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }

                    $shouldDownload = $true
                } else {
                    $r = Show-WpfMessageBox -Message "¿Deseas descargar $toolName?" `
                        -Title "Confirmar descarga" `
                        -Buttons "YesNo" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado confirmación (r) = $r" ([System.ConsoleColor]::Cyan)

                    if ($r -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }

                    $shouldDownload = $true
                }

                if (-not $shouldDownload) { return }

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Creando progress bar (Show-WpfProgressBar)..." ([System.ConsoleColor]::DarkGray)
                $pw = Show-WpfProgressBar -Title "Descargando $toolName" -Message "Iniciando..."
                if ($null -eq $pw -or $null -eq $pw.ProgressBar) {
                    Write-DzDebug "`t[DEBUG]ERROR: No se pudo crear la ventana de progreso." ([System.ConsoleColor]::Red)
                    return
                }

                if (Test-Path $zip) {
                    Write-DzDebug "`t[DEBUG]WARN: Borrando zip anterior: $zip" ([System.ConsoleColor]::Yellow)
                    Remove-Item $zip -Force -ErrorAction SilentlyContinue
                }

                try {
                    Write-DzDebug "`t[DEBUG]PROGRESS: 0% - Iniciando descarga..." ([System.ConsoleColor]::DarkGray)
                    Update-WpfProgressBar -Window $pw -Percent 0 -Message "Preparando descarga..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Preparando descarga..." }

                    $ok = Download-FileWithProgressWpfStream -Url $url -OutFile $zip -Window $pw -OnStatus {
                        param($p, $m)
                        Write-DzDebug "`t[DEBUG]PROGRESS: $p% - $m" ([System.ConsoleColor]::DarkGray)
                        if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = $m }
                    }

                    if (-not $ok) { throw "Descarga fallida." }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Extrayendo..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Extrayendo..." }

                    $exeRunning = Get-Process | Where-Object { $_.Path -and $_.Path -ieq $exePath }

                    if (-not $exeRunning) {
                        if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue }
                    } else {
                        Write-DzDebug "`t[DEBUG]WARN: $exe está en ejecución, se omitió limpieza previa" ([System.ConsoleColor]::Yellow)
                    }

                    Expand-Archive -Path $zip -DestinationPath $extract -Force

                    if (-not (Test-Path $exePath)) { throw "No se encontró $exePath" }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Ejecutando..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Ejecutando..." }

                    Start-Process $exePath
                    Write-DzDebug "`t[DEBUG]INFO: Listo: Ejecutado $toolName" ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Listo: Ejecutado $toolName" }

                } catch {
                    Write-DzDebug "`t[DEBUG]ERROR Download/Run: $($_.Exception.Message)" ([System.ConsoleColor]::Red)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error: $($_.Exception.Message)" }
                } finally {
                    Close-WpfProgressBar -Window $pw
                }

            } catch {
                Write-DzDebug "`t[DEBUG]ERROR TryCatch: $($_.Exception.Message)`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
            }
        })


    $btnDatabase.Add_Click({
            Write-DzDebug "`t[DEBUG]BTN CLICK: Inicio ejecución (Database4)" ([System.ConsoleColor]::DarkGray)

            try {
                $toolName = "Database4"
                $url = "https://fishcodelib.com/files/DatabaseNet4.zip"
                $zip = "C:\Temp\DatabaseNet4.zip"
                $extract = "C:\Temp\Database4"
                $exe = "Database4.exe"
                $exePath = Join-Path $extract $exe

                Write-DzDebug "`t[DEBUG]`t[DEBUG] toolName=$toolName" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] url=$url" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] zip=$zip" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] extract=$extract" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG]`t[DEBUG] exePath=$exePath" ([System.ConsoleColor]::DarkGray)

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Validando si existe exe en $exePath" ([System.ConsoleColor]::DarkGray)

                if (Test-Path $exePath) {
                    Write-DzDebug "`t[DEBUG]INFO: $toolName ya existe en: $exePath" ([System.ConsoleColor]::Cyan)

                    $rExist = Show-WpfMessageBox -Message "$toolName ya existe en:`n$exePath`n`nSí = Abrir local`nNo = Volver a descargar`nCancelar = Cancelar operación" `
                        -Title "Ya existe" `
                        -Buttons "YesNoCancel" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado existe (rExist) = $rExist" ([System.ConsoleColor]::Cyan)

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process $exePath
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Abriendo $toolName existente..." }
                        return
                    }

                    if ($rExist -eq [System.Windows.MessageBoxResult]::Cancel) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }
                } else {
                    $r = Show-WpfMessageBox -Message "¿Deseas descargar $toolName?" `
                        -Title "Confirmar descarga" `
                        -Buttons "YesNo" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado confirmación (r) = $r" ([System.ConsoleColor]::Cyan)

                    if ($r -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                        return
                    }
                }

                $pw = Show-WpfProgressBar -Title "Descargando $toolName" -Message "Iniciando..."
                if ($null -eq $pw -or $null -eq $pw.ProgressBar) {
                    Write-DzDebug "`t[DEBUG]ERROR: No se pudo crear la ventana de progreso." ([System.ConsoleColor]::Red)
                    return
                }

                if (Test-Path $zip) {
                    Write-DzDebug "`t[DEBUG]WARN: Borrando zip anterior: $zip" ([System.ConsoleColor]::Yellow)
                    Remove-Item $zip -Force -ErrorAction SilentlyContinue
                }

                try {
                    Write-DzDebug "`t[DEBUG]PROGRESS: 0% - Iniciando descarga..." ([System.ConsoleColor]::DarkGray)
                    Update-WpfProgressBar -Window $pw -Percent 0 -Message "Preparando descarga..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Preparando descarga..." }

                    $ok = Download-FileWithProgressWpfStream -Url $url -OutFile $zip -Window $pw -OnStatus {
                        param($p, $m)
                        Write-DzDebug "`t[DEBUG]PROGRESS: $p% - $m" ([System.ConsoleColor]::DarkGray)
                        if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = $m }
                    }

                    if (-not $ok) { throw "Descarga fallida." }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Extrayendo..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Extrayendo..." }

                    if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue }
                    Expand-Archive -Path $zip -DestinationPath $extract -Force

                    if (-not (Test-Path $exePath)) { throw "No se encontró $exePath" }

                    Update-WpfProgressBar -Window $pw -Percent 100 -Message "Ejecutando..."
                    if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Ejecutando..." }

                    Start-Process $exePath
                    Write-DzDebug "`t[DEBUG]INFO: Listo: Ejecutado $toolName" ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Listo: Ejecutado $toolName" }

                } catch {
                    Write-DzDebug "`t[DEBUG]ERROR Download/Run: $($_.Exception.Message)" ([System.ConsoleColor]::Red)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error: $($_.Exception.Message)" }
                } finally {
                    Close-WpfProgressBar -Window $pw
                }

            } catch {
                Write-DzDebug "`t[DEBUG]ERROR TryCatch: $($_.Exception.Message)`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
            }
        })

    $btnLectorDPicacls.Add_Click({
            Write-DzDebug "`t[DEBUG]BTN CLICK: Inicio ejecución (Lector DP + icacls)" ([System.ConsoleColor]::DarkGray)

            $pwPs = $null
            $pwDrv = $null

            try {
                $rMain = Show-WpfMessageBox -Message "Este proceso ejecutará cambios de permisos con PsExec (SYSTEM) y puede descargar/instalar un driver.`n`n¿Deseas continuar?" `
                    -Title "Confirmar operación" `
                    -Buttons "YesNo" `
                    -Icon "Warning"

                Write-DzDebug "`t[DEBUG]INFO: Confirmación principal (rMain) = $rMain" ([System.ConsoleColor]::Cyan)
                if ($rMain -ne [System.Windows.MessageBoxResult]::Yes) {
                    Write-DzDebug "`t[DEBUG]INFO: Operación cancelada por el usuario." ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación cancelada." }
                    return
                }

                Write-DzDebug "`t[DEBUG] Iniciando proceso..." ([System.ConsoleColor]::DarkGray)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Iniciando..." }

                $psexecPath = "C:\Temp\PsExec\PsExec.exe"
                $psexecZip = "C:\Temp\PSTools.zip"
                $psexecUrl = "https://download.sysinternals.com/files/PSTools.zip"
                $psexecExtractPath = "C:\Temp\PsExec"

                if (-not (Test-Path $psexecPath)) {
                    Write-DzDebug "`t[DEBUG] PsExec no encontrado. Se descargará PSTools.zip" ([System.ConsoleColor]::Yellow)
                    if (-not (Test-Path "C:\Temp")) { New-Item -Path "C:\Temp" -ItemType Directory | Out-Null }

                    $pwPs = Show-WpfProgressBar -Title "Descargando PsExec" -Message "Preparando descarga..."
                    if ($null -eq $pwPs -or $null -eq $pwPs.ProgressBar) {
                        Write-DzDebug "`t[DEBUG]ERROR: No se pudo crear progress bar para PsExec." ([System.ConsoleColor]::Red)
                        return
                    }

                    if (Test-Path $psexecZip) { Remove-Item $psexecZip -Force -ErrorAction SilentlyContinue }

                    try {
                        Update-WpfProgressBar -Window $pwPs -Percent 0 -Message "Preparando descarga..."
                        if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Descargando PsExec..." }

                        $okPs = Download-FileWithProgressWpfStream -Url $psexecUrl -OutFile $psexecZip -Window $pwPs -OnStatus {
                            param($p, $m)
                            Write-DzDebug "`t[DEBUG]PROGRESS(PsExec): $p% - $m" ([System.ConsoleColor]::DarkGray)
                            if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "PsExec: $m" }
                        }

                        if (-not $okPs) { throw "Descarga de PSTools.zip fallida." }

                        Update-WpfProgressBar -Window $pwPs -Percent 100 -Message "Extrayendo..."
                        if ($txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Extrayendo PsExec..." }

                        if (Test-Path $psexecExtractPath) { Remove-Item $psexecExtractPath -Recurse -Force -ErrorAction SilentlyContinue }
                        Expand-Archive -Path $psexecZip -DestinationPath $psexecExtractPath -Force

                        if (-not (Test-Path $psexecPath)) { throw "No se pudo extraer PsExec.exe." }

                        Write-DzDebug "`t[DEBUG] PsExec descargado y extraído correctamente." ([System.ConsoleColor]::Cyan)
                    } catch {
                        Write-DzDebug "`t[DEBUG]ERROR PsExec: $($_.Exception.Message)" ([System.ConsoleColor]::Red)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error PsExec: $($_.Exception.Message)" }
                        return
                    } finally {
                        if ($null -ne $pwPs) {
                            try { Close-WpfProgressBar -Window $pwPs } catch {}
                            $pwPs = $null
                        }
                    }
                } else {
                    Write-DzDebug "`t[DEBUG] PsExec ya está instalado en: $psexecPath" ([System.ConsoleColor]::Cyan)
                }

                $grupoAdmin = ""
                $gruposLocales = net localgroup | Where-Object { $_ -match "Administrators|Administradores" }
                if ($gruposLocales -match "Administrators") {
                    $grupoAdmin = "Administrators"
                } elseif ($gruposLocales -match "Administradores") {
                    $grupoAdmin = "Administradores"
                } else {
                    Write-DzDebug "`t[DEBUG]ERROR: No se encontró el grupo de administradores en el sistema." ([System.ConsoleColor]::Red)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error: No se encontró grupo Administradores." }
                    return
                }

                Write-DzDebug "`t[DEBUG] Grupo de administradores detectado: $grupoAdmin" ([System.ConsoleColor]::Cyan)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Grupo admin: $grupoAdmin" }

                $comando1 = "icacls C:\Windows\System32\en-us /grant `"$grupoAdmin`":F"
                $comando2 = "icacls C:\Windows\System32\en-us /grant `"NT AUTHORITY\SYSTEM`":F"
                $psexecCmd1 = "`"$psexecPath`" /accepteula /s cmd /c `"$comando1`""
                $psexecCmd2 = "`"$psexecPath`" /accepteula /s cmd /c `"$comando2`""

                Write-DzDebug "`t[DEBUG] Ejecutando primer comando: $comando1" ([System.ConsoleColor]::Yellow)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Ejecutando icacls (1)..." }
                $output1 = & cmd /c $psexecCmd1
                Write-DzDebug ([string]$output1) ([System.ConsoleColor]::Gray)

                Write-DzDebug "`t[DEBUG] Ejecutando segundo comando: $comando2" ([System.ConsoleColor]::Yellow)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Ejecutando icacls (2)..." }
                $output2 = & cmd /c $psexecCmd2
                Write-DzDebug ([string]$output2) ([System.ConsoleColor]::Gray)

                Write-DzDebug "`t[DEBUG] Modificación de permisos completada." ([System.ConsoleColor]::Cyan)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Permisos actualizados." }

                $ResponderDriver = Show-WpfMessageBox -Message "¿Desea descargar e instalar el driver del lector?" `
                    -Title "Descargar Driver" `
                    -Buttons "YesNo" `
                    -Icon "Question"

                Write-DzDebug "`t[DEBUG]INFO: Respuesta driver = $ResponderDriver" ([System.ConsoleColor]::Cyan)

                if ($ResponderDriver -ne [System.Windows.MessageBoxResult]::Yes) {
                    Write-DzDebug "`t[DEBUG]INFO: Usuario decidió NO descargar el driver." ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Driver omitido." }
                    return
                }

                $url = "https://softrestaurant.com/drivers?download=120:dp"
                $zipPath = "C:\Temp\Driver_DP.zip"
                $extractPath = "C:\Temp\Driver_DP"
                $exeName = "x64\Setup.msi"
                $validationPath = "C:\Temp\Driver_DP\x64\Setup.msi"
                $msiPath = Join-Path $extractPath $exeName

                Write-DzDebug "`t[DEBUG] Driver URL=$url" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG] zipPath=$zipPath" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG] extractPath=$extractPath" ([System.ConsoleColor]::DarkGray)
                Write-DzDebug "`t[DEBUG] msiPath=$msiPath" ([System.ConsoleColor]::DarkGray)

                if (Test-Path $validationPath) {
                    $rExistDrv = Show-WpfMessageBox -Message "El driver ya existe en:`n$validationPath`n`nSí = Instalar local`nNo = Volver a descargar`nCancelar = Cancelar operación" `
                        -Title "Driver ya existe" `
                        -Buttons "YesNoCancel" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Resultado existe driver (rExistDrv) = $rExistDrv" ([System.ConsoleColor]::Cyan)

                    if ($rExistDrv -eq [System.Windows.MessageBoxResult]::Yes) {
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Instalando driver existente..." }
                        Start-Process "msiexec.exe" -ArgumentList "/i `"$validationPath`"" -Wait
                        Write-DzDebug "`t[DEBUG]INFO: Driver instalado correctamente." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Driver instalado correctamente." }
                        return
                    }

                    if ($rExistDrv -eq [System.Windows.MessageBoxResult]::Cancel) {
                        Write-DzDebug "`t[DEBUG]INFO: Operación de driver cancelada." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Operación de driver cancelada." }
                        return
                    }
                } else {
                    $rDrv = Show-WpfMessageBox -Message "¿Deseas descargar el driver del lector ahora?" `
                        -Title "Confirmar descarga" `
                        -Buttons "YesNo" `
                        -Icon "Question"

                    Write-DzDebug "`t[DEBUG]INFO: Confirmación descarga driver (rDrv) = $rDrv" ([System.ConsoleColor]::Cyan)
                    if ($rDrv -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-DzDebug "`t[DEBUG]INFO: Descarga de driver cancelada." ([System.ConsoleColor]::Cyan)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Descarga cancelada." }
                        return
                    }
                }

                $pwDrv = Show-WpfProgressBar -Title "Descargando Driver DP" -Message "Preparando descarga..."
                if ($null -eq $pwDrv -or $null -eq $pwDrv.ProgressBar) {
                    Write-DzDebug "`t[DEBUG]ERROR: No se pudo crear progress bar para Driver." ([System.ConsoleColor]::Red)
                    return
                }

                if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }

                try {
                    Update-WpfProgressBar -Window $pwDrv -Percent 0 -Message "Preparando descarga..."
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Descargando driver..." }

                    $okDrv = Download-FileWithProgressWpfStream -Url $url -OutFile $zipPath -Window $pwDrv -OnStatus {
                        param($p, $m)
                        Write-DzDebug "`t[DEBUG]PROGRESS(Driver): $p% - $m" ([System.ConsoleColor]::DarkGray)
                        if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Driver: $m" }
                    }

                    if (-not $okDrv) { throw "Descarga del driver fallida." }

                    Update-WpfProgressBar -Window $pwDrv -Percent 100 -Message "Extrayendo..."
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Extrayendo driver..." }

                    if (Test-Path $extractPath) {
                        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
                    }

                    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

                    if (-not (Test-Path $validationPath)) { throw "No se encontró $validationPath" }

                    Update-WpfProgressBar -Window $pwDrv -Percent 100 -Message "Instalando..."
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Instalando driver..." }

                    try {
                        $pwDrv.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action] {}) | Out-Null
                    } catch {}

                    try {
                        $pwDrv.Dispatcher.Invoke([action] {
                                $pwDrv.Topmost = $false
                                $pwDrv.WindowState = [System.Windows.WindowState]::Minimized
                            }) | Out-Null
                    } catch {}

                    Start-Process "msiexec.exe" -ArgumentList "/i `"$validationPath`" /passive" -Wait

                    try {
                        $pwDrv.Dispatcher.Invoke([action] {
                                $pwDrv.WindowState = [System.Windows.WindowState]::Normal
                            }) | Out-Null
                    } catch {}

                    Write-DzDebug "`t[DEBUG]INFO: Driver instalado correctamente." ([System.ConsoleColor]::Cyan)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Driver instalado correctamente." }

                } catch {
                    Write-DzDebug "`t[DEBUG]ERROR Driver: $($_.Exception.Message)" ([System.ConsoleColor]::Red)
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error Driver: $($_.Exception.Message)" }
                }

            } catch {
                Write-DzDebug "`t[DEBUG]ERROR TryCatch: $($_.Exception.Message)`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
                if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Error: $($_.Exception.Message)" }
            } finally {
                if ($null -ne $pwDrv) {
                    try { Close-WpfProgressBar -Window $pwDrv } catch {}
                    $pwDrv = $null
                }
                if ($null -ne $pwPs) {
                    try { Close-WpfProgressBar -Window $pwPs } catch {}
                    $pwPs = $null
                }
            }
        })


    $btnSQLManager.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            function Get-SQLServerManagers {
                $managers = @()
                $possiblePaths = @(
                    "${env:SystemRoot}\System32\SQLServerManager*.msc",
                    "${env:SystemRoot}\SysWOW64\SQLServerManager*.msc"
                )
                foreach ($path in $possiblePaths) {
                    $foundManagers = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    if ($foundManagers) {
                        $managers += $foundManagers.FullName
                    }
                }
                return $managers
            }
            $managers = Get-SQLServerManagers
            if ($managers.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No se encontró ninguna versión de SQL Server Configuration Manager.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
            Show-SSMSSelectionDialog -Managers $managers
        })
    $btnSQLManagement.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Iniciando búsqueda de SSMS instalados" -Color DarkGray

            function Get-SSMSVersions {
                $ssmsPaths = @()

                # Rutas fijas para SSMS
                $fixedPaths = @(
                    # Instalaciones clásicas (SSMS 2008-2017)
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server\*\Tools\Binn\ManagementStudio\Ssms.exe",

                    # Instalaciones standalone (SSMS 2016+)
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio *\Common7\IDE\Ssms.exe",

                    # SSMS 2019-2022 (x86)
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe",

                    # SSMS 2019-2022 (x64) - versión completa
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe",

                    # SSMS 2022+ en carpeta Release
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 19\Release\Common7\IDE\Ssms.exe",

                    # Búsquedas genéricas
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio *\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio *\Release\Common7\IDE\Ssms.exe"
                )

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Buscando en $($fixedPaths.Count) rutas fijas" -Color DarkGray

                foreach ($path in $fixedPaths) {
                    Write-DzDebug "`t[DEBUG]`t[DEBUG] Buscando patrón: $path" -Color DarkGray
                    $foundPaths = Get-ChildItem -Path $path -ErrorAction SilentlyContinue

                    if ($foundPaths) {
                        foreach ($foundPath in $foundPaths) {
                            if ($ssmsPaths -notcontains $foundPath.FullName) {
                                $ssmsPaths += $foundPath.FullName
                                Write-DzDebug "`t[DEBUG]`t[DEBUG] ✓ Encontrado: $($foundPath.FullName)" -Color Green

                                try {
                                    $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($foundPath.FullName)
                                    Write-DzDebug "`t[DEBUG]`t[DEBUG]   Versión: $($versionInfo.FileVersion)" -Color Cyan
                                    Write-DzDebug "`t[DEBUG]`t[DEBUG]   Producto: $($versionInfo.ProductName)" -Color Cyan
                                } catch {
                                    Write-DzDebug "`t[DEBUG]`t[DEBUG]   No se pudo obtener info de versión" -Color Yellow
                                }
                            } else {
                                Write-DzDebug "`t[DEBUG]`t[DEBUG] - Duplicado (ignorado): $($foundPath.FullName)" -Color DarkGray
                            }
                        }
                    }
                }

                # Búsqueda en el registro solo si no encontramos en las rutas fijas
                if ($ssmsPaths.Count -eq 0) {
                    Write-DzDebug "`t[DEBUG]`t[DEBUG] No se encontró en rutas fijas. Buscando en registro..." -Color DarkGray

                    # Buscar en el registro de manera más específica
                    $registryPaths = @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                    )

                    foreach ($regPath in $registryPaths) {
                        $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
                        Where-Object {
                            $_.DisplayName -like "*SQL Server Management Studio*" -and
                            $_.InstallLocation -and
                            $_.InstallLocation.Trim() -ne ""
                        }

                        foreach ($entry in $entries) {
                            $installPath = $entry.InstallLocation.Trim()

                            # Asegurar que la ruta termine con \
                            if (-not $installPath.EndsWith('\')) {
                                $installPath += '\'
                            }

                            # Verificar si Ssms.exe existe en las subcarpetas comunes
                            $possibleSubPaths = @(
                                "Common7\IDE\Ssms.exe",
                                "Release\Common7\IDE\Ssms.exe"
                            )

                            foreach ($subPath in $possibleSubPaths) {
                                $fullPath = Join-Path $installPath $subPath
                                if (Test-Path $fullPath) {
                                    $resolvedPath = Resolve-Path $fullPath
                                    if ($ssmsPaths -notcontains $resolvedPath.Path) {
                                        $ssmsPaths += $resolvedPath.Path
                                        Write-DzDebug "`t[DEBUG]`t[DEBUG] ✓ Encontrado (registro): $($resolvedPath.Path)" -Color Green
                                    }
                                }
                            }
                        }
                    }
                }

                Write-DzDebug "`t[DEBUG]`t[DEBUG] Total de SSMS encontrados: $($ssmsPaths.Count)" -Color $(if ($ssmsPaths.Count -gt 0) { "Green" } else { "Red" })

                return $ssmsPaths | Sort-Object -Descending
            }

            $ssmsVersions = Get-SSMSVersions

            # Filtrar solo los archivos Ssms.exe reales (no carpetas u otros archivos)
            $filteredVersions = @()
            foreach ($path in $ssmsVersions) {
                $fileName = Split-Path $path -Leaf
                if ($fileName -eq "Ssms.exe" -and (Test-Path $path -PathType Leaf)) {
                    $filteredVersions += $path
                }
            }

            if ($filteredVersions.Count -eq 0) {
                Write-Host "`tNo se encontró ninguna versión de SSMS instalada." -ForegroundColor Red

                # Ofrecer opción para buscar manualmente
                $result = [System.Windows.MessageBox]::Show(
                    "No se encontró SQL Server Management Studio. ¿Desea buscar manualmente?",
                    "SSMS no encontrado",
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question
                )

                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
                    $openFileDialog.Filter = "SSMS Executable (Ssms.exe)|Ssms.exe"
                    $openFileDialog.Title = "Seleccione Ssms.exe"

                    if ($openFileDialog.ShowDialog() -eq $true) {
                        try {
                            Start-Process -FilePath $openFileDialog.FileName
                            Write-Host "`tEjecutando: $($openFileDialog.FileName)" -ForegroundColor Green
                        } catch {
                            [System.Windows.MessageBox]::Show("Error al ejecutar SSMS: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        }
                    }
                } else {
                    # Ofrecer descarga
                    $downloadResult = [System.Windows.MessageBox]::Show(
                        "¿Desea descargar la última versión de SSMS?",
                        "Descargar SSMS",
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Information
                    )

                    if ($downloadResult -eq [System.Windows.MessageBoxResult]::Yes) {
                        Start-Process "https://aka.ms/ssmsfullsetup"
                    }
                }
                return
            }

            Write-Host "`t✓ Se encontraron $($filteredVersions.Count) instalación(es) de SSMS" -ForegroundColor Green

            # SIEMPRE mostrar diálogo de selección, incluso si solo hay una versión
            $dialog = New-Object System.Windows.Window
            $dialog.Title = "Seleccionar versión de SSMS"
            $dialog.Width = 500
            $dialog.Height = 300
            $dialog.WindowStartupLocation = "CenterOwner"
            $dialog.ResizeMode = "NoResize"
            $dialog.Owner = $window  # ¡IMPORTANTE! Establecer el owner para que sea modal

            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $stackPanel.Margin = New-Object System.Windows.Thickness(10)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "Selecciona la versión de SQL Server Management Studio a ejecutar:"
            $label.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $label.FontSize = 12
            $stackPanel.Children.Add($label) | Out-Null

            $listBox = New-Object System.Windows.Controls.ListBox
            $listBox.Height = 150
            $listBox.FontSize = 11

            foreach ($version in $filteredVersions) {
                # Obtener información de versión para mostrar
                try {
                    $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($version)
                    $displayText = "$($versionInfo.ProductName) - $($versionInfo.FileVersion)"
                    $listBox.Items.Add([PSCustomObject]@{Path = $version; Display = $displayText }) | Out-Null
                } catch {
                    $listBox.Items.Add([PSCustomObject]@{Path = $version; Display = $version }) | Out-Null
                }
            }

            # Configurar cómo se muestra el texto
            $listBox.DisplayMemberPath = "Display"
            $listBox.SelectedValuePath = "Path"
            $listBox.SelectedIndex = 0
            $stackPanel.Children.Add($listBox) | Out-Null

            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = "Horizontal"
            $buttonPanel.HorizontalAlignment = "Right"
            $buttonPanel.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)

            $cancelButton = New-Object System.Windows.Controls.Button
            $cancelButton.Content = "Cancelar"
            $cancelButton.Width = 80
            $cancelButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
            $cancelButton.Add_Click({
                    $dialog.DialogResult = $false
                    $dialog.Close()
                })

            $okButton = New-Object System.Windows.Controls.Button
            $okButton.Content = "Ejecutar"
            $okButton.Width = 80
            $okButton.IsDefault = $true
            $okButton.Add_Click({
                    if ($listBox.SelectedValue) {
                        try {
                            Write-Host "`tEjecutando: $($listBox.SelectedValue)" -ForegroundColor Green
                            Start-Process -FilePath $listBox.SelectedValue
                            $dialog.DialogResult = $true
                            $dialog.Close()
                        } catch {
                            [System.Windows.MessageBox]::Show("Error al ejecutar SSMS: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        }
                    }
                })

            $buttonPanel.Children.Add($cancelButton) | Out-Null
            $buttonPanel.Children.Add($okButton) | Out-Null
            $stackPanel.Children.Add($buttonPanel) | Out-Null

            $dialog.Content = $stackPanel

            # Mostrar el diálogo de forma modal
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Mostrando diálogo de selección de SSMS" -Color DarkGray
            $dialog.ShowDialog() | Out-Null
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Diálogo de SSMS cerrado" -Color DarkGray
        })
    $btnForzarActualizacion.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Show-SystemComponents
            $resultado = [System.Windows.MessageBox]::Show("¿Desea forzar la actualización de datos?", "Confirmación", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($resultado -eq [System.Windows.MessageBoxResult]::Yes) {
                Start-SystemUpdate
                [System.Windows.MessageBox]::Show("Actualización completada", "Éxito", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } else {
                Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red
            }
        })
    $btnClearAnyDesk.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $confirmationResult = [System.Windows.MessageBox]::Show("¿Estás seguro de renovar AnyDesk?", "Confirmar Renovación", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($confirmationResult -eq [System.Windows.MessageBoxResult]::Yes) {
                $filesToDelete = @(
                    "C:\ProgramData\AnyDesk\system.conf",
                    "C:\ProgramData\AnyDesk\service.conf",
                    "$env:APPDATA\AnyDesk\system.conf",
                    "$env:APPDATA\AnyDesk\service.conf"
                )
                $deletedFilesCount = 0
                $errors = @()
                try {
                    Write-Host "`tCerrando el proceso AnyDesk..." -ForegroundColor Yellow
                    Stop-Process -Name "AnyDesk" -Force -ErrorAction Stop
                    Write-Host "`tAnyDesk ha sido cerrado correctamente." -ForegroundColor Green
                } catch {
                    Write-Host "`tError al cerrar el proceso AnyDesk: $_" -ForegroundColor Red
                    $errors += "No se pudo cerrar el proceso AnyDesk."
                }
                foreach ($file in $filesToDelete) {
                    try {
                        if (Test-Path $file) {
                            Remove-Item -Path $file -Force -ErrorAction Stop
                            Write-Host "`tArchivo eliminado: $file" -ForegroundColor Green
                            $deletedFilesCount++
                        } else {
                            Write-Host "`tArchivo no encontrado: $file" -ForegroundColor Red
                        }
                    } catch {
                        Write-Host "`nError al eliminar el archivo." -ForegroundColor Red
                    }
                }
                if ($errors.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("$deletedFilesCount archivo(s) eliminado(s) correctamente.", "Éxito", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } else {
                    [System.Windows.MessageBox]::Show("Se encontraron errores. Revisa la consola para más detalles.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            } else {
                Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red
            }
        })
    $btnShowPrinters.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            try {
                $printers = Get-WmiObject -Query "SELECT * FROM Win32_Printer" | ForEach-Object {
                    $printer = $_
                    $isShared = $printer.Shared -eq $true
                    [PSCustomObject]@{
                        Name       = $printer.Name.Substring(0, [Math]::Min(24, $printer.Name.Length))
                        PortName   = $printer.PortName.Substring(0, [Math]::Min(19, $printer.PortName.Length))
                        DriverName = $printer.DriverName.Substring(0, [Math]::Min(19, $printer.DriverName.Length))
                        IsShared   = if ($isShared) { "Sí" } else { "No" }
                    }
                }
                Write-Host "`nImpresoras disponibles en el sistema:"
                if ($printers.Count -gt 0) {
                    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "Nombre", "Puerto", "Driver", "Compartida")
                    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "------", "------", "------", "---------")
                    $printers | ForEach-Object {
                        Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f $_.Name, $_.PortName, $_.DriverName, $_.IsShared)
                    }
                } else {
                    Write-Host "`nNo se encontraron impresoras."
                }
            } catch {
                Write-Host "`nError al obtener impresoras: $_"
            }
        })
    $btnClearPrintJobs.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            try {
                if (-not (Test-Administrator)) {
                    [System.Windows.MessageBox]::Show("Esta acción requiere permisos de administrador.`r`nPor favor, ejecuta Daniel Tools como administrador.", "Permisos insuficientes", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
                if (-not $spooler) {
                    [System.Windows.MessageBox]::Show("No se encontró el servicio 'Cola de impresión (Spooler)' en este equipo.", "Servicio no encontrado", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
                try {
                    Get-Printer -ErrorAction Stop | ForEach-Object {
                        try {
                            Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
                        } catch {
                            Write-Host "`tNo se pudieron limpiar trabajos de la impresora '$($_.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }
                } catch {
                    Write-Host "`tNo se pudieron enumerar impresoras (Get-Printer). ¿Está instalado el módulo PrintManagement?" -ForegroundColor Yellow
                }
                if ($spooler.Status -eq 'Running') {
                    Write-Host "`tDeteniendo servicio Spooler..." -ForegroundColor DarkYellow
                    Stop-Service -Name Spooler -Force -ErrorAction Stop
                } else {
                    Write-Host "`tSpooler no está en estado 'Running' (estado actual: $($spooler.Status))." -ForegroundColor DarkYellow
                }
                $spooler.Refresh()
                if ($spooler.StartType -eq 'Disabled') {
                    [System.Windows.MessageBox]::Show("El servicio 'Cola de impresión (Spooler)' está DESHABILITADO.`r`nHabilítalo manualmente desde services.msc para poder iniciarlo.", "Spooler deshabilitado", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    return
                }
                Write-Host "`tIniciando servicio Spooler..." -ForegroundColor DarkYellow
                Start-Service -Name Spooler -ErrorAction Stop
                [System.Windows.MessageBox]::Show("Los trabajos de impresión han sido eliminados y el servicio de cola de impresión se ha reiniciado correctamente.", "Operación exitosa", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } catch {
                Write-Host "`n[ERROR ClearPrintJobs] $($_.Exception.Message)" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Ocurrió un error al intentar limpiar las impresoras o reiniciar el servicio:`r`n$($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
    $btnCheckPermissions.Add_Click({
            Write-Host "`nRevisando permisos en C:\NationalSoft" -ForegroundColor Yellow
            if (-not (Test-Administrator)) {
                [System.Windows.MessageBox]::Show("Esta acción requiere permisos de administrador.`r`nPor favor, ejecuta Daniel Tools como administrador.", "Permisos insuficientes", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            Check-Permissions
        })
    $btnAplicacionesNS.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $resultados = @()
            function Leer-Ini($filePath) {
                if (Test-Path $filePath) {
                    $content = Get-Content $filePath
                    $dataSource = ($content | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
                    $catalog = ($content | Select-String -Pattern "^Catalog=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
                    $authType = ($content | Select-String -Pattern "^autenticacion=(\d+)").Matches.Groups[1].Value
                    $authUser = if ($authType -eq "2") { "sa" } elseif ($authType -eq "1") { "Windows" } else { "Desconocido" }
                    return @{
                        DataSource = $dataSource
                        Catalog    = $catalog
                        Usuario    = $authUser
                    }
                }
                return $null
            }
            $pathsToCheck = @(
                @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
                @{ Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
                @{ Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
                @{ Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
                @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
                @{ Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
            )
            foreach ($entry in $pathsToCheck) {
                $basePath = $entry.Path
                $mainIni = "$basePath\$($entry.INI)"
                $appName = $entry.Nombre
                if (Test-Path $mainIni) {
                    $iniData = Leer-Ini $mainIni
                    if ($iniData) {
                        $resultados += [PSCustomObject]@{
                            Aplicacion = $appName
                            INI        = $entry.INI
                            DataSource = $iniData.DataSource
                            Catalog    = $iniData.Catalog
                            Usuario    = $iniData.Usuario
                        }
                    }
                } else {
                    $resultados += [PSCustomObject]@{
                        Aplicacion = $appName
                        INI        = "No encontrado"
                        DataSource = "NA"
                        Catalog    = "NA"
                        Usuario    = "NA"
                    }
                }
                $inisFolder = "$basePath\INIS"
                if ($appName -eq "OnTheMinute" -and (Test-Path $inisFolder)) {
                    $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
                    if ($iniFiles.Count -gt 1) {
                        foreach ($iniFile in $iniFiles) {
                            $iniData = Leer-Ini $iniFile.FullName
                            if ($iniData) {
                                $resultados += [PSCustomObject]@{
                                    Aplicacion = $appName
                                    INI        = $iniFile.Name
                                    DataSource = $iniData.DataSource
                                    Catalog    = $iniData.Catalog
                                    Usuario    = $iniData.Usuario
                                }
                            }
                        }
                    }
                } elseif (Test-Path $inisFolder) {
                    $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
                    foreach ($iniFile in $iniFiles) {
                        $iniData = Leer-Ini $iniFile.FullName
                        if ($iniData) {
                            $resultados += [PSCustomObject]@{
                                Aplicacion = $appName
                                INI        = $iniFile.Name
                                DataSource = $iniData.DataSource
                                Catalog    = $iniData.Catalog
                                Usuario    = $iniData.Usuario
                            }
                        }
                    }
                }
            }
            $restCardPath = "C:\NationalSoft\Restcard\RestCard.ini"
            if (Test-Path $restCardPath) {
                $resultados += [PSCustomObject]@{
                    Aplicacion = "Restcard"
                    INI        = "RestCard.ini"
                    DataSource = "existe"
                    Catalog    = "existe"
                    Usuario    = "existe"
                }
            } else {
                $resultados += [PSCustomObject]@{
                    Aplicacion = "Restcard"
                    INI        = "No encontrado"
                    DataSource = "NA"
                    Catalog    = "NA"
                    Usuario    = "NA"
                }
            }
            $columnas = @("Aplicacion", "INI", "DataSource", "Catalog", "Usuario")
            $anchos = @{}
            foreach ($col in $columnas) { $anchos[$col] = $col.Length }
            foreach ($res in $resultados) {
                foreach ($col in $columnas) {
                    if ($res.$col.Length -gt $anchos[$col]) {
                        $anchos[$col] = $res.$col.Length
                    }
                }
            }
            $titulos = $columnas | ForEach-Object { $_.PadRight($anchos[$_] + 2) }
            Write-Host ($titulos -join "") -ForegroundColor Cyan
            $separador = $columnas | ForEach-Object { ("-" * $anchos[$_]).PadRight($anchos[$_] + 2) }
            Write-Host ($separador -join "") -ForegroundColor Cyan
            foreach ($res in $resultados) {
                $fila = $columnas | ForEach-Object { $res.$_.PadRight($anchos[$_] + 2) }
                if ($res.INI -eq "No encontrado") {
                    Write-Host ($fila -join "") -ForegroundColor Red
                } else {
                    Write-Host ($fila -join "")
                }
            }
        })
    $btnCambiarOTM.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $syscfgPath = "C:\Windows\SysWOW64\Syscfg45_2.0.dll"
            $iniPath = "C:\NationalSoft\OnTheMinute4.5"
            if (-not (Test-Path $syscfgPath)) {
                [System.Windows.MessageBox]::Show("El archivo de configuración no existe.", "Error", [System.Windows.MessageBoxButton]::OK)
                Write-Host "`tEl archivo de configuración no existe." -ForegroundColor Red
                return
            }
            $fileContent = Get-Content $syscfgPath
            $isSQL = $fileContent -match "494E5354414C4C=1" -and $fileContent -match "56455253495354454D41=3"
            $isDBF = $fileContent -match "494E5354414C4C=2" -and $fileContent -match "56455253495354454D41=2"
            if (!$isSQL -and !$isDBF) {
                [System.Windows.MessageBox]::Show("No se detectó una configuración válida de SQL o DBF.", "Error", [System.Windows.MessageBoxButton]::OK)
                Write-Host "`tNo se detectó una configuración válida de SQL o DBF." -ForegroundColor Red
                return
            }
            $iniFiles = Get-ChildItem -Path $iniPath -Filter "*.ini"
            if ($iniFiles.Count -eq 0) {
                [System.Windows.MessageBox]::Show("No se encontraron archivos INI en $iniPath.", "Error", [System.Windows.MessageBoxButton]::OK)
                Write-Host "`tNo se encontraron archivos INI en $iniPath." -ForegroundColor Red
                return
            }
            $iniSQLFile = $null
            $iniDBFFile = $null
            foreach ($iniFile in $iniFiles) {
                $content = Get-Content $iniFile.FullName
                if ($content -match "Provider=VFPOLEDB.1" -and -not $iniDBFFile) {
                    $iniDBFFile = $iniFile
                }
                if ($content -match "Provider=SQLOLEDB.1" -and -not $iniSQLFile) {
                    $iniSQLFile = $iniFile
                }
                if ($iniSQLFile -and $iniDBFFile) {
                    break
                }
            }
            if (-not $iniSQLFile -or -not $iniDBFFile) {
                [System.Windows.MessageBox]::Show("No se encontraron los archivos INI esperados.", "Error", [System.Windows.MessageBoxButton]::OK)
                Write-Host "`tNo se encontraron los archivos INI esperados." -ForegroundColor Red
                Write-Host "`tArchivos encontrados:" -ForegroundColor Yellow
                $iniFiles | ForEach-Object { Write-Host "`t- $_.Name" }
                return
            }
            $currentConfig = if ($isSQL) { "SQL" } else { "DBF" }
            $newConfig = if ($isSQL) { "DBF" } else { "SQL" }
            $message = "Actualmente tienes configurado: $currentConfig.`n¿Quieres cambiar a $newConfig?"
            $result = [System.Windows.MessageBox]::Show($message, "Cambiar Configuración", [System.Windows.MessageBoxButton]::YesNo)
            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                if ($newConfig -eq "SQL") {
                    Write-Host "`tCambiando a SQL: C:\Windows\SysWOW64\Syscfg45_2.0.dll" -ForegroundColor Yellow
                    Write-Host "`t494E5354414C4C=1"
                    Write-Host "`t56455253495354454D41=3"
                    (Get-Content $syscfgPath) -replace "494E5354414C4C=2", "494E5354414C4C=1" | Set-Content $syscfgPath
                    (Get-Content $syscfgPath) -replace "56455253495354454D41=2", "56455253495354454D41=3" | Set-Content $syscfgPath
                } else {
                    Write-Host "`tCambiando a DBF: C:\Windows\SysWOW64\Syscfg45_2.0.dll" -ForegroundColor Yellow
                    Write-Host "`t494E5354414C4C=2"
                    Write-Host "`t56455253495354454D41=1"
                    (Get-Content $syscfgPath) -replace "494E5354414C4C=1", "494E5354414C4C=2" | Set-Content $syscfgPath
                    (Get-Content $syscfgPath) -replace "56455253495354454D41=3", "56455253495354454D41=2" | Set-Content $syscfgPath
                }
                if ($newConfig -eq "SQL") {
                    Rename-Item -Path $iniDBFFile.FullName -NewName "checadorsql_DBF_old.ini" -ErrorAction Stop
                    Rename-Item -Path $iniSQLFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
                } else {
                    Rename-Item -Path $iniSQLFile.FullName -NewName "checadorsql_SQL_old.ini" -ErrorAction Stop
                    Rename-Item -Path $iniDBFFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
                }
                [System.Windows.MessageBox]::Show("Configuración cambiada exitosamente.", "Éxito", [System.Windows.MessageBoxButton]::OK)
                Write-Host "Configuración cambiada exitosamente." -ForegroundColor Green
            }
        })
    $LZMAbtnBuscarCarpeta.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
            if (-not (Test-Path $LZMAregistryPath)) {
                Write-Host "`nLa ruta del registro no existe: $LZMAregistryPath" -ForegroundColor Yellow
                [System.Windows.MessageBox]::Show("La ruta del registro no existe:`n$LZMAregistryPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
            try {
                Write-Host "`tLeyendo subcarpetas de LZMA…" -ForegroundColor Gray
                $LZMcarpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }
                if ($LZMcarpetasPrincipales.Count -lt 1) {
                    Write-Host "`tNo se encontraron carpetas principales." -ForegroundColor Yellow
                    [System.Windows.MessageBox]::Show("No se encontraron carpetas principales en la ruta del registro.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
                $instaladores = @()
                foreach ($carpeta in $LZMcarpetasPrincipales) {
                    $subdirs = Get-ChildItem -Path $carpeta.PSPath | Where-Object { $_.PSIsContainer }
                    foreach ($sd in $subdirs) {
                        $instaladores += [PSCustomObject]@{
                            Name = $sd.PSChildName
                            Path = $sd.PSPath
                        }
                    }
                }
                if ($instaladores.Count -lt 1) {
                    Write-Host "`tNo se encontraron subcarpetas." -ForegroundColor Yellow
                    [System.Windows.MessageBox]::Show("No se encontraron subcarpetas en la ruta del registro.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
                $instaladores = $instaladores | Sort-Object -Property Name -Descending
                Show-LZMADialog -Instaladores $instaladores
            } catch {
                Write-Host "`tError accediendo al registro: $_" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Error accediendo al registro:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
    $btnConfigurarIPs.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Show-IPConfigDialog
        })
    $btnAddUser.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Show-AddUserDialog
        })
    $btnCreateAPK.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $dllPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\up.dll"
            $infoPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\info.txt"
            try {
                Write-Host "`nIniciando proceso de creación de APK..." -ForegroundColor Cyan
                if (-not (Test-Path $dllPath)) {
                    Write-Host "Componente necesario no encontrado. Verifique la instalación del Enlace Android." -ForegroundColor Red
                    [System.Windows.MessageBox]::Show("Componente necesario no encontrado. Verifique la instalación del Enlace Android.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
                if (-not (Test-Path $infoPath)) {
                    Write-Host "Archivo de configuración no encontrado. Verifique la instalación del Enlace Android." -ForegroundColor Red
                    [System.Windows.MessageBox]::Show("Archivo de configuración no encontrado. Verifique la instalación del Enlace Android.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
                $jsonContent = Get-Content $infoPath -Raw | ConvertFrom-Json
                $versionApp = $jsonContent.versionApp
                Write-Host "Versión detectada: $versionApp" -ForegroundColor Green
                $confirmation = [System.Windows.MessageBox]::Show("Se creará el APK versión: $versionApp`n¿Desea continuar?", "Confirmación", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
                    Write-Host "Proceso cancelado por el usuario" -ForegroundColor Yellow
                    return
                }
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "Archivo APK (*.apk)|*.apk"
                $saveDialog.FileName = "SRM_$versionApp.apk"
                $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
                if ($saveDialog.ShowDialog() -ne $true) {
                    Write-Host "Guardado cancelado por el usuario" -ForegroundColor Yellow
                    return
                }
                Copy-Item -Path $dllPath -Destination $saveDialog.FileName -Force
                Write-Host "APK generado exitosamente en: ((
(saveDialog.FileName)" -ForegroundColor Green
                [System.Windows.MessageBox]::Show("APK creado correctamente!", "Éxito", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            } catch {
                Write-Host "Error durante el proceso: ((
(_.Exception.Message)" -ForegroundColor Red
                [System.Windows.MessageBox]::Show("Error durante la creación del APK. Consulte la consola para más detalles.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        })
    $btnConnectDb.Add_Click({
            Write-Host "`nConectando a la instancia..." -ForegroundColor Gray
            try {
                if ($null -eq $global:txtServer -or $null -eq $global:txtUser -or $null -eq $global:txtPassword) {
                    throw "Error interno: controles de conexión no inicializados."
                }
                $serverText = $global:txtServer.Text.Trim()
                $userText = $global:txtUser.Text.Trim()
                $passwordText = $global:txtPassword.Password
                Write-DzDebug "`t[DEBUG]`t[DEBUG] | Server='$serverText' User='$userText' PasswordLen=$($passwordText.Length)"
                if ([string]::IsNullOrWhiteSpace($serverText) -or [string]::IsNullOrWhiteSpace($userText) -or [string]::IsNullOrWhiteSpace($passwordText)) {
                    throw "Complete todos los campos de conexión"
                }
                $securePassword = (New-Object System.Net.NetworkCredential('', $passwordText)).SecurePassword
                $credential = New-Object System.Management.Automation.PSCredential($userText, $securePassword)
                $global:server = $serverText
                $global:user = $userText
                $global:password = $passwordText
                $global:dbCredential = $credential
                $databases = Get-SqlDatabases -Server $serverText -Credential $credential
                if (-not $databases -or $databases.Count -eq 0) {
                    throw "Conexión correcta, pero no se encontraron bases de datos disponibles."
                }
                $global:cmbDatabases.Items.Clear()
                foreach ($db in $databases) {
                    [void]$global:cmbDatabases.Items.Add($db)
                }
                $global:cmbDatabases.IsEnabled = $true
                $global:cmbDatabases.SelectedIndex = 0
                $global:lblConnectionStatus.Content = @"
Conectado a:
Servidor: $serverText
Base de datos: ((
(global:database)
"@.Trim()
                $global:lblConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Green
                $global:txtServer.IsEnabled = $false
                $global:txtUser.IsEnabled = $false
                $global:txtPassword.IsEnabled = $false
                $global:btnExecute.IsEnabled = $true
                $global:cmbQueries.IsEnabled = $true
                $global:btnConnectDb.IsEnabled = $false
                $global:btnBackup.IsEnabled = $true
                $global:btnDisconnectDb.IsEnabled = $true
                $global:rtbQuery.IsEnabled = $true
            } catch {
                Write-DzDebug "`t[DEBUG]`t[DEBUG][btnConnectDb] CATCH: ((
(_.Exception.Message)"
                Write-DzDebug "`t[DEBUG]`t[DEBUG][btnConnectDb] Tipo: ((
(_.Exception.GetType().FullName)"
                Write-DzDebug "`t[DEBUG]`t[DEBUG][btnConnectDb] Stack: ((
(_.ScriptStackTrace)"
                [System.Windows.MessageBox]::Show("Error de conexión: ((
(_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                Write-Host "Error | Error de conexión: ((
(_.Exception.Message)" -ForegroundColor Red
            }
        })
    $btnDisconnectDb.Add_Click({
            try {
                if ($global:connection -and $global:connection.State -ne [System.Data.ConnectionState]::Closed) {
                    $global:connection.Close()
                    $global:connection.Dispose()
                }
                $global:connection = $null
                $global:dbCredential = $null
                $global:lblConnectionStatus.Content = "Conectado a BDD: Ninguna"
                $global:lblConnectionStatus.Foreground = [System.Windows.Media.Brushes]::Red
                $global:btnConnectDb.IsEnabled = $true
                $global:btnBackup.IsEnabled = $false
                $global:btnDisconnectDb.IsEnabled = $false
                $global:btnExecute.IsEnabled = $false
                $global:rtbQuery.IsEnabled = $false
                $global:txtServer.IsEnabled = $true
                $global:txtUser.IsEnabled = $true
                $global:txtPassword.IsEnabled = $true
                $global:cmbQueries.IsEnabled = $false
                $global:cmbDatabases.Items.Clear()
                $global:cmbDatabases.IsEnabled = $false
                Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow
            } catch {
                Write-Host "`nError al desconectar: $($_.Exception.Message)" -ForegroundColor Red
            }
        })

    $btnExecute.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            try {
                $selectedDb = $global:cmbDatabases.SelectedItem
                if (-not $selectedDb) { throw "Selecciona una base de datos" }
                $rawQuery = New-Object System.Windows.Documents.TextRange($global:rtbQuery.Document.ContentStart, $global:rtbQuery.Document.ContentEnd)
                $cleanQuery = Remove-SqlComments -Query $rawQuery.Text
                $result = Execute-SqlQuery -server $global:server -database $selectedDb -query $cleanQuery
                if ($result -and $result.ContainsKey('Messages') -and $result.Messages) {
                    if ($result.Messages.Count -gt 0) {
                        Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
                        $result.Messages | ForEach-Object { Write-Host $_ }
                    }
                }
                if ($result -and $result.ContainsKey('DataTable') -and $result.DataTable) {
                    $global:dgvResults.ItemsSource = $result.DataTable.DefaultView
                    $global:dgvResults.IsEnabled = $true
                    Write-Host "`nColumnas obtenidas: $($result.DataTable.Columns.ColumnName -join ', ')" -ForegroundColor Cyan
                    if ($result.DataTable.Rows.Count -eq 0) {
                        Write-Host "La consulta no devolvió resultados" -ForegroundColor Yellow
                    } else {
                        $result.DataTable | Format-Table -AutoSize | Out-String | Write-Host
                    }
                } elseif ($result -and $result.ContainsKey('RowsAffected')) {
                    Write-Host "`nFilas afectadas: $($result.RowsAffected)" -ForegroundColor Green
                    $rowsAffectedTable = New-Object System.Data.DataTable
                    $rowsAffectedTable.Columns.Add("Resultado") | Out-Null
                    $rowsAffectedTable.Rows.Add("Filas afectadas: $($result.RowsAffected)") | Out-Null
                    $global:dgvResults.ItemsSource = $rowsAffectedTable.DefaultView
                    $global:dgvResults.IsEnabled = $true
                } else {
                    Write-Host "`nNo se recibió DataTable ni RowsAffected en el resultado." -ForegroundColor Yellow
                }
            } catch {
                $errorTable = New-Object System.Data.DataTable
                $errorTable.Columns.Add("Tipo") | Out-Null
                $errorTable.Columns.Add("Mensaje") | Out-Null
                $errorTable.Columns.Add("Detalle") | Out-Null
                $rawText = New-Object System.Windows.Documents.TextRange($global:rtbQuery.Document.ContentStart, $global:rtbQuery.Document.ContentEnd)
                $cleanQuery = $rawText.Text -replace '(?s)/\*.*?\*/', '' -replace '(?m)^\s*--.*'
                $shortQuery = if ($cleanQuery.Length -gt 50) { $cleanQuery.Substring(0, 47) + "..." } else { $cleanQuery }
                $errorTable.Rows.Add("ERROR SQL", $_.Exception.Message, $shortQuery) | Out-Null
                $global:dgvResults.ItemsSource = $errorTable.DefaultView
                Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
                Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "Consulta: $shortQuery" -ForegroundColor Cyan
                Write-Host "====================================" -ForegroundColor Red
            }
        })
    function Initialize-ThreadJob {
        [CmdletBinding()]
        param()

        Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Verificando módulo ThreadJob"

        # Verificar si el módulo está disponible
        if (Get-Module -ListAvailable -Name ThreadJob) {
            Import-Module ThreadJob -Force
            Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob importado"
            return $true
        } else {
            Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob no encontrado, intentando instalar"

            try {
                # Intentar instalar desde PSGallery
                Install-Module -Name ThreadJob -Force -Scope CurrentUser -ErrorAction Stop
                Import-Module ThreadJob -Force
                Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob instalado e importado"
                return $true
            } catch {
                Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Error instalando ThreadJob: $_"
                return $false
            }
        }
    }

    # Agregar al inicio de tu script principal
    if (-not (Initialize-ThreadJob)) {
        Write-Host "Advertencia: No se pudo cargar ThreadJob. Las funciones de backup pueden no funcionar correctamente." -ForegroundColor Yellow
    }
    $btnBackup.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso de Backup - - -" -ForegroundColor Gray
            Show-BackupDialog -Server $global:server -User $global:user -Password $global:password -Database $global:cmbDatabases.SelectedItem
        })
    # Crear el scriptblock con GetNewClosure() para capturar el contexto
    $closeWindowScript = {
        Write-Host "Cerrando aplicación..." -ForegroundColor Yellow
        Write-DzDebug "`t[DEBUG]`t[DEBUG] Botón Salir presionado" -Color DarkGray

        try {
            # Método 1: Buscar la ventana desde el sender
            $btn = $args[0]  # El botón que disparó el evento
            $win = [System.Windows.Window]::GetWindow($btn)

            if ($null -ne $win) {
                $win.Close()
                Write-DzDebug "`t[DEBUG]`t[DEBUG] Ventana cerrada (método 1)" -Color DarkGray
            } else {
                # Método 2: Usar la ventana capturada
                $window.Close()
                Write-DzDebug "`t[DEBUG]`t[DEBUG] Ventana cerrada (método 2)" -Color DarkGray
            }
        } catch {
            Write-Host "Error al cerrar: $_" -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Error: $_" -Color Red
        }
    }.GetNewClosure()

    $btnExit.Add_Click($closeWindowScript)
    Write-Host "✓ Formulario WPF creado exitosamente" -ForegroundColor Green
    return $window
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
        [System.Windows.MessageBox]::Show("No se pudo crear la interfaz gráfica. Verifique los logs.", "Error crítico")
        return
    }

    try {
        Write-Host "Mostrando formulario WPF..." -ForegroundColor Yellow
        Write-DzDebug "`t[DEBUG]`t[DEBUG] Mostrando ventana principal" -Color DarkGray

        # ShowDialog() bloquea hasta que se cierre la ventana
        $null = $mainForm.ShowDialog()

        Write-Host "Aplicación finalizada correctamente." -ForegroundColor Green
        Write-DzDebug "`t[DEBUG]`t[DEBUG] Aplicación finalizada" -Color DarkGray

    } catch {
        Write-Host "Error mostrando formulario: $_" -ForegroundColor Red
        Write-DzDebug "`t[DEBUG]`t[DEBUG] Error mostrando formulario: $_" -Color Red
        Write-DzDebug "`t[DEBUG]`t[DEBUG] Stack trace: $($_.ScriptStackTrace)" -Color Red
        [System.Windows.MessageBox]::Show("Error: $_", "Error en la aplicación")
    }
    # NO HAY FINALLY - Dejar que PowerShell maneje el cierre naturalmente
}
try {
    Start-Application
} catch {
    Write-Host "Error fatal: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    pause
    exit 1
}