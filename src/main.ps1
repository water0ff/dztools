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
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "✓ WPF cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando WPF: $_" -ForegroundColor Red
    pause
    exit 1
}
if (Get-Command Set-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}
$script:ProgressActive = $false
$script:LastProgressLen = 0
function Show-GlobalProgress {
    param([int]$Percent, [string]$Status)
    $Percent = [math]::Max(0, [math]::Min(100, $Percent))
    $width = 40
    $filled = [math]::Round(($Percent / 100) * $width)
    $bar = "[" + ("#" * $filled).PadRight($width) + "]"
    $text = "{0} {1,3}% - {2}" -f $bar, $Percent, $Status
    $max = [System.Console]::WindowWidth - 1
    if ($max -gt 10 -and $text.Length -gt $max) { $text = $text.Substring(0, $max) }
    $pad = ""
    if ($script:LastProgressLen -gt $text.Length) { $pad = " " * ($script:LastProgressLen - $text.Length) }
    Write-Host ("`r" + $text + $pad) -NoNewline
    $script:ProgressActive = $true
    $script:LastProgressLen = ($text.Length + $pad.Length)
}
function Stop-GlobalProgress {
    if ($script:ProgressActive) {
        Write-Host ""
        $script:ProgressActive = $false
        $script:LastProgressLen = 0
    }
}
function Write-Log {
    param([Parameter(Mandatory)][string]$Message, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    Stop-GlobalProgress
    Write-Host $Message -ForegroundColor $Color
}
Write-Host "`nImportando módulos..." -ForegroundColor Yellow
$modulesPath = Join-Path $PSScriptRoot "modules"
$modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "SqlTreeView.psm1", "Installers.psm1", "WindowsUtilities.psm1", "NationalUtilities.psm1", "SqlOps.psm1")
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
- SSMS Portable
    * TreeView nuevo!
    * Crear y eliminar bases de datos
    * Attach / detach de base de datos
    * Backup de bases de datos con compresión
    * Restaurar permite seleccionar ruta de archivos logicos
    * Funciones para ver tamaño y reparar base de datos
    * Ejecución de queries
    * Exportar resultados a CSV/Excel
    * Queries predefinidas
    * Mejoras en seguridad y manejo de errores
    * Carga de INIS en la conexión a BDD.
    * Multiples Queries (MultiQuery)
- Instalador de impresoras Generic Text por IP
- Registro y deregistro de Dlls
- Configuraciones de Firewall
    * Buscar reglas existentes "deshabilitada termporalmente"
    * Agregar reglas nuevas
- Nueva interfaz WPF
    * Fuentes y colores actualizados.
    * Modo oscuro
    * Modo debug
- Instalador de herramientas (Choco)
    * Nuevo form para buscar paquetes choco.
    * Permite instalar paquetes.
    * Permite desinstalar paquetes.
    * Ver estado de paquetes instalados.
- Restructura del proceso de Backups (choco).
    * Progressbar animada
    * Se agregó compresión con contraseña de respaldos
- Se agregó consola de cambios y tool tip para botones
- Busqueda avanzada de puertos SQL
- Query Browser para SQL en pestaña: Base de datos
- - Ahora se pueden agregar comentarios con "-" y entre "/* */"
- - Tabla en consola
- - Obtener columnas en consola
- Se agregó botón para limpiar AnyDesk
- Se agregó botón para mostrar impresoras instaladas
- Se agregó botón para limpiar y reiniciar cola de impresión
- Se agregó botón para agregar usuario de Windows
- Se agregó botón para forzar actualización de datos del sistema
- Se agregó botón para buscar instalador LZMA
- Se agregó botón para agregar IPs a adaptadores de red
- Se agregó botón para cambiar OTM a SQL/DBF
- Se agregó botón para permisos en C:\NationalSoft
- Se agregó botón para creación de SRM APK
- Se agregó botón para extractor de instalador
- Se agregó botón para ejecutar SQL Server Management Studio
- Se agregó botón para ejecutar Database4
- Se agregó botón para ejecutar SQL Server Manager
- Se agregó botón para ejecutar ExpressProfiler
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
    } catch {
        Write-Host "Advertencia: No se pudo inicializar la configuración de debug. $_" -ForegroundColor Yellow
    }
    return $true
}
$global:isHighlightingQuery = $false
function New-MainForm {
    Write-Host "`nCreando formulario principal WPF..." -ForegroundColor Yellow
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Gerardo Zermeño Tools"
        Height="650" Width="900" MinHeight="600" MinWidth="1000"
        WindowStartupLocation="CenterScreen" WindowState="Normal"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="{x:Type Control}">
            <Setter Property="FontFamily" Value="{DynamicResource UiFontFamily}"/>
            <Setter Property="FontSize" Value="{DynamicResource UiFontSize}"/>
        </Style>
        <Style TargetType="{x:Type TextBlock}">
            <Setter Property="FontFamily" Value="{DynamicResource UiFontFamily}"/>
            <Setter Property="FontSize" Value="{DynamicResource UiFontSize}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>
        <Style TargetType="{x:Type Label}">
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>
        <Style TargetType="{x:Type TabControl}">
            <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="{x:Type TabItem}">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="2,0,0,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TabItem}">
                        <Border Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6,6,0,0"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource PanelBg}"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource AccentPrimary}"/>
                                <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource AccentPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="{x:Type TextBox}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="{x:Type PasswordBox}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="{x:Type ComboBox}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="{x:Type CheckBox}">
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>
        <Style TargetType="{x:Type RichTextBox}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>
        <Style TargetType="{x:Type Paragraph}">
            <Setter Property="Margin" Value="0"/>
        </Style>
        <Style TargetType="{x:Type DataGrid}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="RowBackground" Value="{DynamicResource ControlBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource PanelBg}"/>
        </Style>
        <Style TargetType="{x:Type DataGridRow}">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="{x:Type DataGridColumnHeader}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        </Style>
        <Style TargetType="{x:Type DataGridCell}">
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        </Style>
        <Style x:Key="InfoHeaderTextBoxStyle" TargetType="{x:Type TextBox}" BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style x:Key="ConsoleTextBoxStyle" TargetType="{x:Type TextBox}" BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="FontFamily" Value="{DynamicResource CodeFontFamily}"/>
            <Setter Property="FontSize" Value="{DynamicResource CodeFontSize}"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
        </Style>
        <Style x:Key="GeneralButtonStyle" TargetType="{x:Type Button}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="Bd"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="8"
                                SnapsToDevicePixels="True">
                            <ContentPresenter
                                Margin="{TemplateBinding Padding}"
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                RecognizesAccessKey="True"
                                TextElement.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{DynamicResource AccentPrimary}"/>
                                <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{DynamicResource AccentPrimary}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.92"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="Column1ButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentMagenta}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentMagentaHover}"/>
                    <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentBlueHover}"/>
                    <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DatabaseButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentDatabase}"/>
            <Setter Property="Foreground" Value="#111111"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentDatabaseHover}"/>
                    <Setter Property="Foreground" Value="#111111"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DbConnectButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
        <Setter Property="Background" Value="{DynamicResource AccentGreen}"/>
        <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="{DynamicResource AccentGreenHover}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            </Trigger>
        </Style.Triggers>
        </Style>
        <Style x:Key="DbDisconnectButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
        <Setter Property="Background" Value="{DynamicResource AccentRed}"/>
        <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="{DynamicResource AccentRedHover}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            </Trigger>
        </Style.Triggers>
        </Style>
        <Style x:Key="NationalSoftButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentOrange}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentOrangeHover}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="TogglePillStyle" TargetType="{x:Type ToggleButton}">
            <Setter Property="Width" Value="84"/>
            <Setter Property="Height" Value="30"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type ToggleButton}">
                        <Grid Width="{TemplateBinding Width}" Height="{TemplateBinding Height}">
                            <Border x:Name="SwitchBorder"
                                    Background="{DynamicResource BorderBrushColor}"
                                    BorderBrush="{DynamicResource BorderBrushColor}"
                                    BorderThickness="1"
                                    CornerRadius="15"/>
                            <Border x:Name="SwitchThumb"
                                    Width="22" Height="22"
                                    Background="{DynamicResource FormBg}"
                                    CornerRadius="11"
                                    HorizontalAlignment="Left"
                                    Margin="4,4,0,4"/>
                            <TextBlock x:Name="SwitchLabel"
                                       Text="OFF"
                                       Foreground="{DynamicResource FormFg}"
                                       FontWeight="Bold"
                                       HorizontalAlignment="Right"
                                       VerticalAlignment="Center"
                                       Margin="0,0,8,0"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="SwitchBorder" Property="Background" Value="{DynamicResource AccentPrimary}"/>
                                <Setter TargetName="SwitchBorder" Property="BorderBrush" Value="{DynamicResource AccentPrimary}"/>
                                <Setter TargetName="SwitchThumb" Property="HorizontalAlignment" Value="Right"/>
                                <Setter TargetName="SwitchThumb" Property="Margin" Value="0,4,4,4"/>
                                <Setter TargetName="SwitchLabel" Property="Text" Value="ON"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="SwitchBorder" Property="BorderBrush" Value="{DynamicResource AccentSecondary}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Background="{DynamicResource FormBg}">
        <TabControl Name="tabControl" Margin="5">
            <TabItem Name="tabAplicaciones">
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="🧩" Margin="0,0,6,0"/>
                        <TextBlock Text="Aplicaciones"/>
                    </StackPanel>
                </TabItem.Header>
                <Grid Background="{DynamicResource PanelBg}">
                    <TextBox Name="lblHostname"
                             Width="220" Height="40" Margin="10,1,0,0"
                             VerticalAlignment="Top" HorizontalAlignment="Left"
                             Style="{StaticResource InfoHeaderTextBoxStyle}"
                             Text="HOSTNAME"
                             Cursor="Hand"
                             IsReadOnly="True"
                             IsReadOnlyCaretVisible="False"
                             TextAlignment="Center"
                             VerticalContentAlignment="Center"
                             HorizontalContentAlignment="Center"
                             Focusable="False"
                             IsTabStop="False"/>
                    <TextBox Name="lblPort"
                             Width="220" Height="40" Margin="250,1,0,0"
                             VerticalAlignment="Top" HorizontalAlignment="Left"
                             Style="{StaticResource InfoHeaderTextBoxStyle}"
                             Text="Puerto: No disponible"/>
                    <TextBox Name="txt_IpAdress"
                             Width="220" Height="40" Margin="490,1,0,0"
                             VerticalAlignment="Top" HorizontalAlignment="Left"
                             Style="{StaticResource InfoHeaderTextBoxStyle}"/>
                    <TextBox Name="txt_AdapterStatus"
                             Width="220" Height="40" Margin="730,1,0,0"
                             VerticalAlignment="Top" HorizontalAlignment="Left"
                             Style="{StaticResource InfoHeaderTextBoxStyle}"/>
                    <Button Content="Instalar Herramientas" Name="btnInstalarHerramientas" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,50,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Ejecutar ExpressProfiler" Name="btnProfiler" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,90,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Ejecutar Database4" Name="btnDatabase" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Ejecutar Manager" Name="btnSQLManager" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,170,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Ejecutar Management" Name="btnSQLManagement" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,210,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Printer Tools" Name="btnPrinterTool" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,250,0,0" Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Clear AnyDesk" Name="btnClearAnyDesk" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,290,0,0"
                            Style="{StaticResource Column1ButtonStyle}"/>
                    <Button Content="Lector DP - Permisos" Name="btnLectorDPicacls" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,50,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Buscar Instalador LZMA" Name="LZMAbtnBuscarCarpeta" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,90,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Agregar IPs" Name="btnConfigurarIPs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,130,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Agregar usuario de Windows" Name="btnAddUser" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,170,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Configuraciones de Firewall" Name="btnFirewallConfig" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,210,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Actualizar datos del sistema" Name="btnForzarActualizacion" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,250,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Mostrar Impresoras" Name="btnShowPrinters" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,290,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Instalar impresora" Name="btnInstallPrinter" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,330,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Limpia y Reinicia Cola de Impresión" Name="btnClearPrintJobs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,370,0,0"
                            Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Aplicaciones National Soft" Name="btnAplicacionesNS" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,50,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Cambiar OTM a SQL/DBF" Name="btnCambiarOTM" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,90,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Permisos C:\NationalSoft" Name="btnCheckPermissions" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,130,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Creación de SRM APK" Name="btnCreateAPK" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,170,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Extractor de Instalador" Name="btnExtractInstaller" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,210,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Instaladores NS" Name="btnInstaladoresNS" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,250,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <Button Content="Registro de dlls" Name="btnRegisterDlls" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,290,0,0" Style="{StaticResource NationalSoftButtonStyle}"/>
                    <TextBox Name="txt_InfoInstrucciones" HorizontalAlignment="Left" VerticalAlignment="Top"
                             Width="220" Height="300" Margin="730,50,0,0" Style="{StaticResource ConsoleTextBoxStyle}"
                             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                    <Border Name="cardQuickSettings"
                            Width="220" Height="150"
                            Margin="730,370,0,0"
                            HorizontalAlignment="Left"
                            VerticalAlignment="Top"
                            Background="{DynamicResource ControlBg}"
                            BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1"
                            CornerRadius="10"
                            Padding="10">
                        <Border.Effect>
                            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
                        </Border.Effect>
                        <StackPanel>
                            <TextBlock Text="Ajustes rápidos"
                                       FontWeight="Bold"
                                       Foreground="{DynamicResource AccentPrimary}"
                                       Margin="0,0,0,8"/>
                            <Grid Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="10"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                            </Grid>
                            <Grid Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="10"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <ToggleButton Name="tglDarkMode" Grid.Column="0" Style="{StaticResource TogglePillStyle}"/>
                                <TextBlock Grid.Column="2" Text="🌙 Dark Mode" VerticalAlignment="Center"/>
                            </Grid>
                            <Grid Margin="0,0,0,6">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="10"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <ToggleButton Name="tglDebugMode" Grid.Column="0" Style="{StaticResource TogglePillStyle}"/>
                                <TextBlock Grid.Column="2" Text="🐞 DEBUG" VerticalAlignment="Center"/>
                            </Grid>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>
<TabItem Name="tabProSql">
  <TabItem.Header>
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="🗄️" Margin="0,0,6,0"/>
      <TextBlock Text="SSMS portable"/>
    </StackPanel>
  </TabItem.Header>
  <Grid Background="{DynamicResource PanelBg}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>  <!-- Conexión -->
      <RowDefinition Height="Auto"/>  <!-- Barra intermedia -->
      <RowDefinition Height="*"/>     <!-- Área principal -->
      <RowDefinition Height="Auto"/>  <!-- StatusBar -->
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Margin="10" Padding="10"
            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
            CornerRadius="6" Background="{DynamicResource ControlBg}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0">
          <Button Name="btnConnectDb"
                Width="60" Height="30" Margin="0,0,8,0"
                Style="{StaticResource DbConnectButtonStyle}"
                ToolTip="Conectar">
        <TextBlock Text="🔌✅" FontFamily="Segoe UI Emoji" FontSize="16" HorizontalAlignment="Center"/>
        </Button>
        <Button Name="btnDisconnectDb"
                Width="60" Height="30"
                Style="{StaticResource DbDisconnectButtonStyle}"
                ToolTip="Desconectar" IsEnabled="False">
        <TextBlock Text="🔌✖" FontFamily="Segoe UI Emoji" FontSize="16" HorizontalAlignment="Center"/>
        </Button>
        </StackPanel>
        <WrapPanel Grid.Column="1">
          <StackPanel Margin="0,0,16,0">
            <TextBlock Text="Instancia SQL:"/>
            <ComboBox Name="txtServer" Width="180" IsEditable="True" Text=".\NationalSoft"/>
          </StackPanel>
          <StackPanel Margin="0,0,16,0">
            <TextBlock Text="Usuario:"/>
            <TextBox Name="txtUser" Width="160" Text="sa"/>
          </StackPanel>
          <StackPanel Margin="0,0,16,0">
            <TextBlock Text="Contraseña:"/>
            <PasswordBox Name="txtPassword" Width="160"/>
          </StackPanel>
          <StackPanel Margin="0,0,16,0">
            <TextBlock Text="Base de datos:"/>
            <ComboBox Name="cmbDatabases" Width="180" IsEnabled="False"/>
          </StackPanel>
        </WrapPanel>
      </Grid>
    </Border>
    <Border Grid.Row="1" Margin="10,0,10,10" Padding="10"
            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
            CornerRadius="6" Background="{DynamicResource ControlBg}">
      <StackPanel Orientation="Horizontal">
        <Button Content="Ejecutar (F5)" Name="btnExecute"
                Width="110" Height="30" Margin="0,0,8,0"
                Style="{StaticResource DatabaseButtonStyle}" IsEnabled="False"/>
        <ComboBox Name="cmbQueries" Width="280" Margin="0,0,8,0"
                  IsEnabled="False" ToolTip="Consultas predefinidas"/>
        <Button Content="Limpiar" Name="btnClearQuery"
                Width="90" Height="30"
                Style="{StaticResource DatabaseButtonStyle}" IsEnabled="False"/>
        <Button Content="Exportar" Name="btnExport"
                Width="100" Height="30" Margin="8,0,0,0"
                Style="{StaticResource DatabaseButtonStyle}" IsEnabled="False"/>
      </StackPanel>
    </Border>
    <Grid Grid.Row="2" Margin="10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="250" MinWidth="200"/>
        <ColumnDefinition Width="5"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Grid.Column="0"
              BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
              CornerRadius="6" Background="{DynamicResource ControlBg}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <TextBlock Grid.Row="0" Text="Explorador de Objetos"
                     Padding="8" FontWeight="Bold"
                     Background="{DynamicResource AccentPrimary}"
                     Foreground="{DynamicResource OnAccentFg}"/>
          <TreeView Grid.Row="1" Name="tvDatabases" Padding="4"/>
        </Grid>
      </Border>
      <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch"
                    Background="{DynamicResource BorderBrushColor}"/>
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="*" MinHeight="150"/>
          <RowDefinition Height="5"/>
          <RowDefinition Height="2*" MinHeight="200"/>
        </Grid.RowDefinitions>
        <TabControl Name="tcQueries" Grid.Row="0" Background="{DynamicResource ControlBg}">
          <TabItem Header="Consulta 1">
            <Border BorderBrush="{DynamicResource BorderBrushColor}"
                    BorderThickness="1" Margin="5" CornerRadius="4">
              <RichTextBox Name="rtbQueryEditor1"
                           VerticalScrollBarVisibility="Auto"
                           AcceptsReturn="True" AcceptsTab="True"
                           FontFamily="Consolas" FontSize="12"/>
            </Border>
          </TabItem>
          <TabItem Header="+" Name="tabAddQuery" IsEnabled="True"/>
        </TabControl>
        <GridSplitter Grid.Row="1" Height="5" HorizontalAlignment="Stretch"
                      Background="{DynamicResource BorderBrushColor}"/>

        <TabControl Name="tcResults" Grid.Row="2" Background="{DynamicResource ControlBg}">
          <TabItem Header="Resultados">
            <DataGrid Name="dgResults" IsReadOnly="True" AutoGenerateColumns="True"
                      CanUserAddRows="False" CanUserDeleteRows="False"/>
          </TabItem>
          <TabItem Header="Mensajes">
            <TextBox Name="txtMessages" IsReadOnly="True"
                     VerticalScrollBarVisibility="Auto"
                     FontFamily="Consolas" Background="Transparent"
                     BorderThickness="0"/>
          </TabItem>
        </TabControl>
      </Grid>
    </Grid>
    <StatusBar Grid.Row="3" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}">
      <StatusBarItem>
        <TextBlock Name="lblConnectionStatus" Text="Desconectado"/>
      </StatusBarItem>
      <Separator/>
            <StatusBarItem>
                <TextBlock Name="lblExecutionTimer" Text="Tiempo: --"/>
            </StatusBarItem>
      <Separator/>
      <StatusBarItem>
        <TextBlock Name="lblRowCount" Text="Filas: --"/>
      </StatusBarItem>
    </StatusBar>
  </Grid>
</TabItem>
        </TabControl>
    </Grid>
</Window>
"@
    [xml]$xaml = $stringXaml
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    try {
        $window = [Windows.Markup.XamlReader]::Load($reader)
        $theme = Get-DzUiTheme
        $global:MainWindow = $window
        Set-DzWpfThemeResources -Window $window -Theme $theme
    } catch {
        Write-Host "`n[XAML ERROR] $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception -and ($_.Exception.PSObject.Properties.Name -contains "LineNumber")) {
            Write-Host "Linea: $($_.Exception.LineNumber)  Col: $($_.Exception.LinePosition)" -ForegroundColor Yellow
        }
        if ($_.Exception.InnerException) {
            Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
        }
        throw
    }
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
    $btnFirewallConfig = $window.FindName("btnFirewallConfig")
    $btnForzarActualizacion = $window.FindName("btnForzarActualizacion")
    $btnClearAnyDesk = $window.FindName("btnClearAnyDesk")
    $btnShowPrinters = $window.FindName("btnShowPrinters")
    $btnInstallPrinter = $window.FindName("btnInstallPrinter")
    $btnClearPrintJobs = $window.FindName("btnClearPrintJobs")
    $btnAplicacionesNS = $window.FindName("btnAplicacionesNS")
    $btnCambiarOTM = $window.FindName("btnCambiarOTM")
    $btnCheckPermissions = $window.FindName("btnCheckPermissions")
    $btnCreateAPK = $window.FindName("btnCreateAPK")
    $btnExtractInstaller = $window.FindName("btnExtractInstaller")
    $btnInstaladoresNS = $window.FindName("btnInstaladoresNS")
    $btnRegisterDlls = $window.FindName("btnRegisterDlls")
    $txtServer = $window.FindName("txtServer")
    $txtUser = $window.FindName("txtUser")
    $txtPassword = $window.FindName("txtPassword")
    $cmbDatabases = $window.FindName("cmbDatabases")
    $btnConnectDb = $window.FindName("btnConnectDb")
    $btnDisconnectDb = $window.FindName("btnDisconnectDb")
    $lblConnectionStatus = $window.FindName("lblConnectionStatus")
    $btnExecute = $window.FindName("btnExecute")
    $btnClearQuery = $window.FindName("btnClearQuery")
    $btnExport = $window.FindName("btnExport")
    $cmbQueries = $window.FindName("cmbQueries")
    $tcQueries = $window.FindName("tcQueries")
    $tcResults = $window.FindName("tcResults")
    $tvDatabases = $window.FindName("tvDatabases")
    $tabAddQuery = $window.FindName("tabAddQuery")
    $tglDarkMode = $window.FindName("tglDarkMode")
    $tglDebugMode = $window.FindName("tglDebugMode")
    $script:predefinedQueries = Get-PredefinedQueries
    $script:sqlKeywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE'
    if ($cmbQueries) {
        $cmbQueries.Items.Clear()
        $cmbQueries.Items.Add("Selecciona una consulta predefinida") | Out-Null
        foreach ($key in ($script:predefinedQueries.Keys | Sort-Object)) {
            $cmbQueries.Items.Add($key) | Out-Null
        }
        $cmbQueries.SelectedIndex = 0
        $cmbQueries.Add_SelectionChanged({
                $selectedQuery = $cmbQueries.SelectedItem
                if (-not $selectedQuery -or $selectedQuery -eq "Selecciona una consulta predefinida") { return }
                if (-not $script:predefinedQueries.ContainsKey($selectedQuery)) { return }
                $rtb = Get-ActiveQueryRichTextBox -TabControl $global:tcQueries
                Write-DzDebug "`t[DEBUG]Insertando consulta predefinida '$selectedQuery' en la pestaña consulta: $($rtb.Name)"
                if (-not $rtb) { return }
                $queryText = $script:predefinedQueries[$selectedQuery]
                $rtb.Document.Blocks.Clear()
                $paragraph = New-Object System.Windows.Documents.Paragraph
                $run = New-Object System.Windows.Documents.Run($queryText)
                $paragraph.Inlines.Add($run)
                $rtb.Document.Blocks.Add($paragraph)
                if ($script:sqlKeywords) {
                    Set-WpfSqlHighlighting -RichTextBox $rtb -Keywords $script:sqlKeywords
                }
                $rtb.Focus() | Out-Null
            })
    }
    if ($tabAddQuery) {
        $tabAddQuery.Add_PreviewMouseLeftButtonDown({
                New-QueryTab -TabControl $tcQueries | Out-Null
                $_.Handled = $true
            })
    }
    if ($tcQueries -and $tcQueries.Items.Count -eq 1) {
        New-QueryTab -TabControl $tcQueries | Out-Null
    }
    if ($btnExecute -and -not $btnExecute.IsEnabled) {
        if ($tcQueries) { $tcQueries.IsEnabled = $false }
        if ($tcResults) { $tcResults.IsEnabled = $false }
    }
    $script:execStopwatch = [System.Diagnostics.Stopwatch]::new()
    $script:execUiTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:execUiTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:execUiTimer.Add_Tick({
            if ($global:lblExecutionTimer) {
                $t = $script:execStopwatch.Elapsed
                $global:lblExecutionTimer.Text = ("Timer: {0:mm\:ss\.f}" -f $t)
            }
        })
    function Start-ExecutionTimer {
        $script:execStopwatch.Restart()
        if (-not $script:execUiTimer.IsEnabled) { $script:execUiTimer.Start() }
    }    function Stop-ExecutionTimer {
        $script:execStopwatch.Stop()
        if ($script:execUiTimer.IsEnabled) { $script:execUiTimer.Stop() }
        if ($global:lblExecutionTime) {
            $t = $script:execStopwatch.Elapsed
            $global:lblExecutionTime.Text = ("Tiempo: {0:mm\:ss\.fff}" -f $t)
        }
    }
    $global:txtServer = $txtServer
    $global:txtUser = $txtUser
    $global:txtPassword = $txtPassword
    $global:cmbDatabases = $cmbDatabases
    $global:btnConnectDb = $btnConnectDb
    $global:btnDisconnectDb = $btnDisconnectDb
    $global:btnExecute = $btnExecute
    $global:btnClearQuery = $btnClearQuery
    $global:btnExport = $btnExport
    $global:cmbQueries = $cmbQueries
    $global:tcQueries = $tcQueries
    $global:tcResults = $tcResults
    $global:tvDatabases = $tvDatabases
    $global:tabAddQuery = $tabAddQuery
    $global:rtbQueryEditor1 = $window.FindName("rtbQueryEditor1")
    $global:dgResults = $window.FindName("dgResults")
    $global:txtMessages = $window.FindName("txtMessages")
    $global:lblExecutionTimer = $window.FindName("lblExecutionTimer")
    $global:lblRowCount = $window.FindName("lblRowCount")
    $global:lblConnectionStatus = $lblConnectionStatus
    if ($global:tcQueries -and $global:tcQueries.Items.Count -gt 0 -and $global:rtbQueryEditor1) {
        $firstTab = $global:tcQueries.Items[0]
        if ($firstTab -is [System.Windows.Controls.TabItem]) {
            if (-not $firstTab.Tag -or $firstTab.Tag.Type -ne 'QueryTab') {
                $title = if ($firstTab.Header) { [string]$firstTab.Header } else { "Consulta 1" }
                $firstTab.Tag = [pscustomobject]@{
                    Type            = "QueryTab"
                    RichTextBox     = $global:rtbQueryEditor1
                    Title           = $title
                    HeaderTextBlock = $null
                    IsDirty         = $false
                }
                if (-not $global:tcQueries.SelectedItem) {
                    $global:tcQueries.SelectedIndex = 0
                }
            }
        }
    }

    $lblHostname.text = [System.Net.Dns]::GetHostName()
    $txt_InfoInstrucciones.Text = $global:defaultInstructions
    $script:initializingToggles = $true
    if ($tglDarkMode) { $tglDarkMode.IsChecked = ((Get-DzUiMode) -eq 'dark') }
    if ($tglDebugMode) { $tglDebugMode.IsChecked = (Get-DzDebugPreference) }
    $script:initializingToggles = $false
    Write-Host "`n==================================================" -ForegroundColor DarkCyan
    Write-Host "       Gerardo Zermeño Tools - Suite de Utilidades       " -ForegroundColor Green
    Write-Host "              Versión: $($global:version)               " -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor DarkCyan
    Write-Host "`nTodos los derechos reservados para Gerardo Zermeño Tools." -ForegroundColor Cyan
    Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan
    Write-Host "O crea un issue en GitHub en:" -ForegroundColor Cyan
    Write-Host "https://github.com/water0ff/dztools/issues/new" -ForegroundColor Cyan
    $script:setInstructionText = {
        param([string]$Message)
        if ($null -ne $txt_InfoInstrucciones) {
            $txt_InfoInstrucciones.Dispatcher.Invoke([action] { $txt_InfoInstrucciones.Text = $Message })
        }
    }.GetNewClosure()
    $script:showRestartNotice = {
        param([string]$settingLabel)
        Show-WpfMessageBoxSafe -Message "Se guardó $settingLabel en dztools.ini.`nReinicia la aplicación para aplicar los cambios." -Title "Reinicio requerido" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
    }.GetNewClosure()
    if ($tglDarkMode) {
        $tglDarkMode.Add_Checked({
                Write-DzDebug "`t[DEBUG]Toggle Dark Mode activado"
                if ($script:initializingToggles) { return }
                Set-DzUiMode -Mode "dark"
                if ($script:showRestartNotice -is [scriptblock]) { $script:showRestartNotice.Invoke("el modo Dark") }
            })
        $tglDarkMode.Add_Unchecked({
                Write-DzDebug "`t[DEBUG]Toggle Dark Mode desactivado"
                if ($script:initializingToggles) { return }
                Set-DzUiMode -Mode "light"
                if ($script:showRestartNotice -is [scriptblock]) { $script:showRestartNotice.Invoke("el modo Light") }
            })
    }
    if ($tglDebugMode) {
        $tglDebugMode.Add_Checked({
                if ($script:initializingToggles) { return }
                Set-DzDebugPreference -Enabled $true
                Write-Host "[DEBUG] Activado" -ForegroundColor Yellow
            })
        $tglDebugMode.Add_Unchecked({
                Write-DzDebug "`t[DEBUG]Toggle DEBUG desactivado"
                if ($script:initializingToggles) { return }
                Set-DzDebugPreference -Enabled $false
                Write-Host "[DEBUG] Desactivado" -ForegroundColor DarkGray
            })
    }
    $global:txt_AdapterStatus = $txt_AdapterStatus
    Initialize-SystemInfo -LblPort $lblPort -LblIpAddress $txt_IpAdress -LblAdapterStatus $txt_AdapterStatus -ModulesPath $modulesPath
    Load-IniConnectionsToComboBox -Combo $txtServer
    $buttonsToUpdate = @($LZMAbtnBuscarCarpeta, $btnInstalarHerramientas, $btnFirewallConfig, $btnProfiler,
        $btnDatabase, $btnSQLManager, $btnSQLManagement, $btnPrinterTool, $btnLectorDPicacls, $btnConfigurarIPs,
        $btnAddUser, $btnForzarActualizacion, $btnClearAnyDesk, $btnShowPrinters, $btnInstallPrinter, $btnClearPrintJobs, $btnAplicacionesNS,
        $btnCheckPermissions, $btnCambiarOTM, $btnCreateAPK, $btnExtractInstaller, $btnInstaladoresNS, $btnRegisterDlls)
    foreach ($button in $buttonsToUpdate) {
        $button.Add_MouseLeave({ if ($script:setInstructionText) { $script:setInstructionText.Invoke($global:defaultInstructions) } })
    }
    $lblHostname.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG] Click en lblHostname - Evento iniciado" -Color DarkGray
            try {
                $hostname = [System.Net.Dns]::GetHostName()
                if ([string]::IsNullOrWhiteSpace($hostname)) { Write-Host "`n[AVISO] No se pudo obtener el hostname." -ForegroundColor Yellow; return }
                $ok = Set-ClipboardTextSafe -Text $hostname -Owner $global:MainWindow
                if ($ok) { Write-Host "`nNombre del equipo copiado: $hostname" -ForegroundColor Green } else { Ui-Error "No se pudo copiar el hostname al portapapeles." $global:MainWindow }
            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                Ui-Error "Error: $($_.Exception.Message)" $global:MainWindow
            } finally {
                $e.Handled = $true
            }
        }.GetNewClosure())
    $lblPort.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG] Click en lblPort - Evento iniciado" -Color DarkGray
            try {
                if ($null -eq $sender) { return }
                $detail = $null
                try { $detail = [string]$sender.Tag } catch { $detail = $null }
                $textToCopy = if (-not [string]::IsNullOrWhiteSpace($detail)) {
                    $detail.Trim()
                } else {
                    try { ([string]$sender.Text).Trim() } catch { "" }
                }
                if ([string]::IsNullOrWhiteSpace($textToCopy) -or
                    $textToCopy -match 'No se encontraron' -or
                    $textToCopy -match 'No encontrado' -or
                    $textToCopy -match 'No hay') {
                    Write-Host "`n[AVISO] No hay puertos SQL para copiar." -ForegroundColor Yellow
                    return
                }
                $ok = Set-ClipboardTextSafe -Text $textToCopy -Owner $global:MainWindow
                if ($ok) {
                    Write-Host "`nPuertos SQL copaido al portapapeles: $textToCopy" -ForegroundColor Green
                } else {
                    Ui-Error "No se pudo copiar la información de puertos al portapapeles." $global:MainWindow
                }
            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                Ui-Error "Error: $($_.Exception.Message)" $global:MainWindow
            } finally {
                $e.Handled = $true
            }
        }.GetNewClosure())
    $txt_IpAdress.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG] Click en txt_IpAdress - Evento iniciado" -Color DarkGray
            try {
                $ipsText = [string]$sender.Text
                Write-DzDebug "`t[DEBUG] Contenido (sender): '$ipsText'" -Color DarkGray
                Write-DzDebug "`t[DEBUG] Length: $($ipsText.Length)" -Color DarkGray
                if ([string]::IsNullOrWhiteSpace($ipsText)) { Write-Host "`n[AVISO] No hay IPs para copiar." -ForegroundColor Yellow; return }
                $textToCopy = $ipsText.TrimEnd()
                $ok = Set-ClipboardTextSafe -Text $textToCopy -Owner $global:MainWindow
                if ($ok) { Write-Host "`nIP's copiadas al portapapeles:`n$textToCopy" -ForegroundColor Green } else { Ui-Error "No se pudieron copiar las IPs al portapapeles." $global:MainWindow }
            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                Ui-Error "Error: $($_.Exception.Message)" $global:MainWindow
            } finally {
                $e.Handled = $true
            }
        }.GetNewClosure())
    $txt_AdapterStatus.Add_PreviewMouseLeftButtonDown({
            param($sender, $e)
            Get-NetConnectionProfile | Where-Object { $_.NetworkCategory -ne 'Private' } | ForEach-Object { Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private }
            Write-Host "Todas las redes se han establecido como Privadas."
            Refresh-AdapterStatus
            $e.Handled = $true
        })
    $btnInstalarHerramientas.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Instalar Herramientas' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Instalar Herramientas' - - -" -ForegroundColor Magenta
            if (-not (Check-Chocolatey)) {
                Write-Host "Chocolatey no está instalado. No se puede abrir el menú de instaladores." -ForegroundColor Red
                return
            }
            Show-ChocolateyInstallerMenu
        })
    $btnProfiler.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Profiler' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Profiler' - - -" -ForegroundColor Magenta
            Invoke-PortableTool -ToolName "ExpressProfiler" -Url "https://github.com/ststeiger/ExpressProfiler/releases/download/1.0/ExpressProfiler20.zip" -ZipPath "C:\Temp\ExpressProfiler22wAddinSigned.zip" -ExtractPath "C:\Temp\ExpressProfiler2" -ExeName "ExpressProfiler.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnDatabase.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Database' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Database' - - -" -ForegroundColor Magenta
            Invoke-PortableTool -ToolName "Database4" -Url "https://fishcodelib.com/files/DatabaseNet4.zip" -ZipPath "C:\Temp\DatabaseNet4.zip" -ExtractPath "C:\Temp\Database4" -ExeName "Database4.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnPrinterTool.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Printer Tool' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Printer Tool' - - -" -ForegroundColor Magenta
            Invoke-PortableTool -ToolName "POS Printer Test" -Url "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip" -ZipPath "C:\Temp\RPT-RPI-Printer-Tool-1.zip" -ExtractPath "C:\Temp\RPT-RPI-Printer-Tool-1" -ExeName "POS Printer Test.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnLectorDPicacls.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Lector DP + icacls' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Lector DP + icacls' - - -" -ForegroundColor Magenta
            Invoke-LectorDP
        })
    $btnSQLManager.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'SQL Manager' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'SQL Manager' - - -" -ForegroundColor Magenta
            function Get-SQLServerManagers {
                $possiblePaths = @("${env:SystemRoot}\System32\SQLServerManager*.msc", "${env:SystemRoot}\SysWOW64\SQLServerManager*.msc")
                $managers = foreach ($pattern in $possiblePaths) { Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | ForEach-Object FullName }
                @($managers) | Where-Object { $_ } | Select-Object -Unique
            }
            $managers = Get-SQLServerManagers
            if (-not $managers -or $managers.Count -eq 0) { Ui-Error "No se encontró ninguna versión de SQL Server Configuration Manager." $global:MainWindow ; return }
            Show-SQLselector -Managers $managers
        })
    $btnSQLManagement.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'SQL Management' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'SQL Management' - - -" -ForegroundColor Magenta
            function Get-SSMSVersions {
                $ssmsPaths = @()
                $fixedPaths = @(
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server\*\Tools\Binn\ManagementStudio\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio *\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 19\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 22\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 21\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 20\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio 19\Release\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio *\Common7\IDE\Ssms.exe",
                    "${env:ProgramFiles}\Microsoft SQL Server Management Studio *\Release\Common7\IDE\Ssms.exe"
                )
                foreach ($pattern in $fixedPaths) {
                    $found = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue
                    foreach ($f in $found) {
                        if ($ssmsPaths -notcontains $f.FullName) {
                            $ssmsPaths += $f.FullName
                            Write-DzDebug "`t[DEBUG] ✓ Encontrado: $($f.FullName)" -Color Green
                        }
                    }
                }
                if ($ssmsPaths.Count -eq 0) {
                    Write-DzDebug "`t[DEBUG] No se encontró en rutas fijas. Buscando en registro..." -Color DarkGray
                    $registryPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
                    foreach ($regPath in $registryPaths) {
                        $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*SQL Server Management Studio*" -and $_.InstallLocation -and $_.InstallLocation.Trim() -ne "" }
                        foreach ($entry in $entries) {
                            $installPath = $entry.InstallLocation.Trim()
                            if (-not $installPath.EndsWith('\')) { $installPath += '\' }
                            foreach ($sub in @("Common7\IDE\Ssms.exe", "Release\Common7\IDE\Ssms.exe")) {
                                $full = Join-Path $installPath $sub
                                if (Test-Path $full) {
                                    $resolved = (Resolve-Path $full).Path
                                    if ($ssmsPaths -notcontains $resolved) {
                                        $ssmsPaths += $resolved
                                        Write-DzDebug "`t[DEBUG] ✓ Encontrado (registro): $resolved" -Color Green
                                    }
                                }
                            }
                        }
                    }
                }
                $ssmsPaths | Sort-Object -Descending
            }
            $ssmsVersions = Get-SSMSVersions
            $filteredVersions = foreach ($p in $ssmsVersions) { if ((Split-Path $p -Leaf) -eq "Ssms.exe" -and (Test-Path $p -PathType Leaf)) { $p } }
            if (-not $filteredVersions -or $filteredVersions.Count -eq 0) {
                Write-Host "`tNo se encontró ninguna versión de SSMS instalada." -ForegroundColor Red
                $wantManual = Ui-Confirm "No se encontró SQL Server Management Studio. ¿Desea buscar manualmente?" "SSMS no encontrado"
                if ($wantManual) {
                    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
                    $openFileDialog.Filter = "SSMS Executable (Ssms.exe)|Ssms.exe"
                    $openFileDialog.Title = "Seleccione Ssms.exe"
                    if ($openFileDialog.ShowDialog() -eq $true) {
                        try {
                            Start-Process -FilePath $openFileDialog.FileName
                            Write-Host "`tEjecutando: $($openFileDialog.FileName)" -ForegroundColor Green
                        } catch { Ui-Error "Error al ejecutar SSMS: $($_.Exception.Message)" $global:MainWindow }
                    }
                } else {
                    $wantDownload = Ui-Confirm "¿Desea descargar la última versión de SSMS?" "Descargar SSMS"
                    if ($wantDownload) { Start-Process "https://aka.ms/ssmsfullsetup" }
                }
                return
            }
            Write-Host "`t✓ Se encontraron $($filteredVersions.Count) instalación(es) de SSMS" -ForegroundColor Green
            Show-SQLselector -SSMSVersions $filteredVersions
        })
    $btnForzarActualizacion.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Forzar Actualización' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Forzar Actualización' - - -" -ForegroundColor Magenta
            Show-SystemComponents
            $ok = Ui-Confirm "¿Desea forzar la actualización de datos?" "Confirmación" $global:MainWindow
            if ($ok) { Start-SystemUpdate ; Ui-Info "Actualización completada" "Éxito" $global:MainWindow } else { Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red }
        })
    $btnClearAnyDesk.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Clear AnyDesk' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Clear AnyDesk' - - -" -ForegroundColor Magenta
            $ok = Ui-Confirm "¿Estás seguro de renovar AnyDesk?" "Confirmar renovación" $global:MainWindow
            if ($ok) {
                $filesToDelete = @("C:\ProgramData\AnyDesk\system.conf", "C:\ProgramData\AnyDesk\service.conf", "$env:APPDATA\AnyDesk\system.conf", "$env:APPDATA\AnyDesk\service.conf")
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
                    } catch { Write-Host "`nError al eliminar el archivo." -ForegroundColor Red }
                }
                if ($errors.Count -eq 0) { Ui-Info "$deletedFilesCount archivo(s) eliminado(s) correctamente." "Éxito" $global:MainWindow } else { Ui-Error "Se encontraron errores. Revisa la consola para más detalles." $global:MainWindow }
            } else { Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red }
        })
    $btnShowPrinters.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Show Printers' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Show Printers' - - -" -ForegroundColor Magenta
            Show-NSPrinters
        })
    $btnInstallPrinter.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Instalar impresora' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Instalar impresora' - - -" -ForegroundColor Magenta
            Show-InstallPrinterDialog
        })
    $btnClearPrintJobs.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Clear Print Jobs' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Clear Print Jobs' - - -" -ForegroundColor Magenta
            Invoke-ClearPrintJobs -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnCheckPermissions.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Revisar Permisos' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Revisar Permisos' - - -" -ForegroundColor Magenta
            if (-not (Test-Administrator)) {
                Ui-Error "Esta acción requiere permisos de administrador.`r`nPor favor, ejecuta Gerardo Zermeño Tools como administrador." $global:MainWindow
                return
            }
            Check-Permissions
        })
    $btnAplicacionesNS.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Aplicaciones NS' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Aplicaciones NS' - - -" -ForegroundColor Magenta
            $res = Get-NSApplicationsIniReport
            Show-NSApplicationsIniReport -Resultados $res
        })
    $btnRegisterDlls.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Registro registro de dlls' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Registro registro de dlls' - - -" -ForegroundColor Magenta
            Show-DllRegistrationDialog
        })
    $btnCambiarOTM.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Cambiar OTM' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Cambiar OTM' - - -" -ForegroundColor Magenta
            Invoke-CambiarOTMConfig -InfoTextBlock $txt_InfoInstrucciones
        })
    $LZMAbtnBuscarCarpeta.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Buscar Instaladores LZMA' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Buscar Instaladores LZMA' - - -" -ForegroundColor Magenta
            Show-LZMADialog
        })
    $btnConfigurarIPs.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Configurar IPs' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Configurar IPs' - - -" -ForegroundColor Magenta
            Show-IPConfigDialog
        })
    $btnAddUser.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Agregar Usuario' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Agregar Usuario' - - -" -ForegroundColor Magenta
            Show-AddUserDialog
        })
    $btnFirewallConfig.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Configuraciones de Firewall' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Configuraciones de Firewall' - - -" -ForegroundColor Magenta
            Show-FirewallConfigDialog
        })
    $btnCreateAPK.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Crear APK' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Crear APK' - - -" -ForegroundColor Magenta
            Invoke-CreateApk -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnExtractInstaller.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Extraer Instalador' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Extraer Instalador' - - -" -ForegroundColor Magenta
            Show-InstallerExtractorDialog
        })
    $btnInstaladoresNS.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Instaladores NS' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`t- - - Abriendo 'Instaladores NS' - - -" -ForegroundColor Gray
            Start-Process "https://nationalsoft-my.sharepoint.com/:f:/g/personal/gerardo_zermeno_softrestaurant_com/IgC3tKgxlNw9S7JmBk935kCrAVq9jkz06CJek9ljNOrr_Hw?e=xf2dFh"
        })
    $btnDisconnectDb.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Desconectar Base de Datos' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Desconectando de la Base de Datos - - -" -ForegroundColor Cyan
            try {
                if ($script:QueryRunning) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Cancelando query en ejecución..."
                    try {
                        if ($script:CurrentQueryPowerShell) {
                            $script:CurrentQueryPowerShell.Stop()
                            $script:CurrentQueryPowerShell.Dispose()
                        }
                        if ($script:CurrentQueryRunspace) {
                            $script:CurrentQueryRunspace.Close()
                            $script:CurrentQueryRunspace.Dispose()
                        }
                    } catch {
                        Write-DzDebug "`t[DEBUG][Disconnect] Error cancelando query: $_"
                    }
                    $script:CurrentQueryPowerShell = $null
                    $script:CurrentQueryRunspace = $null
                    $script:CurrentQueryAsync = $null
                    $script:QueryRunning = $false
                    Write-DzDebug "`t[DEBUG][Disconnect] Query cancelada"
                }
                if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo execUiTimer..."
                    $script:execUiTimer.Stop()
                }
                if ($script:QueryDoneTimer -and $script:QueryDoneTimer.IsEnabled) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo QueryDoneTimer..."
                    $script:QueryDoneTimer.Stop()
                }
                if ($script:execStopwatch) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo stopwatch..."
                    $script:execStopwatch.Stop()
                    $script:execStopwatch = $null
                }
                if ($global:connection) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Cerrando conexión SQL..."
                    try {
                        if ($global:connection.State -ne [System.Data.ConnectionState]::Closed) {
                            $global:connection.Close()
                        }
                        $global:connection.Dispose()
                    } catch {
                        Write-DzDebug "`t[DEBUG][Disconnect] Error cerrando conexión: $_"
                    }
                    $global:connection = $null
                }
                Write-DzDebug "`t[DEBUG][Disconnect] Limpiando variables globales..."
                $global:server = $null
                $global:user = $null
                $global:password = $null
                $global:database = $null
                $global:dbCredential = $null
                if ($global:tvDatabases) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Limpiando TreeView..."
                    $global:tvDatabases.Items.Clear()
                }
                if ($global:cmbDatabases) {
                    Write-DzDebug "`t[DEBUG][Disconnect] Limpiando ComboBox..."
                    $global:cmbDatabases.Items.Clear()
                    $global:cmbDatabases.IsEnabled = $false
                }
                if ($global:lblConnectionStatus) {
                    $global:lblConnectionStatus.Text = "⚫ Desconectado"
                }
                Write-DzDebug "`t[DEBUG][Disconnect] Habilitando controles de conexión..."
                if ($global:txtServer) { $global:txtServer.IsEnabled = $true }
                if ($global:txtUser) { $global:txtUser.IsEnabled = $true }
                if ($global:txtPassword) { $global:txtPassword.IsEnabled = $true }
                if ($global:btnConnectDb) { $global:btnConnectDb.IsEnabled = $true }
                Write-DzDebug "`t[DEBUG][Disconnect] Deshabilitando botones de operaciones..."
                if ($global:btnDisconnectDb) { $global:btnDisconnectDb.IsEnabled = $false }
                if ($global:btnExecute) { $global:btnExecute.IsEnabled = $false }
                if ($global:btnClearQuery) { $global:btnClearQuery.IsEnabled = $false }
                if ($global:btnExport) { $global:btnExport.IsEnabled = $false }
                if ($global:cmbQueries) { $global:cmbQueries.IsEnabled = $false }
                if ($global:tcQueries) { $global:tcQueries.IsEnabled = $false }
                if ($global:tcResults) { $global:tcResults.IsEnabled = $false }
                if ($global:rtbQueryEditor1) { $global:rtbQueryEditor1.IsEnabled = $false }
                if ($global:dgResults) { $global:dgResults.IsEnabled = $false }
                if ($global:txtMessages) { $global:txtMessages.IsEnabled = $false }
                if ($global:txtMessages) {
                    $global:txtMessages.Text = "Desconectado de la base de datos."
                }
                Write-DzDebug "`t[DEBUG][Disconnect] Desconexión completada exitosamente"
                Write-Host "✓ Desconectado exitosamente" -ForegroundColor Green
                if ($global:txtServer) {
                    $global:txtServer.Focus() | Out-Null
                }
            } catch {
                Write-DzDebug "`t[DEBUG][Disconnect] ERROR: $($_.Exception.Message)"
                Write-DzDebug "`t[DEBUG][Disconnect] Stack: $($_.ScriptStackTrace)"
                Write-Host "Error al desconectar: $($_.Exception.Message)" -ForegroundColor Red
                Ui-Error "Error al desconectar:`n`n$($_.Exception.Message)" $global:MainWindow
            }
        })
    $btnConnectDb.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Conectar Base de Datos' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n- - - Comenzando el proceso de 'Conectar Base de Datos' - - -" -ForegroundColor Magenta
            try {
                if ($null -eq $global:txtServer -or $null -eq $global:txtUser -or $null -eq $global:txtPassword) {
                    throw "Error interno: controles de conexión no inicializados."
                }
                $serverText = $global:txtServer.Text.Trim()
                $userText = $global:txtUser.Text.Trim()
                $passwordText = $global:txtPassword.Password
                Write-DzDebug "`t[DEBUG] | Server='$serverText' User='$userText' PasswordLen=$($passwordText.Length)"
                if ([string]::IsNullOrWhiteSpace($serverText) -or [string]::IsNullOrWhiteSpace($userText) -or [string]::IsNullOrWhiteSpace($passwordText)) {
                    throw "Complete todos los campos de conexión"
                }
                $securePassword = (New-Object System.Net.NetworkCredential('', $passwordText)).SecurePassword
                $credential = New-Object System.Management.Automation.PSCredential($userText, $securePassword)
                $global:server = $serverText
                $global:user = $userText
                $global:password = $passwordText
                $global:dbCredential = $credential
                Write-DzDebug "`t[DEBUG] Obteniendo lista de bases de datos..."
                $databases = Get-SqlDatabases -Server $serverText -Credential $credential
                if (-not $databases -or $databases.Count -eq 0) {
                    throw "Conexión correcta, pero no se encontraron bases de datos disponibles."
                }
                Write-DzDebug "`t[DEBUG] Se encontraron $($databases.Count) bases de datos"
                $global:cmbDatabases.Items.Clear()
                foreach ($db in $databases) {
                    [void]$global:cmbDatabases.Items.Add($db)
                }
                $global:cmbDatabases.IsEnabled = $true
                $global:cmbDatabases.SelectedIndex = 0
                $global:database = $global:cmbDatabases.SelectedItem
                $global:lblConnectionStatus.Text = "✓ Conectado a: $serverText | DB: $($global:database)"
                $global:txtServer.IsEnabled = $false
                $global:txtUser.IsEnabled = $false
                $global:txtPassword.IsEnabled = $false
                $global:btnConnectDb.IsEnabled = $false
                $global:btnDisconnectDb.IsEnabled = $true
                $global:btnExecute.IsEnabled = $true
                $global:btnClearQuery.IsEnabled = $true
                $global:cmbQueries.IsEnabled = $true
                $global:btnExport.IsEnabled = $true
                if ($global:tcQueries) { $global:tcQueries.IsEnabled = $true }
                if ($global:tcResults) { $global:tcResults.IsEnabled = $true }
                if ($global:rtbQueryEditor1) { $global:rtbQueryEditor1.IsEnabled = $true }
                if ($global:dgResults) { $global:dgResults.IsEnabled = $true }
                if ($global:txtMessages) { $global:txtMessages.IsEnabled = $true }
                $global:rtbQueryEditor1.Focus() | Out-Null
                Write-DzDebug "`t[DEBUG] Inicializando TreeView..."
                Initialize-SqlTreeView -TreeView $global:tvDatabases -Server $serverText -Credential $credential -User $userText -Password $passwordText -GetCurrentDatabase { $global:database } `
                    -AutoExpand $true `
                    -OnDatabaseSelected {
                    param($dbName)
                    if (-not $global:cmbDatabases) { return }
                    $global:cmbDatabases.SelectedItem = $dbName
                    if (-not $global:cmbDatabases.SelectedItem) {
                        for ($i = 0; $i -lt $global:cmbDatabases.Items.Count; $i++) {
                            if ([string]$global:cmbDatabases.Items[$i] -eq [string]$dbName) {
                                $global:cmbDatabases.SelectedIndex = $i
                                break
                            }
                        }
                    }
                    $global:database = $global:cmbDatabases.SelectedItem
                    if ($global:lblConnectionStatus) {
                        $global:lblConnectionStatus.Text = "✓ Conectado a: $($global:server) | DB: $($global:database)"
                    }
                    Write-DzDebug "`t[DEBUG][TreeView] BD seleccionada: $($global:database)"
                } `
                    -OnDatabasesRefreshed {
                    try {
                        if (-not $global:server -or -not $global:dbCredential) { return }
                        $databases = Get-SqlDatabases -Server $global:server -Credential $global:dbCredential
                        if ($databases -and $databases.Count -gt 0) {
                            $global:cmbDatabases.Items.Clear()
                            foreach ($db in $databases) {
                                [void]$global:cmbDatabases.Items.Add($db)
                            }
                            if ($global:database -and $global:cmbDatabases.Items.Contains($global:database)) {
                                $global:cmbDatabases.SelectedItem = $global:database
                            } elseif ($global:cmbDatabases.Items.Count -gt 0) {
                                $global:cmbDatabases.SelectedIndex = 0
                                $global:database = $global:cmbDatabases.SelectedItem
                            }
                            if ($global:lblConnectionStatus) {
                                $global:lblConnectionStatus.Text = "✓ Conectado a: $($global:server) | DB: $($global:database)"
                            }
                            Write-DzDebug "`t[DEBUG][TreeView] ComboBox actualizado con $($databases.Count) bases de datos"
                        }
                    } catch {
                        Write-DzDebug "`t[DEBUG][OnDatabasesRefreshed] Error: $_"
                    }
                } `
                    -InsertTextHandler {
                    param($text)
                    if ($global:tcQueries) {
                        Insert-TextIntoActiveQuery -TabControl $global:tcQueries -Text $text
                    }
                }
                Write-DzDebug "`t[DEBUG] Conexión establecida exitosamente"
                Write-Host "✓ Conectado exitosamente a: $serverText" -ForegroundColor Green
            } catch {
                Write-DzDebug "`t[DEBUG][btnConnectDb] CATCH: $($_.Exception.Message)"
                Write-DzDebug "`t[DEBUG][btnConnectDb] Tipo: $($_.Exception.GetType().FullName)"
                Write-DzDebug "`t[DEBUG][btnConnectDb] Stack: $($_.ScriptStackTrace)"
                Ui-Error "Error de conexión: $($_.Exception.Message)" $global:MainWindow
                Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
            }
        })

    $cmbDatabases.Add_SelectionChanged({
            param($sender, $e)
            if ($sender.SelectedItem) {
                $selectedItem = $sender.SelectedItem
                $dbName = if ($selectedItem -is [PSCustomObject] -and $selectedItem.DatabaseName) {
                    $selectedItem.DatabaseName
                } elseif ($selectedItem -is [string]) {
                    $selectedItem -replace ' \(.*?\)$', ''
                } else {
                    $selectedItem.ToString() -replace ' \(.*?\)$', ''
                }
                Write-DzDebug "`t[DEBUG] ComboBox DB seleccionada: '$dbName'"
                if ($selectedItem -is [PSCustomObject] -and $selectedItem.State -ne "ONLINE") {
                    Ui-Warn "La base de datos '$dbName' está en estado '$($selectedItem.State)'.`nNo se puede usar hasta que esté ONLINE." "Base de datos no disponible" $global:MainWindow
                    if ($global:database) {
                        for ($i = 0; $i -lt $sender.Items.Count; $i++) {
                            $item = $sender.Items[$i]
                            $itemDbName = if ($item -is [PSCustomObject]) { $item.DatabaseName } else { $item }
                            if ($itemDbName -eq $global:database) {
                                $sender.SelectedIndex = $i
                                break
                            }
                        }
                    }
                    return
                }
                $global:database = $dbName
                if ($global:lblConnectionStatus) {
                    $global:lblConnectionStatus.Text = "✓ Conectado a: $($global:server) | DB: $($global:database)"
                }
                if ($global:tvDatabases) {
                    Select-SqlTreeDatabase -TreeView $global:tvDatabases -Server $global:server -Database $global:database
                }
            }
        })

    $btnExecute.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Ejecutar Consulta' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            Write-Host "`n`t- - - Ejecutando consulta - - -" -ForegroundColor Gray
            if ($script:QueryRunning) {
                Write-DzDebug "`t[DEBUG] Query ya en ejecución, ignorando click"
                return
            }
            $script:QueryRunning = $true
            try {
                $selectedDb = $global:cmbDatabases.SelectedItem
                if (-not $selectedDb) { throw "Selecciona una base de datos" }
                $activeRtb = Get-ActiveQueryRichTextBox -TabControl $global:tcQueries
                if (-not $activeRtb) { throw "No hay una pestaña de consulta activa." }
                $textRange = New-Object System.Windows.Documents.TextRange(
                    $activeRtb.Document.ContentStart,
                    $activeRtb.Document.ContentEnd
                )
                $rawQuery = $textRange.Text
                if ([string]::IsNullOrWhiteSpace($rawQuery)) { throw "La consulta está vacía." }
                $query = Remove-SqlComments -Query $rawQuery
                if ([string]::IsNullOrWhiteSpace($query)) { throw "La consulta está vacía después de limpiar comentarios." }
                if ($global:tcResults) { $global:tcResults.Items.Clear() }
                if ($global:dgResults) { $global:dgResults.ItemsSource = $null }
                if ($global:txtMessages) { $global:txtMessages.Text = "" }
                if ($global:lblRowCount) { $global:lblRowCount.Text = "Filas: --" }
                if ($global:lblExecutionTimer) {
                    $global:lblExecutionTimer.Text = "Tiempo: 00:00.0"
                }
                if (-not $script:execStopwatch) {
                    $script:execStopwatch = [System.Diagnostics.Stopwatch]::new()
                }
                if ($script:execUiTimer) {
                    try {
                        if ($script:execUiTimer.IsEnabled) {
                            $script:execUiTimer.Stop()
                        }
                    } catch {
                        Write-DzDebug "`t[DEBUG] Error deteniendo timer previo: $_"
                    }
                }
                $script:execUiTimer = [System.Windows.Threading.DispatcherTimer]::new()
                $script:execUiTimer.Interval = [TimeSpan]::FromMilliseconds(100)
                $script:execUiTimer.Add_Tick({
                        try {
                            if ($global:lblExecutionTimer -and $script:execStopwatch) {
                                $t = $script:execStopwatch.Elapsed
                                $global:lblExecutionTimer.Text = ("Tiempo: {0:mm\:ss\.f}" -f $t)
                            }
                        } catch {
                            Write-DzDebug "`t[DEBUG][Timer] Error actualizando: $_"
                        }
                    })
                $script:execStopwatch.Restart()
                $script:execUiTimer.Start()
                $btnExecute.IsEnabled = $false
                $server = [string]$global:server
                $db = [string]$selectedDb
                $userText = [string]$global:user
                $passwordTxt = [string]$global:password
                $modulesPath = Join-Path $PSScriptRoot "modules"
                $query = Remove-SqlComments -Query $rawQuery
                if ([string]::IsNullOrWhiteSpace($query)) { throw "La consulta está vacía después de limpiar comentarios." }
                # Mostrar el query en consola
                Write-Host "Query:" -ForegroundColor Cyan
                foreach ($line in ($query -split "`r?`n")) {
                    Write-Host "`t$line" -ForegroundColor Green
                }
                Write-Host "" # Línea en blanco
                if ($global:tcResults) { $global:tcResults.Items.Clear() }
                Write-Host "Ejecutando consulta en '$db'..." -ForegroundColor Cyan
                $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
                $rs.ApartmentState = 'MTA'
                $rs.ThreadOptions = 'ReuseThread'
                $rs.Open()
                $ps = [PowerShell]::Create()
                $ps.Runspace = $rs
                $script:CurrentQueryRunspace = $rs
                $script:CurrentQueryPowerShell = $ps
                $worker = {
                    param($Server, $Database, $Query, $User, $Password, $ModulesPath)
                    try {
                        $utilPath = Join-Path $ModulesPath "Utilities.psm1"
                        $dbPath = Join-Path $ModulesPath "Database.psm1"
                        Write-Host "[WORKER] Importando Utilities.psm1..." -ForegroundColor Magenta
                        try {
                            Import-Module $utilPath -Force -DisableNameChecking -ErrorAction Stop
                            Write-Host "[WORKER] ✓ Utilities.psm1 importado OK" -ForegroundColor Green
                        } catch {
                            Write-Host "[WORKER] ✗ ERROR importando Utilities: $($_.Exception.Message)" -ForegroundColor Red
                            throw "Error en Utilities.psm1: $($_.Exception.Message)"
                        }
                        Write-Host "[WORKER] Importando Database.psm1..." -ForegroundColor Magenta
                        try {
                            Import-Module $dbPath -Force -DisableNameChecking -ErrorAction Stop
                            Write-Host "[WORKER] ✓ Database.psm1 importado OK" -ForegroundColor Green
                        } catch {
                            Write-Host "[WORKER] ✗ ERROR importando Database: $($_.Exception.Message)" -ForegroundColor Red
                            Write-Host "[WORKER] Stack: $($_.ScriptStackTrace)" -ForegroundColor Red
                            throw "Error en Database.psm1: $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                        }
                        Write-Host "[WORKER] Query tipo: $($Query.GetType().FullName)" -ForegroundColor Cyan
                        Write-Host "[WORKER] Query longitud: $($Query.Length)" -ForegroundColor Cyan
                        $secure = New-Object System.Security.SecureString
                        foreach ($ch in $Password.ToCharArray()) { $secure.AppendChar($ch) }
                        $secure.MakeReadOnly()
                        $cred = New-Object System.Management.Automation.PSCredential($User, $secure)
                        $r = Invoke-SqlQueryMultiResultSet -Server $Server -Database $Database -Query $Query -Credential $cred
                        if ($null -eq $r) {
                            return @{
                                Success      = $false
                                ErrorMessage = "La ejecución devolvió NULL."
                                ResultSets   = @()
                                Messages     = @()
                            }
                        }
                        return $r
                    } catch {
                        return @{
                            Success      = $false
                            ErrorMessage = $_.Exception.Message
                            ResultSets   = @()
                            Messages     = @()
                            Type         = "Error"
                        }
                    }
                }
                [void]$ps.AddScript($worker).
                AddArgument($server).
                AddArgument($db).
                AddArgument($query).
                AddArgument($userText).
                AddArgument($passwordTxt).
                AddArgument($modulesPath)
                $async = $ps.BeginInvoke()
                $script:CurrentQueryAsync = $async
                if ($script:QueryDoneTimer) {
                    try {
                        if ($script:QueryDoneTimer.IsEnabled) {
                            $script:QueryDoneTimer.Stop()
                        }
                    } catch {}
                }
                $script:QueryDoneTimer = [System.Windows.Threading.DispatcherTimer]::new()
                $script:QueryDoneTimer.Interval = [TimeSpan]::FromMilliseconds(150)
                $script:QueryDoneTimer.Add_Tick({
                        try {
                            Write-DzDebug "`t[DEBUG][TICK] Verificando query..."
                            if (-not $script:CurrentQueryAsync) {
                                Write-DzDebug "`t[DEBUG][TICK] No hay async"
                                return
                            }
                            if (-not $script:CurrentQueryAsync.IsCompleted) {
                                Write-DzDebug "`t[DEBUG] Aún en ejecución..."
                                return
                            }
                            $script:QueryDoneTimer.Stop()
                            Write-DzDebug "`t[DEBUG][TICK] Query completada, procesando..."
                            $result = $null
                            try {
                                Write-DzDebug "`t[DEBUG][TICK] Llamando EndInvoke..."
                                $result = $script:CurrentQueryPowerShell.EndInvoke($script:CurrentQueryAsync)
                                Write-DzDebug "`t[DEBUG][TICK] EndInvoke OK. Tipo: $($result.GetType().FullName)"
                            } catch {
                                $result = @{
                                    Success      = $false
                                    ErrorMessage = $_.Exception.Message
                                    ResultSets   = @()
                                    Messages     = @()
                                    Type         = "Error"
                                }
                            }
                            Write-DzDebug "`t[DEBUG][TICK] Normalizando resultado..."
                            Write-DzDebug "`t[DEBUG][TICK] result tipo: $($result.GetType().FullName)"
                            if ($result -is [System.Management.Automation.PSDataCollection[psobject]]) {
                                Write-DzDebug "`t[DEBUG][TICK] Es PSDataCollection, Count=$($result.Count)"
                                if ($result.Count -gt 0) {
                                    $result = $result[0]
                                    Write-DzDebug "`t[DEBUG][TICK] Tomando primer elemento del PSDataCollection"
                                } else {
                                    $result = $null
                                    Write-DzDebug "`t[DEBUG][TICK] PSDataCollection vacío"
                                }
                            } elseif ($result -is [System.Array]) {
                                Write-DzDebug "`t[DEBUG][TICK] Es array normal, Count=$($result.Count)"
                                if ($result.Count -gt 0) {
                                    $result = $result[0]
                                } else {
                                    $result = $null
                                }
                            }
                            Write-DzDebug "`t[DEBUG][TICK] Después de normalizar, tipo: $($result.GetType().FullName)"
                            Write-DzDebug "`t[DEBUG][TICK] ===== ESTRUCTURA DEL RESULTADO ====="
                            if ($result) {
                                if ($result -is [hashtable]) {
                                    Write-DzDebug "`t[DEBUG][TICK] Es hashtable, Keys:"
                                    foreach ($key in $result.Keys) {
                                        $value = $result[$key]
                                        $valueType = if ($value) { $value.GetType().Name } else { "null" }
                                        Write-DzDebug "`t[DEBUG][TICK]   $key = $valueType"
                                    }
                                } else {
                                    Write-DzDebug "`t[DEBUG][TICK] Propiedades:"
                                    $result.PSObject.Properties | ForEach-Object {
                                        Write-DzDebug "`t[DEBUG][TICK]   $($_.Name) = $($_.Value)"
                                    }
                                }
                            } else {
                                Write-DzDebug "`t[DEBUG][TICK] result es NULL"
                            }
                            try {
                                if ($script:CurrentQueryPowerShell) {
                                    $script:CurrentQueryPowerShell.Dispose()
                                }
                            } catch {
                                Write-DzDebug "`t[DEBUG] Error disposing PowerShell: $_"
                            }
                            try {
                                if ($script:CurrentQueryRunspace) {
                                    $script:CurrentQueryRunspace.Close()
                                    $script:CurrentQueryRunspace.Dispose()
                                }
                            } catch {
                                Write-DzDebug "`t[DEBUG] Error disposing Runspace: $_"
                            }
                            $script:CurrentQueryPowerShell = $null
                            $script:CurrentQueryRunspace = $null
                            $script:CurrentQueryAsync = $null
                            try {
                                if ($script:execStopwatch) {
                                    $script:execStopwatch.Stop()
                                }
                                if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) {
                                    $script:execUiTimer.Stop()
                                }
                                if ($global:lblExecutionTimer) {
                                    $t = $script:execStopwatch.Elapsed
                                    $global:lblExecutionTimer.Text = ("Tiempo: {0:mm\:ss\.fff}" -f $t)
                                }
                            } catch {
                                Write-DzDebug "`t[DEBUG] Error deteniendo timer: $_"
                            }
                            try {
                                $btnExecute.IsEnabled = $true
                            } catch {}
                            $script:QueryRunning = $false
                            if (-not $result -or -not $result.Success) {
                                $msg = ""
                                try { $msg = [string]$result.ErrorMessage } catch {}
                                if ([string]::IsNullOrWhiteSpace($msg) -and $result -and $result.Messages) {
                                    try { $msg = ($result.Messages -join "`n") } catch {}
                                }
                                if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Error desconocido al ejecutar la consulta." }
                                if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
                                    Write-DzDebug "`t[DEBUG][TICK] Mostrando ResultSets a pesar del error (count: $($result.ResultSets.Count))"
                                    try {
                                        Show-MultipleResultSets -TabControl $global:tcResults -ResultSets $result.ResultSets
                                        Show-ErrorResultTab -ResultsTabControl $global:tcResults -Message $msg -AddWithoutClear
                                        Write-Host "`n=============== ERROR SQL ==============" -ForegroundColor Red
                                        Write-Host "Mensaje: $msg" -ForegroundColor Yellow
                                        Write-Host "====================================" -ForegroundColor Red
                                        if ($global:txtMessages) {
                                            $currentText = $global:txtMessages.Text
                                            $global:txtMessages.Text = "ERROR: $msg`n`n$currentText"
                                        }
                                        if ($global:lblRowCount) {
                                            $totalRows = ($result.ResultSets | Measure-Object -Property RowCount -Sum).Sum
                                            if ($result.ResultSets.Count -eq 1) {
                                                $global:lblRowCount.Text = "Filas: $totalRows (con error)"
                                            } else {
                                                $global:lblRowCount.Text = "Filas: $totalRows ($($result.ResultSets.Count) resultsets, con error)"
                                            }
                                        }
                                    } catch {
                                        Write-DzDebug "`t[DEBUG][TICK] ERROR en Show-MultipleResultSets: $($_.Exception.Message)"
                                        if ($global:txtMessages) {
                                            $global:txtMessages.Text = "Error mostrando resultados: $($_.Exception.Message)"
                                        }
                                    }
                                } else {
                                    Write-Host "`n=============== ERROR SQL ==============" -ForegroundColor Red
                                    Write-Host "Mensaje: $msg" -ForegroundColor Yellow
                                    Write-Host "====================================" -ForegroundColor Red
                                    if ($global:txtMessages) {
                                        $global:txtMessages.Text = $msg
                                    }
                                    if ($global:lblRowCount) {
                                        $global:lblRowCount.Text = "Filas: --"
                                    }
                                    if ($global:tcResults) {
                                        try {
                                            Show-ErrorResultTab -ResultsTabControl $global:tcResults -Message $msg
                                        } catch {}
                                    }
                                }
                                return
                            }
                            if ($result.DebugLog -and $global:txtMessages) {
                                try {
                                    $dbg = ($result.DebugLog -join "`n")
                                    if (-not [string]::IsNullOrWhiteSpace($dbg)) {
                                        $global:txtMessages.Text = $dbg + "`n`n" + $global:txtMessages.Text
                                    }
                                } catch {}
                            }
                            if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
                                Write-DzDebug "`t[DEBUG][TICK] Mostrando $($result.ResultSets.Count) resultsets..."
                                try {
                                    $idx = 0
                                    foreach ($rs in $result.ResultSets) {
                                        $idx++
                                        $dt = $rs.DataTable
                                        $rows = $rs.RowCount
                                        Write-Host ("`t Resultado #{0} | Filas: {1}" -f $idx, $rows) -ForegroundColor Green
                                    }
                                    Show-MultipleResultSets -TabControl $global:tcResults -ResultSets $result.ResultSets
                                    Write-DzDebug "`t[DEBUG][TICK] ResultSets mostrados OK"
                                } catch {
                                    Write-DzDebug "`t[DEBUG][TICK] ERROR en Show-MultipleResultSets: $($_.Exception.Message)"
                                    if ($global:txtMessages) {
                                        $global:txtMessages.Text = "Error mostrando resultados: $($_.Exception.Message)"
                                    }
                                }
                                return
                            }
                            if ($result -and $result.ContainsKey('RowsAffected') -and $result.RowsAffected -ne $null) {
                                Write-DzDebug "`t[DEBUG][TICK] Mostrando RowsAffected: $($result.RowsAffected)"
                                if ($global:tcResults) {
                                    $global:tcResults.Items.Clear()
                                    $tab = New-Object System.Windows.Controls.TabItem
                                    $tab.Header = "Resultado"
                                    $text = New-Object System.Windows.Controls.TextBlock
                                    $text.Text = "Filas afectadas: $($result.RowsAffected)"
                                    $text.Margin = "10"
                                    $text.FontSize = 14
                                    $text.FontWeight = "Bold"
                                    $tab.Content = $text
                                    [void]$global:tcResults.Items.Add($tab)
                                    $global:tcResults.SelectedItem = $tab
                                }
                                if ($global:txtMessages) {
                                    $global:txtMessages.Text = "Filas afectadas: $($result.RowsAffected)"
                                }
                                if ($global:lblRowCount) {
                                    $global:lblRowCount.Text = "Filas afectadas: $($result.RowsAffected)"
                                }
                                Write-Host "`n=============== RESULTADO ==============" -ForegroundColor Green
                                Write-Host "Filas afectadas: $($result.RowsAffected)" -ForegroundColor Yellow
                                Write-Host "====================================" -ForegroundColor Green
                                return
                            }
                            Show-MultipleResultSets -TabControl $global:tcResults -ResultSets @()
                            if ($global:lblRowCount) {
                                $global:lblRowCount.Text = "Filas: 0"
                            }
                        } catch {
                            $err = "[UI][QueryDoneTimer ERROR] $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                            if ($global:txtMessages) {
                                $global:txtMessages.Text = $err
                            }
                            Write-Host $err -ForegroundColor Red
                            try {
                                $btnExecute.IsEnabled = $true
                            } catch {}
                            $script:QueryRunning = $false
                            try {
                                if ($script:execStopwatch) {
                                    $script:execStopwatch.Stop()
                                }
                                if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) {
                                    $script:execUiTimer.Stop()
                                }
                            } catch {}
                        }
                    })
                $script:QueryDoneTimer.Start()
            } catch {
                $msg = $_.Exception.Message
                if ($global:txtMessages) {
                    $global:txtMessages.Text = $msg
                }
                Write-Host "`n[ERROR btnExecute] $msg" -ForegroundColor Red
                try {
                    if ($script:execStopwatch) {
                        $script:execStopwatch.Stop()
                    }
                    if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) {
                        $script:execUiTimer.Stop()
                    }
                } catch {}
                $btnExecute.IsEnabled = $true
                $script:QueryRunning = $false
            }
        })
    $btnClearQuery.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Limpiar Query' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            try {
                $rtb = Get-ActiveQueryRichTextBox -TabControl $global:tcQueries
                if (-not $rtb) {
                    throw "No hay una pestaña de consulta activa."
                }
                $rtb.Document.Blocks.Clear()
                if ($global:dgResults) { $global:dgResults.ItemsSource = $null }
                if ($global:txtMessages) { $global:txtMessages.Text = "" }
                if ($global:lblRowCount) { $global:lblRowCount.Text = "Filas: --" }
                $tab = Get-ActiveQueryTab -TabControl $global:tcQueries
                $tabName = if ($tab -and $tab.Tag -and $tab.Tag.Title) { $tab.Tag.Title } else { "Desconocida" }
                Write-DzDebug "`t[DEBUG] Se limpió la consulta en la pestaña: $tabName"
                Write-Host "Consulta limpiada" -ForegroundColor Cyan
            } catch {
                $msg = $_.Exception.Message
                if ($global:txtMessages) { $global:txtMessages.Text = $msg }
                Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
                Write-Host "Mensaje: $msg" -ForegroundColor Yellow
                Write-Host "====================================" -ForegroundColor Red
            }
        })
    $btnExport.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Exportar resultados' - {0}" -f (Get-Date -Format "HH:mm:ss")) -Color DarkYellow
            try {
                if (-not $global:tcResults) {
                    Ui-Warn "No existe un panel de resultados para exportar." "Atención" $global:MainWindow
                    return
                }
                $resultTabs = Get-ExportableResultTabs -TabControl $global:tcResults
                if (-not $resultTabs -or $resultTabs.Count -eq 0) {
                    Ui-Warn "No existe pestaña con resultados para exportar." "Atención" $global:MainWindow
                    return
                }
                $target = $null
                if ($resultTabs.Count -gt 1) {
                    $items = $resultTabs | ForEach-Object {
                        [pscustomobject]@{
                            Path         = $_
                            Display      = $_.Display
                            DisplayShort = $_.DisplayShort
                        }
                    }
                    $selected = Show-WpfPathSelectionDialog -Title "Exportar resultados" -Prompt "Seleccione la pestaña de resultados a exportar:" -Items $items -ExecuteButtonText "Exportar"
                    if (-not $selected) { return }
                    $target = $selected.Path
                } else {
                    $target = $resultTabs[0]
                }
                $rowCount = $target.RowCount
                $headerText = $target.DisplayShort
                if (-not (Ui-Confirm "Se exportarán $rowCount filas de '$headerText'. ¿Deseas continuar?" "Confirmar exportación" $global:MainWindow)) { return }
                $safeName = ($headerText -replace '[\\/:*?"<>|]', '-')
                if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "resultado" }
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Filter = "CSV (*.csv)|*.csv|Texto delimitado (*.txt)|*.txt"
                $saveDialog.FileName = "$safeName.csv"
                $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
                if ($saveDialog.ShowDialog() -ne $true) { return }
                $filePath = $saveDialog.FileName
                $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
                if ($extension -eq ".txt" -or $saveDialog.FilterIndex -eq 2) {
                    $separator = New-WpfInputDialog -Title "Separador de exportación" -Prompt "Ingrese el separador para el archivo de texto:" -DefaultValue "|"
                    if ($null -eq $separator) { return }
                    Export-ResultSetToDelimitedText -ResultSet $target.DataTable -Path $filePath -Separator $separator
                } else {
                    Export-ResultSetToCsv -ResultSet ([pscustomobject]@{ DataTable = $target.DataTable }) -Path $filePath
                }
                Ui-Info "Exportación completada en:`n$filePath" "Exportación" $global:MainWindow
            } catch {
                Ui-Error "Error al exportar resultados:`n$($_.Exception.Message)" "Error" $global:MainWindow
                Write-DzDebug "`t[DEBUG][btnExport] CATCH: $($_.Exception.Message)" -Color Red
            }
        })
    $window.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'F5' -and $btnExecute.IsEnabled) {
                $btnExecute.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                $e.Handled = $true
            }
        })
    $closeWindowScript = {
        Write-Host "Cerrando aplicación..." -ForegroundColor Yellow
        Write-DzDebug "`t[DEBUG] Botón Salir presionado" -Color DarkGray
        try {
            $btn = $args[0]
            $win = [System.Windows.Window]::GetWindow($btn)
            if ($null -ne $win) {
                $win.Close()
                Write-DzDebug "`t[DEBUG] Ventana cerrada (método 1)" -Color DarkGray
            } else {
                $window.Close()
                Write-DzDebug "`t[DEBUG] Ventana cerrada (método 2)" -Color DarkGray
            }
        } catch {
            Write-Host "Error al cerrar: $_" -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG] Error: $_" -Color Red
        }
    }.GetNewClosure()
    Write-Host "✓ Formulario WPF creado exitosamente" -ForegroundColor Green
    $window.Dispatcher.Add_UnhandledException({
            param($sender, $e)
            $ex = $e.Exception
            Write-Host "`n[WPF Dispatcher ERROR]" -ForegroundColor Red
            Write-Host "Mensaje: $($ex.Message)" -ForegroundColor Yellow
            Write-Host "Tipo   : $($ex.GetType().FullName)" -ForegroundColor Yellow
            Write-Host "Stack  : $($ex.StackTrace)" -ForegroundColor DarkYellow
            if ($ex -is [System.Management.Automation.RuntimeException] -and $ex.ErrorRecord) {
                $er = $ex.ErrorRecord
                if ($er.InvocationInfo) {
                    Write-Host "Archivo : $($er.InvocationInfo.ScriptName)" -ForegroundColor Cyan
                    Write-Host "Línea   : $($er.InvocationInfo.ScriptLineNumber)" -ForegroundColor Cyan
                    Write-Host "Código  : $($er.InvocationInfo.Line)" -ForegroundColor Cyan
                    Write-Host "Pos     : $($er.InvocationInfo.PositionMessage)" -ForegroundColor DarkCyan
                }
                Write-Host "PSScriptStackTrace:" -ForegroundColor Magenta
                Write-Host $er.ScriptStackTrace -ForegroundColor Magenta
            }
            $e.Handled = $true
        })
    return $window
}
function Start-Application {
    Show-GlobalProgress -Percent 0 -Status "Inicializando..."
    if (-not (Initialize-Environment)) { Show-GlobalProgress -Percent 100 -Status "Error inicializando" ; return }
    Show-GlobalProgress -Percent 10 -Status "Entorno listo"
    Show-GlobalProgress -Percent 20 -Status "Cargando módulos..."
    $modulesPath = Join-Path $PSScriptRoot "modules"
    $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "SqlTreeView.psm1", "Installers.psm1", "WindowsUtilities.psm1", "NationalUtilities.psm1", "SqlOps.psm1")
    $i = 0
    foreach ($module in $modules) {
        $i++
        Show-GlobalProgress -Percent (20 + [math]::Round(($i / $modules.Count) * 20)) -Status "Importando $module"
        $modulePath = Join-Path $modulesPath $module
        if (Test-Path $modulePath) { Import-Module $modulePath -Force -DisableNameChecking -ErrorAction Stop }
    }
    Show-GlobalProgress -Percent 45 -Status "Módulos listos"
    Show-GlobalProgress -Percent 80 -Status "Creando formulario..."
    Show-GlobalProgress -Percent 95 -Status "Mostrando GUI..."
    $mainForm = New-MainForm
    Show-GlobalProgress -Percent 100 -Status "¡Listo!`n"
    $null = $mainForm.ShowDialog()
}
try {
    Start-Application
} catch {
    Write-Host "Error fatal: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.InvocationInfo) {
        Write-Host "Archivo : $($_.InvocationInfo.ScriptName)" -ForegroundColor Yellow
        Write-Host "Línea   : $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
        Write-Host "Col     : $($_.InvocationInfo.OffsetInLine)" -ForegroundColor Yellow
        Write-Host "Código  : $($_.InvocationInfo.Line)" -ForegroundColor Yellow
        Write-Host "Pos     : $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkYellow
    }
    Write-Host "ScriptStackTrace:" -ForegroundColor Magenta
    Write-Host $_.ScriptStackTrace -ForegroundColor Magenta
    Write-Host "Stack Trace .NET:" -ForegroundColor DarkRed
    Write-Host $_.Exception.StackTrace -ForegroundColor DarkRed
    pause
    exit 1
}