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
$modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "SqlTreeView.psm1", "MultiQuery.psm1", "Installers.psm1")
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
- Nueva interfaz WPF
    * Fuentes y colores actualizados.
- Switch para modo oscuro
- Switch para modo debug
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
        Write-DzDebug "`t[DEBUG]Configuración de debug cargada (debug=$debugEnabled)" -Color DarkGray
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
        Height="650" Width="1200" MinHeight="600" MinWidth="1000"
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
        <Style x:Key="SystemButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>
        <Style x:Key="NationalSoftButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
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
            <TabItem Header="Aplicaciones" Name="tabAplicaciones">
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
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,50,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Ejecutar ExpressProfiler" Name="btnProfiler" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,90,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Ejecutar Database4" Name="btnDatabase" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,130,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Ejecutar Manager" Name="btnSQLManager" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,170,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Ejecutar Management" Name="btnSQLManagement" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,210,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Printer Tools" Name="btnPrinterTool" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,250,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Lector DP - Permisos" Name="btnLectorDPicacls" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,290,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Buscar Instalador LZMA" Name="LZMAbtnBuscarCarpeta" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,330,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Agregar IPs" Name="btnConfigurarIPs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,370,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Agregar usuario de Windows" Name="btnAddUser" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,410,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Configuraciones de Firewall" Name="btnFirewallConfig" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,450,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Actualizar datos del sistema" Name="btnForzarActualizacion" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,490,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Clear AnyDesk" Name="btnClearAnyDesk" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,50,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Mostrar Impresoras" Name="btnShowPrinters" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,90,0,0" Style="{StaticResource GeneralButtonStyle}"/>
                    <Button Content="Limpia y Reinicia Cola de Impresión" Name="btnClearPrintJobs" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="250,130,0,0" Style="{StaticResource GeneralButtonStyle}"/>
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
            <TabItem Header="Base de datos" Name="tabProSql">
    <Grid Background="{DynamicResource PanelBg}">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- 📊 Barra de Conexión y Comandos -->
        <Border Grid.Row="0" Margin="10" Padding="10" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" CornerRadius="6" Background="{DynamicResource ControlBg}">
            <WrapPanel>
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
                <StackPanel Margin="0,0,16,0" VerticalAlignment="Bottom">
                    <Button Content="Conectar" Name="btnConnectDb" Width="120" Height="30" Margin="0,0,0,6" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Desconectar" Name="btnDisconnectDb" Width="120" Height="30" Margin="0,0,0,6" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
                    <Button Content="Backup" Name="btnBackup" Width="120" Height="30" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
                </StackPanel>
            </WrapPanel>
        </Border>

        <!-- 🪟 Área de Trabajo Principal (TreeView + Consultas) -->
        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="250" MinWidth="200"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- 🌲 Panel izquierdo: TreeView del Explorador de Objetos -->
            <Border Grid.Column="0" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" CornerRadius="6" Background="{DynamicResource ControlBg}">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="Explorador de Objetos" Padding="8" FontWeight="Bold" Background="{DynamicResource AccentPrimary}" Foreground="{DynamicResource OnAccentFg}"/>
                    <TreeView Grid.Row="1" Name="tvDatabases" Padding="4"/>
                </Grid>
            </Border>

            <!-- 🖱️ GridSplitter para redimensionar -->
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="{DynamicResource BorderBrushColor}"/>

            <!-- 📝 Panel derecho: Pestañas de Consultas y Resultados -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*" MinHeight="150"/>
                    <RowDefinition Height="5"/>
                    <RowDefinition Height="2*" MinHeight="200"/>
                </Grid.RowDefinitions>

                <!-- Panel de Consultas -->
                <TabControl Name="tcQueries" Grid.Row="0" Background="{DynamicResource ControlBg}">
                    <TabItem Header="Consulta 1">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <!-- Barra de herramientas de consulta -->
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
                                <Button Content="Ejecutar (F5)" Name="btnExecute" Width="100" Height="28" Margin="2" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
                                <ComboBox Name="cmbQueries" Width="200" Margin="2" IsEnabled="False" ToolTip="Consultas predefinidas"/>
                                <Button Content="Limpiar" Name="btnClearQuery" Width="80" Height="28" Margin="2" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
                                <Button Content="Formato" Name="btnFormat" Width="80" Height="28" Margin="2" Style="{StaticResource GeneralButtonStyle}" IsEnabled="False"/>
                                <Button Content="Comentar" Name="btnComment" Width="80" Height="28" Margin="2" Style="{StaticResource GeneralButtonStyle}" IsEnabled="False"/>
                            </StackPanel>
                            <!-- Editor de consultas (RichTextBox) -->
                            <Border Grid.Row="1" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" Margin="5" CornerRadius="4">
                                <RichTextBox Name="rtbQueryEditor1" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" AcceptsTab="True" FontFamily="Consolas" FontSize="12"/>
                            </Border>
                        </Grid>
                    </TabItem>
                    <!-- Pestaña para agregar nuevas consultas -->
                    <TabItem Header="+" Name="tabAddQuery" IsEnabled="True"/>
                </TabControl>

                <!-- 🖱️ GridSplitter entre consultas y resultados -->
                <GridSplitter Grid.Row="1" Height="5" HorizontalAlignment="Stretch" Background="{DynamicResource BorderBrushColor}"/>

                <!-- 📊 Panel de Resultados y Mensajes -->
                <TabControl Name="tcResults" Grid.Row="2" Background="{DynamicResource ControlBg}">
                    <TabItem Header="Resultados">
                        <DataGrid Name="dgResults" IsReadOnly="True" AutoGenerateColumns="True" CanUserAddRows="False" CanUserDeleteRows="False"/>
                    </TabItem>
                    <TabItem Header="Mensajes">
                        <TextBox Name="txtMessages" IsReadOnly="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" Background="Transparent" BorderThickness="0"/>
                    </TabItem>
                </TabControl>
            </Grid>
        </Grid>

        <!-- 📍 Barra de Estado -->
        <StatusBar Grid.Row="2" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}">
            <StatusBarItem>
                <TextBlock Name="lblConnectionStatus" Text="Desconectado"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem>
                <TextBlock Name="lblExecutionTime" Text="Tiempo: --"/>
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
    $btnClearPrintJobs = $window.FindName("btnClearPrintJobs")
    $btnAplicacionesNS = $window.FindName("btnAplicacionesNS")
    $btnCambiarOTM = $window.FindName("btnCambiarOTM")
    $btnCheckPermissions = $window.FindName("btnCheckPermissions")
    $btnCreateAPK = $window.FindName("btnCreateAPK")
    $btnExtractInstaller = $window.FindName("btnExtractInstaller")
    $txtServer = $window.FindName("txtServer")
    $txtUser = $window.FindName("txtUser")
    $txtPassword = $window.FindName("txtPassword")
    $cmbDatabases = $window.FindName("cmbDatabases")
    $btnConnectDb = $window.FindName("btnConnectDb")
    $btnDisconnectDb = $window.FindName("btnDisconnectDb")
    $btnBackup = $window.FindName("btnBackup")
    $lblConnectionStatus = $window.FindName("lblConnectionStatus")
    $btnExecute = $window.FindName("btnExecute")
    $btnClearQuery = $window.FindName("btnClearQuery")
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
                if ($selectedQuery -and $selectedQuery -ne "Selecciona una consulta predefinida") {
                    $queryText = $script:predefinedQueries[$selectedQuery]
                    # Limpiar y establecer texto en RichTextBox
                    $global:rtbQueryEditor1.Document.Blocks.Clear()
                    $paragraph = New-Object System.Windows.Documents.Paragraph
                    $run = New-Object System.Windows.Documents.Run($queryText)
                    $paragraph.Inlines.Add($run)
                    $global:rtbQueryEditor1.Document.Blocks.Add($paragraph)

                    # Aplicar resaltado de sintaxis
                    if ($global:rtbQueryEditor1) {
                        Set-WpfSqlHighlighting -RichTextBox $global:rtbQueryEditor1 -Keywords $script:sqlKeywords
                    }
                }
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
    $global:txtServer = $txtServer
    $global:txtUser = $txtUser
    $global:txtPassword = $txtPassword
    $global:cmbDatabases = $cmbDatabases
    $global:btnConnectDb = $btnConnectDb
    $global:btnDisconnectDb = $btnDisconnectDb
    $global:btnExecute = $btnExecute
    $global:btnBackup = $btnBackup
    $global:btnClearQuery = $btnClearQuery
    $global:cmbQueries = $cmbQueries
    $global:tcQueries = $tcQueries
    $global:tcResults = $tcResults
    $global:tvDatabases = $tvDatabases
    $global:tabAddQuery = $tabAddQuery
    $global:btnFormat = $window.FindName("btnFormat")
    $global:btnComment = $window.FindName("btnComment")
    $global:rtbQueryEditor1 = $window.FindName("rtbQueryEditor1")
    $global:dgResults = $window.FindName("dgResults")
    $global:txtMessages = $window.FindName("txtMessages")
    $global:lblExecutionTime = $window.FindName("lblExecutionTime")
    $global:lblRowCount = $window.FindName("lblRowCount")
    $global:lblConnectionStatus = $lblConnectionStatus
    $global:txt_AdapterStatus = $txt_AdapterStatus
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
    $ipsWithAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object { $_.OperationalStatus -eq 'Up' } | ForEach-Object {
        $interface = $_
        $interface.GetIPProperties().UnicastAddresses | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '127.0.0.1' } | ForEach-Object {
            @{AdapterName = $interface.Name; IPAddress = $_.Address.ToString() }
        }
    }
    if ($ipsWithAdapters.Count -gt 0) {
        $ipsTextForLabel = $ipsWithAdapters | ForEach-Object { "- $($_.AdapterName) - IP: $($_.IPAddress)" } | Out-String
        $txt_IpAdress.Text = $ipsTextForLabel
        Write-DzDebug "`t[DEBUG] txt_IpAdress poblado con: '$($txt_IpAdress.Text)'"
        Write-DzDebug "`t[DEBUG] txt_IpAdress.Text.Length: $($txt_IpAdress.Text.Length)"
    } else {
        $txt_IpAdress.Text = "No se encontraron direcciones IP"
        Write-DzDebug "`t[DEBUG] No se encontraron IPs" -Color Yellow
    }
    Refresh-AdapterStatus
    Load-IniConnectionsToComboBox -Combo $txtServer
    $buttonsToUpdate = @($LZMAbtnBuscarCarpeta, $btnInstalarHerramientas, $btnProfiler, $btnDatabase, $btnSQLManager, $btnSQLManagement, $btnPrinterTool, $btnLectorDPicacls, $btnConfigurarIPs, $btnAddUser, $btnFirewallConfig, $btnForzarActualizacion, $btnClearAnyDesk, $btnShowPrinters, $btnClearPrintJobs, $btnAplicacionesNS, $btnCheckPermissions, $btnCambiarOTM, $btnCreateAPK, $btnExtractInstaller)
    foreach ($button in $buttonsToUpdate) {
        $button.Add_MouseLeave({ if ($script:setInstructionText) { $script:setInstructionText.Invoke($global:defaultInstructions) } })
    }
    $portsResult = Get-SqlPortWithDebug
    $portsArray = @($portsResult)
    $global:sqlPortsData = @{Ports = $portsArray; Summary = $null; DetailedText = $null; DisplayText = $null }
    if ($portsArray.Count -gt 0) {
        $sortedPorts = $portsArray | Sort-Object -Property Instance
        $displayParts = @()
        foreach ($port in $sortedPorts) {
            $instanceName = if ($port.Instance -eq "MSSQLSERVER") { "Default" } else { $port.Instance }
            $displayParts += "$instanceName : $($port.Port)"
        }
        $global:sqlPortsData.DisplayText = $displayParts -join " | "
        $global:sqlPortsData.DetailedText = $sortedPorts | ForEach-Object {
            $instanceName = if ($_.Instance -eq "MSSQLSERVER") { "Default" } else { $_.Instance }
            "- Instancia: $instanceName | Puerto: $($_.Port) | Tipo: $($_.Type)"
        } | Out-String
        $global:sqlPortsData.Summary = "Total de instancias con puerto encontradas: $($sortedPorts.Count)"
        $lblPort.Text = $global:sqlPortsData.DisplayText
        $lblPort.Tag = $global:sqlPortsData.DetailedText.Trim()
        $lblPort.ToolTip = if ($sortedPorts.Count -eq 1) { "Haz clic para mostrar en consola y copiar al portapapeles" } else { "$($sortedPorts.Count) instancias encontradas. Haz clic para detalles" }
        Write-Host "`n=== RESUMEN DE BÚSQUEDA SQL ===" -ForegroundColor Cyan
        Write-Host $global:sqlPortsData.Summary -ForegroundColor White
        Write-Host "Puertos: " -ForegroundColor White -NoNewline
        foreach ($port in $sortedPorts) {
            $instanceName = if ($port.Instance -eq "MSSQLSERVER") { "Default" } else { $port.Instance }
            Write-Host "$instanceName : " -ForegroundColor White -NoNewline
            Write-Host "$($port.Port) " -ForegroundColor Magenta -NoNewline
            if ($port -ne $sortedPorts[-1]) { Write-Host "| " -ForegroundColor Gray -NoNewline }
        }
        Write-Host ""
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
            Write-DzDebug "`t[DEBUG] Click en lblPort - Evento iniciado" -Color DarkGray
            try {
                $textToCopy = $global:sqlPortsData.DetailedText.Trim()
                Write-DzDebug "`t[DEBUG] Contenido a copiar: '$textToCopy'" -Color DarkGray
                Write-Host "`n=== INFORMACIÓN DE PUERTOS SQL ===" -ForegroundColor Cyan
                if ($global:sqlPortsData.Ports.Count -gt 0) {
                    Write-Host $global:sqlPortsData.Summary -ForegroundColor White
                    Write-Host ""
                    $textToCopy -split "`n" | ForEach-Object { Write-Host $_ -ForegroundColor Green }
                } else {
                    Write-Host $textToCopy -ForegroundColor Red
                }
                Write-Host "=====================================" -ForegroundColor Cyan
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
                            Ui-Error "Error al copiar al portapapeles: $($_.Exception.Message)" $global:MainWindow
                        } else {
                            Start-Sleep -Milliseconds 100
                        }
                    }
                }
                if ($copied) {
                    if ($global:sqlPortsData.Ports.Count -gt 0) {
                        Write-Host "`n[ÉXITO] Información de puertos SQL copiada al portapapeles:" -ForegroundColor Green
                        $textToCopy -split "`n" | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
                    } else {
                        Write-Host "`n[INFORMACIÓN] $textToCopy (copiado al portapapeles)" -ForegroundColor Yellow
                    }
                }
            } catch {
                Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
                Ui-Error "Error: $($_.Exception.Message)" $global:MainWindow
            }
        })
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
            Write-DzDebug ("`t[DEBUG] Click en 'Instalar Herramientas' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Instalar Herramientas' - - -" -ForegroundColor Gray
            if (-not (Check-Chocolatey)) {
                Write-Host "Chocolatey no está instalado. No se puede abrir el menú de instaladores." -ForegroundColor Red
                return
            }
            Show-ChocolateyInstallerMenu
        })
    $btnProfiler.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Profiler' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Profiler' - - -" -ForegroundColor Gray
            Invoke-PortableTool -ToolName "ExpressProfiler" -Url "https://github.com/ststeiger/ExpressProfiler/releases/download/1.0/ExpressProfiler20.zip" -ZipPath "C:\Temp\ExpressProfiler22wAddinSigned.zip" -ExtractPath "C:\Temp\ExpressProfiler2" -ExeName "ExpressProfiler.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnDatabase.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Database' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Database' - - -" -ForegroundColor Gray
            Invoke-PortableTool -ToolName "Database4" -Url "https://fishcodelib.com/files/DatabaseNet4.zip" -ZipPath "C:\Temp\DatabaseNet4.zip" -ExtractPath "C:\Temp\Database4" -ExeName "Database4.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnPrinterTool.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Printer Tool' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Printer Tool' - - -" -ForegroundColor Gray
            Invoke-PortableTool -ToolName "POS Printer Test" -Url "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip" -ZipPath "C:\Temp\RPT-RPI-Printer-Tool-1.zip" -ExtractPath "C:\Temp\RPT-RPI-Printer-Tool-1" -ExeName "POS Printer Test.exe" -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnLectorDPicacls.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Lector DP + icacls' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Lector DP + icacls' - - -" -ForegroundColor Gray
            $pwPs = $null
            $pwDrv = $null
            try {
                $rMain = Show-WpfMessageBox -Message "Este proceso ejecutará cambios de permisos con PsExec (SYSTEM) y puede descargar/instalar un driver.`n`n¿Deseas continuar?" -Title "Confirmar operación" -Buttons "YesNo" -Icon "Warning"
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
                        if ($null -ne $pwPs) { try { Close-WpfProgressBar -Window $pwPs } catch {} ; $pwPs = $null }
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
                $ResponderDriver = Show-WpfMessageBox -Message "¿Desea descargar e instalar el driver del lector?" -Title "Descargar Driver" -Buttons "YesNo" -Icon "Question"
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
                    $rExistDrv = Show-WpfMessageBox -Message "El driver ya existe en:`n$validationPath`n`nSí = Instalar local`nNo = Volver a descargar`nCancelar = Cancelar operación" -Title "Driver ya existe" -Buttons "YesNoCancel" -Icon "Question"
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
                    $rDrv = Show-WpfMessageBox -Message "¿Deseas descargar el driver del lector ahora?" -Title "Confirmar descarga" -Buttons "YesNo" -Icon "Question"
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
                    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue }
                    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
                    if (-not (Test-Path $validationPath)) { throw "No se encontró $validationPath" }
                    Update-WpfProgressBar -Window $pwDrv -Percent 100 -Message "Instalando..."
                    if ($null -ne $txt_InfoInstrucciones) { $txt_InfoInstrucciones.Text = "Instalando driver..." }
                    try { $pwDrv.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action] {}) | Out-Null } catch {}
                    try { $pwDrv.Dispatcher.Invoke([action] { $pwDrv.Topmost = $false; $pwDrv.WindowState = [System.Windows.WindowState]::Minimized }) | Out-Null } catch {}
                    Start-Process "msiexec.exe" -ArgumentList "/i `"$validationPath`" /passive" -Wait
                    try { $pwDrv.Dispatcher.Invoke([action] { $pwDrv.WindowState = [System.Windows.WindowState]::Normal }) | Out-Null } catch {}
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
                if ($null -ne $pwDrv) { try { Close-WpfProgressBar -Window $pwDrv } catch {} ; $pwDrv = $null }
                if ($null -ne $pwPs) { try { Close-WpfProgressBar -Window $pwPs } catch {} ; $pwPs = $null }
            }
        })
    $btnSQLManager.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'SQL Manager' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'SQL Manager' - - -" -ForegroundColor Gray
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
            Write-DzDebug ("`t[DEBUG] Click en 'SQL Management' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'SQL Management' - - -" -ForegroundColor Gray
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
            Write-DzDebug ("`t[DEBUG] Click en 'Forzar Actualización' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Forzar Actualización' - - -" -ForegroundColor Gray
            Show-SystemComponents
            $ok = Ui-Confirm "¿Desea forzar la actualización de datos?" "Confirmación" $global:MainWindow
            if ($ok) { Start-SystemUpdate ; Ui-Info "Actualización completada" "Éxito" $global:MainWindow } else { Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red }
        })
    $btnClearAnyDesk.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Clear AnyDesk' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Clear AnyDesk' - - -" -ForegroundColor Gray
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
            Write-DzDebug ("`t[DEBUG] Click en 'Show Printers' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Show Printers' - - -" -ForegroundColor Gray; Show-NSPrinters })
    $btnClearPrintJobs.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Clear Print Jobs' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Clear Print Jobs' - - -" -ForegroundColor Gray ; Invoke-ClearPrintJobs -InfoTextBlock $txt_InfoInstrucciones })
    $btnCheckPermissions.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Revisar Permisos' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Revisar Permisos' - - -" -ForegroundColor Gray
            if (-not (Test-Administrator)) { Ui-Error "Esta acción requiere permisos de administrador.`r`nPor favor, ejecuta Gerardo Zermeño Tools como administrador." $global:MainWindow ; return }
            Check-Permissions
        })
    $btnAplicacionesNS.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Aplicaciones NS' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Aplicaciones NS' - - -" -ForegroundColor Gray; $res = Get-NSApplicationsIniReport ; Show-NSApplicationsIniReport -Resultados $res })
    $btnCambiarOTM.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Cambiar OTM' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Cambiar OTM' - - -" -ForegroundColor Gray ; Invoke-CambiarOTMConfig -InfoTextBlock $txt_InfoInstrucciones })
    $LZMAbtnBuscarCarpeta.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Buscar Instaladores LZMA' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Buscar Instaladores LZMA' - - -" -ForegroundColor Gray; Show-LZMADialog })
    $btnConfigurarIPs.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Configurar IPs' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Configurar IPs' - - -" -ForegroundColor Gray ; Show-IPConfigDialog })
    $btnAddUser.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Agregar Usuario' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Agregar Usuario' - - -" -ForegroundColor Gray ; Show-AddUserDialog })
    $btnFirewallConfig.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Configuraciones de Firewall' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Configuraciones de Firewall' - - -" -ForegroundColor Gray ; Show-FirewallConfigDialog })
    $btnCreateAPK.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Crear APK' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Crear APK' - - -" -ForegroundColor Gray ; Invoke-CreateApk -InfoTextBlock $txt_InfoInstrucciones })
    $btnExtractInstaller.Add_Click({ Write-DzDebug ("`t[DEBUG] Click en 'Extraer Instalador' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Extraer Instalador' - - -" -ForegroundColor Gray ; Show-InstallerExtractorDialog })
    $btnConnectDb.Add_Click({
            Write-DzDebug ("`t[DEBUG] Click en 'Conectar Base de Datos' - {0}" -f (Get-Date -Format "HH:mm:ss"))
            Write-Host "`t- - - Comenzando el proceso de 'Conectar Base de Datos' - - -" -ForegroundColor Gray

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
                $global:database = $global:cmbDatabases.SelectedItem

                # Actualizar barra de estado
                $global:lblConnectionStatus.Text = "✓ Conectado a: $serverText | DB: $($global:database)"

                $global:txtServer.IsEnabled = $false
                $global:txtUser.IsEnabled = $false
                $global:txtPassword.IsEnabled = $false
                $global:btnExecute.IsEnabled = $true
                $global:btnClearQuery.IsEnabled = $true
                $global:cmbQueries.IsEnabled = $true
                $global:btnConnectDb.IsEnabled = $false
                $global:btnBackup.IsEnabled = $true
                $global:btnDisconnectDb.IsEnabled = $true
                $global:btnFormat.IsEnabled = $true
                $global:btnComment.IsEnabled = $true
                # ✅ Habilitar UI de trabajo al conectar
                if ($global:tcQueries) { $global:tcQueries.IsEnabled = $true }
                if ($global:tcResults) { $global:tcResults.IsEnabled = $true }

                if ($global:rtbQueryEditor1) { $global:rtbQueryEditor1.IsEnabled = $true }
                if ($global:dgResults) { $global:dgResults.IsEnabled = $true }
                if ($global:txtMessages) { $global:txtMessages.IsEnabled = $true }

                # Opcional: enfoque al editor
                $global:rtbQueryEditor1.Focus() | Out-Null
                # IMPORTANTE: Inicializar TreeView con el nuevo RichTextBox
                Initialize-SqlTreeView -TreeView $global:tvDatabases -Server $serverText -Credential $credential -InsertTextHandler {
                    param($text)
                    # Insertar en el RichTextBox actual
                    if ($global:rtbQueryEditor1) {
                        $global:rtbQueryEditor1.Focus()
                        $global:rtbQueryEditor1.CaretPosition.InsertTextInRun($text)
                    }
                }

            } catch {
                Write-DzDebug "`t[DEBUG][btnConnectDb] CATCH: $($_.Exception.Message)"
                Write-DzDebug "`t[DEBUG][btnConnectDb] Tipo: $($_.Exception.GetType().FullName)"
                Write-DzDebug "`t[DEBUG][btnConnectDb] Stack: $($_.ScriptStackTrace)"

                Ui-Error "Error de conexión: $($_.Exception.Message)" $global:MainWindow
                Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
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

                $global:lblConnectionStatus.Text = "Desconectado"
                $global:btnConnectDb.IsEnabled = $true
                $global:btnBackup.IsEnabled = $false
                $global:btnDisconnectDb.IsEnabled = $false
                $global:btnExecute.IsEnabled = $false
                $global:btnClearQuery.IsEnabled = $false
                $global:btnFormat.IsEnabled = $false
                $global:btnComment.IsEnabled = $false

                $global:txtServer.IsEnabled = $true
                $global:txtUser.IsEnabled = $true
                $global:txtPassword.IsEnabled = $true
                $global:cmbQueries.IsEnabled = $false
                $global:rtbQueryEditor1.IsEnabled = $false
                $global:dgResults.IsEnabled = $false
                $global:txtMessages.IsEnabled = $false
                if ($global:tcQueries) { $global:tcQueries.IsEnabled = $false }
                if ($global:tcResults) { $global:tcResults.IsEnabled = $false }
                $global:tvDatabases.Items.Clear()
                $global:cmbDatabases.Items.Clear()
                $global:cmbDatabases.IsEnabled = $false
                $global:dgResults.ItemsSource = $null
                $global:txtMessages.Text = ""

                Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow
            } catch {
                Write-Host "`nError al desconectar: $($_.Exception.Message)" -ForegroundColor Red
            }
        })
    $cmbDatabases.Add_SelectionChanged({
            if ($global:cmbDatabases.SelectedItem) {
                $global:database = $global:cmbDatabases.SelectedItem
                if ($global:lblConnectionStatus.Content -like "Conectado a:*") {
                    $global:lblConnectionStatus.Content = @"
Conectado a:
Servidor: $($global:server)
Base de datos: $($global:database)
"@.Trim()
                }
            }
        })
    $btnExecute.Add_Click({
            Write-Host "`n`t- - - Ejecutando consulta - - -" -ForegroundColor Gray
            try {
                $selectedDb = $global:cmbDatabases.SelectedItem
                if (-not $selectedDb) { throw "Selecciona una base de datos" }

                # Obtener texto del RichTextBox
                $textRange = New-Object System.Windows.Documents.TextRange(
                    $global:rtbQueryEditor1.Document.ContentStart,
                    $global:rtbQueryEditor1.Document.ContentEnd
                )
                $query = $textRange.Text

                if ([string]::IsNullOrWhiteSpace($query)) {
                    throw "La consulta está vacía."
                }

                # Limpiar resultados anteriores
                $global:dgResults.ItemsSource = $null
                $global:txtMessages.Text = ""

                Write-Host "Ejecutando consulta en '$selectedDb'..." -ForegroundColor Cyan

                # Usar la función existente
                $result = Invoke-SqlQueryMultiResultSet -Server $global:server -Database $selectedDb -Query $query -Credential $global:dbCredential

                if (-not $result.Success) {
                    $global:txtMessages.Text = $result.ErrorMessage
                    throw $result.ErrorMessage
                }

                # Mostrar resultados
                if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
                    $global:dgResults.ItemsSource = $result.ResultSets[0].DataTable.DefaultView
                    $global:lblRowCount.Text = "Filas: $($result.ResultSets[0].RowCount)"
                    Write-Host "✓ Consulta ejecutada: $($result.ResultSets[0].RowCount) filas" -ForegroundColor Green
                } elseif ($result.ContainsKey('RowsAffected')) {
                    $global:txtMessages.Text = "Filas afectadas: $($result.RowsAffected)"
                    $global:lblRowCount.Text = "Filas afectadas: $($result.RowsAffected)"
                    Write-Host "✓ Consulta ejecutada: $($result.RowsAffected) filas afectadas" -ForegroundColor Green
                }

            } catch {
                $global:txtMessages.Text = $_.Exception.Message
                Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
                Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Host "====================================" -ForegroundColor Red
            }
        })
    $btnClearQuery.Add_Click({
            $global:rtbQueryEditor1.Document.Blocks.Clear()
            $global:dgResults.ItemsSource = $null
            $global:txtMessages.Text = ""
            $global:lblRowCount.Text = "Filas: --"
            Write-Host "Consulta limpiada" -ForegroundColor Cyan
        })
    $btnBackup.Add_Click({ Write-Host "`n`t- - - Comenzando el proceso de Backup - - -" -ForegroundColor Gray ; Show-BackupDialog -Server $global:server -User $global:user -Password $global:password -Database $global:cmbDatabases.SelectedItem })
    $window.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::F5 -and $global:btnExecute.IsEnabled) {
                $global:btnExecute.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
            }
            if ($e.Key -eq [System.Windows.Input.Key]::T -and [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl)) {
                New-QueryTab -TabControl $global:tcQueries | Out-Null
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
    return $window
}
function Start-Application {
    Show-GlobalProgress -Percent 0 -Status "Inicializando..."
    if (-not (Initialize-Environment)) { Show-GlobalProgress -Percent 100 -Status "Error inicializando" ; return }
    Show-GlobalProgress -Percent 10 -Status "Entorno listo"
    Show-GlobalProgress -Percent 20 -Status "Cargando módulos..."
    $modulesPath = Join-Path $PSScriptRoot "modules"
    $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "SqlTreeView.psm1", "MultiQuery.psm1", "Installers.psm1")
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
    Write-Host "Error fatal: $_" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    pause
    exit 1
}
