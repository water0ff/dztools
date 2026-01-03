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

# Estado de la UI de consola
$script:ProgressActive = $false
$script:LastProgressLen = 0

function Show-GlobalProgress {
    param(
        [int]$Percent,
        [string]$Status
    )

    $Percent = [math]::Max(0, [math]::Min(100, $Percent))

    $width = 40
    $filled = [math]::Round(($Percent / 100) * $width)
    $bar = "[" + ("#" * $filled).PadRight($width) + "]"
    $text = "{0} {1,3}% - {2}" -f $bar, $Percent, $Status

    # Recorta para no rebasar ancho de consola
    $max = [System.Console]::WindowWidth - 1
    if ($max -gt 10 -and $text.Length -gt $max) { $text = $text.Substring(0, $max) }

    # Sobrescribe la misma línea (sin mover cursor por filas)
    $pad = ""
    if ($script:LastProgressLen -gt $text.Length) {
        $pad = " " * ($script:LastProgressLen - $text.Length)
    }

    Write-Host ("`r" + $text + $pad) -NoNewline

    $script:ProgressActive = $true
    $script:LastProgressLen = ($text.Length + $pad.Length)
}

function Stop-GlobalProgress {
    # Termina la línea de progreso para que el siguiente Write-Host no se pegue
    if ($script:ProgressActive) {
        Write-Host ""  # nueva línea
        $script:ProgressActive = $false
        $script:LastProgressLen = 0
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    # Si hay barra activa, ciérrala antes de imprimir logs
    Stop-GlobalProgress
    Write-Host $Message -ForegroundColor $Color
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

    Write-Host "`nCreando formulario principal WPF..." -ForegroundColor Yellow
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Gerardo Zermeño Tools"
        Height="650" Width="1050"
        WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <Style TargetType="{x:Type Label}">
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>

        <!-- TabControl -->
        <Style TargetType="{x:Type TabControl}">
            <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>

        <!-- TabItem -->
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

        <!-- TextBox base -->
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

        <!-- ComboBox -->
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

        <!-- ====== Styles faltantes ====== -->

        <Style x:Key="InfoHeaderTextBoxStyle"
               TargetType="{x:Type TextBox}"
               BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style x:Key="ConsoleTextBoxStyle"
               TargetType="{x:Type TextBox}"
               BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="10"/>
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
        </Style>

        <Style x:Key="SystemButtonStyle"
               TargetType="{x:Type Button}"
               BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        </Style>

        <Style x:Key="NationalSoftButtonStyle"
               TargetType="{x:Type Button}"
               BasedOn="{StaticResource GeneralButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
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
                                       FontSize="10"
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

                    <TextBox Name="lblHostname" HorizontalAlignment="Left" VerticalAlignment="Top"
                           Width="220" Height="40" Margin="10,1,0,0" Background="{DynamicResource ControlBg}"
                            Foreground="{DynamicResource ControlFg}"
                            BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" Cursor="Hand"/>
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
                    <Button Content="Actualizar datos del sistema" Name="btnForzarActualizacion" Width="220" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,450,0,0" Style="{StaticResource SystemButtonStyle}"/>
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
                             Width="220" Height="400" Margin="730,50,0,0" Style="{StaticResource ConsoleTextBoxStyle}"
                             IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                             FontFamily="Courier New" FontSize="10"/>
                    <Border Name="cardQuickSettings"
                            Width="220" Height="150"
                            Margin="730,460,0,0"
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
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                                <TextBlock Text="🌙 Dark Mode" VerticalAlignment="Center"/>
                                <ToggleButton Name="tglDarkMode" Style="{StaticResource TogglePillStyle}" Margin="10,0,0,0"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
                                <TextBlock Text="🐞 DEBUG" VerticalAlignment="Center"/>
                                <ToggleButton Name="tglDebugMode" Style="{StaticResource TogglePillStyle}" Margin="24,0,0,0"/>
                            </StackPanel>
                            <TextBlock Text="Al cambiar, reinicia la app para aplicar."
                                       FontSize="10"
                                       Foreground="{DynamicResource PanelFg}"
                                       TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>
            <TabItem Header="Base de datos" Name="tabProSql">
                  <Grid Background="{DynamicResource PanelBg}">

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
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,220,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Content="Desconectar de BDD" Name="btnDisconnectDb" Width="180" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,260,0,0" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
                    <Button Content="Backup BDD" Name="btnBackup" Width="180" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,300,0,0" Style="{StaticResource SystemButtonStyle}"/>
                    <Label Name="lblConnectionStatus" Content="Conectado a BDD: Ninguna" HorizontalAlignment="Left"
                           VerticalAlignment="Top" Width="180" Height="80" Margin="10,400,0,0"/>
                    <Button Content="Ejecutar" Name="btnExecute" Width="100" Height="30"
                            HorizontalAlignment="Left" VerticalAlignment="Top" Margin="220,20,0,0" Style="{StaticResource SystemButtonStyle}" IsEnabled="False"/>
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
                HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="350,0,0,10" Style="{StaticResource GeneralButtonStyle}"/>
    </Grid>
</Window>
"@

    [xml]$xaml = $stringXaml
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    try {
        $window = [Windows.Markup.XamlReader]::Load($reader)
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $window -Theme $theme

    } catch {
        Write-Host "`n[XAML ERROR] $($_.Exception.Message)" -ForegroundColor Red

        # XamlParseException suele traer LineNumber/LinePosition
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
    $cmbQueries = $window.FindName("cmbQueries")
    $rtbQuery = $window.FindName("rtbQuery")
    $tglDarkMode = $window.FindName("tglDarkMode")
    $tglDebugMode = $window.FindName("tglDebugMode")
    $script:predefinedQueries = Get-PredefinedQueries
    Initialize-PredefinedQueries -ComboQueries $cmbQueries -RichTextBox $rtbQuery -Queries $script:predefinedQueries -Window $window
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
    $lblHostname.text = [System.Net.Dns]::GetHostName()
    $txt_InfoInstrucciones.Text = $global:defaultInstructions
    $script:initializingToggles = $true
    if ($tglDarkMode) {
        $tglDarkMode.IsChecked = ((Get-DzUiMode) -eq 'dark')
    }
    if ($tglDebugMode) {
        $tglDebugMode.IsChecked = (Get-DzDebugPreference)
    }
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
            $txt_InfoInstrucciones.Dispatcher.Invoke([action] {
                    $txt_InfoInstrucciones.Text = $Message
                })
        }
    }.GetNewClosure()
    $showRestartNotice = {
        param([string]$settingLabel)
        Show-WpfMessageBox -Message "Se guardó $settingLabel en dztools.ini.`nReinicia la aplicación para aplicar los cambios." `
            -Title "Reinicio requerido" -Buttons OK -Icon Information | Out-Null
    }.GetNewClosure()

    if ($tglDarkMode) {
        $tglDarkMode.Add_Checked({
                if ($script:initializingToggles) { return }
                Set-DzUiMode -Mode "dark"
                $showRestartNotice.Invoke("el modo Dark")
            })
        $tglDarkMode.Add_Unchecked({
                if ($script:initializingToggles) { return }
                Set-DzUiMode -Mode "light"
                $showRestartNotice.Invoke("el modo Light")
            })
    }

    if ($tglDebugMode) {
        $tglDebugMode.Add_Checked({
                if ($script:initializingToggles) { return }
                Set-DzDebugPreference -Enabled $true
                $showRestartNotice.Invoke("DEBUG activado")
            })
        $tglDebugMode.Add_Unchecked({
                if ($script:initializingToggles) { return }
                Set-DzDebugPreference -Enabled $false
                $showRestartNotice.Invoke("DEBUG desactivado")
            })
    }

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
    $buttonsToUpdate = @(
        $LZMAbtnBuscarCarpeta, $btnInstalarHerramientas, $btnProfiler,
        $btnDatabase, $btnSQLManager, $btnSQLManagement, $btnPrinterTool,
        $btnLectorDPicacls, $btnConfigurarIPs, $btnAddUser, $btnForzarActualizacion,
        $btnClearAnyDesk, $btnShowPrinters, $btnClearPrintJobs, $btnAplicacionesNS,
        $btnCheckPermissions, $btnCambiarOTM, $btnCreateAPK, $btnExtractInstaller
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
        $lblPort.ToolTip = if ($sortedPorts.Count -eq 1) {
            "Haz clic para mostrar en consola y copiar al portapapeles"
        } else {
            "$($sortedPorts.Count) instancias encontradas. Haz clic para detalles"
        }
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
    $lblHostname.Add_PreviewMouseLeftButtonDown({
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
            Invoke-PortableTool `
                -ToolName "ExpressProfiler" `
                -Url "https://github.com/ststeiger/ExpressProfiler/releases/download/1.0/ExpressProfiler20.zip" `
                -ZipPath "C:\Temp\ExpressProfiler22wAddinSigned.zip" `
                -ExtractPath "C:\Temp\ExpressProfiler2" `
                -ExeName "ExpressProfiler.exe" `
                -InfoTextBlock $txt_InfoInstrucciones
        })

    $btnPrinterTool.Add_Click({
            Invoke-PortableTool `
                -ToolName "POS Printer Test" `
                -Url "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip" `
                -ZipPath "C:\Temp\RPT-RPI-Printer-Tool-1.zip" `
                -ExtractPath "C:\Temp\RPT-RPI-Printer-Tool-1" `
                -ExeName "POS Printer Test.exe" `
                -InfoTextBlock $txt_InfoInstrucciones
        })

    $btnDatabase.Add_Click({
            Invoke-PortableTool `
                -ToolName "Database4" `
                -Url "https://fishcodelib.com/files/DatabaseNet4.zip" `
                -ZipPath "C:\Temp\DatabaseNet4.zip" `
                -ExtractPath "C:\Temp\Database4" `
                -ExeName "Database4.exe" `
                -InfoTextBlock $txt_InfoInstrucciones
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
                $possiblePaths = @(
                    "${env:SystemRoot}\System32\SQLServerManager*.msc",
                    "${env:SystemRoot}\SysWOW64\SQLServerManager*.msc"
                )

                $managers = foreach ($pattern in $possiblePaths) {
                    Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | ForEach-Object FullName
                }

                @($managers) | Where-Object { $_ } | Select-Object -Unique
            }
            $managers = Get-SQLServerManagers
            if (-not $managers -or $managers.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "No se encontró ninguna versión de SQL Server Configuration Manager.",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
                return
            }
            Show-SQLselector -Managers $managers
        })

    $btnSQLManagement.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Write-DzDebug "`t[DEBUG]`t[DEBUG] Iniciando búsqueda de SSMS instalados" -Color DarkGray
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
                            Write-DzDebug "`t[DEBUG]`t[DEBUG] ✓ Encontrado: $($f.FullName)" -Color Green
                        }
                    }
                }

                if ($ssmsPaths.Count -eq 0) {
                    Write-DzDebug "`t[DEBUG]`t[DEBUG] No se encontró en rutas fijas. Buscando en registro..." -Color DarkGray
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
                            if (-not $installPath.EndsWith('\')) { $installPath += '\' }
                            foreach ($sub in @("Common7\IDE\Ssms.exe", "Release\Common7\IDE\Ssms.exe")) {
                                $full = Join-Path $installPath $sub
                                if (Test-Path $full) {
                                    $resolved = (Resolve-Path $full).Path
                                    if ($ssmsPaths -notcontains $resolved) {
                                        $ssmsPaths += $resolved
                                        Write-DzDebug "`t[DEBUG]`t[DEBUG] ✓ Encontrado (registro): $resolved" -Color Green
                                    }
                                }
                            }
                        }
                    }
                }

                $ssmsPaths | Sort-Object -Descending
            }

            $ssmsVersions = Get-SSMSVersions

            $filteredVersions = foreach ($p in $ssmsVersions) {
                if ((Split-Path $p -Leaf) -eq "Ssms.exe" -and (Test-Path $p -PathType Leaf)) { $p }
            }

            if (-not $filteredVersions -or $filteredVersions.Count -eq 0) {
                Write-Host "`tNo se encontró ninguna versión de SSMS instalada." -ForegroundColor Red

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
                            [System.Windows.MessageBox]::Show(
                                "Error al ejecutar SSMS: $($_.Exception.Message)",
                                "Error",
                                [System.Windows.MessageBoxButton]::OK,
                                [System.Windows.MessageBoxImage]::Error
                            ) | Out-Null
                        }
                    }
                } else {
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
            # ✅ Aquí ya usamos el mismo selector
            Show-SQLselector -SSMSVersions $filteredVersions
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
            Show-NSPrinters
        })

    $btnClearPrintJobs.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Invoke-ClearPrintJobs -InfoTextBlock $txt_InfoInstrucciones
        })
    $btnCheckPermissions.Add_Click({
            Write-Host "`nRevisando permisos en C:\NationalSoft" -ForegroundColor Yellow
            if (-not (Test-Administrator)) {
                [System.Windows.MessageBox]::Show("Esta acción requiere permisos de administrador.`r`nPor favor, ejecuta Gerardo Zermeño Tools como administrador.", "Permisos insuficientes", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }
            Check-Permissions
        })
    $btnAplicacionesNS.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $res = Get-NSApplicationsIniReport
            Show-NSApplicationsIniReport -Resultados $res
        })

    $btnCambiarOTM.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Invoke-CambiarOTMConfig -InfoTextBlock $txt_InfoInstrucciones
        })


    $LZMAbtnBuscarCarpeta.Add_Click({
            Show-LZMADialog
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
            Invoke-CreateApk -InfoTextBlock $txt_InfoInstrucciones
        })

    $btnExtractInstaller.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            Show-InstallerExtractorDialog
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

    Show-GlobalProgress -Percent 0 -Status "Inicializando..."

    if (-not (Initialize-Environment)) {
        Show-GlobalProgress -Percent 100 -Status "Error inicializando"
        return
    }
    Show-GlobalProgress -Percent 10 -Status "Entorno listo"

    # Importar módulos (solo una vez)
    Show-GlobalProgress -Percent 20 -Status "Cargando módulos..."
    $modulesPath = Join-Path $PSScriptRoot "modules"
    $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Queries.psm1", "Installers.psm1")

    $i = 0
    foreach ($module in $modules) {
        $i++
        Show-GlobalProgress -Percent (20 + [math]::Round(($i / $modules.Count) * 20)) -Status "Importando $module"
        $modulePath = Join-Path $modulesPath $module
        if (Test-Path $modulePath) {
            Import-Module $modulePath -Force -DisableNameChecking -ErrorAction Stop
        }
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
