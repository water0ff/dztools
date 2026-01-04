#requires -Version 5.0

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

#region Funciones Base WPF

function Get-DzUiTheme {
    $iniMode = "dark"
    if (Get-Command Get-DzUiMode -ErrorAction SilentlyContinue) {
        $iniMode = Get-DzUiMode
    }

    $themes = @{
        Light = @{
            FormBackground           = "#FFFFFF"
            FormForeground           = "#333333"

            InfoBackground           = "#F3F3F3"
            InfoForeground           = "#333333"
            InfoHoverBackground      = "#CCE4FF"
            InfoHoverForeground      = "#1A1A1A"

            ControlBackground        = "#FFFFFF"
            ControlForeground        = "#333333"

            BorderColor              = "#E5E5E5"

            ButtonGeneralBackground  = "#E5E5E5"
            ButtonGeneralForeground  = "#333333"
            ButtonSystemBackground   = "#007ACC"
            ButtonSystemForeground   = "#FFFFFF"
            ButtonNationalBackground = "#16825D"
            ButtonNationalForeground = "#FFFFFF"

            ConsoleBackground        = "#FFFFFF"
            ConsoleForeground        = "#111111"

            AccentPrimary            = "#007ACC"
            AccentSecondary          = "#16825D"

            AccentMuted              = "#6A6A6A"

            UiFontFamily             = "Segoe UI"
            UiFontSize               = 13
            CodeFontFamily           = "Consolas"
            CodeFontSize             = 13
        }
        Dark  = @{
            FormBackground           = "#1E1E1E"
            FormForeground           = "#D4D4D4"

            InfoBackground           = "#252526"
            InfoForeground           = "#D4D4D4"
            InfoHoverBackground      = "#264F78"
            InfoHoverForeground      = "#FFFFFF"

            ControlBackground        = "#3C3C3C"
            ControlForeground        = "#D4D4D4"

            BorderColor              = "#454545"

            ButtonGeneralBackground  = "#3C3C3C"
            ButtonGeneralForeground  = "#D4D4D4"
            ButtonSystemBackground   = "#0E639C"
            ButtonSystemForeground   = "#FFFFFF"
            ButtonNationalBackground = "#16825D"
            ButtonNationalForeground = "#FFFFFF"

            ConsoleBackground        = "#012456"
            ConsoleForeground        = "#FFFFFF"

            AccentPrimary            = "#0E639C"
            AccentSecondary          = "#16825D"
            AccentMuted              = "#9B9B9B"

            UiFontFamily             = "Segoe UI"
            UiFontSize               = 13
            CodeFontFamily           = "Consolas"
            CodeFontSize             = 13
        }
    }

    $selectedMode = if ($iniMode -match '^(dark|light)$') {
        ($iniMode.Substring(0, 1).ToUpper() + $iniMode.Substring(1).ToLower())
    } else { 'Dark' }

    return $themes[$selectedMode]
}


function New-WpfWindow {
    param(
        [Parameter(Mandatory)]
        [object]$Xaml,
        [switch]$PassThru
    )

    try {
        # 1) Convertir a string XAML
        $xamlText = switch ($Xaml.GetType().FullName) {
            'System.String' { $Xaml }
            'System.Xml.XmlDocument' { $Xaml.OuterXml }
            default { [string]$Xaml }
        }

        if ([string]::IsNullOrWhiteSpace($xamlText)) {
            throw "XAML vacío o nulo."
        }

        # 2) Crear XmlReader desde string (evita problemas con XmlNodeReader)
        $stringReader = New-Object System.IO.StringReader($xamlText)
        $xmlReaderSettings = New-Object System.Xml.XmlReaderSettings
        $xmlReaderSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $xmlReaderSettings.XmlResolver = $null
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader, $xmlReaderSettings)

        # 3) Cargar ventana WPF
        $window = [Windows.Markup.XamlReader]::Load($xmlReader)

        if ($PassThru) {
            # Para mapear controles por Name necesitamos un XmlDocument (pero SOLO para SelectNodes)
            [xml]$xmlDoc = $xamlText

            $controls = @{}
            $xmlDoc.SelectNodes("//*[@Name]") | ForEach-Object {
                $controls[$_.Name] = $window.FindName($_.Name)
            }
            return @{ Window = $window; Controls = $controls }
        }

        return $window
    } catch {
        Write-Error "Error cargando XAML: $($_.Exception.Message)"
        throw
    } finally {
        if ($xmlReader) { $xmlReader.Close() }
        if ($stringReader) { $stringReader.Close() }
    }
}

function Show-WpfMessageBox {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Title = "Mensaje",
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")]
        [string]$Buttons = "OK",
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [string]$Icon = "Information"
    )

    $theme = Get-DzUiTheme

    # Iconos simples (Unicode). Si quieres, luego los cambiamos por Paths (vector).
    $iconGlyph = switch ($Icon) {
        "Information" { "ℹ" }
        "Warning" { "⚠" }
        "Error" { "⛔" }
        "Question" { "❓" }
        default { "ℹ" }
    }

    # Visibilidad de botones
    $showOK = "Collapsed"
    $showYes = "Collapsed"
    $showNo = "Collapsed"
    $showCancel = "Collapsed"

    switch ($Buttons) {
        "OK" {
            $showOK = "Visible"
        }
        "OKCancel" {
            $showOK = "Visible"
            $showCancel = "Visible"
        }
        "YesNo" {
            $showYes = "Visible"
            $showNo = "Visible"
        }
        "YesNoCancel" {
            $showYes = "Visible"
            $showNo = "Visible"
            $showCancel = "Visible"
        }
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title"
        Height="230" Width="520"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent">
    <Border Background="$($theme.FormBackground)"
            CornerRadius="10"
            BorderBrush="$($theme.AccentPrimary)"
            BorderThickness="2"
            Padding="0">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="14" Opacity="0.30"/>
        </Border.Effect>

        <Grid Margin="14">
            <Grid.RowDefinitions>
                <RowDefinition Height="34"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Header -->
            <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="$Title"
                           Foreground="$($theme.FormForeground)"
                           VerticalAlignment="Center"
                           FontSize="13"
                           FontWeight="SemiBold"/>

                <Button Name="btnClose"
                        Grid.Column="1"
                        Content="✕"
                        Foreground="$($theme.FormForeground)"
                        Width="34" Height="26"
                        Margin="8,0,0,0"
                        ToolTip="Cerrar"
                        Background="Transparent"
                        BorderBrush="Transparent"/>
            </Grid>

            <!-- Body -->
            <Grid Grid.Row="1" Margin="0,10,0,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="42"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Border Width="34" Height="34"
                        CornerRadius="8"
                        Background="$($theme.ControlBackground)"
                        BorderBrush="$($theme.BorderColor)"
                        BorderThickness="1"
                        VerticalAlignment="Top">
                    <TextBlock Text="$iconGlyph"
                               FontSize="18"
                               Foreground="$($theme.AccentPrimary)"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"/>
                </Border>

                <TextBlock Grid.Column="1"
                           Text="$Message"
                           Foreground="$($theme.FormForeground)"
                           FontSize="12"
                           TextWrapping="Wrap"
                           VerticalAlignment="Top"/>
            </Grid>

            <!-- Buttons -->
            <StackPanel Grid.Row="2"
                        Orientation="Horizontal"
                        HorizontalAlignment="Right">

                <Button Name="btnYes"
                        Content="Sí"
                        Visibility="$showYes"
                        Width="110" Height="30"
                        Margin="0,0,10,0"
                        Background="$($theme.ButtonSystemBackground)"
                        Foreground="$($theme.ButtonSystemForeground)"
                        BorderBrush="$($theme.BorderColor)"
                        BorderThickness="1"/>

                <Button Name="btnNo"
                        Content="No"
                        Visibility="$showNo"
                        Width="110" Height="30"
                        Margin="0,0,10,0"
                        Background="$($theme.ButtonGeneralBackground)"
                        Foreground="$($theme.ButtonGeneralForeground)"
                        BorderBrush="$($theme.BorderColor)"
                        BorderThickness="1"/>

                <Button Name="btnCancel"
                        Content="Cancelar"
                        Visibility="$showCancel"
                        Width="110" Height="30"
                        Margin="0,0,10,0"
                        Background="$($theme.ButtonGeneralBackground)"
                        Foreground="$($theme.ButtonGeneralForeground)"
                        BorderBrush="$($theme.BorderColor)"
                        BorderThickness="1"/>

                <Button Name="btnOK"
                        Content="OK"
                        Visibility="$showOK"
                        Width="110" Height="30"
                        Background="$($theme.ButtonSystemBackground)"
                        Foreground="$($theme.ButtonSystemForeground)"
                        BorderBrush="$($theme.BorderColor)"
                        BorderThickness="1"/>
            </StackPanel>

        </Grid>
    </Border>
</Window>
"@

    $ui = New-WpfWindow -Xaml $xaml -PassThru
    $w = $ui.Window
    $c = $ui.Controls

    try { Set-WpfDialogOwner -Dialog $w } catch {}

    # Drag
    $c['HeaderBar'].Add_MouseLeftButtonDown({
            if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() }
        })

    # Default/Cancel para Enter/Esc
    if ($c['btnOK']) { $c['btnOK'].IsDefault = $true }
    if ($c['btnYes']) { $c['btnYes'].IsDefault = $true }
    if ($c['btnCancel']) { $c['btnCancel'].IsCancel = $true }

    # Resultado
    $result = [System.Windows.MessageBoxResult]::None

    $c['btnClose'].Add_Click({ $result = [System.Windows.MessageBoxResult]::Cancel; $w.Close() })

    if ($c['btnOK']) {
        $c['btnOK'].Add_Click({ $result = [System.Windows.MessageBoxResult]::OK; $w.Close() })
    }
    if ($c['btnYes']) {
        $c['btnYes'].Add_Click({ $result = [System.Windows.MessageBoxResult]::Yes; $w.Close() })
    }
    if ($c['btnNo']) {
        $c['btnNo'].Add_Click({ $result = [System.Windows.MessageBoxResult]::No; $w.Close() })
    }
    if ($c['btnCancel']) {
        $c['btnCancel'].Add_Click({ $result = [System.Windows.MessageBoxResult]::Cancel; $w.Close() })
    }

    $null = $w.ShowDialog()
    return $result
}


function Show-WpfProgressBar {
    <#
    .SYNOPSIS
        Muestra una barra de progreso WPF no bloqueante.
    #>
    param(
        [string]$Title = "Procesando",
        [string]$Message = "Por favor espere..."
    )

    $theme = Get-DzUiTheme

    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title"
        Height="220" Width="500"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="False"
        FontFamily="$($theme.UiFontFamily)"
        FontSize="$($theme.UiFontSize)">
    <Border Background="$($theme.FormBackground)"
            CornerRadius="8"
            BorderBrush="$($theme.AccentPrimary)"
            BorderThickness="2"
            Padding="20">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="5" BlurRadius="15" Opacity="0.3"/>
        </Border.Effect>
        <StackPanel>
            <TextBlock Text="$Title"
                       FontSize="18"
                       FontWeight="Bold"
                       Foreground="$($theme.AccentPrimary)"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,20"/>

            <TextBlock Name="lblMessage"
                       Text="$Message"
                       FontSize="$($theme.UiFontSize)"
                       Foreground="$($theme.FormForeground)"
                       TextAlignment="Center"
                       TextWrapping="Wrap"
                       Margin="0,0,0,15"
                       MinHeight="30"/>

            <ProgressBar Name="progressBar"
                         Height="25"
                         Minimum="0"
                         Maximum="100"
                         Value="0"
                         Foreground="$($theme.AccentSecondary)"
                         Background="$($theme.ControlBackground)"
                         Margin="0,0,0,10"/>

            <TextBlock Name="lblPercent"
                       Text="0%"
                       FontSize="14"
                       FontWeight="Bold"
                       Foreground="$($theme.AccentPrimary)"
                       HorizontalAlignment="Center"/>
        </StackPanel>
    </Border>
</Window>
"@

    try {
        $result = New-WpfWindow -Xaml $stringXaml -PassThru
        $window = $result.Window

        $window | Add-Member -MemberType NoteProperty -Name ProgressBar   -Value $result.Controls['progressBar'] | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name MessageLabel  -Value $result.Controls['lblMessage']  | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name PercentLabel  -Value $result.Controls['lblPercent']  | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name IsClosed      -Value $false                          | Out-Null
        $window.Add_Closed({
                $window.IsClosed = $true
            })
        $window.Show()

        # Forzar render inicial (siempre con el dispatcher de la misma ventana)
        $window.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action] {}
        ) | Out-Null
        return $window
    } catch {
        Write-Error "Error al crear barra de progreso: $_"
        return $null
    }
}
function Set-WpfDialogOwner {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Window]$Dialog
    )

    try {
        if ($Global:window -is [System.Windows.Window]) { $Dialog.Owner = $Global:window; return }
    } catch {}

    try {
        if ($script:window -is [System.Windows.Window]) { $Dialog.Owner = $script:window; return }
    } catch {}
}

function Update-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)] $Window,
        [Parameter(Mandatory = $true)][ValidateRange(0, 100)][int] $Percent,
        [string] $Message = $null
    )

    if ($null -eq $Window) { return }
    if (-not ($Window -is [System.Windows.Window])) { return }
    if ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed) { return }

    # Capturar valores (sin cambiar variables externas para compatibilidad)
    $pLocal = [Math]::Min($Percent, 100)
    $mLocal = $Message

    # Copia local del window para pasarla como argumento al delegate
    $wLocal = $Window

    try {
        # Tipado del delegate para poder pasar parámetros (evita closures raros en PS5)
        $action = [Action[object, int, string]] {
            param($w, $p, $m)

            if ($null -eq $w) { return }
            if ($w.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $w.IsClosed) { return }

            # ProgressBar
            if ($w.PSObject.Properties.Match('ProgressBar').Count -gt 0 -and $w.ProgressBar) {
                $w.ProgressBar.IsIndeterminate = $false
                $w.ProgressBar.Value = $p
            }

            # Porcentaje
            if ($w.PSObject.Properties.Match('PercentLabel').Count -gt 0 -and $w.PercentLabel) {
                $w.PercentLabel.Text = "$p%"
            }

            # Mensaje
            if (-not [string]::IsNullOrWhiteSpace($m) -and
                $w.PSObject.Properties.Match('MessageLabel').Count -gt 0 -and $w.MessageLabel) {
                $w.MessageLabel.Text = $m
            }

            $w.UpdateLayout()
        }

        # Invoke con prioridad de render, pasando args en lugar de capturar variables
        $wLocal.Dispatcher.Invoke(
            $action,
            [System.Windows.Threading.DispatcherPriority]::Render,
            $wLocal, $pLocal, $mLocal
        ) | Out-Null

    } catch {
        Write-Warning "Error actualizando barra de progreso: $($_.Exception.Message)"
    }
}

function Close-WpfProgressBar {
    param([Parameter(Mandatory = $true)] $Window)

    if ($null -eq $Window) { return }

    if (-not ($Window -is [System.Windows.Window])) {
        Write-Warning "Close-WpfProgressBar: El objeto recibido NO es WPF Window. Tipo: $($Window.GetType().FullName)"
        return
    }

    if ($Window.IsClosed) { return }

    if ($null -eq $Window.Dispatcher -or
        $Window.Dispatcher.HasShutdownStarted -or
        $Window.Dispatcher.HasShutdownFinished) { return }

    try {
        $Window.Dispatcher.Invoke([action] {
                if (-not $Window.IsClosed) { $Window.Close() }
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
    } catch {
        Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)"
    }
}
function Show-ProgressBar {
    return Show-WpfProgressBar -Title "Progreso de Actualización" -Message "Iniciando proceso..."
}
function Set-WpfControlEnabled {
    param(
        [Parameter(Mandatory = $true)]
        $Control,
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    if ($null -eq $Control) {
        Write-Warning "Control es null."
        return
    }

    try {
        if ($Control.Dispatcher.CheckAccess()) {
            $Control.IsEnabled = $Enabled
        } else {
            $Control.Dispatcher.Invoke([action] {
                    $Control.IsEnabled = $Enabled
                })
        }
    } catch {
        Write-Warning "Error cambiando estado del control: $_"
    }
}
function New-WpfInputDialog {
    param(
        [string]$Title = "Entrada",
        [string]$Prompt = "Ingrese un valor:",
        [string]$DefaultValue = ""
    )

    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title"
        Height="180" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="$($theme.FormBackground)"
        FontFamily="$($theme.UiFontFamily)"
        FontSize="$($theme.UiFontSize)">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <StackPanel Margin="20" Background="$($theme.FormBackground)">
        <TextBlock Text="$Prompt" FontSize="$($theme.UiFontSize)" Margin="0,0,0,10"/>
        <TextBox Name="txtInput" Text="$DefaultValue" FontSize="$($theme.UiFontSize)" Padding="5" Margin="0,0,0,20"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="Aceptar" Width="80" Margin="0,0,10,0" IsDefault="True" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnCancel" Content="Cancelar" Width="80" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $result = New-WpfWindow -Xaml $stringXaml -PassThru
    $window = $result.Window
    $controls = $result.Controls

    $script:inputValue = $null

    $controls['btnOK'].Add_Click({
            $script:inputValue = $controls['txtInput'].Text
            $window.DialogResult = $true
            $window.Close()
        })

    $controls['btnCancel'].Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })

    $dialogResult = $window.ShowDialog()

    if ($dialogResult) {
        return $script:inputValue
    }

    return $null
}

function Get-WpfPasswordBoxText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.PasswordBox]$PasswordBox
    )
    return $PasswordBox.Password
}
function Show-WpfFolderDialog {
    param(
        [string]$Description = "Seleccione una carpeta",
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop')
    )
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        Write-Warning "Show-WpfFolderDialog requiere STA. Ejecuta PowerShell con -STA."
        return $null
    }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.SelectedPath = $InitialDirectory
    $dialog.ShowNewFolderButton = $true

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}


function Set-BrushResource {
    param(
        [Parameter(Mandatory)]
        [System.Windows.ResourceDictionary]$Resources,

        [Parameter(Mandatory)]
        [string]$Key,

        [Parameter(Mandatory)]
        [string]$Hex
    )

    if ([string]::IsNullOrWhiteSpace($Hex)) {
        throw "Theme error: el color para '$Key' llegó vacío/nulo."
    }

    # Acepta #RRGGBB o #AARRGGBB
    if ($Hex -notmatch '^#([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$') {
        throw "Theme error: el color para '$Key' no es HEX válido: '$Hex' (usa #RRGGBB o #AARRGGBB)."
    }

    # ✅ Esto regresa un Brush (SolidColorBrush) listo para Background/Foreground/BorderBrush
    $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)

    # Congelar si se puede (mejor rendimiento)
    if ($brush -is [System.Windows.Freezable] -and $brush.CanFreeze) { $brush.Freeze() }

    $Resources[$Key] = $brush
}


function Set-DzWpfThemeResources {
    param(
        [Parameter(Mandatory)] [System.Windows.Window]$Window,
        [Parameter(Mandatory)] $Theme
    )

    # Mapeo: keys WPF (FormBg, PanelBg...) -> propiedades reales del Theme
    Set-BrushResource -Resources $Window.Resources -Key "FormBg"           -Hex $Theme.FormBackground
    Set-BrushResource -Resources $Window.Resources -Key "FormFg"           -Hex $Theme.FormForeground

    # PanelBg/PanelFg: puedes usar InfoBackground/InfoForeground para paneles
    Set-BrushResource -Resources $Window.Resources -Key "PanelBg"          -Hex $Theme.InfoBackground
    Set-BrushResource -Resources $Window.Resources -Key "PanelFg"          -Hex $Theme.InfoForeground

    Set-BrushResource -Resources $Window.Resources -Key "ControlBg"        -Hex $Theme.ControlBackground
    Set-BrushResource -Resources $Window.Resources -Key "ControlFg"        -Hex $Theme.ControlForeground

    Set-BrushResource -Resources $Window.Resources -Key "BorderBrushColor" -Hex $Theme.BorderColor

    Set-BrushResource -Resources $Window.Resources -Key "AccentPrimary"    -Hex $Theme.AccentPrimary
    Set-BrushResource -Resources $Window.Resources -Key "AccentSecondary"  -Hex $Theme.AccentSecondary

    $Window.Resources["UiFontFamily"] = [System.Windows.Media.FontFamily]::new($Theme.UiFontFamily)
    $Window.Resources["UiFontSize"] = [double]$Theme.UiFontSize
    $Window.Resources["CodeFontFamily"] = [System.Windows.Media.FontFamily]::new($Theme.CodeFontFamily)
    $Window.Resources["CodeFontSize"] = [double]$Theme.CodeFontSize
}


Export-ModuleMember -Function @(
    'Get-DzUiTheme',
    'New-WpfWindow',
    'Show-WpfMessageBox',
    'New-WpfInputDialog',
    'Show-WpfProgressBar',
    'Update-WpfProgressBar',
    'Close-WpfProgressBar',
    'Set-WpfControlEnabled',
    'Get-WpfPasswordBoxText',
    'Show-WpfFolderDialog',
    'Show-ProgressBar',
    'Set-WpfDialogOwner',
    'Set-DzWpfThemeResources'
)
