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
            FormBackground            = "#000000"
            FormForeground            = "#FFFFFF"
            InfoBackground            = "#1E1E1E"
            InfoForeground            = "#FFFFFF"
            InfoHoverBackground       = "#FF8C00"
            ButtonGeneralBackground   = "#2F2F2F"
            ButtonGeneralForeground   = "#FFFFFF"
            ButtonSystemBackground    = "#96C8FF"
            ButtonSystemForeground    = "#000000"
            ButtonNationalBackground  = "#FFC896"
            ButtonNationalForeground  = "#000000"
            ConsoleBackground         = "#012456"
            ConsoleForeground         = "#FFFFFF"
            BorderColor               = "#4C4C4C"
            AccentPrimary             = "#2196F3"
            AccentSecondary           = "#4CAF50"
            AccentMuted               = "#757575"
            ControlBackground         = "#1C1C1C"
            ControlForeground         = "#FFFFFF"
        }
        Dark  = @{
            FormBackground            = "#000000"
            FormForeground            = "#FFFFFF"
            InfoBackground            = "#1E1E1E"
            InfoForeground            = "#FFFFFF"
            InfoHoverBackground       = "#FF8C00"
            ButtonGeneralBackground   = "#2F2F2F"
            ButtonGeneralForeground   = "#FFFFFF"
            ButtonSystemBackground    = "#96C8FF"
            ButtonSystemForeground    = "#000000"
            ButtonNationalBackground  = "#FFC896"
            ButtonNationalForeground  = "#000000"
            ConsoleBackground         = "#012456"
            ConsoleForeground         = "#FFFFFF"
            BorderColor               = "#4C4C4C"
            AccentPrimary             = "#2196F3"
            AccentSecondary           = "#4CAF50"
            AccentMuted               = "#757575"
            ControlBackground         = "#1C1C1C"
            ControlForeground         = "#FFFFFF"
        }
    }

    $selectedMode = "Light"
    return $themes[$selectedMode]
}

function New-WpfWindow {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xaml,
        [switch]$PassThru
    )
    try {
        $reader = New-Object System.Xml.XmlNodeReader $Xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)

        if ($PassThru) {
            $controls = @{}
            $Xaml.SelectNodes("//*[@Name]") | ForEach-Object {
                $controls[$_.Name] = $window.FindName($_.Name)
            }
            return @{
                Window   = $window
                Controls = $controls
            }
        }
        return $window
    } catch {
        Write-Error "Error cargando XAML: $_"
        throw
    }
}

function Show-WpfMessageBox {
    param(
        [string]$Message,
        [string]$Title = "Mensaje",
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")]
        [string]$Buttons = "OK",
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [string]$Icon = "Information"
    )

    $buttonMap = @{
        "OK"          = [System.Windows.MessageBoxButton]::OK
        "OKCancel"    = [System.Windows.MessageBoxButton]::OKCancel
        "YesNo"       = [System.Windows.MessageBoxButton]::YesNo
        "YesNoCancel" = [System.Windows.MessageBoxButton]::YesNoCancel
    }

    $iconMap = @{
        "Information" = [System.Windows.MessageBoxImage]::Information
        "Warning"     = [System.Windows.MessageBoxImage]::Warning
        "Error"       = [System.Windows.MessageBoxImage]::Error
        "Question"    = [System.Windows.MessageBoxImage]::Question
    }

    return [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        $buttonMap[$Buttons],
        $iconMap[$Icon]
    )
}

#endregion

#region Barra de Progreso WPF Mejorada

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

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title"
        Height="220" Width="500"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="False">
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
                       FontSize="12"
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
        $result = New-WpfWindow -Xaml $xaml -PassThru
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
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title"
        Height="180" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <StackPanel Margin="20" Background="$($theme.FormBackground)">
        <TextBlock Text="$Prompt" FontSize="12" Margin="0,0,0,10"/>
        <TextBox Name="txtInput" Text="$DefaultValue" FontSize="12" Padding="5" Margin="0,0,0,20"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="Aceptar" Width="80" Margin="0,0,10,0" IsDefault="True" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnCancel" Content="Cancelar" Width="80" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $result = New-WpfWindow -Xaml $xaml -PassThru
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
    'Set-WpfDialogOwner'
)
