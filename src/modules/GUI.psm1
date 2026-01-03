#requires -Version 5.0

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
function Get-DzUiTheme {
    $iniMode = "dark"
    if (Get-Command Get-DzUiMode -ErrorAction SilentlyContinue) {
        $iniMode = Get-DzUiMode
    }
    $themes = @{
        Light = @{
            FormBackground           = "#F4F6F8"
            FormForeground           = "#111111"
            InfoBackground           = "#FFFFFF"
            InfoForeground           = "#111111"
            InfoHoverBackground      = "#FF8C00"
            InfoHoverForeground      = "#111111"   # <- importante
            ControlBackground        = "#FFFFFF"
            ControlForeground        = "#111111"
            BorderColor              = "#CFCFCF"
            ButtonGeneralBackground  = "#E6E6E6"
            ButtonGeneralForeground  = "#111111"
            ButtonSystemBackground   = "#96C8FF"
            ButtonSystemForeground   = "#111111"
            ButtonNationalBackground = "#FFC896"
            ButtonNationalForeground = "#111111"
            ConsoleBackground        = "#FFFFFF"
            ConsoleForeground        = "#111111"
            AccentPrimary            = "#1976D2"
            AccentSecondary          = "#2E7D32"
            AccentMuted              = "#6B7280"
        }
        Dark  = @{
            FormBackground           = "#000000"
            FormForeground           = "#FFFFFF"
            InfoBackground           = "#1E1E1E"
            InfoForeground           = "#FFFFFF"
            InfoHoverBackground      = "#FF8C00"
            InfoHoverForeground      = "#000000"   # <- importante (naranja + negro = legible)
            ControlBackground        = "#1C1C1C"
            ControlForeground        = "#FFFFFF"
            BorderColor              = "#4C4C4C"
            ButtonGeneralBackground  = "#2F2F2F"
            ButtonGeneralForeground  = "#FFFFFF"
            ButtonSystemBackground   = "#96C8FF"
            ButtonSystemForeground   = "#000000"
            ButtonNationalBackground = "#FFC896"
            ButtonNationalForeground = "#000000"
            ConsoleBackground        = "#012456"
            ConsoleForeground        = "#FFFFFF"
            AccentPrimary            = "#2196F3"
            AccentSecondary          = "#4CAF50"
            AccentMuted              = "#9CA3AF"
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
        $xamlText = switch ($Xaml.GetType().FullName) {
            'System.String' { $Xaml }
            'System.Xml.XmlDocument' { $Xaml.OuterXml }
            default { [string]$Xaml }
        }
        if ([string]::IsNullOrWhiteSpace($xamlText)) {
            throw "XAML vacío o nulo."
        }
        $stringReader = New-Object System.IO.StringReader($xamlText)
        $xmlReaderSettings = New-Object System.Xml.XmlReaderSettings
        $xmlReaderSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $xmlReaderSettings.XmlResolver = $null
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader, $xmlReaderSettings)
        $window = [Windows.Markup.XamlReader]::Load($xmlReader)
        if ($PassThru) {
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
function Show-WpfMessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Message,
        [string]$Title = "Mensaje",
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")]
        [string]$Buttons = "OK",
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [string]$Icon = "Information",
        [System.Windows.Window]$Owner
    )

    $safeTitle = [Security.SecurityElement]::Escape($Title)
    $safeMsg = [Security.SecurityElement]::Escape($Message)

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$safeTitle"
        SizeToContent="WidthAndHeight"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        ShowInTaskbar="False"
        Topmost="True">

    <Border Background="{DynamicResource FormBg}"
            BorderBrush="{DynamicResource BorderBrushColor}"
            BorderThickness="1"
            CornerRadius="14"
            Padding="16">

        <Grid Width="430">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <DockPanel Grid.Row="0" Margin="0,0,0,10">
                <TextBlock Text="$safeTitle"
                           FontSize="15"
                           FontWeight="SemiBold"
                           Foreground="{DynamicResource FormFg}"
                           DockPanel.Dock="Left"/>

                <Button Name="btnClose"
                        Content="✕"
                        Width="34"
                        Height="28"
                        Margin="10,0,0,0"
                        DockPanel.Dock="Right"
                        Background="{DynamicResource ControlBg}"
                        Foreground="{DynamicResource ControlFg}"
                        BorderBrush="{DynamicResource BorderBrushColor}"
                        BorderThickness="1"
                        Cursor="Hand"/>
            </DockPanel>

            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Border Width="40" Height="40"
                        CornerRadius="10"
                        Background="{DynamicResource PanelBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}"
                        BorderThickness="1"
                        VerticalAlignment="Top">
                    <TextBlock Name="txtIcon"
                               Text="i"
                               FontSize="20"
                               FontWeight="Bold"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"
                               Foreground="{DynamicResource AccentPrimary}"/>
                </Border>

                <TextBlock Grid.Column="1"
                           Name="txtMessage"
                           Text="$safeMsg"
                           TextWrapping="Wrap"
                           Margin="12,2,0,0"
                           Foreground="{DynamicResource PanelFg}"/>
            </Grid>

            <StackPanel Grid.Row="2"
                        Orientation="Horizontal"
                        HorizontalAlignment="Right"
                        Margin="0,16,0,0">
                <Button Name="btn1" Width="120" Height="34" Margin="0,0,10,0"/>
                <Button Name="btn2" Width="120" Height="34" Margin="0,0,10,0"/>
                <Button Name="btn3" Width="140" Height="34"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@

    $result = New-WpfWindow -Xaml $xaml -PassThru
    $w = $result.Window

    if (-not ($w -is [System.Windows.Window])) {
        return [System.Windows.MessageBoxResult]::None
    }

    $w | Add-Member -MemberType NoteProperty -Name DzResult -Value ([System.Windows.MessageBoxResult]::None) -Force | Out-Null

    if ($Owner) {
        $w.Owner = $Owner
        $w.WindowStartupLocation = "CenterOwner"
    } else {
        $w.WindowStartupLocation = "CenterScreen"
    }

    $theme = Get-DzUiTheme
    Set-DzWpfThemeResources -Window $w -Theme $theme

    $btnClose = $w.FindName("btnClose")
    $btn1 = $w.FindName("btn1")
    $btn2 = $w.FindName("btn2")
    $btn3 = $w.FindName("btn3")
    $txtIcon = $w.FindName("txtIcon")

    if ($txtIcon) {
        switch ($Icon) {
            "Information" { $txtIcon.Text = "i" }
            "Warning" { $txtIcon.Text = "!" }
            "Error" { $txtIcon.Text = "×" }
            "Question" { $txtIcon.Text = "?" }
        }
    }

    foreach ($b in @($btn1, $btn2, $btn3)) {
        if ($null -eq $b) { continue }
        $b.Background = $w.FindResource("ControlBg")
        $b.Foreground = $w.FindResource("ControlFg")
        $b.BorderBrush = $w.FindResource("BorderBrushColor")
        $b.BorderThickness = [System.Windows.Thickness]::new(1)
        $b.Cursor = "Hand"
        $b.Visibility = "Collapsed"
    }

    if ($btn3) {
        $btn3.Background = $w.FindResource("AccentPrimary")
        $btn3.Foreground = $w.FindResource("FormFg")
        $btn3.BorderThickness = [System.Windows.Thickness]::new(0)
    }

    $setBtn = {
        param(
            [Parameter(Mandatory)] $btn,
            [Parameter(Mandatory)] [string]$text,
            [Parameter(Mandatory)] [System.Windows.MessageBoxResult]$resultValue,
            [Parameter(Mandatory)] [bool]$isPrimary,
            [Parameter(Mandatory)] [System.Windows.Window]$win
        )

        $btn.Content = $text
        $btn.Visibility = "Visible"

        if ($isPrimary) {
            $btn.Background = $win.FindResource("AccentPrimary")
            $btn.Foreground = $win.FindResource("FormFg")
            $btn.BorderThickness = [System.Windows.Thickness]::new(0)
        } else {
            $btn.Background = $win.FindResource("ControlBg")
            $btn.Foreground = $win.FindResource("ControlFg")
            $btn.BorderBrush = $win.FindResource("BorderBrushColor")
            $btn.BorderThickness = [System.Windows.Thickness]::new(1)
        }

        $localResult = $resultValue
        $btn.Add_Click({
                param($sender, $e)
                $win.DzResult = $localResult
                try { $win.DialogResult = $true } catch {}
                $win.Close()
            }.GetNewClosure())
    }

    switch ($Buttons) {
        "OK" {
            & $setBtn $btn3 "OK" ([System.Windows.MessageBoxResult]::OK) $true $w
        }
        "OKCancel" {
            & $setBtn $btn2 "Cancelar" ([System.Windows.MessageBoxResult]::Cancel) $false $w
            & $setBtn $btn3 "OK"       ([System.Windows.MessageBoxResult]::OK)     $true  $w
        }
        "YesNo" {
            & $setBtn $btn2 "No" ([System.Windows.MessageBoxResult]::No)  $false $w
            & $setBtn $btn3 "Sí" ([System.Windows.MessageBoxResult]::Yes) $true  $w
        }
        "YesNoCancel" {
            & $setBtn $btn1 "Cancelar" ([System.Windows.MessageBoxResult]::Cancel) $false $w
            & $setBtn $btn2 "No"       ([System.Windows.MessageBoxResult]::No)     $false $w
            & $setBtn $btn3 "Sí"       ([System.Windows.MessageBoxResult]::Yes)    $true  $w
        }
    }

    if ($btnClose) {
        $btnClose.Add_Click({
                param($sender, $e)
                $w.DzResult = [System.Windows.MessageBoxResult]::Cancel
                try { $w.DialogResult = $false } catch {}
                $w.Close()
            })
    }

    $w.Add_Closing({
            param($sender, $e)
            if ($sender -is [System.Windows.Window]) {
                if ($null -eq $sender.DzResult -or $sender.DzResult -eq [System.Windows.MessageBoxResult]::None) {
                    switch ($Buttons) {
                        "YesNo" { $sender.DzResult = [System.Windows.MessageBoxResult]::No }
                        "YesNoCancel" { $sender.DzResult = [System.Windows.MessageBoxResult]::Cancel }
                        "OKCancel" { $sender.DzResult = [System.Windows.MessageBoxResult]::Cancel }
                        default { $sender.DzResult = [System.Windows.MessageBoxResult]::OK }
                    }
                }
            }
        })

    $w.Add_MouseLeftButtonDown({
            param($sender, $e)
            try { $w.DragMove() } catch {}
        })

    try {
        $null = $w.ShowDialog()
    } catch {
        return [System.Windows.MessageBoxResult]::None
    }

    return $w.DzResult
}
function ConvertTo-MessageBoxResult {
    param([object]$Value)

    if ($null -eq $Value) { return [System.Windows.MessageBoxResult]::None }

    # Si ya viene bien
    if ($Value -is [System.Windows.MessageBoxResult]) { return $Value }

    # Si viene numérico (WPF: Yes=6 No=7 OK=1 Cancel=2)
    if ($Value -is [int]) {
        switch ($Value) {
            6 { return [System.Windows.MessageBoxResult]::Yes }
            7 { return [System.Windows.MessageBoxResult]::No }
            2 { return [System.Windows.MessageBoxResult]::Cancel }
            1 { return [System.Windows.MessageBoxResult]::OK }
            default { return [System.Windows.MessageBoxResult]::None }
        }
    }

    # Si viene string
    $s = ($Value.ToString()).Trim().ToLowerInvariant()
    switch ($s) {
        'yes' { return [System.Windows.MessageBoxResult]::Yes }
        'no' { return [System.Windows.MessageBoxResult]::No }
        'cancel' { return [System.Windows.MessageBoxResult]::Cancel }
        'ok' { return [System.Windows.MessageBoxResult]::OK }
        'true' { return [System.Windows.MessageBoxResult]::Yes }
        'false' { return [System.Windows.MessageBoxResult]::No }
        default { return [System.Windows.MessageBoxResult]::None }
    }
}

function Show-WpfMessageBoxSafe {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")] [string]$Buttons,
        [Parameter(Mandatory)] [ValidateSet("Information", "Warning", "Error", "Question")] [string]$Icon,
        [System.Windows.Window]$Owner
    )
    try {
        $raw = Show-WpfMessageBox -Message $Message -Title $Title -Buttons $Buttons -Icon $Icon -Owner $Owner
        $norm = ConvertTo-MessageBoxResult $raw
        return $norm
    } catch {
        return [System.Windows.MessageBoxResult]::None
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
    $pLocal = [Math]::Min($Percent, 100)
    $mLocal = $Message
    $wLocal = $Window
    try {
        $action = [Action[object, int, string]] {
            param($w, $p, $m)

            if ($null -eq $w) { return }
            if ($w.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $w.IsClosed) { return }
            if ($w.PSObject.Properties.Match('ProgressBar').Count -gt 0 -and $w.ProgressBar) {
                $w.ProgressBar.IsIndeterminate = $false
                $w.ProgressBar.Value = $p
            }
            if ($w.PSObject.Properties.Match('PercentLabel').Count -gt 0 -and $w.PercentLabel) {
                $w.PercentLabel.Text = "$p%"
            }
            if (-not [string]::IsNullOrWhiteSpace($m) -and
                $w.PSObject.Properties.Match('MessageLabel').Count -gt 0 -and $w.MessageLabel) {
                $w.MessageLabel.Text = $m
            }
            $w.UpdateLayout()
        }
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
        Background="$($theme.FormBackground)">
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
        <TextBlock Text="$Prompt" FontSize="12" Margin="0,0,0,10"/>
        <TextBox Name="txtInput" Text="$DefaultValue" FontSize="12" Padding="5" Margin="0,0,0,20"/>
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
    if ($Hex -notmatch '^#([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$') {
        throw "Theme error: el color para '$Key' no es HEX válido: '$Hex' (usa #RRGGBB o #AARRGGBB)."
    }
    $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    if ($brush -is [System.Windows.Freezable] -and $brush.CanFreeze) { $brush.Freeze() }
    $Resources[$Key] = $brush
}
function Set-DzWpfThemeResources {
    param(
        [Parameter(Mandatory)] [System.Windows.Window]$Window,
        [Parameter(Mandatory)] $Theme
    )
    Set-BrushResource -Resources $Window.Resources -Key "FormBg"           -Hex $Theme.FormBackground
    Set-BrushResource -Resources $Window.Resources -Key "FormFg"           -Hex $Theme.FormForeground
    Set-BrushResource -Resources $Window.Resources -Key "PanelBg"          -Hex $Theme.InfoBackground
    Set-BrushResource -Resources $Window.Resources -Key "PanelFg"          -Hex $Theme.InfoForeground
    Set-BrushResource -Resources $Window.Resources -Key "ControlBg"        -Hex $Theme.ControlBackground
    Set-BrushResource -Resources $Window.Resources -Key "ControlFg"        -Hex $Theme.ControlForeground
    Set-BrushResource -Resources $Window.Resources -Key "BorderBrushColor" -Hex $Theme.BorderColor
    Set-BrushResource -Resources $Window.Resources -Key "AccentPrimary"    -Hex $Theme.AccentPrimary
    Set-BrushResource -Resources $Window.Resources -Key "AccentSecondary"  -Hex $Theme.AccentSecondary
    Set-BrushResource -Resources $Window.Resources -Key "AccentMuted"      -Hex $Theme.AccentMuted
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
    'Set-DzWpfThemeResources',
    'Show-WpfMessageBoxSafe',
    'ConvertTo-MessageBoxResult'
)