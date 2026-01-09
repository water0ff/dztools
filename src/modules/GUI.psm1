#requires -Version 5.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
function Get-DzUiTheme {
    $iniMode = "dark"
    if (Get-Command Get-DzUiMode -ErrorAction SilentlyContinue) { $iniMode = Get-DzUiMode }
    $themes = @{
        Light = @{
            FormBackground = "#F4F6F8"; FormForeground = "#111111"
            InfoBackground = "#FFFFFF"; InfoForeground = "#111111"
            InfoHoverBackground = "#FF8C00"; InfoHoverForeground = "#111111"
            ControlBackground = "#FFFFFF"; ControlForeground = "#111111"
            BorderColor = "#CFCFCF"
            ButtonGeneralBackground = "#E6E6E6"; ButtonGeneralForeground = "#111111"
            ButtonSystemBackground = "#96C8FF"; ButtonSystemForeground = "#111111"
            ButtonNationalBackground = "#FFC896"; ButtonNationalForeground = "#111111"
            ConsoleBackground = "#FFFFFF"; ConsoleForeground = "#111111"
            AccentPrimary = "#1976D2"; AccentSecondary = "#2E7D32"; AccentMuted = "#6B7280"
            UiFontFamily = "Segoe UI"; UiFontSize = 12
            CodeFontFamily = "Consolas"; CodeFontSize = 12
            #OnAccentForeground = "#000000"
            AccentBlue = "#007ACC"
            AccentBlueHover = "#006BB3"
            AccentOrange = "#CE9178"
            AccentOrangeHover = "#D19A66"
            AccentMagenta = "#C586C0"
            AccentMagentaHover = "#B267B3"
            OnAccentForeground = "#FFFFFF"
        }
        Dark  = @{
            FormBackground = "#000000"; FormForeground = "#FFFFFF"
            InfoBackground = "#1E1E1E"; InfoForeground = "#FFFFFF"
            InfoHoverBackground = "#FF8C00"; InfoHoverForeground = "#000000"
            ControlBackground = "#1C1C1C"; ControlForeground = "#FFFFFF"
            BorderColor = "#4C4C4C"
            ButtonGeneralBackground = "#2F2F2F"; ButtonGeneralForeground = "#FFFFFF"
            ButtonSystemBackground = "#96C8FF"; ButtonSystemForeground = "#000000"
            ButtonNationalBackground = "#FFC896"; ButtonNationalForeground = "#000000"
            ConsoleBackground = "#012456"; ConsoleForeground = "#FFFFFF"
            AccentPrimary = "#2196F3"; AccentSecondary = "#4CAF50"; AccentMuted = "#9CA3AF"
            UiFontFamily = "Segoe UI"; UiFontSize = 12
            CodeFontFamily = "Consolas"; CodeFontSize = 12
            #OnAccentForeground = "#000000"
            AccentBlue = "#007ACC"
            AccentBlueHover = "#006BB3"
            AccentOrange = "#CE9178"
            AccentOrangeHover = "#D19A66"
            AccentMagenta = "#C586C0"
            AccentMagentaHover = "#B267B3"
            OnAccentForeground = "#FFFFFF"
        }
    }
    $selectedMode = if ($iniMode -match '^(dark|light)$') { ($iniMode.Substring(0, 1).ToUpper() + $iniMode.Substring(1).ToLower()) } else { "Dark" }
    $themes[$selectedMode]
}

function New-WpfWindow {
    param([Parameter(Mandatory)][object]$Xaml, [switch]$PassThru)
    $xmlReader = $null; $stringReader = $null
    try {
        $xamlText = switch ($Xaml.GetType().FullName) {
            'System.String' { $Xaml }
            'System.Xml.XmlDocument' { $Xaml.OuterXml }
            default { [string]$Xaml }
        }
        if ([string]::IsNullOrWhiteSpace($xamlText)) { throw "XAML vacío o nulo." }
        $stringReader = New-Object System.IO.StringReader($xamlText)
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $xmlReader = [System.Xml.XmlReader]::Create($stringReader, $settings)
        $window = [Windows.Markup.XamlReader]::Load($xmlReader)
        if ($PassThru) {
            [xml]$xmlDoc = $xamlText
            $controls = @{}

            $nodes = $xmlDoc.SelectNodes("//*[@Name]")
            foreach ($node in $nodes) {
                $n = $node.GetAttribute("Name")
                if (-not [string]::IsNullOrWhiteSpace($n)) {
                    $controls[$n] = $window.FindName($n)
                }
            }

            return @{ Window = $window; Controls = $controls }
        }
        $window
    } catch {
        Write-Error "Error cargando XAML: $($_.Exception.Message)"
        throw
    } finally {
        if ($xmlReader) { $xmlReader.Close() }
        if ($stringReader) { $stringReader.Close() }
    }
}
function Set-BrushResource {
    param([Parameter(Mandatory)][System.Windows.ResourceDictionary]$Resources, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { throw "Theme error: el color para '$Key' llegó vacío/nulo." }
    if ($Hex -notmatch '^#([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$') { throw "Theme error: el color para '$Key' no es HEX válido: '$Hex' (usa #RRGGBB o #AARRGGBB)." }
    $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
    if ($brush -is [System.Windows.Freezable] -and $brush.CanFreeze) { $brush.Freeze() }
    $Resources[$Key] = $brush
}
function Set-DzWpfThemeResources {
    param([Parameter(Mandatory)][System.Windows.Window]$Window, [Parameter(Mandatory)]$Theme)
    Set-BrushResource -Resources $Window.Resources -Key "FormBg" -Hex $Theme.FormBackground
    Set-BrushResource -Resources $Window.Resources -Key "FormFg" -Hex $Theme.FormForeground
    Set-BrushResource -Resources $Window.Resources -Key "PanelBg" -Hex $Theme.InfoBackground
    Set-BrushResource -Resources $Window.Resources -Key "PanelFg" -Hex $Theme.InfoForeground
    Set-BrushResource -Resources $Window.Resources -Key "ControlBg" -Hex $Theme.ControlBackground
    Set-BrushResource -Resources $Window.Resources -Key "ControlFg" -Hex $Theme.ControlForeground
    Set-BrushResource -Resources $Window.Resources -Key "BorderBrushColor" -Hex $Theme.BorderColor
    Set-BrushResource -Resources $Window.Resources -Key "AccentPrimary" -Hex $Theme.AccentPrimary
    Set-BrushResource -Resources $Window.Resources -Key "AccentSecondary" -Hex $Theme.AccentSecondary
    Set-BrushResource -Resources $Window.Resources -Key "AccentMuted" -Hex $Theme.AccentMuted
    # ComboBox: forzar esquema claro SIEMPRE (dark y light)
    Set-BrushResource -Resources $Window.Resources -Key "ComboBoxBg"      -Hex "#FFFFFF"
    Set-BrushResource -Resources $Window.Resources -Key "ComboBoxFg"      -Hex "#111111"
    Set-BrushResource -Resources $Window.Resources -Key "ComboBoxBorder"  -Hex $Theme.BorderColor
    Set-BrushResource -Resources $Window.Resources -Key "ComboBoxDropBg"  -Hex "#FFFFFF"
    Set-BrushResource -Resources $Window.Resources -Key "ComboBoxDropFg"  -Hex "#111111"
    # VS Code accents (new)
    Set-BrushResource -Resources $Window.Resources -Key "AccentBlue"         -Hex $Theme.AccentBlue
    Set-BrushResource -Resources $Window.Resources -Key "AccentBlueHover"    -Hex $Theme.AccentBlueHover
    Set-BrushResource -Resources $Window.Resources -Key "AccentOrange"       -Hex $Theme.AccentOrange
    Set-BrushResource -Resources $Window.Resources -Key "AccentOrangeHover"  -Hex $Theme.AccentOrangeHover
    Set-BrushResource -Resources $Window.Resources -Key "AccentMagenta"      -Hex $Theme.AccentMagenta
    Set-BrushResource -Resources $Window.Resources -Key "AccentMagentaHover" -Hex $Theme.AccentMagentaHover
    $onAccent = "#000000"
    if ($Theme -and $Theme.PSObject -and ($Theme.PSObject.Properties.Match('OnAccentForeground').Count -gt 0)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$Theme.OnAccentForeground)) { $onAccent = [string]$Theme.OnAccentForeground }
    }
    Set-BrushResource -Resources $Window.Resources -Key "OnAccentFg" -Hex $onAccent
    $Window.Resources["UiFontFamily"] = [System.Windows.Media.FontFamily]::new($Theme.UiFontFamily)
    $Window.Resources["UiFontSize"] = [double]$Theme.UiFontSize
    $Window.Resources["CodeFontFamily"] = [System.Windows.Media.FontFamily]::new($Theme.CodeFontFamily)
    $Window.Resources["CodeFontSize"] = [double]$Theme.CodeFontSize
    # ------------------------------------------------------------
    # Global ComboBox styling (fix dropdown/items readability)
    # ------------------------------------------------------------
    try {
        $cbType = [Type]'System.Windows.Controls.ComboBox'
        $cbiType = [Type]'System.Windows.Controls.ComboBoxItem'
        $tbType = [Type]'System.Windows.Controls.TextBox'

        $bg = $Window.Resources["ComboBoxBg"]
        $fg = $Window.Resources["ComboBoxFg"]
        $brd = $Window.Resources["ComboBoxBorder"]
        $dropBg = $Window.Resources["ComboBoxDropBg"]
        $dropFg = $Window.Resources["ComboBoxDropFg"]

        # Limpia estilos previos si existen (para evitar duplicados)
        if ($Window.Resources.Contains($cbType)) { $Window.Resources.Remove($cbType) }
        if ($Window.Resources.Contains($cbiType)) { $Window.Resources.Remove($cbiType) }

        # ComboBox (la caja visible)
        $comboStyle = New-Object System.Windows.Style($cbType)
        [void]$comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $bg)))
        [void]$comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $fg)))
        [void]$comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderBrushProperty, $brd)))
        [void]$comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderThicknessProperty, [System.Windows.Thickness]::new(1))))
        #[void]$comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.FrameworkElement]::MarginProperty, [System.Windows.Thickness]::new(2))))
        # IMPORTANTE: si el ComboBox es editable, el TextBox interno a veces queda con Foreground "del tema"
        # Le aplicamos estilo al TextBox interno SOLO cuando vive dentro de un ComboBox editable.
        $editTbStyle = New-Object System.Windows.Style($tbType)
        $editTbStyle.BasedOn = $Window.TryFindResource($tbType)  # respeta tu estilo global de TextBox si existe
        [void]$editTbStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $bg)))
        [void]$editTbStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $fg)))
        [void]$editTbStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderThicknessProperty, [System.Windows.Thickness]::new(0))))
        if ([System.Windows.Controls.ComboBox]::TextBoxStyleProperty -ne $null) {
            $comboStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.ComboBox]::TextBoxStyleProperty, $editTbStyle)))
        }
        # Dropdown items
        $itemStyle = New-Object System.Windows.Style($cbiType)
        [void]$itemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $dropBg)))
        [void]$itemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $dropFg)))
        [void]$itemStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::PaddingProperty, [System.Windows.Thickness]::new(6, 4, 6, 4))))

        # Hover/Selected un poco más moderno sin meternos a ControlTemplate
        $tHover = New-Object System.Windows.Trigger
        $tHover.Property = [System.Windows.Controls.ComboBoxItem]::IsHighlightedProperty
        $tHover.Value = $true
        [void]$tHover.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $Window.Resources["AccentPrimary"])))
        [void]$tHover.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $Window.Resources["OnAccentFg"])))

        $tSel = New-Object System.Windows.Trigger
        $tSel.Property = [System.Windows.Controls.Primitives.Selector]::IsSelectedProperty
        $tSel.Value = $true
        [void]$tSel.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $Window.Resources["AccentPrimary"])))
        [void]$tSel.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $Window.Resources["OnAccentFg"])))

        [void]$itemStyle.Triggers.Add($tHover)
        [void]$itemStyle.Triggers.Add($tSel)

        $Window.Resources.Add($cbType, $comboStyle)
        $Window.Resources.Add($cbiType, $itemStyle)
    } catch {
        Write-DzDebug "`t[DEBUG] ComboBox unified (light-like) style failed: $($_.Exception.Message)" -Color DarkGray
    }
}
function Set-WpfDialogOwner {
    param([Parameter(Mandatory)][System.Windows.Window]$Dialog)
    try { if ($Global:window -is [System.Windows.Window]) { $Dialog.Owner = $Global:window; return } } catch {}
    try { if ($Global:MainWindow -is [System.Windows.Window]) { $Dialog.Owner = $Global:MainWindow; return } } catch {}
    try { if ($script:window -is [System.Windows.Window]) { $Dialog.Owner = $script:window; return } } catch {}
}

function ConvertTo-MessageBoxResult {
    param([object]$Value)
    if ($null -eq $Value) { return [System.Windows.MessageBoxResult]::None }
    if ($Value -is [System.Windows.MessageBoxResult]) { return $Value }
    if ($Value -is [int]) {
        switch ($Value) {
            6 { return [System.Windows.MessageBoxResult]::Yes }
            7 { return [System.Windows.MessageBoxResult]::No }
            2 { return [System.Windows.MessageBoxResult]::Cancel }
            1 { return [System.Windows.MessageBoxResult]::OK }
            default { return [System.Windows.MessageBoxResult]::None }
        }
    }
    $s = ($Value.ToString()).Trim().ToLowerInvariant()
    switch ($s) {
        'yes' { [System.Windows.MessageBoxResult]::Yes }
        'no' { [System.Windows.MessageBoxResult]::No }
        'cancel' { [System.Windows.MessageBoxResult]::Cancel }
        'ok' { [System.Windows.MessageBoxResult]::OK }
        'true' { [System.Windows.MessageBoxResult]::Yes }
        'false' { [System.Windows.MessageBoxResult]::No }
        default { [System.Windows.MessageBoxResult]::None }
    }
}
function Show-WpfMessageBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = "Mensaje",
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")][string]$Buttons = "OK",
        [ValidateSet("Information", "Warning", "Error", "Question")][string]$Icon = "Information",
        [System.Windows.Window]$Owner
    )
    try {
        $safeTitle = [Security.SecurityElement]::Escape($Title)
        $safeMsg = [Security.SecurityElement]::Escape($Message)
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="$safeTitle" SizeToContent="WidthAndHeight" ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True" Background="Transparent" ShowInTaskbar="False" Topmost="True" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Border Background="{DynamicResource FormBg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" CornerRadius="14" Padding="16">
    <Grid Width="430">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <DockPanel Grid.Row="0" Margin="0,0,0,10">
        <TextBlock Text="$safeTitle" FontSize="15" FontWeight="SemiBold" Foreground="{DynamicResource FormFg}" DockPanel.Dock="Left"/>
        <Button Name="btnClose" Content="✕" Width="34" Height="28" Margin="10,0,0,0" DockPanel.Dock="Right" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" Cursor="Hand"/>
      </DockPanel>
      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Border Width="40" Height="40" CornerRadius="10" Background="{DynamicResource PanelBg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" VerticalAlignment="Top">
          <TextBlock Name="txtIcon" Text="i" FontSize="20" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource AccentPrimary}"/>
        </Border>
        <TextBlock Grid.Column="1" Name="txtMessage" Text="$safeMsg" TextWrapping="Wrap" Margin="12,2,0,0" Foreground="{DynamicResource PanelFg}"/>
      </Grid>
      <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
        <Button Name="btn1" Width="120" Height="34" Margin="0,0,10,0"/>
        <Button Name="btn2" Width="120" Height="34" Margin="0,0,10,0"/>
        <Button Name="btn3" Width="140" Height="34"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"@
        $result = New-WpfWindow -Xaml $xaml -PassThru
        if (-not $result -or -not $result.Window) { throw "Show-WpfMessageBox: ventana no creada." }
        $w = $result.Window
        $c = $result.Controls
        $dzState = @{ Result = [System.Windows.MessageBoxResult]::None; Completed = $false }
        $theme = Get-DzUiTheme
        try { Set-DzWpfThemeResources -Window $w -Theme $theme } catch {}
        if ($Owner) { try { $w.Owner = $Owner; $w.WindowStartupLocation = "CenterOwner" } catch { $w.WindowStartupLocation = "CenterScreen" } } else { $w.WindowStartupLocation = "CenterScreen" }
        $btnClose = $c['btnClose']
        $btn1 = $c['btn1']
        $btn2 = $c['btn2']
        $btn3 = $c['btn3']
        $txtIcon = $c['txtIcon']
        if ($txtIcon) {
            switch ($Icon) {
                "Information" { $txtIcon.Text = "i" }
                "Warning" { $txtIcon.Text = "!" }
                "Error" { $txtIcon.Text = "×" }
                "Question" { $txtIcon.Text = "?" }
            }
        }
        $allButtons = @($btn1, $btn2, $btn3, $btnClose) | Where-Object { $_ }
        foreach ($b in @($btn1, $btn2, $btn3)) {
            if (-not $b) { continue }
            try {
                $b.Background = $w.FindResource("ControlBg")
                $b.Foreground = $w.FindResource("ControlFg")
                $b.BorderBrush = $w.FindResource("BorderBrushColor")
                $b.BorderThickness = [System.Windows.Thickness]::new(1)
                $b.Cursor = "Hand"
                $b.Visibility = "Collapsed"
            } catch {}
        }
        $w.Add_Closing({
                if (-not $dzState.Completed) {
                    $dzState.Result = [System.Windows.MessageBoxResult]::Cancel
                    $dzState.Completed = $true
                }
            }.GetNewClosure())
        $finish = {
            param([System.Windows.MessageBoxResult]$rv)
            if ($dzState.Completed) { return }
            $dzState.Result = $rv
            $dzState.Completed = $true
            foreach ($b in $allButtons) { try { $b.IsEnabled = $false } catch {} }
            try {
                $w.Dispatcher.Invoke([action] {
                        try { $w.DialogResult = $true } catch {}
                        try { $w.Close() } catch {}
                    })
            } catch {
                try { $w.Close() } catch {}
            }
        }.GetNewClosure()
        $finishSb = $finish
        $setBtn = {
            param($btn, $text, [System.Windows.MessageBoxResult]$rv, [bool]$isPrimary)
            if (-not $btn) { return }
            $btn.Content = $text
            $btn.Visibility = "Visible"
            $btn.IsDefault = $false
            $btn.IsCancel = $false
            if ($isPrimary) { $btn.IsDefault = $true }
            if ($isPrimary) {
                try {
                    $btn.Background = $w.FindResource("AccentPrimary")
                    $btn.Foreground = $w.FindResource("FormFg")
                    $btn.BorderThickness = [System.Windows.Thickness]::new(0)
                } catch {}
            }
            $localText = $text
            $localRv = $rv
            $localFinish = $finishSb
            $btn.Add_Click({
                    Write-DzDebug "`t[DEBUG] Show-WpfMessageBox: BTN '$localText' handler ejecutado. result=$localRv" -Color DarkGray
                    $localFinish.Invoke($localRv)
                }.GetNewClosure())
        }.GetNewClosure()
        switch ($Buttons) {
            "OK" { $setBtn.Invoke($btn3, "OK", ([System.Windows.MessageBoxResult]::OK), $true) }
            "OKCancel" { $setBtn.Invoke($btn2, "Cancelar", ([System.Windows.MessageBoxResult]::Cancel), $false); $setBtn.Invoke($btn3, "OK", ([System.Windows.MessageBoxResult]::OK), $true) }
            "YesNo" { $setBtn.Invoke($btn2, "No", ([System.Windows.MessageBoxResult]::No), $false); $setBtn.Invoke($btn3, "Sí", ([System.Windows.MessageBoxResult]::Yes), $true) }
            "YesNoCancel" { $setBtn.Invoke($btn1, "Cancelar", ([System.Windows.MessageBoxResult]::Cancel), $false); $setBtn.Invoke($btn2, "No", ([System.Windows.MessageBoxResult]::No), $false); $setBtn.Invoke($btn3, "Sí", ([System.Windows.MessageBoxResult]::Yes), $true) }
        }
        if ($btnClose) {
            $localFinish2 = $finishSb
            $btnClose.Add_Click({
                    Write-DzDebug "`t[DEBUG] Show-WpfMessageBox: btnClose clicked." -Color DarkGray
                    $localFinish2.Invoke([System.Windows.MessageBoxResult]::Cancel)
                }.GetNewClosure())
        }
        $null = $w.ShowDialog()
        return [System.Windows.MessageBoxResult]$dzState.Result
    } catch {
        Write-Warning "Show-WpfMessageBox falló: $($_.Exception.Message)"
        switch ($Buttons) {
            "YesNo" { return [System.Windows.MessageBoxResult]::No }
            "YesNoCancel" { return [System.Windows.MessageBoxResult]::Cancel }
            "OKCancel" { return [System.Windows.MessageBoxResult]::Cancel }
            default { return [System.Windows.MessageBoxResult]::OK }
        }
    }
}

function Show-WpfMessageBoxSafe {
    param(
        [Parameter(Mandatory)][string]$Message,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")][string]$Buttons,
        [Parameter(Mandatory)][ValidateSet("Information", "Warning", "Error", "Question")][string]$Icon,
        [System.Windows.Window]$Owner
    )
    ConvertTo-MessageBoxResult (Show-WpfMessageBox -Message $Message -Title $Title -Buttons $Buttons -Icon $Icon -Owner $Owner)
}

function Show-WpfProgressBar {
    param([string]$Title = "Procesando", [string]$Message = "Por favor espere...")
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="$Title" Height="220" Width="500" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True" Background="Transparent" Topmost="True" ShowInTaskbar="False" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Border Background="{DynamicResource FormBg}" CornerRadius="8" BorderBrush="{DynamicResource AccentPrimary}" BorderThickness="2" Padding="20">
    <Border.Effect><DropShadowEffect Color="Black" Direction="270" ShadowDepth="5" BlurRadius="15" Opacity="0.3"/></Border.Effect>
    <StackPanel>
      <TextBlock Text="$Title" FontWeight="Bold" Foreground="{DynamicResource AccentPrimary}" HorizontalAlignment="Center" Margin="0,0,0,20"/>
      <TextBlock Name="lblMessage" Text="$Message" Foreground="{DynamicResource FormFg}" TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,15" MinHeight="30"/>
      <ProgressBar Name="progressBar" Height="25" Minimum="0" Maximum="100"
             IsIndeterminate="True" Value="0"
             Foreground="{DynamicResource AccentSecondary}"
             Background="{DynamicResource ControlBg}" Margin="0,0,0,10"/>
      <TextBlock Name="lblPercent" Text="0%" FontWeight="Bold" Foreground="{DynamicResource AccentPrimary}" HorizontalAlignment="Center"/>
    </StackPanel>
  </Border>
</Window>
"@
    try {
        $result = New-WpfWindow -Xaml $stringXaml -PassThru
        $window = $result.Window
        Set-DzWpfThemeResources -Window $window -Theme $theme
        $window | Add-Member -MemberType NoteProperty -Name ProgressBar -Value $result.Controls['progressBar'] | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name MessageLabel -Value $result.Controls['lblMessage'] | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name PercentLabel -Value $result.Controls['lblPercent'] | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name IsClosed -Value $false | Out-Null
        $window.Add_Closed({ $window.IsClosed = $true })
        $window.Show()
        $window.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action] {}) | Out-Null
        $window
    } catch {
        Write-Error "Error al crear barra de progreso: $_"
        $null
    }
}
function Update-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Window]$Window,
        [Parameter(Mandatory = $true)][ValidateRange(0, 100)][int]$Percent,
        [string]$Message = $null
    )
    if ($null -eq $Window) {
        return
    }
    if ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed) {
        return
    }
    if ($null -eq $Window.Dispatcher -or $Window.Dispatcher.HasShutdownStarted) {
        return
    }
    $p = [Math]::Min($Percent, 100)
    $m = $Message
    $doUpdate = {
        try {
            if ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed) {
                return
            }
            if ($Window.PSObject.Properties.Match('ProgressBar').Count -gt 0 -and $Window.ProgressBar) {
                $Window.ProgressBar.IsIndeterminate = $false
                $Window.ProgressBar.Value = $p
            }
            if ($Window.PSObject.Properties.Match('PercentLabel').Count -gt 0 -and $Window.PercentLabel) {
                $Window.PercentLabel.Text = "$p%"
            }
            if (-not [string]::IsNullOrWhiteSpace($m) -and
                $Window.PSObject.Properties.Match('MessageLabel').Count -gt 0 -and
                $Window.MessageLabel) {
                $Window.MessageLabel.Text = $m
            }
        } catch {
            if (-not ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed)) {
                Write-Warning "Error actualizando ProgressBar: $($_.Exception.Message)"
            }
        }
    }
    try {
        if ($Window.Dispatcher.CheckAccess()) {
            & $doUpdate
        } else {
            $Window.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]::Normal,
                $doUpdate
            )
        }
    } catch {
        if (-not ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed)) {
            Write-DzDebug "ERROR actualizando ProgressBar: $($_.Exception.Message)" -Color Red
        }
    }
}
function Close-WpfProgressBar {
    param([Parameter(Mandatory = $true)]$Window)
    if ($null -eq $Window) { return }
    if (-not ($Window -is [System.Windows.Window])) { Write-Warning "Close-WpfProgressBar: El objeto recibido NO es WPF Window. Tipo: $($Window.GetType().FullName)"; return }
    if ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed) { return }
    if ($null -eq $Window.Dispatcher -or $Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
    try { $Window.Dispatcher.Invoke([action] { if (-not $Window.IsClosed) { $Window.Close() } }, [System.Windows.Threading.DispatcherPriority]::Normal) } catch { Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)" }
}
function Show-ProgressBar { Show-WpfProgressBar -Title "Progreso de Actualización" -Message "Iniciando proceso..." }
function Set-WpfControlEnabled {
    param([Parameter(Mandatory = $true)]$Control, [Parameter(Mandatory = $true)][bool]$Enabled)
    if ($null -eq $Control) { Write-Warning "Control es null."; return }
    try {
        if ($Control.Dispatcher.CheckAccess()) { $Control.IsEnabled = $Enabled }
        else { $Control.Dispatcher.Invoke([action] { $Control.IsEnabled = $Enabled }) }
    } catch {
        Write-Warning "Error cambiando estado del control: $_"
    }
}
function New-WpfInputDialog {
    param([string]$Title = "Entrada", [string]$Prompt = "Ingrese un valor:", [string]$DefaultValue = "")
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="$Title" Height="180" Width="400" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" ShowInTaskbar="False" Background="{DynamicResource FormBg}" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource FormFg}"/></Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="6,4"/>
    </Style>
    <Style x:Key="SystemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
      <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,6"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="20" Background="{DynamicResource FormBg}">
    <TextBlock Text="$Prompt" Margin="0,0,0,10"/>
    <TextBox Name="txtInput" Text="$DefaultValue" Margin="0,0,0,20"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnOK" Content="Aceptar" Width="90" Margin="0,0,10,0" IsDefault="True" Style="{StaticResource SystemButtonStyle}"/>
      <Button Name="btnCancel" Content="Cancelar" Width="90" IsCancel="True" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" Padding="10,6"/>
    </StackPanel>
  </StackPanel>
</Window>
"@
    $result = New-WpfWindow -Xaml $stringXaml -PassThru
    $window = $result.Window
    Set-DzWpfThemeResources -Window $window -Theme $theme
    $controls = $result.Controls
    $script:inputValue = $null
    $script:__inputOk = $false
    $controls['btnOK'].Add_Click({ $script:inputValue = $controls['txtInput'].Text; $script:__inputOk = $true; $window.Close() })
    $controls['btnCancel'].Add_Click({ $script:__inputOk = $false; $window.Close() })
    $null = $window.ShowDialog()
    if ($script:__inputOk) { return $script:inputValue }
    return $null
}
function Get-WpfPasswordBoxText {
    param([Parameter(Mandatory = $true)][System.Windows.Controls.PasswordBox]$PasswordBox)
    $PasswordBox.Password
}
function Show-WpfFolderDialog {
    param([string]$Description = "Seleccione una carpeta", [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'))
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { Write-Warning "Show-WpfFolderDialog requiere STA. Ejecuta PowerShell con -STA."; return $null }
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.SelectedPath = $InitialDirectory
    $dialog.ShowNewFolderButton = $true
    $r = $dialog.ShowDialog()
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) { $dialog.SelectedPath } else { $null }
}
function Ui-Info { param([string]$Message, [string]$Title = "Información", [System.Windows.Window]$Owner) Show-WpfMessageBoxSafe -Message $Message -Title $Title -Buttons "OK" -Icon "Information" -Owner $Owner | Out-Null }
function Ui-Warn { param([string]$Message, [string]$Title = "Atención", [System.Windows.Window]$Owner) Show-WpfMessageBoxSafe -Message $Message -Title $Title -Buttons "OK" -Icon "Warning" -Owner $Owner | Out-Null }
function Ui-Error { param([string]$Message, [string]$Title = "Error", [System.Windows.Window]$Owner) Show-WpfMessageBoxSafe -Message $Message -Title $Title -Buttons "OK" -Icon "Error" -Owner $Owner | Out-Null }
function Ui-Confirm { param([string]$Message, [string]$Title = "Confirmar", [System.Windows.Window]$Owner) (Show-WpfMessageBoxSafe -Message $Message -Title $Title -Buttons "YesNo" -Icon "Question" -Owner $Owner) -eq [System.Windows.MessageBoxResult]::Yes }

Export-ModuleMember -Function @(
    'Get-DzUiTheme', 'New-WpfWindow', 'Show-WpfMessageBox', 'Show-WpfMessageBoxSafe', 'ConvertTo-MessageBoxResult',
    'New-WpfInputDialog', 'Show-WpfProgressBar', 'Update-WpfProgressBar', 'Close-WpfProgressBar', 'Show-ProgressBar',
    'Set-WpfControlEnabled', 'Get-WpfPasswordBoxText', 'Show-WpfFolderDialog', 'Set-WpfDialogOwner', 'Set-DzWpfThemeResources',
    'Ui-Info', 'Ui-Warn', 'Ui-Error', 'Ui-Confirm'
)