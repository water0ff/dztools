#requires -Version 5.0
function Show-DllRegistrationDialog {
    [CmdletBinding()]
    param()
    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] INICIO"
    $theme = Get-DzUiTheme
    $defaultList = @(
        "#Asi se ve uno deshabilitado: C:\Windows\SysWOW64\slicensing.dll (no agregar comillas)"
        "C:\Windows\SysWOW64\Nslicensing.dll"
        "C:\Windows\SysWOW64\Nslicensingr1.dll"
        "C:\Windows\SysWOW64\Nsservices1_0.dll"
        "C:\Windows\SysWOW64\Nsservices1_0r1.dll"
    ) -join "`r`n"
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Registro de DLLs"
        Height="400" Width="760"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True">
  <Window.Resources>
    <Style x:Key="SystemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
      <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
    </Style>
  </Window.Resources>
  <Border Background="{DynamicResource FormBg}"
          BorderBrush="{DynamicResource BorderBrushColor}"
          BorderThickness="1"
          CornerRadius="8"
          Padding="12">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Row="0"
                 Text="Escribe las rutas a registrar o deregistrar (una por línea)."
                 Margin="0,0,0,6"
                 FontWeight="SemiBold"/>
      <TextBlock Grid.Row="1"
                 Text="Nota: si un renglón empieza con #, se omitirá."
                 Foreground="{DynamicResource AccentMuted}"
                 Margin="0,0,0,8"/>
      <RichTextBox Grid.Row="2" Name="rtbDlls"
                   VerticalScrollBarVisibility="Auto"
                   Background="{DynamicResource ControlBg}"
                   Foreground="{DynamicResource ControlFg}"
                   BorderBrush="{DynamicResource BorderBrushColor}"
                   BorderThickness="1"/>
      <Grid Grid.Row="3" Margin="0,10,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <RadioButton Name="rbRegister" Content="Registrar (regsvr32)" IsChecked="True" Margin="0,0,16,0"/>
          <RadioButton Name="rbUnregister" Content="Deregistrar (regsvr32 /u)"/>
        </StackPanel>
        <Button Grid.Column="1" Name="btnRun" Content="Ejecutar" Width="120" Height="32" Margin="0,0,10,0"
                Style="{StaticResource SystemButtonStyle}"/>
        <Button Grid.Column="2" Name="btnClose" Content="Cerrar" Width="110" Height="32"
                Style="{StaticResource SystemButtonStyle}"/>
      </Grid>
    </Grid>
  </Border>
</Window>
"@
    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        $w = $ui.Window
        $c = $ui.Controls
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $w -Theme $theme
        try { Set-WpfDialogOwner -Dialog $w } catch {}
        $rtb = $c['rtbDlls']
        $textRange = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
        $textRange.Text = $defaultList
        $script:dllFormatting = $false
        $GetTextPointerFromOffset = {
            param(
                [System.Windows.Documents.FlowDocument]$Document,
                [int]$Offset
            )
            $navigator = $Document.ContentStart
            $count = 0
            while ($navigator -ne $null) {
                $context = $navigator.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward)
                if ($context -eq [System.Windows.Documents.TextPointerContext]::Text) {
                    $text = $navigator.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
                    if (($count + $text.Length) -ge $Offset) { return $navigator.GetPositionAtOffset($Offset - $count) }
                    $count += $text.Length
                    $navigator = $navigator.GetPositionAtOffset($text.Length)
                } else {
                    $navigator = $navigator.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
                }
            }
            return $Document.ContentEnd
        }.GetNewClosure()
        $applyHighlight = {
            if ($script:dllFormatting) { return }
            $script:dllFormatting = $true
            try {
                $caret = $rtb.CaretPosition

                $range = New-Object System.Windows.Documents.TextRange(
                    $rtb.Document.ContentStart,
                    $rtb.Document.ContentEnd
                )

                $text = $range.Text -replace "`r", ""

                $rtb.Document.Blocks.Clear()
                $p = New-Object System.Windows.Documents.Paragraph
                $p.Margin = "0"

                foreach ($line in $text -split "`n", -1) {
                    $run = New-Object System.Windows.Documents.Run($line)
                    if ($line.TrimStart().StartsWith("#")) {
                        $run.Foreground = [System.Windows.Media.Brushes]::Gray
                    }
                    $p.Inlines.Add($run)
                    $p.Inlines.Add((New-Object System.Windows.Documents.LineBreak))
                }

                $rtb.Document.Blocks.Add($p)
                $rtb.CaretPosition = $caret
            } finally {
                $script:dllFormatting = $false
            }
        }.GetNewClosure()

        & $applyHighlight
        $rtb.Add_TextChanged({ & $applyHighlight })
        $c['btnRun'].Add_Click({
                Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] Ejecutando registro"
                if (-not (Test-Administrator)) {
                    Show-WpfMessageBox -Message "Esta acción requiere permisos de administrador." -Title "Permisos requeridos" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                $range = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
                $rawText = ($range.Text -replace "`r", "")
                $lines = $rawText -split "`n"
                $paths = @()
                $currentDir = $null
                foreach ($line in $lines) {
                    $trim = $line.Trim()
                    if ([string]::IsNullOrWhiteSpace($trim)) { continue }
                    if ($trim.StartsWith("#")) { continue }
                    if (Test-Path -LiteralPath $trim -PathType Container) {
                        $currentDir = $trim
                        continue
                    }
                    $candidate = $trim
                    if (-not [System.IO.Path]::IsPathRooted($candidate) -and $currentDir) { $candidate = Join-Path $currentDir $candidate }
                    $paths += $candidate
                }
                if (-not $paths -or $paths.Count -eq 0) {
                    Show-WpfMessageBox -Message "No se encontraron rutas válidas para procesar." -Title "Sin datos" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                $isUnregister = ($c['rbUnregister'].IsChecked -eq $true)
                $success = @()
                $errors = @()
                foreach ($path in $paths) {
                    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                        $errors += "No existe: $path"
                        continue
                    }
                    $args = if ($isUnregister) { "/s /u `"$path`"" } else { "/s `"$path`"" }
                    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] regsvr32 $args"
                    try {
                        $proc = Start-Process -FilePath "regsvr32.exe" -ArgumentList $args -NoNewWindow -Wait -PassThru
                        if ($proc.ExitCode -eq 0) { $success += $path } else { $errors += "Error ($($proc.ExitCode)) en: $path" }
                    } catch {
                        $errors += "Error en $path`: $($_.Exception.Message)"
                    }
                }
                $summary = @()
                if ($success.Count -gt 0) { $summary += "Procesados correctamente: $($success.Count)" }
                if ($errors.Count -gt 0) { $summary += "Errores: $($errors.Count)" }
                $detail = ($summary -join "`n")
                if ($errors.Count -gt 0) { $detail += "`n`nDetalles:`n$($errors -join "`n")" }
                $title = if ($errors.Count -gt 0) { "Proceso con errores" } else { "Proceso completado" }
                $icon = if ($errors.Count -gt 0) { "Warning" } else { "Information" }
                Show-WpfMessageBox -Message $detail -Title $title -Buttons OK -Icon $icon | Out-Null
            })
        $c['btnClose'].Add_Click({ $w.Close() })
        $w.ShowDialog() | Out-Null
    } catch {
        Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana de registro de DLLs." -Title "Error" -Buttons OK -Icon Error | Out-Null
    }
    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] FIN"
}
function Check-Permissions {
    [CmdletBinding()]
    param(
        [string]$folderPath = "C:\NationalSoft",
        [System.Windows.Window]$Owner
    )
    $uiInfo = {
        param([string]$msg, [string]$title = "Información")
        if (Get-Command Ui-Info -ErrorAction SilentlyContinue) { Ui-Info -Message $msg -Title $title -Owner $Owner; return }
        if (Get-Command Show-WpfMessageBox -ErrorAction SilentlyContinue) { Show-WpfMessageBox -Message $msg -Title $title -Buttons OK -Icon Information -Owner $Owner | Out-Null; return }
        Write-Host $msg -ForegroundColor Cyan
    }.GetNewClosure()
    $uiWarn = {
        param([string]$msg, [string]$title = "Atención")
        if (Get-Command Ui-Warn -ErrorAction SilentlyContinue) { Ui-Warn -Message $msg -Title $title -Owner $Owner; return }
        if (Get-Command Show-WpfMessageBox -ErrorAction SilentlyContinue) { Show-WpfMessageBox -Message $msg -Title $title -Buttons OK -Icon Warning -Owner $Owner | Out-Null; return }
        Write-Host $msg -ForegroundColor Yellow
    }.GetNewClosure()
    $uiError = {
        param([string]$msg, [string]$title = "Error")
        if (Get-Command Ui-Error -ErrorAction SilentlyContinue) { Ui-Error -Message $msg -Title $title -Owner $Owner; return }
        if (Get-Command Show-WpfMessageBox -ErrorAction SilentlyContinue) { Show-WpfMessageBox -Message $msg -Title $title -Buttons OK -Icon Error -Owner $Owner | Out-Null; return }
        Write-Host $msg -ForegroundColor Red
    }.GetNewClosure()
    $uiConfirm = {
        param([string]$msg, [string]$title = "Confirmar")
        if (Get-Command Ui-Confirm -ErrorAction SilentlyContinue) { return (Ui-Confirm -Message $msg -Title $title -Owner $Owner) }
        if (Get-Command Show-WpfMessageBox -ErrorAction SilentlyContinue) {
            return ((Show-WpfMessageBox -Message $msg -Title $title -Buttons YesNo -Icon Question -Owner $Owner) -eq [System.Windows.MessageBoxResult]::Yes)
        }
        return $false
    }.GetNewClosure()
    Write-DzDebug "`t[DEBUG][Check-Permissions] INICIO folderPath='$folderPath'"
    if (-not (Test-Path -LiteralPath $folderPath)) {
        $msg = "La carpeta '$folderPath' no existe en este equipo.`r`nCrea la carpeta o corrige la ruta antes de continuar."
        Write-Host $msg -ForegroundColor Red
        & $uiWarn $msg "Carpeta no encontrada"
        return
    }
    try {
        $directoryInfo = [System.IO.DirectoryInfo]::new($folderPath)
        $acl = $directoryInfo.GetAccessControl()
    } catch {
        $msg1 = $_.Exception.Message
        if ($msg1 -match "Get-Acl.*m[oó]dulo 'Microsoft\.PowerShell\.Security'.*no pudo cargarse" -or $msg1 -match "Microsoft\.PowerShell\.Security") {
            try {
                Import-Module Microsoft.PowerShell.Security -ErrorAction Stop | Out-Null
            } catch {
                $msg2 = $_.Exception.Message
                if ($msg2 -match "TypeData" -or $msg2 -match "Ya existe el miembro" -or $msg2 -match "extended type data") {
                    try {
                        Import-Module Microsoft.PowerShell.Security -DisableNameChecking -ErrorAction Stop | Out-Null
                    } catch {
                        $errMsg = "No se pudo cargar el módulo Microsoft.PowerShell.Security.`r`n`r`n$($_.Exception.Message)"
                        Write-Host "Error cargando Microsoft.PowerShell.Security: $($_.Exception.Message)" -ForegroundColor Red
                        & $uiError $errMsg "Error de PowerShell"
                        return
                    }
                } else {
                    $errMsg = "No se pudo cargar el módulo Microsoft.PowerShell.Security.`r`n`r`n$msg2"
                    Write-Host "Error cargando Microsoft.PowerShell.Security: $msg2" -ForegroundColor Red
                    & $uiError $errMsg "Error de PowerShell"
                    return
                }
            }
            try {
                $acl = Get-Acl -LiteralPath $folderPath -ErrorAction Stop
            } catch {
                $errMsg = "Error obteniendo permisos de '$folderPath':`r`n$($_.Exception.Message)"
                Write-Host "Error obteniendo ACL de $folderPath : $($_.Exception.Message)" -ForegroundColor Red
                & $uiError $errMsg "Error al obtener permisos"
                return
            }
        } else {
            $errMsg = "Error obteniendo permisos de '$folderPath':`r`n$msg1"
            Write-Host "Error obteniendo ACL de $folderPath : $msg1" -ForegroundColor Red
            & $uiError $errMsg "Error al obtener permisos"
            return
        }
    }
    $permissions = @()
    $everyoneSid = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::WorldSid, $null)
    $everyonePermissions = @()
    $everyoneHasFullControl = $false
    $rules = $acl.GetAccessRules($true, $true, [System.Security.Principal.NTAccount])
    foreach ($access in $rules) {
        $permissions += [PSCustomObject]@{
            Usuario = $access.IdentityReference.Value
            Permiso = $access.FileSystemRights
            Tipo    = $access.AccessControlType
        }
        if ($access.IdentityReference.Value -match '^(Everyone|Todos)$') {
            $everyonePermissions += $access.FileSystemRights
            if ($access.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl) { $everyoneHasFullControl = $true }
        }
    }
    Write-Host ""
    Write-Host "Permisos en $folderPath :" -ForegroundColor Cyan
    $permissions | ForEach-Object { Write-Host "`t$($_.Usuario) - $($_.Tipo) - $($_.Permiso)" -ForegroundColor Green }
    if ($everyonePermissions.Count -gt 0) { Write-Host "`tEveryone tiene: $($everyonePermissions -join ', ')" -ForegroundColor Green } else { Write-Host "`tNo hay permisos para 'Everyone'" -ForegroundColor Red }
    if (-not $everyoneHasFullControl) {
        $ask = "El usuario 'Everyone' no tiene permisos de 'Full Control'. ¿Deseas concederlo?"
        $doIt = & $uiConfirm $ask "Permisos 'Everyone'"
        if ($doIt) {
            try {
                if (-not (Test-Administrator)) {
                    & $uiWarn "Esta acción requiere permisos de administrador." "Permisos requeridos"
                    return
                }
                $directoryInfo = New-Object System.IO.DirectoryInfo($folderPath)
                $dirAcl = $directoryInfo.GetAccessControl()
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $everyoneSid,
                    [System.Security.AccessControl.FileSystemRights]::FullControl,
                    [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit",
                    [System.Security.AccessControl.PropagationFlags]::None,
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $dirAcl.AddAccessRule($accessRule)
                $dirAcl.SetAccessRuleProtection($false, $true)
                $directoryInfo.SetAccessControl($dirAcl)
                Write-Host "Se ha concedido 'Full Control' a 'Everyone'." -ForegroundColor Green
                & $uiInfo "Se ha concedido 'Full Control' a 'Everyone' en '$folderPath'." "Permisos actualizados"
            } catch {
                $errMsg = "Error aplicando permisos a '$folderPath':`r`n$($_.Exception.Message)"
                Write-Host "Error aplicando permisos: $($_.Exception.Message)" -ForegroundColor Red
                & $uiError $errMsg "Error al aplicar permisos"
            }
        }
    } else {
        $msg = "El usuario 'Everyone' ya tiene permisos de 'Full Control' en '$folderPath'."
        Write-Host $msg -ForegroundColor Green
        & $uiInfo $msg "Permisos OK"
    }
    Write-DzDebug "`t[DEBUG][Check-Permissions] FIN"
}
function Show-InstallerExtractorDialog {
    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] INICIO"
    $theme = Get-DzUiTheme

    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Extractor de instalador"
        Height="400" Width="720"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True"
        Topmost="True"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
        </Style>

        <Style x:Key="BaseButtonStyle" TargetType="Button">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="8"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="1"/>
                    <Setter Property="Cursor" Value="Arrow"/>
                    <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                    <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ActionButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentMagenta}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentMagentaHover}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                    <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                    <Setter Property="BorderThickness" Value="1"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="OutlineButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                    <Setter Property="BorderThickness" Value="0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="CloseButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Width" Value="34"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Content" Value="×"/>
        </Style>
    </Window.Resources>

    <Grid Background="{DynamicResource FormBg}" Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0"
                Name="brdTitleBar"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="12"
                Margin="0,0,0,10">
            <DockPanel LastChildFill="True">
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="Extractor de instalador"
                               Foreground="{DynamicResource FormFg}"
                               FontSize="16"
                               FontWeight="SemiBold"/>
                    <TextBlock Text="Seleccione el instalador y el destino de extracción."
                               Foreground="{DynamicResource PanelFg}"
                               Margin="0,2,0,0"/>
                </StackPanel>
                <Button DockPanel.Dock="Right"
                        Name="btnClose"
                        Style="{StaticResource CloseButtonStyle}"/>
            </DockPanel>
        </Border>

        <Border Grid.Row="1"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="12"
                Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="Instalador (.exe)" Margin="0,0,0,6"/>
                <Grid Grid.Row="1" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Name="btnPickInstaller"
                            Content="📁"
                            Width="42" Height="34"
                            Margin="0,0,8,0"
                            ToolTip="Seleccionar instalador"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <TextBox Name="txtInstallerPath"
                             Grid.Column="1"
                             Height="34"
                             IsReadOnly="True"
                             VerticalContentAlignment="Center"
                             Text=""/>
                </Grid>

                <Border Grid.Row="2"
                        Background="{DynamicResource FormBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}"
                        BorderThickness="1"
                        CornerRadius="8"
                        Padding="10"
                        Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Name="lblVersionInfo" Text="Versión: -"/>
                        <TextBlock Name="lblLastWrite" Text="Última modificación: -" Margin="0,4,0,0"/>
                    </StackPanel>
                </Border>

                <TextBlock Grid.Row="3" Text="Destino" Margin="0,0,0,6"/>
                <Grid Grid.Row="4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Name="btnPickDestination"
                            Content="📁"
                            Width="42" Height="34"
                            Margin="0,0,8,0"
                            ToolTip="Seleccionar destino"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <TextBox Name="txtDestinationPath"
                             Grid.Column="1"
                             Height="34"
                             IsReadOnly="False"
                             VerticalContentAlignment="Center"
                             Text=""/>
                </Grid>
            </Grid>
        </Border>
        <Border Grid.Row="2"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="10"
                Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0"
                           Text="Enter: Extraer   |   Esc: Cerrar"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button Name="btnCancel"
                            Content="Cancelar"
                            Width="120"
                            Height="34"
                            Margin="0,0,10,0"
                            IsCancel="True"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <Button Name="btnExtract"
                            Content="Extraer"
                            Width="140"
                            Height="34"
                            IsDefault="True"
                            Style="{StaticResource ActionButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>

    </Grid>
</Window>
"@

    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        $window = $ui.Window
        Set-DzWpfThemeResources -Window $window -Theme $theme
        try { Set-WpfDialogOwner -Dialog $window } catch {}
    } catch {
        Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana del extractor." -Title "Error" -Buttons OK -Icon Error | Out-Null
        return
    }

    $c = $ui.Controls

    $installerPath = $null
    $defaultFolderName = $null
    $destinationManuallySet = $false

    if ($c.ContainsKey('btnClose') -and $c['btnClose']) { $c['btnClose'].Add_Click({ $window.Close() }) }

    $brdTitleBar = $window.FindName("brdTitleBar")
    if ($brdTitleBar) {
        $brdTitleBar.Add_MouseLeftButtonDown({
                param($sender, $e)
                if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) {
                    try { $window.DragMove() } catch {}
                }
            })
    }

    $updateStatus = {
        param([string]$msg)
        if ($c.ContainsKey('txtStatus') -and $c['txtStatus']) { $c['txtStatus'].Text = $msg }
    }

    $updateInstallerInfo = {
        param([string]$path)
        $c['lblVersionInfo'].Text = "Versión: -"
        $c['lblLastWrite'].Text = "Última modificación: -"
        Set-Variable -Name defaultFolderName -Value $null -Scope 1
        if ([string]::IsNullOrWhiteSpace($path)) { return }
        try {
            $file = Get-Item -Path $path -ErrorAction Stop
            $fileInfo = (Get-ItemProperty $file.FullName).VersionInfo
            $creationDate = $file.LastWriteTime
            $formattedDate = $creationDate.ToString("yyMMdd")
            $versionText = if ($fileInfo.FileVersion) { $fileInfo.FileVersion } else { "N/D" }
            $c['lblVersionInfo'].Text = "Versión: $versionText"
            $c['lblLastWrite'].Text = "Última modificación: $($creationDate.ToString('dd/MM/yyyy HH:mm')) ($formattedDate)"
            $computedDefault = if ($versionText -ne "N/D") { "$versionText`_$formattedDate" } else { $formattedDate }
            Set-Variable -Name defaultFolderName -Value $computedDefault -Scope 1
            if (-not $destinationManuallySet) { $c['txtDestinationPath'].Text = Join-Path "C:\Temp" $computedDefault }
        } catch {
            Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR leyendo info del instalador: $($_.Exception.Message)" Red
            Show-WpfMessageBox -Message "No se pudo leer la información del instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
        }
    }

    $c['btnPickInstaller'].Add_Click({
            try {
                $dialog = New-Object Microsoft.Win32.OpenFileDialog
                $dialog.Filter = "Instalador (*.exe)|*.exe"
                $dialog.Title = "Seleccionar instalador"
                $dialog.Multiselect = $false
                if ($dialog.ShowDialog() -ne $true) { return }
                $selectedPath = $dialog.FileName
                if ([System.IO.Path]::GetExtension($selectedPath).ToLowerInvariant() -ne ".exe") {
                    Show-WpfMessageBox -Message "El instalador debe ser un archivo .EXE." -Title "Formato inválido" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                Set-Variable -Name installerPath -Value $selectedPath -Scope 1
                $c['txtInstallerPath'].Text = $selectedPath
                & $updateInstallerInfo $selectedPath
                & $updateStatus "Instalador seleccionado."
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR seleccionando instalador: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "Error al seleccionar el instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnPickDestination'].Add_Click({
            try {
                $initialDir = "C:\Temp"
                if (-not [string]::IsNullOrWhiteSpace($c['txtDestinationPath'].Text)) {
                    $initialDir = Split-Path -Path $c['txtDestinationPath'].Text -Parent
                }
                $selectedFolder = Show-WpfFolderDialog -Description "Seleccionar destino de extracción" -InitialDirectory $initialDir
                if (-not $selectedFolder) { return }
                Set-Variable -Name destinationManuallySet -Value $true -Scope 1
                if ($defaultFolderName) {
                    $c['txtDestinationPath'].Text = Join-Path $selectedFolder $defaultFolderName
                } else {
                    $c['txtDestinationPath'].Text = $selectedFolder
                }
                & $updateStatus "Destino seleccionado."
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR seleccionando destino: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "Error al seleccionar el destino." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnExtract'].Add_Click({
            try {
                if (-not $installerPath) {
                    Show-WpfMessageBox -Message "Seleccione un instalador primero." -Title "Falta instalador" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                if ([System.IO.Path]::GetExtension($installerPath).ToLowerInvariant() -ne ".exe") {
                    Show-WpfMessageBox -Message "El instalador debe ser un archivo .EXE." -Title "Formato inválido" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                $destinationPath = $c['txtDestinationPath'].Text.Trim()
                if ([string]::IsNullOrWhiteSpace($destinationPath)) {
                    Show-WpfMessageBox -Message "Seleccione un destino válido." -Title "Falta destino" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                if (-not (Test-Path -Path $destinationPath)) { New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null }
                $arguments = "/extract `"$destinationPath`""
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] Ejecutando: '$installerPath' $arguments"
                & $updateStatus "Extrayendo..."
                $proc = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
                if ($proc.ExitCode -ne 0) {
                    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ExitCode: $($proc.ExitCode)" Yellow
                    Show-WpfMessageBox -Message "El instalador devolvió código de salida $($proc.ExitCode)." -Title "Atención" -Buttons OK -Icon Warning | Out-Null
                    & $updateStatus "El instalador devolvió código $($proc.ExitCode)."
                    return
                }
                Show-WpfMessageBox -Message "Extracción completada en:`n$destinationPath" -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                try {
                    if (Test-Path -Path $destinationPath) { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$destinationPath`"" }
                } catch {
                    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] No se pudo abrir Explorer: $($_.Exception.Message)" Yellow
                }
                $window.Close()
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR extracción: $($_.Exception.Message)" Red
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] Stack: $($_.ScriptStackTrace)" Red
                Show-WpfMessageBox -Message "Error al extraer el instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnCancel'].Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] FIN"
}

function Invoke-CreateApk {
    [CmdletBinding()]
    param(
        [string]$DllPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\up.dll",
        [string]$InfoPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\info.txt",
        [System.Windows.Controls.TextBlock]$InfoTextBlock
    )
    Write-DzDebug "`t[DEBUG][Invoke-CreateApk] INICIO" ([System.ConsoleColor]::DarkGray)
    try {
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Validando componentes..." }
        Write-Host "`nIniciando proceso de creación de APK..." -ForegroundColor Cyan
        if (-not (Test-Path -LiteralPath $DllPath)) {
            $msg = "Componente necesario no encontrado:`n$DllPath`n`nVerifique la instalación del Enlace Android."
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe up.dll" }
            return $false
        }
        if (-not (Test-Path -LiteralPath $InfoPath)) {
            $msg = "Archivo de configuración no encontrado:`n$InfoPath`n`nVerifique la instalación del Enlace Android."
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe info.txt" }
            return $false
        }
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Leyendo versión..." }
        $jsonContent = Get-Content -LiteralPath $InfoPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $versionApp = [string]$jsonContent.versionApp
        if ([string]::IsNullOrWhiteSpace($versionApp)) {
            $msg = "No se pudo leer 'versionApp' desde:`n$InfoPath"
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: versionApp vacío" }
            return $false
        }
        Write-Host "Versión detectada: $versionApp" -ForegroundColor Green
        $confirmMsg = "Se creará el APK versión: $versionApp`n¿Desea continuar?"
        $confirmation = Show-WpfMessageBox -Message $confirmMsg -Title "Confirmación" -Buttons "YesNo" -Icon "Question"
        if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Host "Proceso cancelado por el usuario" -ForegroundColor Yellow
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Cancelado." }
            return $false
        }
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Selecciona dónde guardar..." }
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Archivo APK (*.apk)|*.apk"
        $saveDialog.FileName = "SRM_$versionApp.apk"
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($saveDialog.ShowDialog() -ne $true) {
            Write-Host "Guardado cancelado por el usuario" -ForegroundColor Yellow
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Guardado cancelado." }
            return $false
        }
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Copiando APK..." }
        Copy-Item -LiteralPath $DllPath -Destination $saveDialog.FileName -Force -ErrorAction Stop
        $okMsg = "APK generado exitosamente en:`n$($saveDialog.FileName)"
        Write-Host $okMsg -ForegroundColor Green
        Show-WpfMessageBox -Message "APK creado correctamente!" -Title "Éxito" -Buttons "OK" -Icon "Information" | Out-Null
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: APK creado." }
        return $true
    } catch {
        $err = $_.Exception.Message
        Write-Host "Error durante el proceso: $err" -ForegroundColor Red
        Write-DzDebug "`t[DEBUG][Invoke-CreateApk] ERROR: $err`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
        Show-WpfMessageBox -Message "Error durante la creación del APK:`n$err" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
        return $false
    }
}
function Invoke-CambiarOTMConfig {
    [CmdletBinding()]
    param(
        [string]$SyscfgPath = "C:\Windows\SysWOW64\Syscfg45_2.0.dll",
        [string]$IniPath = "C:\NationalSoft\OnTheMinute4.5",
        [System.Windows.Controls.TextBlock]$InfoTextBlock
    )
    Write-DzDebug "`t[DEBUG][Invoke-CambiarOTMConfig] INICIO" ([System.ConsoleColor]::DarkGray)
    try {
        if (-not (Test-Path -LiteralPath $SyscfgPath)) {
            Show-WpfMessageBox -Message "El archivo de configuración no existe:`n$SyscfgPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] Syscfg no existe: $SyscfgPath" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe Syscfg." }
            return $false
        }
        if (-not (Test-Path -LiteralPath $IniPath)) {
            Show-WpfMessageBox -Message "No existe la ruta de INIs:`n$IniPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] IniPath no existe: $IniPath" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe ruta INI." }
            return $false
        }
        $current = Get-OtmConfigFromSyscfg -SyscfgPath $SyscfgPath
        if (-not $current) {
            Show-WpfMessageBox -Message "No se detectó una configuración válida de SQL o DBF en:`n$SyscfgPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] No se detectó config válida" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: config inválida." }
            return $false
        }
        $ini = Get-OtmIniFiles -IniPath $IniPath
        if (-not $ini) {
            Show-WpfMessageBox -Message "No se encontraron los archivos INI esperados en:`n$IniPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] No se encontraron INIs esperados" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: faltan INIs." }
            return $false
        }
        $new = if ($current -eq "SQL") { "DBF" } else { "SQL" }
        $msg = "Actualmente tienes configurado: $current.`n¿Quieres cambiar a $new?"
        $res = Show-WpfMessageBox -Message $msg -Title "Cambiar Configuración" -Buttons "YesNo" -Icon "Question"
        if ($res -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-DzDebug "`t[DEBUG][OTM] Usuario canceló" ([System.ConsoleColor]::Cyan)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Operación cancelada." }
            return $false
        }
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Aplicando cambios..." }
        Set-OtmSyscfgConfig -SyscfgPath $SyscfgPath -Target $new
        Rename-OtmIniForTarget -Target $new -IniSqlFile $ini.SQL -IniDbfFile $ini.DBF
        Show-WpfMessageBox -Message "Configuración cambiada exitosamente a $new." -Title "Éxito" -Buttons "OK" -Icon "Information" | Out-Null
        Write-DzDebug "`t[DEBUG][OTM] OK cambiado a $new" ([System.ConsoleColor]::Green)
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: cambiado a $new." }
        return $true
    } catch {
        $err = $_.Exception.Message
        Write-DzDebug "`t[DEBUG][Invoke-CambiarOTMConfig] ERROR: $err`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
        Show-WpfMessageBox -Message "Error: $err" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
        return $false
    }
}
function Show-NSApplicationsIniReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Resultados
    )
    $columnas = @("Aplicacion", "INI", "DataSource", "Catalog", "Usuario")
    $anchos = @{}
    foreach ($col in $columnas) { $anchos[$col] = $col.Length }
    foreach ($res in $Resultados) {
        foreach ($col in $columnas) {
            $val = [string]$res.$col
            if ($val.Length -gt $anchos[$col]) { $anchos[$col] = $val.Length }
        }
    }
    $titulos = $columnas | ForEach-Object { $_.PadRight($anchos[$_] + 2) }
    Write-Host ($titulos -join "") -ForegroundColor Cyan
    $separador = $columnas | ForEach-Object { ("-" * $anchos[$_]).PadRight($anchos[$_] + 2) }
    Write-Host ($separador -join "") -ForegroundColor Cyan
    foreach ($res in $Resultados) {
        $fila = $columnas | ForEach-Object { ([string]$res.$_).PadRight($anchos[$_] + 2) }
        if ($res.INI -eq "No encontrado") { Write-Host ($fila -join "") -ForegroundColor Red } else { Write-Host ($fila -join "") }
    }
    $found = @($Resultados | Where-Object { $_.INI -ne "No encontrado" })
    if (-not $found -or $found.Count -eq 0) { return }
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Aplicaciones National Soft"
        Height="400" Width="860"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True"
        Topmost="False"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="BaseButtonStyle" TargetType="Button">
      <Setter Property="OverridesDefaultStyle" Value="True"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="1"/>
          <Setter Property="Cursor" Value="Arrow"/>
          <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="GeneralButtonStyle" TargetType="Button">
      <Setter Property="OverridesDefaultStyle" Value="True"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter Property="Background" Value="{DynamicResource AccentMuted}"/>
          <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
          <Setter Property="BorderThickness" Value="0"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="1"/>
          <Setter Property="Cursor" Value="Arrow"/>
          <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="ActionButtonStyle" TargetType="Button">
      <Setter Property="OverridesDefaultStyle" Value="True"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Background" Value="{DynamicResource AccentMagenta}"/>
      <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    CornerRadius="8"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{DynamicResource AccentMagentaHover}"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="1"/>
          <Setter Property="Cursor" Value="Arrow"/>
          <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
          <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
          <Setter Property="BorderThickness" Value="1"/>
          <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="OutlineButtonStyle" TargetType="Button">
      <Setter Property="OverridesDefaultStyle" Value="True"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
          <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
          <Setter Property="BorderThickness" Value="0"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="1"/>
          <Setter Property="Cursor" Value="Arrow"/>
          <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="{DynamicResource PanelBg}"/>
      <Setter Property="AlternatingRowBackground" Value="{DynamicResource ControlBg}"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="VerticalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="CanUserResizeRows" Value="False"/>
      <Setter Property="CanUserSortColumns" Value="True"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
          <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Padding" Value="8,0,8,0"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>
  <Border Background="{DynamicResource FormBg}"
          BorderBrush="{DynamicResource BorderBrushColor}"
          BorderThickness="1"
          CornerRadius="12"
          Padding="12"
          Margin="10">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Grid.Row="0"
              Name="brdTitleBar"
              Background="{DynamicResource PanelBg}"
              BorderBrush="{DynamicResource BorderBrushColor}"
              BorderThickness="1"
              CornerRadius="10"
              Padding="10"
              Margin="0,0,0,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock Name="txtTitleBar"
                     Text="Aplicaciones National Soft"
                     Foreground="{DynamicResource FormFg}"
                     FontWeight="SemiBold"
                     FontSize="14"
                     VerticalAlignment="Center"
                     IsHitTestVisible="False"/>
          <Button Grid.Column="1"
                  Name="btnClose"
                  Content="✕"
                  Width="34"
                  Height="28"
                  Margin="10,0,0,0"
                  Background="{DynamicResource ControlBg}"
                  Foreground="{DynamicResource ControlFg}"
                  BorderBrush="{DynamicResource BorderBrushColor}"
                  BorderThickness="1"
                  Cursor="Hand"/>
        </Grid>
      </Border>
      <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,8">
        <Button Name="btnCopyApp" Content="Copiar Aplicación" Width="140" Height="30" Margin="0,0,8,0" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Name="btnCopyIni" Content="Copiar INI" Width="110" Height="30" Margin="0,0,8,0" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Name="btnCopyDataSource" Content="Copiar DataSource" Width="150" Height="30" Margin="0,0,8,0" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Name="btnCopyCatalog" Content="Copiar Catalog" Width="130" Height="30" Margin="0,0,8,0" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Name="btnCopyUsuario" Content="Copiar Usuario" Width="130" Height="30" Style="{StaticResource GeneralButtonStyle}"/>
      </StackPanel>
      <DataGrid Grid.Row="2"
                Name="dgApps"
                AutoGenerateColumns="False"
                IsReadOnly="True"
                CanUserAddRows="False"
                SelectionMode="Single">
        <DataGrid.Columns>
          <DataGridTextColumn Header="Aplicación" Binding="{Binding Aplicacion}"/>
          <DataGridTextColumn Header="INI" Binding="{Binding INI}"/>
          <DataGridTextColumn Header="DataSource" Binding="{Binding DataSource}"/>
          <DataGridTextColumn Header="Catalog" Binding="{Binding Catalog}"/>
          <DataGridTextColumn Header="Usuario" Binding="{Binding Usuario}"/>
        </DataGrid.Columns>
      </DataGrid>
    </Grid>
  </Border>
</Window>
"@
    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        $w = $ui.Window
        $c = $ui.Controls
        Set-DzWpfThemeResources -Window $w -Theme $theme
        $brdTitleBar = $w.FindName("brdTitleBar")
        $btnClose = $w.FindName("btnClose")
        if ($btnClose) {
            $winRef = $w
            $btnClose.Add_Click({
                    try {
                        if ($winRef -and ($winRef -is [System.Windows.Window])) { $winRef.Dispatcher.Invoke([action] { $winRef.Close() }) }
                    } catch {
                        try { $winRef.Close() } catch {}
                    }
                }.GetNewClosure())
        }
        $drag = {
            param($sender, $e)
            if ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
                try { $e.Handled = $true; $w.DragMove() } catch {}
            }
        }.GetNewClosure()
        if ($brdTitleBar) { $brdTitleBar.Add_MouseLeftButtonDown($drag) }
        $tbGrid = $brdTitleBar.Child
        if ($tbGrid) { try { $tbGrid.Add_MouseLeftButtonDown($drag) } catch {} }
        try { Set-WpfDialogOwner -Dialog $w } catch {}
        $c['dgApps'].ItemsSource = $found
        try {
            $culture = [System.Globalization.CultureInfo]::CurrentCulture
            $flow = [System.Windows.FlowDirection]::LeftToRight
            $typeface = New-Object System.Windows.Media.Typeface($w.FontFamily, $w.FontStyle, $w.FontWeight, $w.FontStretch)
            $measure = {
                param([string]$text)
                if ($null -eq $text) { $text = "" }
                $ft = New-Object System.Windows.Media.FormattedText($text, $culture, $flow, $typeface, [double]$w.FontSize, [System.Windows.Media.Brushes]::Black)
                return $ft.WidthIncludingTrailingWhitespace
            }.GetNewClosure()
            foreach ($col in $c['dgApps'].Columns) {
                $path = $null
                try { $path = $col.Binding.Path.Path } catch {}
                $header = [string]$col.Header
                $max = & $measure $header
                if ($path) {
                    foreach ($item in $found) {
                        $v = ""
                        try { $v = [string]$item.$path } catch { $v = "" }
                        $w1 = & $measure $v
                        if ($w1 -gt $max) { $max = $w1 }
                    }
                }
                $target = [Math]::Ceiling($max + 28)
                if ($target -lt 70) { $target = 70 }
                if ($target -gt 520) { $target = 520 }
                $col.Width = New-Object System.Windows.Controls.DataGridLength($target)
            }
        } catch {}
        $copyColumn = {
            param($name)
            $values = $found | ForEach-Object { [string]$_.($name) }
            Set-ClipboardTextSafe -Text ($values -join "`r`n") -Owner $w | Out-Null
        }.GetNewClosure()
        $c['btnCopyApp'].Add_Click({ & $copyColumn 'Aplicacion' })
        $c['btnCopyIni'].Add_Click({ & $copyColumn 'INI' })
        $c['btnCopyDataSource'].Add_Click({ & $copyColumn 'DataSource' })
        $c['btnCopyCatalog'].Add_Click({ & $copyColumn 'Catalog' })
        $c['btnCopyUsuario'].Add_Click({ & $copyColumn 'Usuario' })
        $w.Show() | Out-Null
    } catch {
        Write-DzDebug "`t[DEBUG][Show-NSApplicationsIniReport] ERROR creando ventana: $($_.Exception.Message)" Red
    }
}

Export-ModuleMember -Function @('Show-DllRegistrationDialog', 'Check-Permissions', 'Show-InstallerExtractorDialog', 'Invoke-CreateApk', 'Invoke-CambiarOTMConfig', 'Show-NSApplicationsIniReport')