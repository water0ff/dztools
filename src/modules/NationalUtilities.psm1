#requires -Version 5.0
function Show-DllRegistrationDialog {
    [CmdletBinding()]
    param()
    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] INICIO"
    $theme = Get-DzUiTheme
    $defaultList = @(
        "#Asi se ve uno deshabilitado: C:\Windows\SysWOW64\slicensing.dll"
        "C:\Windows\SysWOW64\slicensing.dll"
        "C:\Windows\SysWOW64\slicensingr1.dll"
        "C:\Windows\SysWOW64\sservices1_0.dll"
        "C:\Windows\SysWOW64\sservices1_0r1.dll"
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
                $caretOffset = $rtb.Document.ContentStart.GetOffsetToPosition($rtb.CaretPosition)
                $rawRange = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
                $rawText = ($rawRange.Text -replace "`r", "")
                $rtb.Document.Blocks.Clear()
                foreach ($line in $rawText -split "`n", -1) {
                    $paragraph = New-Object System.Windows.Documents.Paragraph
                    $paragraph.Margin = "0"
                    $run = New-Object System.Windows.Documents.Run($line)
                    if ($line.TrimStart().StartsWith("#")) { $run.Foreground = [System.Windows.Media.Brushes]::Gray }
                    $paragraph.Inlines.Add($run)
                    $rtb.Document.Blocks.Add($paragraph)
                }
                $pointer = & $GetTextPointerFromOffset $rtb.Document $caretOffset
                if ($pointer) { $rtb.CaretPosition = $pointer }
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
                        $proc = Start-Process -FilePath "regsvr32.exe" -ArgumentList $args -NoNewWindow -Wait -PassThru -WindowStyle Hidden
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
    $rules = $acl.GetAccessRules(
        $true,   # includeExplicit
        $true,   # includeInherited
        [System.Security.Principal.NTAccount]
    )
    foreach ($access in $rules) {
        $permissions += [PSCustomObject]@{
            Usuario = $access.IdentityReference.Value
            Permiso = $access.FileSystemRights
            Tipo    = $access.AccessControlType
        }
        if ($access.IdentityReference.Value -match '^(Everyone|Todos)$') {
            $everyonePermissions += $access.FileSystemRights
            if ($access.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl) {
                $everyoneHasFullControl = $true
            }
        }
    }
    Write-Host ""
    Write-Host "Permisos en $folderPath :" -ForegroundColor Cyan
    $permissions | ForEach-Object { Write-Host "`t$($_.Usuario) - $($_.Tipo) - $($_.Permiso)" -ForegroundColor Green }
    if ($everyonePermissions.Count -gt 0) {
        Write-Host "`tEveryone tiene: $($everyonePermissions -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "`tNo hay permisos para 'Everyone'" -ForegroundColor Red
    }
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
        Height="360" Width="640"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="GeneralButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonGeneralBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonGeneralForeground)"/>
        </Style>
        <Style x:Key="NationalSoftButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
            <Setter Property="Foreground" Value="$($theme.ButtonNationalForeground)"/>
        </Style>
    </Window.Resources>
    <Border Background="{DynamicResource FormBg}"
            CornerRadius="10"
            BorderBrush="{DynamicResource AccentPrimary}"
            BorderThickness="2"
            Padding="0">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
        </Border.Effect>

        <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="36"/>   <!-- Header -->
            <RowDefinition Height="Auto"/> <!-- Instrucción -->
            <RowDefinition Height="Auto"/> <!-- Label instalador -->
            <RowDefinition Height="Auto"/> <!-- picker instalador -->
            <RowDefinition Height="Auto"/> <!-- info instalador -->
            <RowDefinition Height="Auto"/> <!-- label destino -->
            <RowDefinition Height="Auto"/> <!-- picker destino -->
            <RowDefinition Height="Auto"/> <!-- botones -->
        </Grid.RowDefinitions>
            <!-- Header custom (arrastrable) -->
            <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Extractor de instalador"
                           VerticalAlignment="Center"
                           FontWeight="SemiBold"/>

                <Button Name="btnClose"
                        Grid.Column="1"
                        Content="✕"
                        Width="34" Height="26"
                        Margin="8,0,0,0"
                        ToolTip="Cerrar"
                        Background="Transparent"
                        BorderBrush="Transparent"/>
            </Grid>

            <TextBlock Grid.Row="1"
                       Text="Seleccione el instalador y el destino de extracción."
                       FontWeight="SemiBold"
                       Margin="0,0,0,12"/>

            <TextBlock Grid.Row="2" Text="Instalador (.exe)" Margin="0,0,0,6"/>
            <Grid Grid.Row="3" Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnPickInstaller"
                        Content="📁"
                        Width="36" Height="32"
                        Margin="0,0,8,0"
                        ToolTip="Seleccionar instalador"
                        Style="{StaticResource GeneralButtonStyle}"/>
                <TextBox Name="txtInstallerPath"
                         Grid.Column="1"
                         Height="32"
                         IsReadOnly="True"
                         VerticalContentAlignment="Center"
                         Text=""/>
            </Grid>

            <StackPanel Grid.Row="4" Margin="0,0,0,12">
                <TextBlock Name="lblVersionInfo" Text="Versión: -"/>
                <TextBlock Name="lblLastWrite" Text="Última modificación: -" Margin="0,4,0,0"/>
            </StackPanel>

            <TextBlock Grid.Row="5" Text="Destino" Margin="0,0,0,6"/>
            <Grid Grid.Row="6" Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnPickDestination"
                        Content="📁"
                        Width="36" Height="32"
                        Margin="0,0,8,0"
                        ToolTip="Seleccionar destino"
                        Style="{StaticResource GeneralButtonStyle}"/>
                <TextBox Name="txtDestinationPath"
                         Grid.Column="1"
                         Height="32"
                         IsReadOnly="False"
                         VerticalContentAlignment="Center"
                         Text=""/>
            </Grid>

            <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnCancel" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0" IsCancel="True" Style="{StaticResource GeneralButtonStyle}"/>
                <Button Name="btnExtract" Content="Extraer" Width="110" Height="30" Style="{StaticResource NationalSoftButtonStyle}"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@

    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        Set-DzWpfThemeResources -Window $ui.Window -Theme $theme
    } catch {
        Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana del extractor." -Title "Error" -Buttons OK -Icon Error | Out-Null
        return
    }

    $window = $ui.Window
    $c = $ui.Controls
    if ($c.ContainsKey('btnClose') -and $c['btnClose']) {
        $c['btnClose'].Add_Click({ $window.Close() })
    }
    if ($c.ContainsKey('HeaderBar') -and $c['HeaderBar']) {
        $c['HeaderBar'].Add_MouseLeftButtonDown({
                if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
                    $window.DragMove()
                }
            })
    }
    try {
        if ($Global:window -is [System.Windows.Window]) {
            $window.Owner = $Global:window
        }
    } catch {
        Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] No se pudo asignar owner: $($_.Exception.Message)" Yellow
    }

    $installerPath = $null
    $defaultFolderName = $null
    $destinationManuallySet = $false

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

            $computedDefault = if ($versionText -ne "N/D") {
                "$versionText`_$formattedDate"
            } else {
                $formattedDate
            }

            Set-Variable -Name defaultFolderName -Value $computedDefault -Scope 1

            if (-not $destinationManuallySet) {
                $c['txtDestinationPath'].Text = Join-Path "C:\Temp" $computedDefault
            }
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

                # IMPORTANTE: guardar en el scope del function
                Set-Variable -Name installerPath -Value $selectedPath -Scope 1

                $c['txtInstallerPath'].Text = $selectedPath
                & $updateInstallerInfo $selectedPath
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

                if (-not (Test-Path -Path $destinationPath)) {
                    New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                }

                $arguments = "/extract `"$destinationPath`""
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] Ejecutando: '$installerPath' $arguments"
                $proc = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop

                if ($proc.ExitCode -ne 0) {
                    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ExitCode: $($proc.ExitCode)" Yellow
                    Show-WpfMessageBox -Message "El instalador devolvió código de salida $($proc.ExitCode)." -Title "Atención" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                Show-WpfMessageBox -Message "Extracción completada en:`n$destinationPath" -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                # Abrir la carpeta destino en el explorador
                try {
                    if (Test-Path -Path $destinationPath) {
                        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$destinationPath`""
                    }
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

    $c['btnCancel'].Add_Click({
            $window.Close()
        })

    $window.ShowDialog() | Out-Null
    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] FIN"
}
Export-ModuleMember -Function @('Show-DllRegistrationDialog', 'Check-Permissions', 'Show-InstallerExtractorDialog')