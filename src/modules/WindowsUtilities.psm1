#requires -Version 5.0
function Start-SystemUpdate {
    param([int]$DiskCleanupTimeoutMinutes = 3)
    $progressWindow = $null
    Write-DzDebug "`t[DEBUG]Start-SystemUpdate: INICIO (Timeout=$DiskCleanupTimeoutMinutes min)"
    try {
        $progressWindow = Show-WpfProgressBar -Title "Actualización del Sistema" -Message "Iniciando proceso..."
        $totalSteps = 5
        $currentStep = 0
        Write-Host "`nIniciando proceso de actualización..." -ForegroundColor Cyan
        Write-DzDebug "`t[DEBUG]Paso 1: Deteniendo servicio winmgmt..."
        Write-Host "`n[Paso 1/$totalSteps] Deteniendo servicio winmgmt..." -ForegroundColor Yellow
        $service = Get-Service -Name "winmgmt" -ErrorAction Stop
        Write-DzDebug "`t[DEBUG]Paso 1: Estado actual winmgmt=$($service.Status)"
        if ($service.Status -eq "Running") {
            Update-WpfProgressBar -Window $progressWindow -Percent 10 -Message "Deteniendo servicio WMI..."
            Stop-Service -Name "winmgmt" -Force -ErrorAction Stop
            Write-DzDebug "`t[DEBUG]Paso 1: winmgmt detenido OK"
            Write-Host "`n`tServicio detenido correctamente." -ForegroundColor Green
        }
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 20 -Message "Servicio WMI detenido"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 2: Renombrando carpeta Repository..."
        Write-Host "`n[Paso 2/$totalSteps] Renombrando carpeta Repository..." -ForegroundColor Yellow
        Update-WpfProgressBar -Window $progressWindow -Percent 30 -Message "Renombrando Repository..."
        try {
            $repoPath = Join-Path $env:windir "System32\Wbem\Repository"
            Write-DzDebug "`t[DEBUG]Paso 2: repoPath=$repoPath"
            if (Test-Path $repoPath) {
                $newName = "Repository_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Rename-Item -Path $repoPath -NewName $newName -Force -ErrorAction Stop
                Write-DzDebug "`t[DEBUG]Paso 2: Carpeta renombrada a $newName"
                Write-Host "`n`tCarpeta renombrada: $newName" -ForegroundColor Green
            }
        } catch {
            Write-DzDebug "`t[DEBUG]Paso 2: EXCEPCIÓN: $($_.Exception.Message)" Red
            Write-Host "`n`tAdvertencia: No se pudo renombrar Repository. Continuando..." -ForegroundColor Yellow
        }
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 40 -Message "Repository renovado"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 3: Reiniciando servicio winmgmt..."
        Write-Host "`n[Paso 3/$totalSteps] Reiniciando servicio winmgmt..." -ForegroundColor Yellow
        Update-WpfProgressBar -Window $progressWindow -Percent 50 -Message "Reiniciando servicio WMI..."
        net start winmgmt *>&1 | Write-Host
        Write-DzDebug "`t[DEBUG]Paso 3: winmgmt reiniciado"
        Write-Host "`n`tServicio WMI reiniciado." -ForegroundColor Green
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 60 -Message "Servicio WMI reiniciado"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 4: Limpiando archivos temporales..."
        Write-Host "`n[Paso 4/$totalSteps] Limpiando archivos temporales..." -ForegroundColor Cyan
        Update-WpfProgressBar -Window $progressWindow -Percent 70 -Message "Limpiando archivos temporales..."
        $cleanupResult = Clear-TemporaryFiles
        Write-DzDebug "`t[DEBUG]Paso 4: FilesDeleted=$($cleanupResult.FilesDeleted) SpaceFreedMB=$($cleanupResult.SpaceFreedMB)"
        Write-Host "`n`tTotal archivos eliminados: $($cleanupResult.FilesDeleted)" -ForegroundColor Green
        Write-Host "`n`tEspacio liberado: $($cleanupResult.SpaceFreedMB) MB" -ForegroundColor Green
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 80 -Message "Archivos temporales limpiados"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 5: Ejecutando Liberador de espacio..."
        Write-Host "`n[Paso 5/$totalSteps] Ejecutando Liberador de espacio..." -ForegroundColor Cyan
        Update-WpfProgressBar -Window $progressWindow -Percent 90 -Message "Preparando limpieza de disco..."
        Invoke-DiskCleanup -Wait -TimeoutMinutes $DiskCleanupTimeoutMinutes -ProgressWindow $progressWindow
        Write-DzDebug "`t[DEBUG]Paso 5: Liberador completado"
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 100 -Message "Proceso completado exitosamente"
        Start-Sleep -Seconds 1
        if ($progressWindow -ne $null -and $progressWindow.IsVisible) {
            Close-WpfProgressBar -Window $progressWindow
            $progressWindow = $null
        }
        Write-Host "`n`n============================================" -ForegroundColor Green
        Write-Host "   Proceso de actualización completado" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "`nSe recomienda REINICIAR el equipo" -ForegroundColor Yellow
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Mostrando diálogo de reinicio"
        $result = [System.Windows.MessageBox]::Show("El proceso de actualización se completó exitosamente.`n`n" + "Se recomienda REINICIAR el equipo para completar la actualización del sistema WMI.`n`n" + "¿Desea reiniciar ahora?", "Actualización completada", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Write-Host "`n`tReiniciando equipo en 10 segundos..." -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Usuario eligió reiniciar"
            Start-Sleep -Seconds 3
            shutdown /r /t 10 /c "Reinicio para completar actualización de sistema WMI"
        } else {
            Write-Host "`n`tRecuerde reiniciar el equipo más tarde." -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Usuario canceló reinicio"
        }
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: FIN OK"
        return $true
    } catch {
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: EXCEPCIÓN: $($_.Exception.Message)" Red
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: ScriptStackTrace: $($_.ScriptStackTrace)" Red
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show("Error durante la actualización: $($_.Exception.Message)`n`n" + "Revise los logs y considere reiniciar manualmente el equipo.", "Error en actualización", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $false
    } finally {
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: FINALLY (cerrando progressWindow)"
        if ($progressWindow -ne $null -and $progressWindow.IsVisible) { Close-WpfProgressBar -Window $progressWindow }
    }
}
function Show-SystemComponents {
    param([switch]$SkipOnError)
    Write-Host "`n=== Componentes del sistema detectados ===" -ForegroundColor Cyan
    $os = $null
    $maxAttempts = 3
    $retryDelaySeconds = 2
    for ($attempt = 1; $attempt -le $maxAttempts -and -not $os; $attempt++) {
        try { $os = Get-CimInstance -ClassName CIM_OperatingSystem -ErrorAction Stop }catch {
            if ($attempt -lt $maxAttempts) {
                $msg = "Show-SystemComponents: ERROR intento {0}: {1}. Reintento en {2}s" -f $attempt, $_.Exception.Message, $retryDelaySeconds
                Write-Host $msg -ForegroundColor DarkYellow
                Start-Sleep -Seconds $retryDelaySeconds
            } else {
                Write-Host "`n[Windows]" -ForegroundColor Yellow
                Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
                if (-not $SkipOnError) { throw "No se pudo obtener información crítica del sistema" }else { Write-Host "Continuando sin información del sistema..." -ForegroundColor Yellow; return }
            }
        }
    }
    if (-not $os) {
        if (-not $SkipOnError) { throw "No se pudo obtener información crítica del sistema" }
        return
    }
    Write-Host "`n[Windows]" -ForegroundColor Yellow
    Write-Host "Versión: $($os.Caption) (Build $($os.Version))" -ForegroundColor White
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_Processor..."
        $procesador = Get-CimInstance -ClassName CIM_Processor -ErrorAction Stop
        Write-Host "`n[Procesador]" -ForegroundColor Yellow
        Write-Host "Modelo: $($procesador.Name)" -ForegroundColor White
        Write-Host "Núcleos: $($procesador.NumberOfCores)" -ForegroundColor White
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo procesador" Red
        Write-Host "`n[Procesador]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_PhysicalMemory..."
        $memoria = Get-CimInstance -ClassName CIM_PhysicalMemory -ErrorAction Stop
        Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
        $memoria | ForEach-Object { Write-Host "Módulo: $([math]::Round($_.Capacity/1GB,2)) GB $($_.Manufacturer) ($($_.Speed) MHz)" -ForegroundColor White }
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo memoria" Red
        Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_DiskDrive..."
        $discos = Get-CimInstance -ClassName CIM_DiskDrive -ErrorAction Stop
        Write-Host "`n[Discos duros]" -ForegroundColor Yellow
        $discos | ForEach-Object { Write-Host "Disco: $($_.Model) ($([math]::Round($_.Size/1GB,2)) GB)" -ForegroundColor White }
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo discos" Red
        Write-Host "`n[Discos duros]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-DzDebug "`t[DEBUG]Show-SystemComponents: FIN"
}
function Show-NSPrinters {
    [CmdletBinding()]param()
    Write-Host "`nImpresoras disponibles en el sistema:"
    $printers = Get-NSPrinters
    if (-not $printers -or $printers.Count -eq 0) { Write-Host "`nNo se encontraron impresoras."; return }
    $view = $printers | ForEach-Object {
        [PSCustomObject]@{
            Name       = ([string]$_.Name).Substring(0, [Math]::Min(24, ([string]$_.Name).Length))
            PortName   = ([string]$_.PortName).Substring(0, [Math]::Min(19, ([string]$_.PortName).Length))
            DriverName = ([string]$_.DriverName).Substring(0, [Math]::Min(19, ([string]$_.DriverName).Length))
            IsShared   = if ($_.Shared) { "Sí" }else { "No" }
        }
    }
    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "Nombre", "Puerto", "Driver", "Compartida")
    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "------", "------", "------", "---------")
    $view | ForEach-Object { Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f $_.Name, $_.PortName, $_.DriverName, $_.IsShared) }
}
function Invoke-ClearPrintJobs {
    [CmdletBinding()]param([System.Windows.Controls.TextBlock]$InfoTextBlock)
    try {
        if (-not (Test-Administrator)) {
            Show-WpfMessageBox -Message "Esta acción requiere permisos de administrador.`nPor favor, ejecuta Gerardo Zermeño Tools como administrador." -Title "Permisos insuficientes" -Buttons "OK" -Icon "Warning" | Out-Null
            return $false
        }
        $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
        if (-not $spooler) {
            Show-WpfMessageBox -Message "No se encontró el servicio 'Cola de impresión (Spooler)' en este equipo." -Title "Servicio no encontrado" -Buttons "OK" -Icon "Error" | Out-Null
            return $false
        }
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Limpiando trabajos de impresión..." }
        try {
            Get-Printer -ErrorAction Stop | ForEach-Object {
                try { Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue }catch { Write-Host "`tNo se pudieron limpiar trabajos de '$($_.Name)': $($_.Exception.Message)" -ForegroundColor Yellow }
            }
        } catch {
            Write-Host "`tNo se pudieron enumerar impresoras (Get-Printer). ¿Está instalado el módulo PrintManagement?" -ForegroundColor Yellow
        }
        if ($spooler.Status -eq 'Running') {
            Write-Host "`tDeteniendo servicio Spooler..." -ForegroundColor DarkYellow
            Stop-Service -Name Spooler -Force -ErrorAction Stop
        } else {
            Write-Host "`tSpooler no está en 'Running' (estado actual: $($spooler.Status))." -ForegroundColor DarkYellow
        }
        $spooler.Refresh()
        if ($spooler.StartType -eq 'Disabled') {
            Show-WpfMessageBox -Message "El servicio 'Cola de impresión (Spooler)' está DESHABILITADO.`nHabilítalo manualmente desde services.msc para poder iniciarlo." -Title "Spooler deshabilitado" -Buttons "OK" -Icon "Warning" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Spooler deshabilitado." }
            return $false
        }
        Write-Host "`tIniciando servicio Spooler..." -ForegroundColor DarkYellow
        Start-Service -Name Spooler -ErrorAction Stop
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: cola de impresión reiniciada." }
        Show-WpfMessageBox -Message "Los trabajos de impresión han sido eliminados y el servicio de cola de impresión se reinició correctamente." -Title "Operación exitosa" -Buttons "OK" -Icon "Information" | Out-Null
        return $true
    } catch {
        $err = $_.Exception.Message
        Write-Host "`n[ERROR Invoke-ClearPrintJobs] $err" -ForegroundColor Red
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
        Show-WpfMessageBox -Message "Ocurrió un error al intentar limpiar impresoras o reiniciar el servicio:`n$err" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
        return $false
    }
}
function Show-AddUserDialog {
    Write-DzDebug "`t[DEBUG][Show-AddUserDialog] INICIO"
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Crear Usuario de Windows" Height="420" Width="640" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" ShowInTaskbar="False" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <Style TargetType="TextBlock"><Setter Property="Foreground" Value="$($theme.FormForeground)"/></Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="PasswordBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="RadioButton"><Setter Property="Foreground" Value="$($theme.FormForeground)"/></Style>
    <Style TargetType="ToggleButton">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Style.Triggers>
        <Trigger Property="IsChecked" Value="True">
          <Setter Property="Background" Value="$($theme.AccentPrimary)"/>
          <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="SystemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
    </Style>
  </Window.Resources>
  <Border Background="{DynamicResource FormBg}" CornerRadius="10" BorderBrush="{DynamicResource AccentPrimary}" BorderThickness="2" Padding="0">
    <Border.Effect><DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/></Border.Effect>
    <Grid Margin="16">
      <Grid.RowDefinitions>
        <RowDefinition Height="36"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Text="Crear Usuario de Windows" VerticalAlignment="Center" FontWeight="SemiBold"/>
        <Button Name="btnClose" Grid.Column="1" Content="✕" Width="34" Height="26" Margin="8,0,0,0" ToolTip="Cerrar" Background="Transparent" BorderBrush="Transparent"/>
      </Grid>
      <TextBlock Grid.Row="1" Text="Crea un usuario local y asígnalo al grupo correspondiente." FontWeight="SemiBold" Margin="0,0,0,12"/>
      <Grid Grid.Row="2" Margin="0,0,0,12">
        <Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Nombre de usuario" VerticalAlignment="Center" Margin="0,0,10,8"/>
        <TextBox Name="txtUsername" Grid.Row="0" Grid.Column="1" Height="32" VerticalContentAlignment="Center" Margin="0,0,0,8"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Contraseña" VerticalAlignment="Center" Margin="0,0,10,8"/>
        <Grid Grid.Row="1" Grid.Column="1" Margin="0,0,0,8">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <PasswordBox Name="pwdPassword" Grid.Column="0" Height="32" Padding="6,0,6,0"/>
          <TextBox Name="txtPasswordVisible" Grid.Column="0" Height="32" VerticalContentAlignment="Center" Visibility="Collapsed"/>
          <ToggleButton Name="tglShowPassword" Grid.Column="1" Content="👁" Width="40" Height="32" Margin="8,0,0,0" ToolTip="Mostrar/Ocultar contraseña"/>
        </Grid>
        <TextBlock Grid.Row="2" Grid.Column="0" Text="Tipo de usuario" VerticalAlignment="Center"/>
        <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <RadioButton Name="rbStandard" Content="Usuario estándar" IsChecked="True" Margin="0,0,12,0"/>
          <RadioButton Name="rbAdmin" Content="Administrador"/>
        </StackPanel>
        <Button Name="btnShowUsers" Grid.Row="2" Grid.Column="2" Content="Ver usuarios" Width="110" Height="30" Margin="12,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
      </Grid>
      <Border Grid.Row="3" Background="{DynamicResource ControlBg}" CornerRadius="8" Padding="12">
        <StackPanel>
          <TextBlock Text="Requisitos:" FontWeight="SemiBold" Margin="0,0,0,6"/>
          <TextBlock Text="• Nombre: sin espacios (ej. soporte01)" Margin="0,0,0,2"/>
          <TextBlock Text="• Contraseña: mínimo 8 caracteres" Margin="0,0,0,2"/>
          <TextBlock Text="• Administrador: úsalo solo si es necesario"/>
        </StackPanel>
      </Border>
      <Grid Grid.Row="4" Margin="0,12,0,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Name="lblStatus" Grid.Column="0" Text="Listo." Foreground="#2E7D32" VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,10,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button Name="btnCancel" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
          <Button Name="btnCreate" Content="Crear usuario" Width="130" Height="30" Style="{StaticResource SystemButtonStyle}" IsEnabled="False" IsDefault="True"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"@
    try { $ui = New-WpfWindow -Xaml $stringXaml -PassThru }catch { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR creando ventana: $($_.Exception.Message)" Red; Show-WpfMessageBox -Message "No se pudo crear la ventana de usuario." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
    $w = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme
    try { Set-WpfDialogOwner -Dialog $w }catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }
    $script:__dlgResult = $false
    $c['btnClose'].Add_Click({ $w.DialogResult = $false; $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({ if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() } })
    try { $adminGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-544').Name; $userGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-545').Name }catch { Show-WpfMessageBox -Message "No se pudieron obtener los grupos locales (requiere permisos).`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null; $w.Close(); return }
    function Set-Status { param([string]$Text, [string]$Level = "Ok"); switch ($Level) { "Ok" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::ForestGreen }"Warn" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::DarkGoldenrod }"Error" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::Firebrick } }; $c['lblStatus'].Text = $Text }
    function Get-PasswordText { if ($c['txtPasswordVisible'].Visibility -eq 'Visible') { return [string]$c['txtPasswordVisible'].Text }; return [string]$c['pwdPassword'].Password }
    $WriteUsersTableConsoleSb = {
        param([Parameter(Mandatory)]$Rows)
        Write-Host ""
        Write-Host "Usuarios locales" -ForegroundColor Cyan
        $lines = @()
        $lines += ("{0,-28} {1,-14} {2,-14}" -f "Usuario", "Administrador", "Estado")
        $lines += ("{0,-28} {1,-14} {2,-14}" -f ("-" * 28), ("-" * 14), ("-" * 14))
        foreach ($r in $Rows) { $lines += ("{0,-28} {1,-14} {2,-14}" -f $r.Usuario, $r.Administrador, $r.Estado) }
        $lines | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        Write-Host ""
    }.GetNewClosure()
    $getUsersRowsSb = {
        $adminGroupName = (Get-LocalGroup | Where-Object SID -eq 'S-1-5-32-544').Name
        $adminMembers = @{}
        try {
            Get-LocalGroupMember -Group $adminGroupName -ErrorAction Stop | ForEach-Object {
                $n = $_.Name
                if ($n -match '\\') { $n = ($n -split '\\')[-1] }
                $adminMembers[$n.ToLowerInvariant()] = $true
            }
        } catch {
            Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] No se pudieron obtener miembros admin: $($_.Exception.Message)" Yellow
        }
        Get-LocalUser | Sort-Object Name | ForEach-Object {
            $uname = $_.Name
            $isAdmin = $false
            if ($adminMembers.Count -gt 0) { $isAdmin = $adminMembers.ContainsKey($uname.ToLowerInvariant()) }
            [pscustomobject]@{Usuario = $uname; Administrador = if ($isAdmin) { "Sí" }else { "No" }; Estado = if ($_.Enabled) { "Habilitado" }else { "Deshabilitado" } }
        }
    }.GetNewClosure()
    $ShowUsersTableDialogSb = {
        param([Parameter(Mandatory)][System.Windows.Window]$Owner, [Parameter(Mandatory)]$Rows, [Parameter(Mandatory)][scriptblock]$GetUsersRowsSb, [Parameter(Mandatory)][scriptblock]$WriteUsersTableConsoleSb)
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Usuarios locales" Height="500" Width="700" WindowStartupLocation="CenterOwner" ResizeMode="CanResize" ShowInTaskbar="False" Background="{DynamicResource FormBg}">
  <Grid Margin="12">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="Usuarios locales" Foreground="{DynamicResource FormFg}" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
    <Border Grid.Row="1" Background="{DynamicResource PanelBg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" CornerRadius="10" Padding="8">
      <DataGrid Name="dgUsers" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="None" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" RowHeight="28" AlternationCount="2" SelectionMode="Single">
        <DataGrid.ColumnHeaderStyle>
          <Style TargetType="{x:Type DataGridColumnHeader}">
            <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
          </Style>
        </DataGrid.ColumnHeaderStyle>
        <DataGrid.RowStyle>
          <Style TargetType="{x:Type DataGridRow}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
              <Trigger Property="ItemsControl.AlternationIndex" Value="1"><Setter Property="Background" Value="{DynamicResource PanelBg}"/></Trigger>
              <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{DynamicResource AccentPrimary}"/><Setter Property="Foreground" Value="{DynamicResource FormFg}"/></Trigger>
            </Style.Triggers>
          </Style>
        </DataGrid.RowStyle>
        <DataGrid.CellStyle>
          <Style TargetType="{x:Type DataGridCell}">
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="Padding" Value="10,4"/>
          </Style>
        </DataGrid.CellStyle>
        <DataGrid.Columns>
          <DataGridTextColumn Header="Usuario" Binding="{Binding Usuario}" Width="*"/>
          <DataGridTextColumn Header="Administrador" Binding="{Binding Administrador}" Width="140"/>
          <DataGridTextColumn Header="Estado" Binding="{Binding Estado}" Width="150"/>
        </DataGrid.Columns>
      </DataGrid>
    </Border>
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button Name="btnDelete" Content="Eliminar" Width="110" Height="32" Margin="0,0,10,0" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"/>
      <Button Name="btnCopy" Content="Copiar" Width="110" Height="32" Margin="0,0,10,0" Background="{DynamicResource ControlBg}" Foreground="{DynamicResource ControlFg}" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"/>
      <Button Name="btnClose" Content="Cerrar" Width="110" Height="32" Background="{DynamicResource AccentPrimary}" Foreground="{DynamicResource FormFg}" BorderThickness="0"/>
    </StackPanel>
  </Grid>
</Window>
"@
        $ui2 = New-WpfWindow -Xaml $xaml -PassThru
        $win = $ui2.Window
        $ctrl = $ui2.Controls
        $theme2 = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $win -Theme $theme2
        if ($Owner) { $win.Owner = $Owner; $win.WindowStartupLocation = "CenterOwner" }else { $win.WindowStartupLocation = "CenterScreen" }
        $Rows = @($Rows)
        $ctrl['dgUsers'].ItemsSource = $Rows
        $refresh = {
            Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Refresh: obteniendo usuarios..."
            $newRows = @(& $GetUsersRowsSb)
            Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Refresh: rows=$($newRows.Count)"
            $Rows = $newRows
            $ctrl['dgUsers'].ItemsSource = $null
            $ctrl['dgUsers'].ItemsSource = $Rows
            & $WriteUsersTableConsoleSb -Rows $Rows
        }.GetNewClosure()
        $ctrl['btnCopy'].Add_Click({
                try {
                    $tsv = ($Rows | ForEach-Object { "{0}`t{1}`t{2}" -f $_.Usuario, $_.Administrador, $_.Estado }) -join "`r`n"
                    [System.Windows.Clipboard]::SetText($tsv)
                } catch {}
            }.GetNewClosure())
        $ctrl['btnDelete'].Add_Click({
                try {
                    $sel = $ctrl['dgUsers'].SelectedItem
                    if (-not $sel -or [string]::IsNullOrWhiteSpace([string]$sel.Usuario)) { Show-WpfMessageBox -Message "Selecciona un usuario para eliminar." -Title "Aviso" -Buttons OK -Icon Warning -Owner $win | Out-Null; return }
                    $u = [string]$sel.Usuario
                    Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Solicitud eliminar usuario='$u'"
                    if ($u.ToLowerInvariant() -in @("administrator", "administrador", "guest", "invitado", "defaultaccount", "wdagutilityaccount")) { Show-WpfMessageBox -Message "Este usuario está protegido y no se eliminará." -Title "Aviso" -Buttons OK -Icon Warning -Owner $win | Out-Null; return }
                    $profile = $null
                    try { $profile = (Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { $_.LocalPath -and ($_.LocalPath -match "\\Users\\$([regex]::Escape($u))$") } | Select-Object -First 1) }catch {}
                    $profilePath = if ($profile -and $profile.LocalPath) { [string]$profile.LocalPath }else { "C:\Users\$u" }
                    $warn = "Se eliminará el usuario:`n`n$u`n`nImportante:`nLa carpeta de perfil, si existe, normalmente está en:`n$profilePath`n`nLa eliminación del usuario no siempre borra esa carpeta."
                    $c1 = Show-WpfMessageBoxSafe -Message $warn -Title "Confirmar eliminación" -Buttons YesNo -Icon Warning -Owner $win
                    if ($c1 -ne [System.Windows.MessageBoxResult]::Yes) { return }
                    $c2 = Show-WpfMessageBoxSafe -Message "¿Eliminar definitivamente el usuario '$u'?" -Title "Confirmación final" -Buttons YesNo -Icon Warning -Owner $win
                    if ($c2 -ne [System.Windows.MessageBoxResult]::Yes) { return }
                    Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Eliminando usuario='$u'..."
                    try { Remove-LocalUser -Name $u -ErrorAction Stop }catch { Show-WpfMessageBox -Message "No se pudo eliminar el usuario:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error -Owner $win | Out-Null; return }
                    Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Usuario eliminado OK usuario='$u'"
                    Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Ejecutando refresh tras eliminar usuario='$u'..."
                    & $refresh
                    Write-DzDebug "`t[DEBUG][Show-UsersTableDialog] Refresh OK tras eliminar usuario='$u'"
                } catch {
                    Show-WpfMessageBox -Message "Error al eliminar usuario:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error -Owner $win | Out-Null
                }
            }.GetNewClosure())
        $ctrl['btnClose'].Add_Click({ $win.Close() })
        $null = $win.ShowDialog()
    }.GetNewClosure()
    function Validate-Form {
        $username = ([string]$c['txtUsername'].Text).Trim()
        $pass = (Get-PasswordText).Trim()
        Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Validate username='$username' passLen=$($pass.Length)"
        if ([string]::IsNullOrWhiteSpace($username)) { Set-Status "Escriba un nombre de usuario." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ($username -match "\s") { Set-Status "El nombre no debe contener espacios." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ([string]::IsNullOrWhiteSpace($pass)) { Set-Status "Escriba una contraseña." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ($pass.Length -lt 8) { Set-Status "La contraseña debe tener al menos 8 caracteres." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        try { $exists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue; if ($exists) { Set-Status "El usuario '$username' ya existe." "Error"; $c['btnCreate'].IsEnabled = $false; return } }catch {
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Validate existencia falló: $($_.Exception.Message)" Yellow
            Set-Status "Aviso: no se pudo validar si el usuario ya existe (permisos)." "Warn"
        }
        Set-Status "Listo para crear usuario." "Ok"
        $c['btnCreate'].IsEnabled = $true
    }
    $c['tglShowPassword'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowPassword ON"; $c['txtPasswordVisible'].Text = [string]$c['pwdPassword'].Password; $c['pwdPassword'].Visibility = 'Collapsed'; $c['txtPasswordVisible'].Visibility = 'Visible'; Validate-Form }.GetNewClosure())
    $c['tglShowPassword'].Add_Unchecked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowPassword OFF"; $c['pwdPassword'].Password = [string]$c['txtPasswordVisible'].Text; $c['txtPasswordVisible'].Visibility = 'Collapsed'; $c['pwdPassword'].Visibility = 'Visible'; Validate-Form }.GetNewClosure())
    $c['txtUsername'].Add_TextChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] txtUsername changed"; Validate-Form }.GetNewClosure())
    $c['pwdPassword'].Add_PasswordChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] pwdPassword changed"; Validate-Form }.GetNewClosure())
    $c['txtPasswordVisible'].Add_TextChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] txtPasswordVisible changed"; Validate-Form }.GetNewClosure())
    $c['rbStandard'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Tipo=Standard"; Validate-Form }.GetNewClosure())
    $c['rbAdmin'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Tipo=Admin"; Validate-Form }.GetNewClosure())
    $c['btnShowUsers'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnShowUsers click (tabla)"
            try {
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnShowUsers: getUsersRowsSb=$([bool]$getUsersRowsSb)"
                $rows = @(& $getUsersRowsSb)
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnShowUsers: rows=$($rows.Count)"
                & $WriteUsersTableConsoleSb -Rows $rows
                & $ShowUsersTableDialogSb -Owner $w -Rows $rows -GetUsersRowsSb $getUsersRowsSb -WriteUsersTableConsoleSb $WriteUsersTableConsoleSb | Out-Null
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Tabla de usuarios mostrada OK"
            } catch {
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR btnShowUsers: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "No se pudieron listar usuarios:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        }.GetNewClosure())
    $c['btnCreate'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnCreate click"
            $username = ([string]$c['txtUsername'].Text).Trim()
            $password = Get-PasswordText
            $isAdmin = $false; try { $isAdmin = [bool]$c['rbAdmin'].IsChecked }catch {}
            $tipo = if ($isAdmin) { "Administrador" }else { "Usuario estándar" }
            $group = if ($isAdmin) { $adminGroup }else { $userGroup }
            $confirmMsg = "Se creará el usuario:`n`n$username`n`nTipo: $tipo`nGrupo: $group"
            $conf = Show-WpfMessageBoxSafe -Message $confirmMsg -Title "Confirmar creación" -Buttons YesNo -Icon Warning -Owner $w
            if ($conf -ne [System.Windows.MessageBoxResult]::Yes) { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Creación cancelada"; return }
            try {
                if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) { Set-Status "El usuario '$username' ya existe." "Error"; Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Ya existe: $username" Yellow; return }
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                New-LocalUser -Name $username -Password $securePassword -AccountNeverExpires -PasswordNeverExpires | Out-Null
                Add-LocalGroupMember -Group $group -Member $username
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Usuario creado: $username Grupo: $group"
                Show-WpfMessageBox -Message "Usuario '$username' creado y agregado al grupo '$group'." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $w.DialogResult = $true; $w.Close()
            } catch {
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR creando usuario: $($_.Exception.Message)" Red
                Set-Status "Error: $($_.Exception.Message)" "Error"
                Show-WpfMessageBox -Message "Error al crear usuario:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        }.GetNewClosure())
    $c['btnCancel'].Add_Click({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnCancel"; $w.DialogResult = $false; $w.Close() }.GetNewClosure())
    Validate-Form
    try { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowDialog()"; $w.ShowDialog() | Out-Null }catch { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR ShowDialog: $($_.Exception.Message)" Red; throw }
    Write-DzDebug "`t[DEBUG][Show-AddUserDialog] FIN"
}
function Show-IPConfigDialog {
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    $theme = Get-DzUiTheme
    function Test-IPv4 {
        param([string]$Ip)
        if ([string]::IsNullOrWhiteSpace($Ip)) { return $false }
        $Ip = $Ip.Trim()
        return [System.Net.IPAddress]::TryParse($Ip, [ref]([System.Net.IPAddress]$null)) -and ($Ip -match '^\d{1,3}(\.\d{1,3}){3}$')
    }
    function Get-AdapterIpsText {
        param([string]$Alias)
        try {
            $ips = Get-NetIPAddress -InterfaceAlias $Alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress
            if ($ips) { return "IPs asignadas: " + ($ips -join ", ") }
        } catch {}
        return "IPs asignadas: -"
    }
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Asignación de IPs"
        Height="250" Width="560"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="{DynamicResource FormBg}"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="12" Background="{DynamicResource FormBg}">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0"
                   Text="Seleccione el adaptador de red:"
                   Margin="0,0,0,8"/>
        <ComboBox Name="cmbAdapters"
                  Grid.Row="1"
                  Height="28"
                  Margin="0,0,0,10"/>
        <TextBlock Name="lblIps"
                   Grid.Row="2"
                   Text="IPs asignadas: -"
                   TextWrapping="Wrap"
                   Margin="0,0,0,12"
                   MinHeight="36"/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnAssignIp" Content="Asignar Nueva IP" Width="140" Height="32" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnDhcp"     Content="Cambiar a DHCP"  Width="140" Height="32" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnClose"    Content="Cerrar"         Width="110" Height="32" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $ui = New-WpfWindow -Xaml $stringXaml -PassThru
    $window = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $window -Theme $theme
    try {
        if ($Global:window -is [System.Windows.Window]) {
            $window.Owner = $Global:window
            $window.WindowStartupLocation = "CenterOwner"
        } else { $window.WindowStartupLocation = "CenterScreen" }
    } catch { $window.WindowStartupLocation = "CenterScreen" }
    $adapters = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "Up" })
    $c['cmbAdapters'].Items.Clear()
    $c['cmbAdapters'].Items.Add("Selecciona 1 adaptador de red") | Out-Null
    foreach ($a in $adapters) { $c['cmbAdapters'].Items.Add($a.Name) | Out-Null }
    $c['cmbAdapters'].SelectedIndex = 0
    $updateUi = {
        $sel = [string]$c['cmbAdapters'].SelectedItem
        $valid = ($sel -and $sel -ne "Selecciona 1 adaptador de red")
        $c['btnAssignIp'].IsEnabled = $valid
        $c['btnDhcp'].IsEnabled = $valid
        if ($valid) { $c['lblIps'].Text = (Get-AdapterIpsText -Alias $sel) }else { $c['lblIps'].Text = "IPs asignadas: -" }
    }
    $c['cmbAdapters'].Add_SelectionChanged({ & $updateUi })
    & $updateUi
    $c['btnAssignIp'].Add_Click({
            $alias = [string]$c['cmbAdapters'].SelectedItem
            if (-not $alias -or $alias -eq "Selecciona 1 adaptador de red") { Show-WpfMessageBox -Message "Por favor, selecciona un adaptador de red." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
            $current = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $current) { Show-WpfMessageBox -Message "No se pudo obtener la configuración IPv4 del adaptador." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
            $prefixLen = $current.PrefixLength
            $newIp = New-WpfInputDialog -Title "Nueva IP" -Prompt "Ingrese la nueva dirección IP IPv4:" -DefaultValue ""
            if ([string]::IsNullOrWhiteSpace($newIp)) { return }
            if (-not (Test-IPv4 -Ip $newIp)) { Show-WpfMessageBox -Message "La IP '$newIp' no es válida." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
            $exists = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -eq $newIp.Trim() }
            if ($exists) { Show-WpfMessageBox -Message "La IP $newIp ya está asignada a $alias." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
            try {
                New-NetIPAddress -IPAddress $newIp.Trim() -PrefixLength $prefixLen -InterfaceAlias $alias -ErrorAction Stop | Out-Null
                Show-WpfMessageBox -Message "Se agregó la IP $newIp al adaptador $alias." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $c['lblIps'].Text = (Get-AdapterIpsText -Alias $alias)
            } catch {
                Show-WpfMessageBox -Message "Error al agregar IP:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnDhcp'].Add_Click({
            $alias = [string]$c['cmbAdapters'].SelectedItem
            if (-not $alias -or $alias -eq "Selecciona 1 adaptador de red") { Show-WpfMessageBox -Message "Por favor, selecciona un adaptador de red." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
            $any = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($any -and $any.PrefixOrigin -eq "Dhcp") { Show-WpfMessageBox -Message "El adaptador ya está en DHCP." -Title "Información" -Buttons OK -Icon Information | Out-Null; return }
            $conf = Show-WpfMessageBox -Message "¿Está seguro de que desea cambiar a DHCP?" -Title "Confirmación" -Buttons YesNo -Icon Question
            if ($conf -ne [System.Windows.MessageBoxResult]::Yes) { return }
            try {
                $manualIps = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.PrefixOrigin -eq "Manual" }
                foreach ($ip in $manualIps) { Remove-NetIPAddress -IPAddress $ip.IPAddress -PrefixLength $ip.PrefixLength -InterfaceAlias $alias -Confirm:$false -ErrorAction SilentlyContinue }
                Set-NetIPInterface -InterfaceAlias $alias -Dhcp Enabled -ErrorAction Stop | Out-Null
                Set-DnsClientServerAddress -InterfaceAlias $alias -ResetServerAddresses -ErrorAction SilentlyContinue | Out-Null
                Show-WpfMessageBox -Message "Se cambió a DHCP en el adaptador $alias." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $c['lblIps'].Text = "Generando IP por DHCP. Seleccione de nuevo."
            } catch {
                Show-WpfMessageBox -Message "Error al cambiar a DHCP:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnClose'].Add_Click({ $window.Close() })
    $window.ShowDialog() | Out-Null
}
function Show-LZMADialog {
    param([array]$Instaladores)
    Write-DzDebug "`t[DEBUG][Show-LZMADialog] INICIO"
    $theme = Get-DzUiTheme
    $UiConfirm = { param([string]$m, [string]$t = "Confirmar")(Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "YesNo" -Icon "Question" -Owner $w) -eq [System.Windows.MessageBoxResult]::Yes }.GetNewClosure()
    function Convert-RegProviderPathToDisplay {
        param([Parameter(Mandatory)][string]$ProviderPath)
        $p = $ProviderPath -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
        if ($p -like 'HKEY_LOCAL_MACHINE\*') { return ('HKLM:\' + $p.Substring('HKEY_LOCAL_MACHINE\'.Length)) }
        if ($p -like 'HKEY_CURRENT_USER\*') { return ('HKCU:\' + $p.Substring('HKEY_CURRENT_USER\'.Length)) }
        return $p
    }
    function Get-LzmaInstallerItems {
        $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
        if (-not (Test-Path $LZMAregistryPath)) { return @() }
        $carpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
        $tmp = @()
        foreach ($carpeta in $carpetasPrincipales) {
            $subdirs = Get-ChildItem -Path $carpeta.PSPath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
            foreach ($sd in $subdirs) {
                $tmp += [PSCustomObject]@{Name = [string]$sd.PSChildName; ProviderPath = [string]$sd.PSPath; DisplayPath = (Convert-RegProviderPathToDisplay -ProviderPath ([string]$sd.PSPath)) }
            }
        }
        $tmp | Sort-Object Name -Descending
    }
    if (-not $Instaladores -or $Instaladores.Count -eq 0) {
        $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
        if (-not (Test-Path $LZMAregistryPath)) {
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] No existe la clave LZMA: $LZMAregistryPath" Yellow
            Show-WpfMessageBox -Message "No se encontró Advanced Installer (LZMA) en este equipo.`n`nRuta no existe:`n$LZMAregistryPath" -Title "Sin instaladores" -Buttons OK -Icon Information | Out-Null
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Fin - sin clave LZMA"
            return
        }
        try {
            $carpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }
            if (-not $carpetasPrincipales -or $carpetasPrincipales.Count -lt 1) {
                Write-DzDebug "`t[DEBUG][Show-LZMADialog] No se encontraron carpetas principales." Yellow
                Show-WpfMessageBox -Message "No se encontraron carpetas principales en la ruta del registro." -Title "Sin resultados" -Buttons OK -Icon Information | Out-Null
                return
            }
            $tmp = @()
            foreach ($carpeta in $carpetasPrincipales) {
                $subdirs = Get-ChildItem -Path $carpeta.PSPath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
                foreach ($sd in $subdirs) { $tmp += [PSCustomObject]@{Name = $sd.PSChildName; Path = $sd.PSPath } }
            }
            if (-not $tmp -or $tmp.Count -lt 1) {
                Write-DzDebug "`t[DEBUG][Show-LZMADialog] No se encontraron subcarpetas." Yellow
                Show-WpfMessageBox -Message "No se encontraron instaladores (subcarpetas) en la ruta del registro." -Title "Sin resultados" -Buttons OK -Icon Information | Out-Null
                return
            }
            $Instaladores = $tmp | Sort-Object Name -Descending
        } catch {
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Error accediendo al registro: $($_.Exception.Message)" Red
            Show-WpfMessageBox -Message "Error accediendo al registro:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            return
        }
    }
    $items = foreach ($i in $Instaladores) {
        if (-not $i) { continue }
        $provider = if ($null -ne $i.PSObject.Properties['ProviderPath']) { [string]$i.ProviderPath }else { [string]$i.Path }
        $displayPath = if ($null -ne $i.PSObject.Properties['DisplayPath'] -and -not [string]::IsNullOrWhiteSpace([string]$i.DisplayPath)) { [string]$i.DisplayPath }else { Convert-RegProviderPathToDisplay -ProviderPath $provider }
        [PSCustomObject]@{Name = [string]$i.Name; ProviderPath = $provider; DisplayPath = $displayPath; Display = ("{0}  |  {1}" -f [string]$i.Name, $displayPath) }
    }
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Carpetas LZMA" Height="290" Width="760" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" ShowInTaskbar="False" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#2A2A2A"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}" Color="#2A2A2A"/>
    <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="ComboItemHoverBg" Color="#2A2A2A"/>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Background" Value="$($theme.ControlBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,4"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" SnapsToDevicePixels="True">
                <DockPanel>
                  <Border DockPanel.Dock="Right" Width="32" Background="Transparent">
                    <Path Data="M 0 0 L 6 6 L 12 0 Z" Fill="{TemplateBinding Foreground}" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,2,0,0"/>
                  </Border>
                  <ContentPresenter x:Name="ContentSite" Margin="{TemplateBinding Padding}" VerticalAlignment="Center" HorizontalAlignment="Left" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" RecognizesAccessKey="True"/>
                </DockPanel>
              </Border>
              <ToggleButton x:Name="DropDownToggle" Background="Transparent" BorderThickness="0" Focusable="False" ClickMode="Press" IsChecked="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>
              <Popup x:Name="Popup" Placement="Bottom" PlacementTarget="{Binding ElementName=Bd}" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                <Border Background="$($theme.ControlBackground)" BorderBrush="$($theme.BorderColor)" BorderThickness="1" CornerRadius="8" SnapsToDevicePixels="True" Margin="0,6,0,0">
                  <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Bd" Property="Opacity" Value="0.60"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="$($theme.AccentPrimary)"/>
              </Trigger>
              <Trigger Property="IsDropDownOpen" Value="True">
                <Setter TargetName="Bd" Property="BorderBrush" Value="$($theme.AccentPrimary)"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="{StaticResource ComboItemHoverBg}"/>
          <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
          <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="0.55"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style x:Key="SystemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
      <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
    </Style>
  </Window.Resources>
  <Border Background="{DynamicResource FormBg}" CornerRadius="10" BorderBrush="{DynamicResource AccentPrimary}" BorderThickness="2" Padding="0">
    <Border.Effect>
      <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
    </Border.Effect>
    <Grid Margin="16">
      <Grid.RowDefinitions>
        <RowDefinition Height="36"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Carpetas LZMA" VerticalAlignment="Center" FontWeight="SemiBold"/>
        <Button Name="btnClose" Grid.Column="1" Content="✕" Width="34" Height="26" Margin="8,0,0,0" ToolTip="Cerrar" Background="Transparent" BorderBrush="Transparent"/>
      </Grid>
      <TextBlock Grid.Row="1" Text="Seleccione el instalador (registro) que desea renombrar." FontWeight="SemiBold" Margin="0,0,0,10"/>
      <ComboBox Name="cmbInstallers" Grid.Row="2" Height="30" Margin="0,0,0,10" DisplayMemberPath="Display" SelectedValuePath="ProviderPath"/>
      <Border Grid.Row="3" Background="{DynamicResource ControlBg}" CornerRadius="8" Padding="10" Margin="0,0,0,8" MinHeight="78">
        <StackPanel>
          <TextBlock Text="AI_ExePath:" Margin="0,0,0,4"/>
          <TextBlock Name="lblExePath" Text="-" Foreground="{DynamicResource AccentMuted}" TextWrapping="Wrap" TextTrimming="None" MaxHeight="48"/>
        </StackPanel>
      </Border>
      <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,0">
        <Button Name="btnRename" Content="Renombrar" Width="110" Height="30" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
        <Button Name="btnExit" Content="Salir" Width="110" Height="30" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"@
    try { $ui = New-WpfWindow -Xaml $stringXaml -PassThru }catch { Write-DzDebug "`t[DEBUG][Show-LZMADialog] ERROR creando ventana: $($_.Exception.Message)" Red; Show-WpfMessageBox -Message "No se pudo crear la ventana LZMA." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
    $w = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme
    try { Set-WpfDialogOwner -Dialog $w }catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }
    $script:__dlgResult = $false
    $script:__allowClose = $false
    $w.Add_Closing({ param($sender, $e)if (-not $script:__allowClose) { $e.Cancel = $true } })
    $c['btnClose'].Add_Click({ $script:__dlgResult = $false; $script:__allowClose = $true; $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({ if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() } })
    $placeholder = [PSCustomObject]@{Name = ""; ProviderPath = ""; Display = "Selecciona instalador a renombrar" }
    $c['cmbInstallers'].ItemsSource = @($placeholder) + @($items)
    $c['cmbInstallers'].SelectedIndex = 0
    $updateUi = {
        $idx = $c['cmbInstallers'].SelectedIndex
        $c['btnRename'].IsEnabled = ($idx -gt 0)
        if ($idx -le 0) { $c['lblExePath'].Text = "-"; return }
        $it = $c['cmbInstallers'].SelectedItem
        if (-not $it -or [string]::IsNullOrWhiteSpace($it.ProviderPath)) { $c['lblExePath'].Text = "No encontrado"; return }
        try {
            $prop = Get-ItemProperty -Path $it.ProviderPath -Name "AI_ExePath" -ErrorAction SilentlyContinue
            if ($prop -and $prop.AI_ExePath) { $c['lblExePath'].Text = [string]$prop.AI_ExePath }else { $c['lblExePath'].Text = "No encontrado" }
        } catch { $c['lblExePath'].Text = "Error leyendo AI_ExePath" }
    }
    $c['cmbInstallers'].Add_SelectionChanged({ & $updateUi })
    & $updateUi
    $c['btnRename'].Add_Click({
            $idx = $c['cmbInstallers'].SelectedIndex
            if ($idx -le 0) { return }
            $it = $c['cmbInstallers'].SelectedItem
            if (-not $it -or [string]::IsNullOrWhiteSpace($it.ProviderPath)) { return }
            $rutaVieja = [string]$it.ProviderPath
            $nombreViejo = [string]$it.Name
            $nuevoNombre = "$nombreViejo.backup"
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Renombrar solicitado:" -Color DarkGray
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Nombre viejo: $nombreViejo" -Color DarkGray
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Nuevo nombre: $nuevoNombre" -Color DarkGray
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] ProviderPath viejo: $rutaVieja" -Color DarkGray
            $msg = "¿Está seguro de renombrar el registro?`n`n$rutaVieja`n`nA:`n$nuevoNombre"
            $ok = & $UiConfirm $msg "Confirmar renombrado"
            if (-not $ok) { Write-DzDebug "t[DEBUG][Show-LZMADialog] Usuario canceló renombrado." -Color DarkGray; return }
            try {
                Rename-Item -Path $rutaVieja -NewName $nuevoNombre -ErrorAction Stop
                Show-WpfMessageBox -Message "Registro renombrado correctamente." -Title "Éxito" -Buttons OK -Icon Information -Owner $w | Out-Null
                $fresh = Get-LzmaInstallerItems
                $freshItems = foreach ($i in $fresh) { [PSCustomObject]@{Name = [string]$i.Name; ProviderPath = [string]$i.ProviderPath; DisplayPath = [string]$i.DisplayPath; Display = ("{0}  |  {1}" -f [string]$i.Name, [string]$i.DisplayPath) } }
                $placeholder2 = [PSCustomObject]@{Name = ""; ProviderPath = ""; Display = "Selecciona instalador a renombrar" }
                $c['cmbInstallers'].ItemsSource = @($placeholder2) + @($freshItems)
                $c['cmbInstallers'].SelectedIndex = 0
                & $updateUi
            } catch {
                Show-WpfMessageBox -Message "Error al renombrar:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error -Owner $w | Out-Null
            }
        })
    $c['btnExit'].Add_Click({ $script:__dlgResult = $false; $script:__allowClose = $true; $w.Close() })
    $null = $w.ShowDialog()
    Write-DzDebug "`t[DEBUG][Show-LZMADialog] FIN"
    return $script:__dlgResult
}
function Show-FirewallConfigDialog {
    Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] INICIO"
    if (-not (Test-Administrator)) {
        Show-WpfMessageBox -Message "Esta acción requiere permisos de administrador.`n`nPor favor, ejecuta Gerardo Zermeño Tools como administrador." -Title "Permisos requeridos" -Buttons OK -Icon Warning | Out-Null
        return
    }
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Configuraciones de Firewall" Height="560" Width="760" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" ShowInTaskbar="False" WindowStyle="None" AllowsTransparency="True" Background="Transparent" FontFamily="{DynamicResource UiFontFamily}" FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <Style TargetType="TextBlock"><Setter Property="Foreground" Value="$($theme.FormForeground)"/></Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="6,2"/>
    </Style>
    <Style TargetType="CheckBox"><Setter Property="Foreground" Value="$($theme.FormForeground)"/></Style>
    <Style TargetType="Button">
      <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
  </Window.Resources>
  <Border Background="{DynamicResource FormBg}" CornerRadius="12" BorderBrush="{DynamicResource AccentPrimary}" BorderThickness="2" Padding="0">
    <Border.Effect><DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="14" Opacity="0.25"/></Border.Effect>
    <Grid Margin="18">
      <Grid.RowDefinitions>
        <RowDefinition Height="36"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Text="Configuraciones de Firewall" VerticalAlignment="Center" FontWeight="SemiBold"/>
        <Button Name="btnClose" Grid.Column="1" Content="✕" Width="34" Height="26" Margin="8,0,0,0" ToolTip="Cerrar" Background="Transparent" BorderBrush="Transparent"/>
      </Grid>
      <TextBlock Grid.Row="1" Text="Busca y agrega puertos al Firewall de Windows (reglas de entrada/salida)." FontWeight="SemiBold" Margin="0,0,0,12"/>
      <Border Grid.Row="2" Background="{DynamicResource ControlBg}" CornerRadius="10" Padding="12" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1">
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <TextBlock Text="Buscar puerto" FontWeight="SemiBold" Margin="0,0,0,8"/>
          <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Puerto" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" Name="txtSearchPort" Height="30" IsEnabled="False" VerticalContentAlignment="Center"/>
            <Button Grid.Column="2" Name="btnSearch" IsEnabled="False" Content="Buscar" Width="110" Height="30" Margin="10,0,0,0"/>
          </Grid>
          <ListBox Grid.Row="2" Name="lbResults" Height="140" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" Background="{DynamicResource PanelBg}" Foreground="{DynamicResource FormFg}"/>
        </Grid>
      </Border>
      <Border Grid.Row="3" Background="{DynamicResource ControlBg}" CornerRadius="10" Padding="12" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1" Margin="0,12,0,0">
        <Grid>
            <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Text="Agregar regla de puerto" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Puerto" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" Name="txtAddPort" Height="30" VerticalContentAlignment="Center"/>
            <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="10,0,0,0" VerticalAlignment="Center">
                <CheckBox Name="chkInbound" Content="Entrada" Margin="0,0,10,0" IsChecked="True"/>
                <CheckBox Name="chkOutbound" Content="Salida" IsChecked="True"/>
            </StackPanel>
            </Grid>
            <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Nombre de regla" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" Name="txtRuleName" Height="30" VerticalContentAlignment="Center" Text="Regla de puerto"/>
            </Grid>
        </Grid>
        </Border>
      <Grid Grid.Row="4" Margin="0,12,0,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Name="lblStatus" Grid.Column="0" Text="Listo." Foreground="#2E7D32" VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,10,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button Name="btnAdd" Content="Agregar reglas" Width="140" Height="32" Margin="0,0,10,0"/>
          <Button Name="btnCloseFooter" Content="Cerrar" Width="110" Height="32"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"@
    try { $ui = New-WpfWindow -Xaml $stringXaml -PassThru }catch {
        Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana de firewall." -Title "Error" -Buttons OK -Icon Error | Out-Null
        return
    }
    $w = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme
    try { Set-WpfDialogOwner -Dialog $w }catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }
    $SetStatus = {
        param([string]$Text, [string]$Level = "Ok")
        switch ($Level) {
            "Ok" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::ForestGreen }
            "Warn" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::DarkGoldenrod }
            "Error" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::Firebrick }
        }
        $c['lblStatus'].Text = $Text
    }.GetNewClosure()
    $GetPortValue = {
        param([string]$Text)
        $trim = ([string]$Text).Trim()
        $val = 0
        if (-not [int]::TryParse($trim, [ref]$val)) { return $null }
        if ($val -lt 1 -or $val -gt 65535) { return $null }
        return $val
    }.GetNewClosure()
    $TestPortMatch = {
        param([string]$LocalPort, [int]$Port)
        if ([string]::IsNullOrWhiteSpace($LocalPort)) { return $false }
        if ($LocalPort -eq "Any") { return $true }
        foreach ($segment in ($LocalPort -split ',')) {
            $s = $segment.Trim()
            if ([string]::IsNullOrWhiteSpace($s)) { continue }
            if ($s -match '^\d+$') {
                if ([int]$s -eq $Port) { return $true }
            } elseif ($s -match '^(?<start>\d+)\s*-\s*(?<end>\d+)$') {
                $start = [int]$Matches.start
                $end = [int]$Matches.end
                if ($Port -ge $start -and $Port -le $end) { return $true }
            }
        }
        return $false
    }.GetNewClosure()
    $GetFirewallPortMatchesAsync = {
        param([int]$Port, [System.Windows.Window]$ProgressWindow, [scriptblock]$OnComplete)
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = "STA"
        $rs.ThreadOptions = "ReuseThread"
        $rs.Open()
        $rs.SessionStateProxy.SetVariable("Port", $Port)
        $rs.SessionStateProxy.SetVariable("ProgressWindow", $ProgressWindow)
        $rs.SessionStateProxy.SetVariable("OnComplete", $OnComplete)
        $rs.SessionStateProxy.SetVariable("TestPortMatchCode", $TestPortMatch)
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
                function Update-ProgressSafe {
                    param($Window, $Percent, $Message)
                    if (-not $Window -or $Window.IsClosed) { return }
                    try {
                        $Window.Dispatcher.Invoke([action] {
                                if (-not $Window.IsClosed) {
                                    if ($Window.ProgressBar) {
                                        $Window.ProgressBar.IsIndeterminate = $false
                                        $Window.ProgressBar.Value = $Percent
                                    }
                                    if ($Window.PercentLabel) { $Window.PercentLabel.Text = "$Percent%" }
                                    if ($Window.MessageLabel -and $Message) { $Window.MessageLabel.Text = $Message }
                                    $Window.UpdateLayout()
                                }
                            }, [System.Windows.Threading.DispatcherPriority]::Normal)
                    } catch {}
                }
                try {
                    Update-ProgressSafe -Window $ProgressWindow -Percent 5 -Message "Obteniendo reglas de Firewall..."
                    Start-Sleep -Milliseconds 100
                    $rules = Get-NetFirewallRule -ErrorAction Stop
                    Update-ProgressSafe -Window $ProgressWindow -Percent 30 -Message "Filtrando puertos..."
                    Start-Sleep -Milliseconds 100
                    $filters = $rules | Get-NetFirewallPortFilter -ErrorAction Stop
                    Update-ProgressSafe -Window $ProgressWindow -Percent 70 -Message "Procesando coincidencias..."
                    Start-Sleep -Milliseconds 100
                    $matches = @()
                    $total = $filters.Count
                    $current = 0
                    $lastUpdate = [DateTime]::Now
                    foreach ($f in $filters) {
                        $current++
                        $now = [DateTime]::Now
                        if (($now - $lastUpdate).TotalMilliseconds -gt 200 -or $current -eq $total) {
                            $percent = [Math]::Min(70 + [int](($current / $total) * 25), 95)
                            Update-ProgressSafe -Window $ProgressWindow -Percent $percent -Message "Procesando $current de $total reglas..."
                            $lastUpdate = $now
                        }
                        if ($f.Protocol -notin @('TCP', 'UDP')) { continue }
                        $testResult = & $TestPortMatchCode -LocalPort $f.LocalPort -Port $Port
                        if (-not $testResult) { continue }
                        $r = $f.AssociatedNetFirewallRule
                        if (-not $r) { continue }
                        $matches += [pscustomobject]@{Direction = [string]$r.Direction; Name = [string]$r.DisplayName; Enabled = [string]$r.Enabled; Action = [string]$r.Action; Profile = [string]$r.Profile; Protocol = [string]$f.Protocol; LocalPort = [string]$f.LocalPort }
                    }
                    Update-ProgressSafe -Window $ProgressWindow -Percent 100 -Message "Listo."
                    if ($OnComplete) { $ProgressWindow.Dispatcher.Invoke([action] { & $OnComplete $matches }, [System.Windows.Threading.DispatcherPriority]::Normal) }
                } catch {
                    $errMsg = $_.Exception.Message
                    if ($OnComplete) { $ProgressWindow.Dispatcher.Invoke([action] { & $OnComplete @{Error = $errMsg } }, [System.Windows.Threading.DispatcherPriority]::Normal) }
                }
            })
        $handle = $ps.BeginInvoke()
        return @{PowerShell = $ps; Runspace = $rs; Handle = $handle }
    }.GetNewClosure()
    $GetFirewallPortMatchesSync = {
        param([int]$Port)
        $oldPref = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            $rules = Get-NetFirewallRule -ErrorAction Stop
            $filters = $rules | Get-NetFirewallPortFilter -ErrorAction Stop
        } finally { $ProgressPreference = $oldPref }
        $matches = @()
        foreach ($f in $filters) {
            if ($f.Protocol -notin @('TCP', 'UDP')) { continue }
            if (-not (& $TestPortMatch -LocalPort $f.LocalPort -Port $Port)) { continue }
            $r = $f.AssociatedNetFirewallRule
            if (-not $r) { continue }
            $matches += [pscustomobject]@{Direction = [string]$r.Direction; Name = [string]$r.DisplayName; Enabled = [string]$r.Enabled; Action = [string]$r.Action; Profile = [string]$r.Profile; Protocol = [string]$f.Protocol; LocalPort = [string]$f.LocalPort }
        }
        return $matches
    }.GetNewClosure()
    $RenderResults = {
        param([array]$Matches, [int]$Port)
        $c['lbResults'].Items.Clear()
        if (-not $Matches -or $Matches.Count -eq 0) {
            $c['lbResults'].Items.Add("No se encontraron reglas para el puerto $Port.") | Out-Null
            & $SetStatus "Sin coincidencias para el puerto $Port." "Warn"
            return
        }
        foreach ($m in $Matches) {
            $label = "{0} | {1} | {2} | {3} | {4} | Puertos: {5}" -f $m.Direction, $m.Action, $m.Profile, $m.Protocol, $m.Name, $m.LocalPort
            $c['lbResults'].Items.Add($label) | Out-Null
        }
        & $SetStatus "Se encontraron $($Matches.Count) regla(s) para el puerto $Port." "Ok"
    }.GetNewClosure()
    $c['btnClose'].Add_Click({ $w.Close() })
    $c['btnCloseFooter'].Add_Click({ $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({ if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() } })
    $c['btnSearch'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] btnSearch click a las $([DateTime]::Now.ToString("o"))"
            $port = & $GetPortValue $c['txtSearchPort'].Text
            Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Buscar puerto='$($c['txtSearchPort'].Text)' port=$port"
            if ($null -eq $port) { & $SetStatus "Ingresa un puerto válido (1-65535)." "Warn"; return }
            $pb = Show-WpfProgressBar -Title "Buscando reglas de Firewall" -Message "Iniciando..."
            if (-not $pb) { & $SetStatus "Error creando barra de progreso." "Error"; return }
            $c['btnSearch'].IsEnabled = $false
            $c['btnAdd'].IsEnabled = $false
            $onComplete = {
                param($result)
                try {
                    if ($result -is [hashtable] -and $result.ContainsKey('Error')) {
                        Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Error buscando reglas: $($result.Error)" Red
                        & $SetStatus "Error al buscar reglas: $($result.Error)" "Error"
                    } else {
                        $matches = $result
                        Write-DzDebug "`t[DEBUG] Encontradas $($matches.Count) coincidencias para puerto $port"
                        & $RenderResults $matches $port
                    }
                } catch {
                    Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Error procesando resultados: $($_.Exception.Message)" Red
                    & $SetStatus "Error procesando resultados: $($_.Exception.Message)" "Error"
                } finally {
                    if ($pb) { Close-WpfProgressBar -Window $pb }
                    $c['btnSearch'].IsEnabled = $true
                    $c['btnAdd'].IsEnabled = $true
                }
            }.GetNewClosure()
            try { $job = & $GetFirewallPortMatchesAsync $port $pb $onComplete }catch {
                Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Error iniciando búsqueda async: $($_.Exception.Message)" Red
                & $SetStatus "Error al iniciar búsqueda: $($_.Exception.Message)" "Error"
                if ($pb) { Close-WpfProgressBar -Window $pb }
                $c['btnSearch'].IsEnabled = $true
                $c['btnAdd'].IsEnabled = $true
            }
        }.GetNewClosure())
    $c['btnAdd'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] btnAdd click a las $([DateTime]::Now.ToString("o"))"
            $port = & $GetPortValue $c['txtAddPort'].Text
            if ($null -eq $port) { & $SetStatus "Ingresa un puerto válido (1-65535)." "Warn"; return }
            $directions = @()
            if ($c['chkInbound'].IsChecked) { $directions += "Inbound" }
            if ($c['chkOutbound'].IsChecked) { $directions += "Outbound" }
            if ($directions.Count -eq 0) { & $SetStatus "Selecciona al menos una dirección (Entrada/Salida)." "Warn"; return }
            $ruleName = $c['txtRuleName'].Text.Trim()
            if ([string]::IsNullOrWhiteSpace($ruleName)) { $ruleName = "Puerto $port" }
            try {
                $created = 0
                $createdRules = @()
                $errors = @()
                foreach ($dir in $directions) {
                    $dirLabel = if ($dir -eq "Inbound") { "Entrada" }else { "Salida" }
                    $finalRuleName = "$ruleName ($dirLabel)"
                    $desc = "Regla creada para puerto $port ($dirLabel)."
                    try {
                        New-NetFirewallRule -DisplayName $finalRuleName -Direction $dir -Action Allow -Protocol TCP -LocalPort $port -Profile Any -Description $desc -ErrorAction Stop | Out-Null
                        $created++
                        $createdRules += $finalRuleName
                        Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Regla creada: $finalRuleName"
                    } catch {
                        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*ya existe*") {
                            Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Ya existe regla $dir para puerto $port"
                        } else {
                            $errors += "Error en $dirLabel`: $($_.Exception.Message)"
                            Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Error creando regla $dir`: $($_.Exception.Message)" Red
                        }
                    }
                }
                if ($errors.Count -gt 0) {
                    & $SetStatus "Errores al crear reglas: $($errors -join '; ')" "Error"
                } elseif ($created -gt 0) {
                    & $SetStatus "Se agregaron $created regla(s) para el puerto $port." "Ok"
                    $dirText = ($directions | ForEach-Object { if ($_ -eq "Inbound") { "Entrada" }else { "Salida" } }) -join " y "
                    $msg = "Regla(s) agregada(s) exitosamente:`n`nNombre: $($createdRules -join ', ')`nPuerto: $port`nDirección: $dirText"
                    $result = Show-WpfMessageBox -Message $msg -Title "Regla agregada" -Buttons "OK" -Icon "Information" -Owner $w
                    if ($result -eq [System.Windows.MessageBoxResult]::OK) { $w.Close() }
                } else {
                    & $SetStatus "No se crearon reglas nuevas: ya existen reglas con ese nombre." "Warn"
                }
            } catch {
                Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] Error agregando reglas: $($_.Exception.Message)" Red
                & $SetStatus "Error al agregar reglas: $($_.Exception.Message)" "Error"
            }
        }.GetNewClosure())
    try {
        Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] ShowDialog()"
        $w.ShowDialog() | Out-Null
    } catch {
        Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] ERROR ShowDialog: $($_.Exception.Message)" Red
        throw
    }
    Write-DzDebug "`t[DEBUG][Show-FirewallConfigDialog] FIN"
}
Export-ModuleMember -Function @('Show-SystemComponents', 'Start-SystemUpdate', 'show-NSPrinters', 'Invoke-ClearPrintJobs', 'Show-AddUserDialog', 'Show-IPConfigDialog', 'Show-LZMADialog', 'Show-FirewallConfigDialog')