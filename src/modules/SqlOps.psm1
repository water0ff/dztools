if ($PSVersionTable.PSVersion.Major -lt 5) { throw "Se requiere PowerShell 5.0 o superior." }
function Show-RestoreDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][string]$User, [Parameter(Mandatory = $true)][string]$Password, [Parameter(Mandatory = $true)][string]$Database, [Parameter(Mandatory = $false)][scriptblock]$OnRestoreCompleted)
    $script:RestoreRunning = $false
    $script:RestoreDone = $false
    $defaultPath = "C:\NationalSoft\DATABASES"
    if (-not (Test-Path -Path $defaultPath)) {
        New-Item -Path $defaultPath -ItemType Directory -Force | Out-Null
        Write-Host "Directorio creado: $defaultPath" -ForegroundColor Green
    } else {
        Write-DzDebug "`t[DEBUG][Show-RestoreDialog] El directorio $defaultPath ya existe" -Color DarkYellow
    }
    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null }
    function Ui-Warn([string]$m, [string]$t = "Atención", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Warning" -Owner $o | Out-Null }
    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null }
    function Ui-Confirm([string]$m, [string]$t = "Confirmar", [System.Windows.Window]$o) { (Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "YesNo" -Icon "Question" -Owner $o) -eq [System.Windows.MessageBoxResult]::Yes }
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] INICIO"
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Server='$Server' Database='$Database' User='$User'"
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    $theme = Get-DzUiTheme
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Opciones de Restauración" Height="540" Width="620" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="$($theme.AccentSecondary)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20" Background="$($theme.FormBackground)">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label Grid.Row="0" Content="Archivo de respaldo (.bak):"/>
        <Grid Grid.Row="1" Margin="0,5,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtBackupPath" Grid.Column="0" Height="25"/>
            <Button x:Name="btnBrowseBackup" Grid.Column="1" Content="Examinar..." Width="90" Margin="5,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
        </Grid>
        <Label Grid.Row="2" Content="Nombre destino:"/>
        <TextBox x:Name="txtDestino" Grid.Row="3" Height="25" Margin="0,5,0,10"/>
        <Label Grid.Row="4" Content="Ruta MDF (datos):"/>
        <TextBox x:Name="txtMdfPath" Grid.Row="5" Height="25" Margin="0,5,0,10"/>
        <Label Grid.Row="6" Content="Ruta LDF (log):"/>
        <TextBox x:Name="txtLdfPath" Grid.Row="7" Height="25" Margin="0,5,0,10"/>
        <GroupBox Grid.Row="8" Header="Progreso" Margin="0,0,0,10">
            <StackPanel>
                <ProgressBar x:Name="pbRestore" Height="20" Margin="5" Minimum="0" Maximum="100" Value="0"/>
                <TextBlock x:Name="txtProgress" Text="Esperando..." Margin="5,5,5,10" TextWrapping="Wrap"/>
            </StackPanel>
        </GroupBox>
        <GroupBox Grid.Row="9" Header="Log">
            <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Height="140"/>
        </GroupBox>
        <StackPanel Grid.Row="10" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnAceptar" Content="Iniciar Restauración" Width="140" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
            <Button x:Name="btnCerrar" Content="Cerrar" Width="80" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) { Write-DzDebug "`t[DEBUG][Show-RestoreDialog] ERROR: window=NULL"; throw "No se pudo crear la ventana (XAML)." }
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Ventana creada OK"
    $txtBackupPath = $window.FindName("txtBackupPath")
    $btnBrowseBackup = $window.FindName("btnBrowseBackup")
    $txtDestino = $window.FindName("txtDestino")
    $txtMdfPath = $window.FindName("txtMdfPath")
    $txtLdfPath = $window.FindName("txtLdfPath")
    $pbRestore = $window.FindName("pbRestore")
    $txtProgress = $window.FindName("txtProgress")
    $txtLog = $window.FindName("txtLog")
    $btnAceptar = $window.FindName("btnAceptar")
    $btnCerrar = $window.FindName("btnCerrar")
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Controles: txtBackupPath=$([bool]$txtBackupPath) btnBrowseBackup=$([bool]$btnBrowseBackup) txtDestino=$([bool]$txtDestino) txtMdfPath=$([bool]$txtMdfPath) txtLdfPath=$([bool]$txtLdfPath) pbRestore=$([bool]$pbRestore) txtProgress=$([bool]$txtProgress) txtLog=$([bool]$txtLog) btnAceptar=$([bool]$btnAceptar) btnCerrar=$([bool]$btnCerrar)"
    if (-not $txtBackupPath -or -not $btnBrowseBackup -or -not $txtDestino -or -not $txtMdfPath -or -not $txtLdfPath -or -not $pbRestore -or -not $txtProgress -or -not $txtLog -or -not $btnAceptar -or -not $btnCerrar) { Write-DzDebug "`t[DEBUG][Show-RestoreDialog] ERROR: controles NULL"; throw "Controles WPF incompletos (FindName devolvió NULL)." }
    $defaultRestoreFolder = "C:\NationalSoft\DATABASES"
    $txtDestino.Text = $Database
    function Normalize-RestoreFolder {
        param([string]$BasePath)
        if ([string]::IsNullOrWhiteSpace($BasePath)) { return $BasePath }
        $trimmed = $BasePath.Trim()
        if ($trimmed.EndsWith('\')) { return $trimmed.TrimEnd('\') }
        $trimmed
    }
    function Update-RestorePaths {
        param([string]$DatabaseName)
        if ([string]::IsNullOrWhiteSpace($DatabaseName)) { return }
        $baseFolder = Normalize-RestoreFolder -BasePath $defaultRestoreFolder
        $txtMdfPath.Text = Join-Path $baseFolder "$DatabaseName.mdf"
        $txtLdfPath.Text = Join-Path $baseFolder "$DatabaseName.ldf"
    }
    Update-RestorePaths -DatabaseName $txtDestino.Text
    $logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $progressQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'
    function Paint-Progress { param([int]$Percent, [string]$Message) $pbRestore.Value = $Percent; $txtProgress.Text = $Message }
    function Add-Log { param([string]$Message) $logQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)) }
    function New-SafeCredential { param([string]$Username, [string]$PlainPassword) $secure = New-Object System.Security.SecureString; foreach ($ch in $PlainPassword.ToCharArray()) { $secure.AppendChar($ch) }; $secure.MakeReadOnly(); New-Object System.Management.Automation.PSCredential($Username, $secure) }
    function Start-RestoreWorkAsync {
        param(
            [string]$Server,
            [string]$DatabaseName,
            [string]$RestoreQuery,
            [System.Management.Automation.PSCredential]$Credential,
            [System.Collections.Concurrent.ConcurrentQueue[string]]$LogQueue,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$ProgressQueue
        )
        Write-DzDebug "`t[DEBUG][Start-RestoreWorkAsync] Preparando runspace..."
        $worker = {
            param($Server, $DatabaseName, $RestoreQuery, $Credential, $LogQueue, $ProgressQueue)
            function EnqLog([string]$m) { $LogQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) }
            function EnqProg([int]$p, [string]$m) { $ProgressQueue.Enqueue(@{Percent = $p; Message = $m }) }
            function Invoke-SqlQueryLite {
                param([string]$Server, [string]$Database, [string]$Query, [System.Management.Automation.PSCredential]$Credential, [scriptblock]$InfoMessageCallback)
                $connection = $null
                $passwordBstr = [IntPtr]::Zero
                $plainPassword = $null
                $hasError = $false
                $errorMessages = @()
                try {
                    $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
                    $cs = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
                    $connection = New-Object System.Data.SqlClient.SqlConnection($cs)
                    $state = [pscustomobject]@{ HasError = $false; Errors = New-Object System.Collections.Generic.List[string] }
                    if ($InfoMessageCallback) {
                        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
                            param($sender, $e)
                            try {
                                $msg = [string]$e.Message
                                if ([string]::IsNullOrWhiteSpace($msg)) { return }
                                $isProgressMsg = $msg -match '(?i)\b\d{1,3}\s*(%|percent|porcentaje|por\s+ciento)\b'
                                $isSuccessMsg = $msg -match '(?i)(successfully processed|procesad[oa]\s+correctamente|se proces[oó]\s+correctamente)'
                                if (-not $isProgressMsg -and -not $isSuccessMsg) {
                                    if ($msg -match '(?i)(abnormal termination|not compatible|no es compatible|cannot restore|failed|restore failed|imposible|\berror\b)') {
                                        $state.HasError = $true
                                        $null = $state.Errors.Add($msg)
                                    }
                                }
                                & $InfoMessageCallback $msg $state.HasError
                            } catch { }
                        }
                        $connection.add_InfoMessage($handler)
                        $connection.FireInfoMessageEventOnUserErrors = $true
                    }
                    $connection.Open()
                    $cmd = $connection.CreateCommand()
                    $cmd.CommandText = $Query
                    $cmd.CommandTimeout = 0
                    [void]$cmd.ExecuteNonQuery()
                    if ($state.HasError) {
                        $combinedErrors = $state.Errors -join "`n"
                        return @{ Success = $false; ErrorMessage = $combinedErrors; IsInfoMessageError = $true }
                    }
                    @{ Success = $true }
                } catch {
                    @{ Success = $false; ErrorMessage = $_.Exception.Message; IsInfoMessageError = $false }
                } finally {
                    if ($plainPassword) { $plainPassword = $null }
                    if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                    if ($connection) { try { $connection.Close() } catch { } ; try { $connection.Dispose() } catch { } }
                }
            }
            try {
                EnqLog "Enviando comando de restauración a SQL Server..."
                EnqProg 10 "Iniciando restauración..."
                $hasErrorFlag = $false
                $errorDetails = ""
                $allErrors = @()
                $progressCb = {
                    param([string]$Message, [bool]$IsError)
                    $m = ($Message -replace '\s+', ' ').Trim()
                    if (-not $m) { return }
                    if ($m -match '(?i)\b(\d{1,3})\s*(%|percent|porcentaje|por\s+ciento)\b') {
                        $p = [int]$Matches[1]
                        $p = [math]::Max(0, [math]::Min(100, $p))
                        EnqProg $p ("Progreso restauración: $p%")
                        EnqLog ("[SQL] Progreso: $p%  ($m)")
                        return
                    }
                    if ($m -match '(?i)\b(processed\s+\d{1,3}\s*percent\b|se proces[oó] correctamente|completad[oa])') {
                        EnqProg 100 "Restauración completada."
                        EnqLog "✅ Restauración finalizada"
                        return
                    }
                    $isCriticalError = $m -match '(?i)(abnormal termination|failed to|error|cannot restore|not compatible|no es compatible|imposible|terminado an[oó]malo)'
                    if ($IsError -or $isCriticalError) {
                        $script:hasErrorFlag = $true
                        $script:allErrors += $m
                        if (-not $script:errorDetails -or $m.Length -gt $script:errorDetails.Length) { $script:errorDetails = $m }
                        EnqLog ("[SQL ERROR] $m")
                        return
                    }
                    EnqLog ("[SQL] $m")
                }
                $r = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $RestoreQuery -Credential $Credential -InfoMessageCallback $progressCb
                if (-not $r.Success) {
                    EnqProg 0 "❌ Error en restauración"
                    EnqLog ("❌ Error de SQL: {0}" -f $r.ErrorMessage)
                    EnqLog "ERROR_RESULT|$($r.ErrorMessage)"
                    EnqLog "__DONE__"
                    return
                }
                if ($hasErrorFlag) {
                    EnqProg 0 "❌ Restauración falló"
                    EnqLog "❌ La restauración no se completó correctamente"
                    if ($allErrors.Count -gt 0) {
                        if ($allErrors.Count -le 3) {
                            $finalErrorMsg = $allErrors -join "`n`n"
                        } else {
                            $uniqueErrors = $allErrors | Select-Object -Unique | Select-Object -First 3
                            $finalErrorMsg = $uniqueErrors -join "`n`n"
                        }
                        EnqLog "ERROR_RESULT|$finalErrorMsg"
                    } else {
                        EnqLog "ERROR_RESULT|Error detectado durante la restauración. Revisa el log para más detalles."
                    }
                    EnqLog "__DONE__"
                    return
                }
                EnqProg 100 "Restauración finalizada."
                EnqLog "✅ Comando RESTORE finalizó correctamente"
                EnqLog "SUCCESS_RESULT|Base de datos restaurada exitosamente"
                EnqLog "__DONE__"
            } catch {
                EnqProg 0 "Error"
                EnqLog ("❌ Error inesperado (worker): {0}" -f $_.Exception.Message)
                EnqLog "ERROR_RESULT|$($_.Exception.Message)"
                EnqLog "__DONE__"
            }
        }
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($worker).AddArgument($Server).AddArgument($DatabaseName).AddArgument($RestoreQuery).AddArgument($Credential).AddArgument($LogQueue).AddArgument($ProgressQueue)
        $null = $ps.BeginInvoke()
        Write-DzDebug "`t[DEBUG][Start-RestoreWorkAsync] Worker lanzado"
    }
    $logTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $logTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $logTimer.Add_Tick({
            try {
                $count = 0
                $doneThisTick = $false
                $finalResult = $null
                while ($count -lt 50) {
                    $line = $null
                    if (-not $logQueue.TryDequeue([ref]$line)) { break }
                    if ($line -like "*SUCCESS_RESULT|*") { $finalResult = @{ Success = $true; Message = $line -replace '^.*SUCCESS_RESULT\|', '' } }
                    if ($line -like "*ERROR_RESULT|*") { $finalResult = @{ Success = $false; Message = $line -replace '^.*ERROR_RESULT\|', '' } }
                    if ($line -notmatch '(SUCCESS_RESULT|ERROR_RESULT)') { $txtLog.Text = "$line`n" + $txtLog.Text }
                    if ($line -like "*__DONE__*") {
                        Write-DzDebug "`t[DEBUG][UI] Señal DONE recibida (restore)"
                        $doneThisTick = $true
                        $script:RestoreRunning = $false
                        $btnAceptar.IsEnabled = $true
                        $btnAceptar.Content = "Iniciar Restauración"
                        $txtBackupPath.IsEnabled = $true
                        $btnBrowseBackup.IsEnabled = $true
                        $txtDestino.IsEnabled = $true
                        $txtMdfPath.IsEnabled = $true
                        $txtLdfPath.IsEnabled = $true
                        $tmp = $null
                        while ($progressQueue.TryDequeue([ref]$tmp)) { }
                        Paint-Progress -Percent 100 -Message "Completado"
                        $script:RestoreDone = $true
                        if ($finalResult) {
                            $window.Dispatcher.Invoke([action] {
                                    if ($finalResult.Success) {
                                        Ui-Info "Base de datos '$($txtDestino.Text)' restaurada con éxito.`n`n$($finalResult.Message)" "✓ Restauración Exitosa" $window
                                        if ($OnRestoreCompleted) {
                                            Write-DzDebug "`t[DEBUG][Show-RestoreDialog] OnRestoreCompleted: DB='$($txtDestino.Text)'"
                                            try { & $OnRestoreCompleted $txtDestino.Text } catch { Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Error OnRestoreCompleted: $($_.Exception.Message)" }
                                        }
                                    } else {
                                        Ui-Error "La restauración falló:`n`n$($finalResult.Message)" "✗ Error en Restauración" $window
                                    }
                                }, [System.Windows.Threading.DispatcherPriority]::Normal)
                        }
                    }
                    $count++
                }
                if ($count -gt 0) { $txtLog.ScrollToLine(0) }
                if (-not $doneThisTick) {
                    $last = $null
                    while ($true) {
                        $p = $null
                        if (-not $progressQueue.TryDequeue([ref]$p)) { break }
                        $last = $p
                    }
                    if ($last) { Paint-Progress -Percent $last.Percent -Message $last.Message }
                }
            } catch { Write-DzDebug "`t[DEBUG][UI][logTimer][restore] ERROR: $($_.Exception.Message)" }
            if ($script:RestoreDone) {
                $tmpLine = $null
                $tmpProg = $null
                if (-not $logQueue.TryPeek([ref]$tmpLine) -and -not $progressQueue.TryPeek([ref]$tmpProg)) { $logTimer.Stop(); $script:RestoreDone = $false }
            }
        })
    $logTimer.Start()
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] logTimer iniciado"
    $btnBrowseBackup.Add_Click({
            try {
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Filter = "SQL Backup (*.bak)|*.bak|Todos los archivos (*.*)|*.*"
                $dlg.Multiselect = $false
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $txtBackupPath.Text = $dlg.FileName
                    if ([string]::IsNullOrWhiteSpace($txtDestino.Text)) { $txtDestino.Text = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName) }
                }
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error btnBrowseBackup: $($_.Exception.Message)"
                Ui-Error "No se pudo abrir el selector de archivos: $($_.Exception.Message)" "Error" $window
            }
        })
    $txtDestino.Add_TextChanged({
            try { Update-RestorePaths -DatabaseName $txtDestino.Text } catch { Write-DzDebug "`t[DEBUG][UI] Error actualizando rutas: $($_.Exception.Message)" }
        })
    $btnAceptar.Add_Click({
            Write-DzDebug "`t[DEBUG][UI] btnAceptar Restore Click"
            if ($script:RestoreRunning) { return }
            $script:RestoreDone = $false
            if (-not $logTimer.IsEnabled) { $logTimer.Start() }
            try {
                $btnAceptar.IsEnabled = $false
                $btnAceptar.Content = "Procesando..."
                $txtLog.Text = ""
                $pbRestore.Value = 0
                $txtProgress.Text = "Esperando..."
                Add-Log "Iniciando proceso de restauración..."
                $backupPath = $txtBackupPath.Text.Trim()
                $destName = $txtDestino.Text.Trim()
                $mdfPath = $txtMdfPath.Text.Trim()
                $ldfPath = $txtLdfPath.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($backupPath)) { Ui-Warn "Selecciona el archivo .bak a restaurar." "Atención" $window; Reset-RestoreUI -ProgressText "Archivo de respaldo requerido"; return }
                if ([string]::IsNullOrWhiteSpace($destName)) { Ui-Warn "Indica el nombre destino de la base de datos." "Atención" $window; Reset-RestoreUI -ProgressText "Nombre destino requerido"; return }
                if ([string]::IsNullOrWhiteSpace($mdfPath) -or [string]::IsNullOrWhiteSpace($ldfPath)) { Ui-Warn "Indica las rutas de destino para MDF y LDF." "Atención" $window; Reset-RestoreUI -ProgressText "Rutas MDF/LDF requeridas"; return }
                Add-Log "Servidor: $Server"
                Add-Log "Base de datos destino: $destName"
                Add-Log "Backup: $backupPath"
                Add-Log "MDF: $mdfPath"
                Add-Log "LDF: $ldfPath"
                $credential = New-SafeCredential -Username $User -PlainPassword $Password
                Add-Log "✓ Credenciales listas"
                $escapedBackup = $backupPath -replace "'", "''"
                $escapedMdf = $mdfPath -replace "'", "''"
                $escapedLdf = $ldfPath -replace "'", "''"
                $escapedDest = $destName -replace "'", "''"
                $destNameSafe = $destName -replace ']', ']]'
                Paint-Progress -Percent 5 -Message "Leyendo metadatos del backup..."
                $fileListQuery = "RESTORE FILELISTONLY FROM DISK = N'$escapedBackup'"
                function Get-BackupFileList {
                    param([string]$Server, [string]$Query, [System.Management.Automation.PSCredential]$Credential)
                    $connection = $null
                    $passwordBstr = [IntPtr]::Zero
                    $plainPassword = $null
                    try {
                        $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
                        $cs = "Server=$Server;Database=master;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
                        $connection = New-Object System.Data.SqlClient.SqlConnection($cs)
                        $connection.Open()
                        $cmd = $connection.CreateCommand()
                        $cmd.CommandText = $Query
                        $cmd.CommandTimeout = 60
                        $reader = $cmd.ExecuteReader()
                        $dataTable = New-Object System.Data.DataTable
                        $dataTable.Load($reader)
                        $reader.Close()
                        return @{ Success = $true; DataTable = $dataTable }
                    } catch {
                        return @{ Success = $false; ErrorMessage = $_.Exception.Message; DataTable = $null }
                    } finally {
                        if ($plainPassword) { $plainPassword = $null }
                        if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                        if ($connection) { try { $connection.Close() } catch { } ; try { $connection.Dispose() } catch { } }
                    }
                }
                $fileListResult = Get-BackupFileList -Server $Server -Query $fileListQuery -Credential $credential
                if (-not $fileListResult.Success) {
                    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Error FILELISTONLY: $($fileListResult.ErrorMessage)"
                    Ui-Error "No se pudo leer el contenido del backup: $($fileListResult.ErrorMessage)" "Error" $window
                    Reset-RestoreUI -ProgressText "Error leyendo backup"
                    return
                }
                if (-not $fileListResult.DataTable -or $fileListResult.DataTable.Rows.Count -eq 0) {
                    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] FILELISTONLY no devolvió filas"
                    Ui-Error "El archivo de backup no contiene información de archivos válida." "Error" $window
                    Reset-RestoreUI -ProgressText "Backup sin información"
                    return
                }
                Add-Log "Archivos en backup: $($fileListResult.DataTable.Rows.Count)"
                $logicalData = $null
                $logicalLog = $null
                foreach ($row in $fileListResult.DataTable.Rows) {
                    $type = [string]$row["Type"]
                    $logicalName = [string]$row["LogicalName"]
                    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Archivo: LogicalName='$logicalName' Type='$type'"
                    if (-not $logicalData -and $type -eq "D") { $logicalData = $logicalName; Add-Log "Encontrado archivo de datos: $logicalName" }
                    if (-not $logicalLog -and $type -eq "L") { $logicalLog = $logicalName; Add-Log "Encontrado archivo de log: $logicalName" }
                }
                if (-not $logicalData -or -not $logicalLog) {
                    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Logical names missing. Data='$logicalData' Log='$logicalLog'"
                    Ui-Error "No se encontraron nombres lógicos válidos en el backup.`n`nData: $logicalData`nLog: $logicalLog" "Error" $window
                    Reset-RestoreUI -ProgressText "Error en nombres lógicos"
                    return
                }
                Add-Log ("✓ Logical Data: {0}" -f $logicalData)
                Add-Log ("✓ Logical Log: {0}" -f $logicalLog)
                $restoreQuery = @"
IF DB_ID(N'$escapedDest') IS NOT NULL
BEGIN
    ALTER DATABASE [$destNameSafe] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
END
RESTORE DATABASE [$destNameSafe]
FROM DISK = N'$escapedBackup'
WITH MOVE N'$logicalData' TO N'$escapedMdf',
     MOVE N'$logicalLog' TO N'$escapedLdf',
     REPLACE, RECOVERY, STATS = 1;
IF DB_ID(N'$escapedDest') IS NOT NULL
BEGIN
    ALTER DATABASE [$destNameSafe] SET MULTI_USER;
END
"@
                Paint-Progress -Percent 10 -Message "Conectando a SQL Server..."
                Write-DzDebug "`t[DEBUG][UI] Llamando Start-RestoreWorkAsync"
                Start-RestoreWorkAsync -Server $Server -DatabaseName $destName -RestoreQuery $restoreQuery -Credential $credential -LogQueue $logQueue -ProgressQueue $progressQueue
                $script:RestoreRunning = $true
                $txtBackupPath.IsEnabled = $false
                $btnBrowseBackup.IsEnabled = $false
                $txtDestino.IsEnabled = $false
                $txtMdfPath.IsEnabled = $false
                $txtLdfPath.IsEnabled = $false
            } catch {
                Write-DzDebug "`t[DEBUG][UI] ERROR btnAceptar Restore: $($_.Exception.Message)"
                Add-Log "❌ Error: $($_.Exception.Message)"
                Reset-RestoreUI -ProgressText "Error inesperado"
            }
        })
    $btnCerrar.Add_Click({
            Write-DzDebug "`t[DEBUG][UI] btnCerrar Restore Click"
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch { }
            $window.DialogResult = $false
            $window.Close()
        })
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Antes de ShowDialog()"
    $null = $window.ShowDialog()
    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Después de ShowDialog()"
}
function Reset-RestoreUI {
    param([string]$ButtonText = "Iniciar Restauración", [string]$ProgressText = "Esperando...")
    $script:RestoreRunning = $false
    $btnAceptar.IsEnabled = $true
    $btnAceptar.Content = $ButtonText
    $txtBackupPath.IsEnabled = $true
    $btnBrowseBackup.IsEnabled = $true
    $txtDestino.IsEnabled = $true
    $txtMdfPath.IsEnabled = $true
    $txtLdfPath.IsEnabled = $true
    $txtProgress.Text = $ProgressText
}
function Show-AttachDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$User,
        [Parameter(Mandatory = $true)][string]$Password,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$ModulesPath,
        [Parameter(Mandatory = $false)][scriptblock]$OnAttachCompleted
    )
    $script:AttachRunning = $false
    $script:AttachDone = $false
    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null }
    function Ui-Warn([string]$m, [string]$t = "Atención", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Warning" -Owner $o | Out-Null }
    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null }
    function Ui-Confirm([string]$m, [string]$t = "Confirmar", [System.Windows.Window]$o) { (Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "YesNo" -Icon "Question" -Owner $o) -eq [System.Windows.MessageBoxResult]::Yes }
    Write-DzDebug "`t[DEBUG][Show-AttachDialog] INICIO"
    Write-DzDebug "`t[DEBUG][Show-AttachDialog] Server='$Server' User='$User'"
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    $theme = Get-DzUiTheme
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Adjuntar base de datos" Height="600" Width="660" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="$($theme.AccentSecondary)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20" Background="$($theme.FormBackground)">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label Grid.Row="0" Content="Archivo MDF (datos):"/>
        <Grid Grid.Row="1" Margin="0,5,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtMdfPath" Grid.Column="0" Height="25"/>
            <Button x:Name="btnBrowseMdf" Grid.Column="1" Content="Examinar..." Width="90" Margin="5,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
        </Grid>
        <Label Grid.Row="2" Content="Archivo LDF (log):"/>
        <Grid Grid.Row="3" Margin="0,5,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtLdfPath" Grid.Column="0" Height="25"/>
            <Button x:Name="btnBrowseLdf" Grid.Column="1" Content="Examinar..." Width="90" Margin="5,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
        </Grid>
        <StackPanel Grid.Row="4" Margin="0,0,0,10">
            <CheckBox x:Name="chkRebuildLog" Content="Reconstruir archivo de log si no existe"/>
            <CheckBox x:Name="chkReadOnly" Content="Adjuntar como solo lectura"/>
        </StackPanel>
        <Label Grid.Row="5" Content="Nombre de la base de datos (Attach As):"/>
        <TextBox x:Name="txtDbName" Grid.Row="6" Height="25" Margin="0,5,0,10"/>
        <Label Grid.Row="7" Content="Owner (opcional):"/>
        <TextBox x:Name="txtOwner" Grid.Row="8" Height="25" Margin="0,5,0,10"/>
        <GroupBox Grid.Row="9" Header="Progreso" Margin="0,0,0,10">
            <StackPanel>
                <ProgressBar x:Name="pbAttach" Height="20" Margin="5" Minimum="0" Maximum="100" Value="0"/>
                <TextBlock x:Name="txtProgress" Text="Esperando..." Margin="5,5,5,10" TextWrapping="Wrap"/>
            </StackPanel>
        </GroupBox>
        <GroupBox Grid.Row="10" Header="Log">
            <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Height="140"/>
        </GroupBox>
        <StackPanel Grid.Row="11" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnAttach" Content="Adjuntar" Width="120" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
            <Button x:Name="btnClose" Content="Cerrar" Width="80" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) { Write-DzDebug "`t[DEBUG][Show-AttachDialog] ERROR: window=NULL"; throw "No se pudo crear la ventana (XAML)." }
    $txtMdfPath = $window.FindName("txtMdfPath")
    $btnBrowseMdf = $window.FindName("btnBrowseMdf")
    $txtLdfPath = $window.FindName("txtLdfPath")
    $btnBrowseLdf = $window.FindName("btnBrowseLdf")
    $chkRebuildLog = $window.FindName("chkRebuildLog")
    $chkReadOnly = $window.FindName("chkReadOnly")
    $txtDbName = $window.FindName("txtDbName")
    $txtOwner = $window.FindName("txtOwner")
    $pbAttach = $window.FindName("pbAttach")
    $txtProgress = $window.FindName("txtProgress")
    $txtLog = $window.FindName("txtLog")
    $btnAttach = $window.FindName("btnAttach")
    $btnClose = $window.FindName("btnClose")
    if (-not $txtMdfPath -or -not $btnBrowseMdf -or -not $txtLdfPath -or -not $btnBrowseLdf -or -not $chkRebuildLog -or -not $chkReadOnly -or -not $txtDbName -or -not $txtOwner -or -not $pbAttach -or -not $txtProgress -or -not $txtLog -or -not $btnAttach -or -not $btnClose) { Write-DzDebug "`t[DEBUG][Show-AttachDialog] ERROR: controles NULL"; throw "Controles WPF incompletos (FindName devolvió NULL)." }
    $logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $progressQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'
    function Paint-Progress { param([int]$Percent, [string]$Message) $pbAttach.Value = $Percent; $txtProgress.Text = $Message }
    function Add-Log { param([string]$Message) $logQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)) }
    function New-SafeCredential { param([string]$Username, [string]$PlainPassword) $secure = New-Object System.Security.SecureString; foreach ($ch in $PlainPassword.ToCharArray()) { $secure.AppendChar($ch) }; $secure.MakeReadOnly(); New-Object System.Management.Automation.PSCredential($Username, $secure) }
    function Set-LogControlsState { param([bool]$Enabled) $txtLdfPath.IsEnabled = $Enabled; $btnBrowseLdf.IsEnabled = $Enabled }
    Set-LogControlsState -Enabled $true
    $chkRebuildLog.Add_Checked({ Set-LogControlsState -Enabled $false })
    $chkRebuildLog.Add_Unchecked({ Set-LogControlsState -Enabled $true })
    function Start-AttachWorkAsync {
        param(
            [string]$Server,
            [string]$AttachQuery,
            [System.Management.Automation.PSCredential]$Credential,
            [System.Collections.Concurrent.ConcurrentQueue[string]]$LogQueue,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$ProgressQueue,
            [Parameter(Mandatory)][string]$UtilitiesModulePath,
            [Parameter(Mandatory)][string]$DatabaseModulePath
        )
        $worker = {
            param($Server, $AttachQuery, $Credential, $LogQueue, $ProgressQueue, $UtilitiesModulePath, $DatabaseModulePath)
            function EnqLog([string]$m) { $LogQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) }
            function EnqProg([int]$p, [string]$m) { $ProgressQueue.Enqueue(@{Percent = $p; Message = $m }) }
            try {
                if (-not (Test-Path -LiteralPath $DatabaseModulePath)) { throw "No se encontró el módulo requerido: $DatabaseModulePath" }
                Import-Module $UtilitiesModulePath -Force -DisableNameChecking -ErrorAction Stop
                Import-Module $DatabaseModulePath -Force -DisableNameChecking -ErrorAction Stop
                EnqProg 10 "Conectando a SQL Server..."
                EnqLog "Ejecutando ATTACH..."
                $r = Invoke-SqlQuery -Server $Server -Database "master" -Query $AttachQuery -Credential $Credential
                if (-not $r.Success) {
                    EnqProg 0 "❌ Error en adjuntar"
                    EnqLog ("❌ Error de SQL: {0}" -f $r.ErrorMessage)
                    EnqLog "ERROR_RESULT|$($r.ErrorMessage)"
                    EnqLog "__DONE__"
                    return
                }
                EnqProg 100 "Adjuntar completado."
                EnqLog "✅ Base de datos adjuntada"
                EnqLog "SUCCESS_RESULT|Base de datos adjuntada exitosamente"
                EnqLog "__DONE__"
            } catch {
                EnqProg 0 "Error"
                EnqLog ("❌ Error inesperado (worker): {0}" -f $_.Exception.Message)
                EnqLog "ERROR_RESULT|$($_.Exception.Message)"
                EnqLog "__DONE__"
            }
        }
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($worker).AddArgument($Server).AddArgument($AttachQuery).AddArgument($Credential).AddArgument($LogQueue).AddArgument($ProgressQueue).AddArgument($UtilitiesModulePath).AddArgument($DatabaseModulePath)
        $null = $ps.BeginInvoke()
    }
    $logTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $logTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $logTimer.Add_Tick({
            try {
                $count = 0
                $doneThisTick = $false
                $finalResult = $null
                while ($count -lt 50) {
                    $line = $null
                    if (-not $logQueue.TryDequeue([ref]$line)) { break }
                    if ($line -like "*SUCCESS_RESULT|*") { $finalResult = @{ Success = $true; Message = $line -replace '^.*SUCCESS_RESULT\|', '' } }
                    if ($line -like "*ERROR_RESULT|*") { $finalResult = @{ Success = $false; Message = $line -replace '^.*ERROR_RESULT\|', '' } }
                    if ($line -notmatch '(SUCCESS_RESULT|ERROR_RESULT)') { $txtLog.Text = "$line`n" + $txtLog.Text }
                    if ($line -like "*__DONE__*") {
                        $doneThisTick = $true
                        $script:AttachRunning = $false
                        $btnAttach.IsEnabled = $true
                        $btnAttach.Content = "Adjuntar"
                        $tmp = $null
                        while ($progressQueue.TryDequeue([ref]$tmp)) { }
                        Paint-Progress -Percent 100 -Message "Completado"
                        $script:AttachDone = $true
                        if ($finalResult) {
                            $window.Dispatcher.Invoke([action] {
                                    if ($finalResult.Success) {
                                        Ui-Info "Base de datos '$($txtDbName.Text)' adjuntada con éxito.`n`n$($finalResult.Message)" "✓ Adjuntar exitoso" $window
                                        if ($OnAttachCompleted) { try { & $OnAttachCompleted $txtDbName.Text } catch { Write-DzDebug "`t[DEBUG][Show-AttachDialog] Error OnAttachCompleted: $($_.Exception.Message)" } }
                                    } else {
                                        Write-DzDebug "`t[DEBUG][Adjuntar] Adjuntar falló: $($finalResult.Message)"
                                        Ui-Error "No se pudo adjuntar la base de datos:`n`n$($finalResult.Message)" "✗ Error al adjuntar" $window
                                    }
                                }, [System.Windows.Threading.DispatcherPriority]::Normal)
                        }
                    }
                    $count++
                }
                if ($count -gt 0) { $txtLog.ScrollToLine(0) }
                if (-not $doneThisTick) {
                    $last = $null
                    while ($true) {
                        $p = $null
                        if (-not $progressQueue.TryDequeue([ref]$p)) { break }
                        $last = $p
                    }
                    if ($last) { Paint-Progress -Percent $last.Percent -Message $last.Message }
                }
            } catch { Write-DzDebug "`t[DEBUG][UI][logTimer][attach] ERROR: $($_.Exception.Message)" }
            if ($script:AttachDone) {
                $tmpLine = $null
                $tmpProg = $null
                if (-not $logQueue.TryPeek([ref]$tmpLine) -and -not $progressQueue.TryPeek([ref]$tmpProg)) { $logTimer.Stop(); $script:AttachDone = $false }
            }
        })
    $logTimer.Start()
    $btnBrowseMdf.Add_Click({
            try {
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Filter = "SQL Data (*.mdf)|*.mdf|Todos los archivos (*.*)|*.*"
                $dlg.Multiselect = $false
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $txtMdfPath.Text = $dlg.FileName
                    if ([string]::IsNullOrWhiteSpace($txtDbName.Text)) { $txtDbName.Text = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName) }
                    if (-not $chkRebuildLog.IsChecked -and [string]::IsNullOrWhiteSpace($txtLdfPath.Text)) {
                        $dir = [System.IO.Path]::GetDirectoryName($dlg.FileName)
                        $base = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
                        $candidate = Join-Path $dir "$base`_log.ldf"
                        if (-not (Test-Path -LiteralPath $candidate)) { $candidate = [System.IO.Path]::ChangeExtension($dlg.FileName, ".ldf") }
                        $txtLdfPath.Text = $candidate
                    }
                }
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error btnBrowseMdf: $($_.Exception.Message)"
                Ui-Error "No se pudo abrir el selector de archivos: $($_.Exception.Message)" "Error" $window
            }
        })
    $btnBrowseLdf.Add_Click({
            try {
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Filter = "SQL Log (*.ldf)|*.ldf|Todos los archivos (*.*)|*.*"
                $dlg.Multiselect = $false
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtLdfPath.Text = $dlg.FileName }
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error btnBrowseLdf: $($_.Exception.Message)"
                Ui-Error "No se pudo abrir el selector de archivos: $($_.Exception.Message)" "Error" $window
            }
        })
    $btnAttach.Add_Click({
            Write-DzDebug "`t[DEBUG][UI] btnAttach Click"
            if ($script:AttachRunning) { return }
            $script:AttachDone = $false
            if (-not $logTimer.IsEnabled) { $logTimer.Start() }
            try {
                $btnAttach.IsEnabled = $false
                $btnAttach.Content = "Procesando..."
                $txtLog.Text = ""
                $pbAttach.Value = 0
                $txtProgress.Text = "Esperando..."
                Add-Log "Iniciando proceso de adjuntar..."
                $mdfPath = $txtMdfPath.Text.Trim()
                $ldfPath = $txtLdfPath.Text.Trim()
                $dbName = $txtDbName.Text.Trim()
                $owner = $txtOwner.Text.Trim()
                $rebuildLog = $chkRebuildLog.IsChecked -eq $true
                $readOnly = $chkReadOnly.IsChecked -eq $true
                if ([string]::IsNullOrWhiteSpace($mdfPath)) { Ui-Warn "Selecciona el archivo MDF a adjuntar." "Atención" $window; Reset-AttachUI -ProgressText "Archivo MDF requerido"; return }
                if (-not (Test-Path -LiteralPath $mdfPath)) { Ui-Warn "El archivo MDF no existe.`n`nRuta: $mdfPath" "Atención" $window; Reset-AttachUI -ProgressText "Archivo MDF no encontrado"; return }
                if ([string]::IsNullOrWhiteSpace($dbName)) { Ui-Warn "Indica el nombre de la base de datos (Attach As)." "Atención" $window; Reset-AttachUI -ProgressText "Nombre requerido"; return }
                if (-not $rebuildLog) {
                    if ([string]::IsNullOrWhiteSpace($ldfPath)) { Ui-Warn "Selecciona el archivo LDF o habilita la reconstrucción del log." "Atención" $window; Reset-AttachUI -ProgressText "Archivo LDF requerido"; return }
                    if (-not (Test-Path -LiteralPath $ldfPath)) { Ui-Warn "El archivo LDF no existe.`n`nRuta: $ldfPath" "Atención" $window; Reset-AttachUI -ProgressText "Archivo LDF no encontrado"; return }
                }
                $credential = New-SafeCredential -Username $User -PlainPassword $Password
                $safeDbName = $dbName -replace ']', ']]'
                $escapedDb = $dbName -replace "'", "''"
                $checkQuery = "SELECT 1 FROM sys.databases WHERE name = N'$escapedDb'"
                $check = Invoke-SqlQuery -Server $Server -Database "master" -Query $checkQuery -Credential $credential
                if ($check.Success -and $check.DataTable -and $check.DataTable.Rows.Count -gt 0) {
                    Ui-Error "Ya existe una base de datos con ese nombre.`n`nNombre: $dbName" "Error" $window
                    Reset-AttachUI -ProgressText "Nombre ya existe"
                    return
                }
                $escapedMdf = $mdfPath -replace "'", "''"
                $query = "CREATE DATABASE [$safeDbName] ON (FILENAME = N'$escapedMdf')"
                if (-not $rebuildLog) {
                    $escapedLdf = $ldfPath -replace "'", "''"
                    $query += ", (FILENAME = N'$escapedLdf')"
                    $query += " FOR ATTACH;"
                } else {
                    $query += " FOR ATTACH_REBUILD_LOG;"
                }
                if ($readOnly) { $query += "`nALTER DATABASE [$safeDbName] SET READ_ONLY WITH NO_WAIT;" }
                if (-not [string]::IsNullOrWhiteSpace($owner)) {
                    $safeOwner = $owner -replace ']', ']]'
                    $query += "`nALTER AUTHORIZATION ON DATABASE::[$safeDbName] TO [$safeOwner];"
                }
                Paint-Progress -Percent 10 -Message "Conectando a SQL Server..."
                $dbModulePath = Join-Path $ModulesPath "Database.psm1"
                $utilModulePath = Join-Path $ModulesPath "Utilities.psm1"
                $dbModulePath = Join-Path -Path $ModulesPath -ChildPath "Database.psm1"
                Add-Log "ModulesPath: '$ModulesPath'"
                Add-Log "DatabaseModulePath: '$dbModulePath'"
                if ([string]::IsNullOrWhiteSpace($ModulesPath)) {
                    Ui-Error "ModulesPath viene vacío/null. No se puede continuar." "Error" $window
                    Reset-AttachUI -ProgressText "ModulesPath inválido"
                    return
                }
                if ([string]::IsNullOrWhiteSpace($dbModulePath)) {
                    Ui-Error "DatabaseModulePath viene vacío/null. No se puede continuar." "Error" $window
                    Reset-AttachUI -ProgressText "Ruta de módulo inválida"
                    return
                }
                if (-not (Test-Path -LiteralPath $dbModulePath)) {
                    Ui-Error "No se encontró Database.psm1 en:`n$dbModulePath" "Error" $window
                    Reset-AttachUI -ProgressText "Módulo no encontrado"
                    return
                }
                Start-AttachWorkAsync -Server $Server -AttachQuery $query -Credential $credential -LogQueue $logQueue -ProgressQueue $progressQueue -DatabaseModulePath $dbModulePath -UtilitiesModulePath $utilModulePath
                $script:AttachRunning = $true
            } catch {
                Write-DzDebug "`t[DEBUG][UI] ERROR btnAttach: $($_.Exception.Message)"
                Add-Log "❌ Error: $($_.Exception.Message)"
                Reset-AttachUI -ProgressText "Error inesperado"
            }
        })
    $btnClose.Add_Click({
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch { }
            $window.DialogResult = $false
            $window.Close()
        })
    $null = $window.ShowDialog()
}
function Reset-AttachUI {
    param([string]$ButtonText = "Adjuntar", [string]$ProgressText = "Esperando...")
    $script:AttachRunning = $false
    $btnAttach.IsEnabled = $true
    $btnAttach.Content = $ButtonText
    $txtProgress.Text = $ProgressText
}
function Show-BackupDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][string]$User, [Parameter(Mandatory = $true)][string]$Password, [Parameter(Mandatory = $true)][string]$Database)
    $script:BackupRunning = $false
    $script:BackupDone = $false
    $script:EnableThreadJob = $false
    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null }
    function Ui-Warn([string]$m, [string]$t = "Atención", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Warning" -Owner $o | Out-Null }
    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null }
    function Ui-Confirm([string]$m, [string]$t = "Confirmar", [System.Windows.Window]$o) { (Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "YesNo" -Icon "Question" -Owner $o) -eq [System.Windows.MessageBoxResult]::Yes }
    function Initialize-ThreadJob {
        [CmdletBinding()]
        param()
        Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Verificando módulo ThreadJob"
        if (Get-Module -ListAvailable -Name ThreadJob) {
            Import-Module ThreadJob -Force
            Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob importado"
            return $true
        } else {
            Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob no encontrado, intentando instalar"
            try {
                Install-Module -Name ThreadJob -Force -Scope CurrentUser -ErrorAction Stop
                Import-Module ThreadJob -Force
                Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Módulo ThreadJob instalado e importado"
                return $true
            } catch {
                Write-DzDebug "`t[DEBUG][Initialize-ThreadJob] Error instalando ThreadJob: $_"
                return $false
            }
        }
    }
    if ($script:EnableThreadJob) { if (-not (Initialize-ThreadJob)) { Write-Host "Advertencia: No se pudo cargar ThreadJob..." -ForegroundColor Yellow } }
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] INICIO"
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] Server='$Server' Database='$Database' User='$User'"
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    $theme = Get-DzUiTheme
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Opciones de Respaldo" Height="500" Width="600" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="PasswordBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="$($theme.AccentSecondary)"/>
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="20" Background="$($theme.FormBackground)"><Grid.RowDefinitions><RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <CheckBox x:Name="chkRespaldo" Grid.Row="0" IsChecked="True" IsEnabled="False" Margin="0,0,0,10"><TextBlock Text="Respaldar" FontWeight="Bold"/>
    </CheckBox><Label Grid.Row="1" Content="Nombre del respaldo:"/><TextBox x:Name="txtNombre" Grid.Row="2" Height="25" Margin="0,5,0,10"/>
    <CheckBox x:Name="chkComprimir" Grid.Row="3" Margin="0,0,0,10"><TextBlock Text="Comprimir (requiere Chocolatey)" FontWeight="Bold"/>
    </CheckBox><Label x:Name="lblPassword" Grid.Row="4" Content="Contraseña (opcional) para ZIP:"/><Grid Grid.Row="5" Margin="0,5,0,10">
    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions><PasswordBox x:Name="txtPassword" Grid.Column="0" Height="25"/>
    <Button x:Name="btnTogglePassword" Grid.Column="1" Content="👁" Width="30" Margin="5,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
    </Grid><CheckBox x:Name="chkSubir" Grid.Row="6" Margin="0,0,0,20" IsEnabled="False">
    <TextBlock Text="Subir a Mega.nz (opción deshabilitada)" FontWeight="Bold" Foreground="$($theme.FormForeground)"/>
    </CheckBox>
    <GroupBox Grid.Row="7" Header="Progreso" Margin="0,0,0,10">
    <StackPanel><ProgressBar x:Name="pbBackup" Height="20" Margin="5" Minimum="0" Maximum="100" Value="0"/>
    <TextBlock x:Name="txtProgress" Text="Esperando..." Margin="5,5,5,10" TextWrapping="Wrap"/>
    </StackPanel></GroupBox><GroupBox Grid.Row="8" Header="Log">
    <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Height="160"/>
    </GroupBox><StackPanel Grid.Row="9" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0"><Button x:Name="btnAceptar" Content="Iniciar Respaldo" Width="120" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
    <Button x:Name="btnAbrirCarpeta" Content="Abrir Carpeta" Width="100" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/>
    <Button x:Name="btnCerrar" Content="Cerrar" Width="80" Height="30" Margin="5,0" Style="{StaticResource SystemButtonStyle}"/></StackPanel></Grid>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    if (-not $window) { Write-DzDebug "`t[DEBUG][Show-BackupDialog] ERROR: window=NULL"; throw "No se pudo crear la ventana (XAML)." }
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] Ventana creada OK"
    $chkRespaldo = $window.FindName("chkRespaldo")
    $txtNombre = $window.FindName("txtNombre")
    $chkComprimir = $window.FindName("chkComprimir")
    $txtPassword = $window.FindName("txtPassword")
    $lblPassword = $window.FindName("lblPassword")
    $chkSubir = $window.FindName("chkSubir")
    $pbBackup = $window.FindName("pbBackup")
    $txtProgress = $window.FindName("txtProgress")
    $txtLog = $window.FindName("txtLog")
    $btnAceptar = $window.FindName("btnAceptar")
    $btnAbrirCarpeta = $window.FindName("btnAbrirCarpeta")
    $btnCerrar = $window.FindName("btnCerrar")
    $btnTogglePassword = $window.FindName("btnTogglePassword")
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] Controles: chkRespaldo=$([bool]$chkRespaldo) txtNombre=$([bool]$txtNombre) chkComprimir=$([bool]$chkComprimir) txtPassword=$([bool]$txtPassword) lblPassword=$([bool]$lblPassword) chkSubir=$([bool]$chkSubir) pbBackup=$([bool]$pbBackup) txtProgress=$([bool]$txtProgress) txtLog=$([bool]$txtLog) btnAceptar=$([bool]$btnAceptar) btnAbrirCarpeta=$([bool]$btnAbrirCarpeta) btnCerrar=$([bool]$btnCerrar) btnTogglePassword=$([bool]$btnTogglePassword)"
    if (-not $txtNombre -or -not $chkComprimir -or -not $txtPassword -or -not $lblPassword -or -not $chkSubir -or -not $pbBackup -or -not $txtProgress -or -not $txtLog -or -not $btnAceptar -or -not $btnAbrirCarpeta -or -not $btnCerrar) { Write-DzDebug "`t[DEBUG][Show-BackupDialog] ERROR: uno o más controles son NULL. Cerrando..."; throw "Controles WPF incompletos (FindName devolvió NULL)." }
    $timestampsDefault = Get-Date -Format 'yyyyMMdd-HHmmss'
    $txtNombre.Text = ("$Database-$timestampsDefault.bak")
    $txtPassword.IsEnabled = $false
    $lblPassword.IsEnabled = $false
    $chkSubir.IsEnabled = $false
    $chkSubir.IsChecked = $false
    $logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $progressQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'
    function Paint-Progress { param([int]$Percent, [string]$Message) $pbBackup.Value = $Percent; $txtProgress.Text = $Message }
    function Add-Log { param([string]$Message) $logQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)) }
    function Disable-CompressionUI { $chkComprimir.IsChecked = $false; $txtPassword.IsEnabled = $false; $lblPassword.IsEnabled = $false; $txtPassword.Password = "" }
    function New-SafeCredential { param([string]$Username, [string]$PlainPassword) $secure = New-Object System.Security.SecureString; foreach ($ch in $PlainPassword.ToCharArray()) { $secure.AppendChar($ch) }; $secure.MakeReadOnly(); New-Object System.Management.Automation.PSCredential($Username, $secure) }
    function Start-BackupWorkAsync {
        param(
            [string]$Server,
            [string]$Database,
            [string]$BackupQuery,
            [string]$ScriptBackupPath,
            [bool]$DoCompress,
            [string]$ZipPassword,
            [System.Management.Automation.PSCredential]$Credential,
            [System.Collections.Concurrent.ConcurrentQueue[string]]$LogQueue,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$ProgressQueue
        )
        Write-DzDebug "`t[DEBUG][Start-BackupWorkAsync] Preparando runspace..."
        $worker = {
            param($Server, $Database, $BackupQuery, $ScriptBackupPath, $DoCompress, $ZipPassword, $Credential, $LogQueue, $ProgressQueue)
            function EnqLog([string]$m) { $LogQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) }
            function EnqProg([int]$p, [string]$m) { $ProgressQueue.Enqueue(@{Percent = $p; Message = $m }) }
            function Invoke-SqlQueryLite {
                param([string]$Server, [string]$Database, [string]$Query, [System.Management.Automation.PSCredential]$Credential, [scriptblock]$InfoMessageCallback)
                $connection = $null
                $passwordBstr = [IntPtr]::Zero
                $plainPassword = $null
                try {
                    $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
                    $cs = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
                    $connection = New-Object System.Data.SqlClient.SqlConnection($cs)
                    if ($InfoMessageCallback) { $connection.add_InfoMessage({ param($sender, $e) try { & $InfoMessageCallback $e.Message } catch { } }); $connection.FireInfoMessageEventOnUserErrors = $true }
                    $connection.Open()
                    $cmd = $connection.CreateCommand()
                    $cmd.CommandText = $Query
                    $cmd.CommandTimeout = 0
                    [void]$cmd.ExecuteNonQuery()
                    @{ Success = $true }
                } catch { @{ Success = $false; ErrorMessage = $_.Exception.Message } } finally {
                    if ($plainPassword) { $plainPassword = $null }
                    if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                    if ($connection) { try { $connection.Close() } catch { } ; try { $connection.Dispose() } catch { } }
                }
            }
            function Get-7ZipPath {
                $c = @("$env:ProgramFiles\7-Zip\7z.exe", "${env:ProgramFiles(x86)}\7-Zip\7z.exe")
                foreach ($p in $c) { if (Test-Path $p) { return $p } }
                $cmd = Get-Command 7z.exe -ErrorAction SilentlyContinue
                if ($cmd) { return $cmd.Source }
                $null
            }
            try {
                EnqLog "Enviando comando a SQL Server..."
                EnqProg 10 "Iniciando backup..."
                $progressCb = {
                    param([string]$Message)
                    $m = ($Message -replace '\s+', ' ').Trim()
                    if ($m) { EnqLog ("[SQL] {0}" -f $m) }
                    $backupMax = 100
                    if ($DoCompress) { $backupMax = 90 }
                    if ($Message -match '(?i)\b(\d{1,3})\s*(percent|porcentaje|por\s+ciento)\b') {
                        $p = [int]$Matches[1]
                        if ($p -gt 100) { $p = 100 }
                        if ($p -lt 0) { $p = 0 }
                        $scaled = [int][math]::Floor(($p * $backupMax) / 100)
                        if ($DoCompress -and $scaled -ge 100) { $scaled = 90 }
                        EnqProg $scaled ("Progreso backup: {0}%" -f $p)
                        EnqLog ("Progreso backup: {0}%" -f $p)
                        return
                    }
                    if ($Message -match '(?i)\b(successfully processed|procesad[oa]\s+correctamente|completad[oa])\b') {
                        if ($DoCompress) { EnqProg 90 "Backup listo. Iniciando compresión..." } else { EnqProg 100 "¡Backup completado!" }
                        EnqLog "✅ Backup completado (mensaje SQL)"
                        return
                    }
                }
                $r = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $BackupQuery -Credential $Credential -InfoMessageCallback $progressCb
                if (-not $r.Success) { EnqProg 0 "Error en backup"; EnqLog ("❌ Error de SQL: {0}" -f $r.ErrorMessage); EnqLog "__DONE__"; return }
                if ($DoCompress) { EnqProg 90 "Backup terminado. Iniciando compresión..." } else { EnqProg 100 "Backup terminado." }
                EnqLog "✅ Comando BACKUP finalizó (ExecuteNonQuery)"
                Start-Sleep -Milliseconds 500
                if (Test-Path $ScriptBackupPath) { $sizeMB = [math]::Round((Get-Item $ScriptBackupPath).Length / 1MB, 2); EnqLog ("📊 Tamaño del archivo: {0} MB" -f $sizeMB); EnqLog ("📁 Ubicación: {0}" -f $ScriptBackupPath) } else { EnqLog ("⚠️ No se encontró el archivo en: {0}" -f $ScriptBackupPath) }
                if ($DoCompress) {
                    EnqProg 90 "Backup listo. Preparando compresión..."
                    EnqLog "🗜️ Iniciando compresión ZIP..."
                    $inputBak = $ScriptBackupPath
                    $zipPath = "$ScriptBackupPath.zip"
                    if (-not (Test-Path $inputBak)) { EnqProg 0 "Error: no existe BAK"; EnqLog ("⚠️ No existe el BAK accesible: {0}" -f $inputBak); EnqLog "__DONE__"; return }
                    $sevenZip = Get-7ZipPath
                    if (-not $sevenZip -or -not (Test-Path $sevenZip)) { EnqProg 0 "Error: no se encontró 7-Zip"; EnqLog "❌ No se encontró 7z.exe. No se puede comprimir."; EnqLog "__DONE__"; return }
                    try {
                        if (Test-Path $zipPath) { Remove-Item $zipPath -Force -ErrorAction SilentlyContinue }
                        EnqProg 92 "Comprimiendo (ZIP)..."
                        if ($ZipPassword -and $ZipPassword.Trim().Length -gt 0) { & $sevenZip a -tzip -p"$($ZipPassword.Trim())" -mem=AES256 $zipPath $inputBak | Out-Null } else { & $sevenZip a -tzip $zipPath $inputBak | Out-Null }
                        EnqProg 97 "Finalizando compresión..."
                        Start-Sleep -Milliseconds 300
                        if (Test-Path $zipPath) { $zipMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2); EnqProg 99 "ZIP creado. Cerrando..."; EnqLog ("✅ ZIP creado ({0} MB): {1}" -f $zipMB, $zipPath) } else { EnqProg 0 "Error: ZIP no generado"; EnqLog ("❌ Se ejecutó 7-Zip pero NO se generó el ZIP: {0}" -f $zipPath) }
                    } catch { EnqProg 0 "Error al comprimir"; EnqLog ("❌ Error al comprimir: {0}" -f $_.Exception.Message) }
                }
                EnqLog "__DONE__"
            } catch {
                EnqProg 0 "Error"
                EnqLog ("❌ Error inesperado (worker): {0}" -f $_.Exception.Message)
                EnqLog "__DONE__"
            }
        }
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()
        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($worker).AddArgument($Server).AddArgument($Database).AddArgument($BackupQuery).AddArgument($ScriptBackupPath).AddArgument($DoCompress).AddArgument($ZipPassword).AddArgument($Credential).AddArgument($LogQueue).AddArgument($ProgressQueue)
        $null = $ps.BeginInvoke()
        Write-DzDebug "`t[DEBUG][Start-BackupWorkAsync] Worker lanzado"
    }
    $stopwatch = $null
    $timer = $null
    $logTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $logTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $logTimer.Add_Tick({
            try {
                $count = 0
                $doneThisTick = $false
                while ($count -lt 50) {
                    $line = $null
                    if (-not $logQueue.TryDequeue([ref]$line)) { break }
                    $txtLog.Text = "$line`n" + $txtLog.Text
                    if ($line -like "*__DONE__*") {
                        Write-DzDebug "`t[DEBUG][UI] Señal DONE recibida"
                        $doneThisTick = $true
                        $script:BackupRunning = $false
                        $btnAceptar.IsEnabled = $true
                        $btnAceptar.Content = "Iniciar Respaldo"
                        $txtNombre.IsEnabled = $true
                        $chkComprimir.IsEnabled = $true
                        if ($chkComprimir.IsChecked -eq $true) { $txtPassword.IsEnabled = $true }
                        $btnAbrirCarpeta.IsEnabled = $true
                        $tmp = $null
                        while ($progressQueue.TryDequeue([ref]$tmp)) { }
                        Paint-Progress -Percent 100 -Message "Completado"
                        $script:BackupDone = $true
                    }
                    $count++
                }
                if ($count -gt 0) { $txtLog.ScrollToLine(0) }
                if (-not $doneThisTick) {
                    $last = $null
                    while ($true) {
                        $p = $null
                        if (-not $progressQueue.TryDequeue([ref]$p)) { break }
                        $last = $p
                    }
                    if ($last) { Paint-Progress -Percent $last.Percent -Message $last.Message }
                }
            } catch { Write-DzDebug "`t[DEBUG][UI][logTimer] ERROR: $($_.Exception.Message)" }
            if ($script:BackupDone) {
                $tmpLine = $null
                $tmpProg = $null
                if (-not $logQueue.TryPeek([ref]$tmpLine) -and -not $progressQueue.TryPeek([ref]$tmpProg)) { $logTimer.Stop(); $script:BackupDone = $false }
            }
        })
    $logTimer.Start()
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] logTimer iniciado"
    $chkComprimir.Add_Checked({
            Write-DzDebug "`t[DEBUG][UI] chkComprimir CHECKED"
            try {
                if (-not (Test-ChocolateyInstalled)) {
                    $msg = @"
Chocolatey es necesario SOLAMENTE si deseas:
✓ Comprimir el respaldo (ZIP con contraseña)
Si solo necesitas crear el respaldo básico (.BAK), NO es necesario instalarlo.
¿Deseas instalar Chocolatey ahora?
"@
                    $wantChoco = Ui-Confirm $msg "Chocolatey requerido" $window
                    if (-not $wantChoco) { Ui-Warn "Compresión deshabilitada (Chocolatey no instalado)." "Atención" $window; Disable-CompressionUI; return }
                    $okChoco = Install-Chocolatey
                    if (-not $okChoco -or -not (Test-ChocolateyInstalled)) { Ui-Warn "No se pudo instalar Chocolatey. Compresión deshabilitada." "Atención" $window; Disable-CompressionUI; return }
                }
                if (-not (Test-7ZipInstalled)) {
                    $want7z = Ui-Confirm "Para comprimir se requiere 7-Zip. ¿Deseas instalarlo ahora con Chocolatey?" "7-Zip requerido" $window
                    if (-not $want7z) { Ui-Warn "Compresión deshabilitada (7-Zip no instalado)." "Atención" $window; Disable-CompressionUI; return }
                    $ok7z = Install-7ZipWithChoco
                    if (-not $ok7z) { Ui-Warn "No se pudo instalar 7-Zip. Compresión deshabilitada." "Atención" $window; Disable-CompressionUI; return }
                }
                $txtPassword.IsEnabled = $true
                $lblPassword.IsEnabled = $true
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error chkComprimir CHECKED: $($_.Exception.Message)"
                Ui-Error "Error validando requisitos de compresión: $($_.Exception.Message)" "Error" $window
                Disable-CompressionUI
            }
        })
    $chkComprimir.Add_Unchecked({ Write-DzDebug "`t[DEBUG][UI] chkComprimir UNCHECKED"; Disable-CompressionUI })
    $btnAceptar.Add_Click({
            if ($script:BackupRunning) { return }
            $script:BackupDone = $false
            if ($script:BackupRunning) { return }
            $script:BackupDone = $false
            if (-not $logTimer.IsEnabled) { $logTimer.Start() }
            if (-not $logTimer.IsEnabled) { $logTimer.Start() }
            try {
                $btnAceptar.IsEnabled = $false
                $btnAceptar.Content = "Procesando..."
                $txtLog.Text = ""
                $pbBackup.Value = 0
                $txtProgress.Text = "Esperando..."
                Add-Log "Iniciando proceso de backup..."
                $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $timer = [System.Windows.Threading.DispatcherTimer]::new()
                $timer.Interval = [TimeSpan]::FromSeconds(1)
                $timer.Start()
                Write-DzDebug "`t[DEBUG][UI] Timer iniciado OK"
                $machinePart = $Server.Split('\')[0]
                $machineName = $machinePart.Split(',')[0]
                if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }
                $sameHost = ($env:COMPUTERNAME -ieq $machineName)
                $backupFileName = $txtNombre.Text
                $sqlBackupFolder = "C:\Temp\SQLBackups"
                $sqlBackupPath = Join-Path $sqlBackupFolder $backupFileName
                if ($sameHost) { $scriptBackupPath = $sqlBackupPath } else { $scriptBackupPath = "\\$machineName\C$\Temp\SQLBackups\$backupFileName" }
                Add-Log "Servidor: $Server"
                Add-Log "Base de datos: $Database"
                Add-Log "Usuario: $User"
                Add-Log "Ruta SQL (donde escribe el motor): $sqlBackupPath"
                Add-Log "Ruta accesible desde esta PC: $scriptBackupPath"
                if (-not $backupFileName.ToLower().EndsWith(".bak")) { $backupFileName = "$backupFileName.bak"; $txtNombre.Text = $backupFileName }
                $invalid = [System.IO.Path]::GetInvalidFileNameChars()
                if ($backupFileName.IndexOfAny($invalid) -ge 0) { Show-WarnDialog -Message "El nombre contiene caracteres no válidos..." -Title "Nombre inválido" -Owner $window; Reset-BackupUI -ProgressText "Nombre inválido"; return }
                $backupExists = $false
                try {
                    $backupExists = Test-Path -LiteralPath $scriptBackupPath -ErrorAction Stop
                } catch {
                    Write-DzDebug "`t[DEBUG][Show-BackupDialog] No se pudo validar ruta backup: $scriptBackupPath | $($_.Exception.Message)"
                    $backupExists = $false
                }
                if ($backupExists) {
                    $choice = Ui-Confirm "Ya existe un respaldo con ese nombre en:`n$scriptBackupPath`n`n¿Deseas sobrescribirlo?" "Archivo existente" $window
                    if (-not $choice) { $timestampsDefault = Get-Date -Format 'yyyyMMdd-HHmmss'; $txtNombre.Text = "$Database-$timestampsDefault.bak"; Add-Log "⚠️ Operación cancelada: el archivo ya existe. Se sugirió un nuevo nombre."; Reset-BackupUI -ProgressText "Cancelado (elige otro nombre y vuelve a intentar)"; return }
                }
                if ($sameHost) {
                    if (-not (Test-Path -LiteralPath $sqlBackupFolder)) { New-Item -ItemType Directory -Path $sqlBackupFolder -Force | Out-Null }
                } else {
                    $uncFolder = "\\$machineName\C$\Temp\SQLBackups"
                    $uncExists = $false
                    try {
                        $uncExists = Test-Path -LiteralPath $uncFolder -ErrorAction Stop
                    } catch {
                        Write-DzDebug "`t[DEBUG][Show-BackupDialog] Sin permisos para validar UNC: $uncFolder | $($_.Exception.Message)"
                        $uncExists = $false
                    }
                    if (-not $uncExists) { Add-Log "⚠️ No pude validar la carpeta UNC: $uncFolder (puede ser permisos). SQL intentará escribir en $sqlBackupFolder en el servidor." }
                }
                $credential = New-SafeCredential -Username $User -PlainPassword $Password
                Add-Log "✓ Credenciales listas"
                $backupQuery = @"
BACKUP DATABASE [$Database]
TO DISK = '$sqlBackupPath'
WITH CHECKSUM, STATS = 1, FORMAT, INIT
"@
                Paint-Progress -Percent 5 -Message "Conectando a SQL Server..."
                Write-DzDebug "`t[DEBUG][UI] Llamando Start-BackupWorkAsync"
                Start-BackupWorkAsync -Server $Server -Database $Database -BackupQuery $backupQuery -ScriptBackupPath $scriptBackupPath -DoCompress ($chkComprimir.IsChecked -eq $true) -ZipPassword $txtPassword.Password -Credential $credential -LogQueue $logQueue -ProgressQueue $progressQueue
            } catch {
                Write-DzDebug "`t[DEBUG][UI] ERROR btnAceptar: $($_.Exception.Message)"
                Add-Log "❌ Error: $($_.Exception.Message)"
                $btnAceptar.IsEnabled = $true
                $btnAceptar.Content = "Iniciar Respaldo"
                if ($timer -and $timer.IsEnabled) { $timer.Stop() }
                if ($stopwatch) { $stopwatch.Stop() }
            }
        })
    $btnAbrirCarpeta.Add_Click({
            Write-DzDebug "`t[DEBUG][UI] btnAbrirCarpeta Click"
            $machinePart = $Server.Split('\')[0]
            $machineName = $machinePart.Split(',')[0]
            if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }
            $backupFolder = "\\$machineName\C$\Temp\SQLBackups"
            if (Test-Path $backupFolder) { Start-Process explorer.exe $backupFolder } else { Ui-Warn "La carpeta de respaldos no existe todavía.`n`nRuta: $backupFolder" "Atención" $window }
        })
    $btnCerrar.Add_Click({
            Write-DzDebug "`t[DEBUG][UI] btnCerrar Click"
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch { }
            try { if ($timer -and $timer.IsEnabled) { $timer.Stop() } } catch { }
            try { if ($stopwatch) { $stopwatch.Stop() } } catch { }
            $window.DialogResult = $false
            $window.Close()
        })
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] Antes de ShowDialog()"
    $null = $window.ShowDialog()
    Write-DzDebug "`t[DEBUG][Show-BackupDialog] Después de ShowDialog()"
}
function Reset-BackupUI {
    param([string]$ButtonText = "Iniciar Respaldo", [string]$ProgressText = "Esperando...")
    $script:BackupRunning = $false
    $btnAceptar.IsEnabled = $true
    $btnAceptar.Content = $ButtonText
    $txtNombre.IsEnabled = $true
    $chkComprimir.IsEnabled = $true
    if ($chkComprimir.IsChecked -eq $true) { $txtPassword.IsEnabled = $true; $lblPassword.IsEnabled = $true } else { $txtPassword.IsEnabled = $false; $lblPassword.IsEnabled = $false }
    $btnAbrirCarpeta.IsEnabled = $true
    $txtProgress.Text = $ProgressText
}
function Show-DetachDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)]$ParentNode
    )

    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) {
        Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null
    }

    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) {
        Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null
    }

    Write-DzDebug "`t[DEBUG][DetachDB] INICIO: Server='$Server' Database='$Database'"

    Add-Type -AssemblyName PresentationFramework

    $safeDb = [Security.SecurityElement]::Escape($Database)

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Separar Base de Datos"
        Height="420" Width="580"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True"
        Topmost="True">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
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
                    <TextBlock Text="📎 Separar Base de Datos (Detach)"
                               Foreground="{DynamicResource FormFg}"
                               FontSize="16"
                               FontWeight="SemiBold"/>
                    <TextBlock Text="Los archivos MDF y LDF quedarán disponibles en el sistema de archivos."
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
                Padding="12">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0"
                           Text="Estás a punto de separar la siguiente base de datos:"
                           TextWrapping="Wrap"
                           Margin="0,0,0,10"/>

                <Border Grid.Row="1"
                        Background="{DynamicResource ControlBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}"
                        BorderThickness="1"
                        CornerRadius="8"
                        Padding="10"
                        Margin="0,0,0,12">
                    <TextBlock Text="🗄️ $safeDb"
                               FontSize="14"
                               FontWeight="SemiBold"
                               Foreground="{DynamicResource AccentPrimary}"/>
                </Border>

                <StackPanel Grid.Row="2" Margin="0,0,0,10">
                    <CheckBox x:Name="chkUpdateStatistics" IsChecked="True" Margin="0,0,0,8">
                        <TextBlock Text="Actualizar estadísticas antes de separar" TextWrapping="Wrap"/>
                    </CheckBox>
                    <CheckBox x:Name="chkCloseConnections" IsChecked="True">
                        <TextBlock Text="Forzar cierre de conexiones existentes (SINGLE_USER + ROLLBACK IMMEDIATE)" TextWrapping="Wrap"/>
                    </CheckBox>
                </StackPanel>

                <Border Grid.Row="3"
                        Background="{DynamicResource FormBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}"
                        BorderThickness="1"
                        CornerRadius="8"
                        Padding="10">
                    <TextBlock FontSize="11"
                               Foreground="{DynamicResource AccentMuted}"
                               TextWrapping="Wrap">
                        ℹ️ Información importante:
• Al separar la base de datos, esta dejará de estar disponible en SQL Server
• Los archivos físicos (MDF y LDF) permanecerán en el disco
• Puedes volver a adjuntar la base de datos posteriormente
• Si hay conexiones activas, usa la opción 'Forzar cierre' para cerrarlas automáticamente
                    </TextBlock>
                </Border>
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
                           Text="Enter: Separar   |   Esc: Cerrar"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btnCancelar"
                            Content="Cancelar"
                            Width="120"
                            Height="34"
                            Margin="0,0,10,0"
                            IsCancel="True"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <Button x:Name="btnSeparar"
                            Content="Separar"
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
        $ui = New-WpfWindow -Xaml $xaml -PassThru
        $window = $ui.Window
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $window -Theme $theme
        try { Set-WpfDialogOwner -Dialog $window } catch {}

        $brdTitleBar = $window.FindName("brdTitleBar")
        if ($brdTitleBar) {
            $brdTitleBar.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) {
                        try { $window.DragMove() } catch {}
                    }
                })
        }
    } catch {
        Write-DzDebug "`t[DEBUG][DetachDB] ERROR creando ventana: $($_.Exception.Message)" -Color Red
        throw "No se pudo crear la ventana (XAML). $($_.Exception.Message)"
    }

    $chkUpdateStatistics = $window.FindName("chkUpdateStatistics")
    $chkCloseConnections = $window.FindName("chkCloseConnections")
    $btnSeparar = $window.FindName("btnSeparar")
    $btnCancelar = $window.FindName("btnCancelar")
    $btnClose = $window.FindName("btnClose")

    $btnClose.Add_Click({
            Write-DzDebug "`t[DEBUG][DetachDB] btnClose Click"
            $window.DialogResult = $false
            $window.Close()
        })

    $btnCancelar.Add_Click({
            Write-DzDebug "`t[DEBUG][DetachDB] btnCancelar Click"
            $window.DialogResult = $false
            $window.Close()
        })

    $window.Add_PreviewKeyDown({
            param($sender, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                $window.DialogResult = $false
                $window.Close()
            }
        })

    $btnSeparar.Add_Click({
            Write-DzDebug "`t[DEBUG][DetachDB] btnSeparar Click"

            $updateStats = $chkUpdateStatistics.IsChecked -eq $true
            $closeConnections = $chkCloseConnections.IsChecked -eq $true

            try {
                $escapedDb = $Database -replace "'", "''"
                $safeName = $Database -replace ']', ']]'

                # Paso 1: Actualizar estadísticas (opcional)
                if ($updateStats) {
                    Write-DzDebug "`t[DEBUG][DetachDB] Paso 1: Actualizando estadísticas"
                    $updateQuery = "USE [$safeName]; EXEC sp_updatestats"
                    $result1 = Invoke-SqlQuery -Server $Server -Database $Database -Query $updateQuery -Credential $Credential

                    if (-not $result1.Success) {
                        Write-DzDebug "`t[DEBUG][DetachDB] Error actualizando estadísticas: $($result1.ErrorMessage)"
                        Ui-Error "Error al actualizar estadísticas:`n`n$($result1.ErrorMessage)" "Error" $window
                        return
                    }
                    Write-DzDebug "`t[DEBUG][DetachDB] Estadísticas actualizadas OK"
                }

                # Paso 2: Cerrar conexiones (opcional)
                if ($closeConnections) {
                    Write-DzDebug "`t[DEBUG][DetachDB] Paso 2: Cerrando conexiones"
                    $closeQuery = "ALTER DATABASE [$safeName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
                    $result2 = Invoke-SqlQuery -Server $Server -Database "master" -Query $closeQuery -Credential $Credential

                    if (-not $result2.Success) {
                        Write-DzDebug "`t[DEBUG][DetachDB] Error cerrando conexiones: $($result2.ErrorMessage)"
                        Ui-Error "Error al cerrar conexiones existentes:`n`n$($result2.ErrorMessage)" "Error" $window
                        return
                    }
                    Write-DzDebug "`t[DEBUG][DetachDB] Conexiones cerradas OK"
                }

                # Paso 3: Separar la base de datos
                Write-DzDebug "`t[DEBUG][DetachDB] Paso 3: Separando base de datos"
                $detachQuery = "EXEC sp_detach_db @dbname = N'$escapedDb', @skipchecks = 'false'"
                $result3 = Invoke-SqlQuery -Server $Server -Database "master" -Query $detachQuery -Credential $Credential

                if (-not $result3.Success) {
                    Write-DzDebug "`t[DEBUG][DetachDB] Error separando BD: $($result3.ErrorMessage)"

                    # Si falla, intentar restaurar MULTI_USER si se había cerrado conexiones
                    if ($closeConnections) {
                        try {
                            $restoreQuery = "ALTER DATABASE [$safeName] SET MULTI_USER"
                            Invoke-SqlQuery -Server $Server -Database "master" -Query $restoreQuery -Credential $Credential | Out-Null
                        } catch {
                            Write-DzDebug "`t[DEBUG][DetachDB] No se pudo restaurar MULTI_USER"
                        }
                    }

                    Ui-Error "Error al separar la base de datos:`n`n$($result3.ErrorMessage)" "Error" $window
                    return
                }

                Write-DzDebug "`t[DEBUG][DetachDB] Base de datos separada OK"

                Ui-Info "La base de datos '$Database' ha sido separada exitosamente.`n`nLos archivos físicos permanecen en el disco y pueden ser adjuntados nuevamente cuando lo necesites." "✓ Éxito" $window

                # Refrescar el TreeView
                $serverNode = $ParentNode.Parent
                if ($serverNode -and $serverNode.Tag.Type -eq "Server") {
                    if ($serverNode.Tag.OnDatabasesRefreshed) {
                        try {
                            Write-DzDebug "`t[DEBUG][DetachDB] Llamando a OnDatabasesRefreshed"
                            & $serverNode.Tag.OnDatabasesRefreshed
                        } catch {
                            Write-DzDebug "`t[DEBUG][DetachDB] Error en OnDatabasesRefreshed: $($_.Exception.Message)"
                        }
                    }
                    Refresh-SqlTreeServerNode -ServerNode $serverNode
                }

                # Remover el nodo del TreeView
                if ($ParentNode.Parent -is [System.Windows.Controls.ItemsControl]) {
                    $window.Dispatcher.Invoke([action] {
                            try {
                                [void]$ParentNode.Parent.Items.Remove($ParentNode)
                                Write-DzDebug "`t[DEBUG][DetachDB] Nodo removido del TreeView"
                            } catch {
                                Write-DzDebug "`t[DEBUG][DetachDB] Error removiendo nodo: $($_.Exception.Message)"
                            }
                        })
                }

                $window.DialogResult = $true
                $window.Close()
            } catch {
                Write-DzDebug "`t[DEBUG][DetachDB] Excepción: $($_.Exception.Message)"
                Ui-Error "Error inesperado al separar la base de datos:`n`n$($_.Exception.Message)" "Error" $window
            }
        })

    Write-DzDebug "`t[DEBUG][DetachDB] Mostrando ventana"
    $null = $window.ShowDialog()
}
function Show-DatabaseSizeDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )

    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) {
        Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null
    }

    Write-DzDebug "`t[DEBUG][DBSize] INICIO: Server='$Server' Database='$Database'"

    Add-Type -AssemblyName PresentationFramework

    # Escapar nombre de base de datos
    $safeDbName = $Database -replace ']', ']]'

    # Consultar tamaños - SIN usar USE, directamente en la query
    $sizeQuery = @"
SELECT
    name AS FileName,
    physical_name AS FilePath,
    type_desc AS FileType,
    CAST(size * 8.0 / 1024 AS DECIMAL(18,2)) AS SizeMB,
    CAST(FILEPROPERTY(name, 'SpaceUsed') * 8.0 / 1024 AS DECIMAL(18,2)) AS UsedMB,
    CAST((size - FILEPROPERTY(name, 'SpaceUsed')) * 8.0 / 1024 AS DECIMAL(18,2)) AS FreeMB,
    CAST(FILEPROPERTY(name, 'SpaceUsed') * 100.0 / size AS DECIMAL(5,2)) AS PercentUsed
FROM [$safeDbName].sys.database_files
ORDER BY type, name;
"@

    Write-DzDebug "`t[DEBUG][DBSize] Ejecutando query..."
    $result = Invoke-SqlQuery -Server $Server -Database $Database -Query $sizeQuery -Credential $Credential

    if (-not $result.Success) {
        Write-DzDebug "`t[DEBUG][DBSize] ERROR en query: $($result.ErrorMessage)"
        Ui-Error "Error al consultar el tamaño de la base de datos:`n`n$($result.ErrorMessage)" "Error" $null
        return
    }

    # Verificar que hay datos
    if (-not $result.DataTable -or $result.DataTable.Rows.Count -eq 0) {
        Write-DzDebug "`t[DEBUG][DBSize] No se obtuvieron filas"
        Ui-Error "No se pudo obtener información de tamaño de la base de datos." "Error" $null
        return
    }

    Write-DzDebug "`t[DEBUG][DBSize] Se obtuvieron $($result.DataTable.Rows.Count) archivos"

    # Construir filas dinámicas del XAML
    $dataRows = New-Object System.Collections.ArrayList
    $totalSizeMB = 0
    $totalUsedMB = 0
    $rowIndex = 1

    foreach ($row in $result.DataTable.Rows) {
        $fileName = [string]$row.FileName
        $fileType = [string]$row.FileType
        $sizeMB = [decimal]$row.SizeMB
        $usedMB = [decimal]$row.UsedMB
        $percentUsed = [decimal]$row.PercentUsed

        $totalSizeMB += $sizeMB
        $totalUsedMB += $usedMB

        $typeIcon = if ($fileType -eq "ROWS") { "📄" } else { "📋" }

        # Escapar para XML
        $safeFileName = [Security.SecurityElement]::Escape($fileName)
        $safeFileType = [Security.SecurityElement]::Escape($fileType)

        $rowXaml = @"
                    <!-- Row $rowIndex -->
                    <Border Grid.Row="$rowIndex" Grid.Column="0" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,0,1,1" Padding="8">
                        <TextBlock Text="$typeIcon $safeFileName" FontWeight="SemiBold"/>
                    </Border>
                    <Border Grid.Row="$rowIndex" Grid.Column="1" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,0,1,1" Padding="8">
                        <TextBlock Text="$safeFileType"/>
                    </Border>
                    <Border Grid.Row="$rowIndex" Grid.Column="2" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,0,1,1" Padding="8" Background="{DynamicResource ControlBg}">
                        <TextBlock Text="$($sizeMB.ToString('N2')) MB" HorizontalAlignment="Right" FontFamily="Consolas"/>
                    </Border>
                    <Border Grid.Row="$rowIndex" Grid.Column="3" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,0,1,1" Padding="8">
                        <TextBlock Text="$($usedMB.ToString('N2')) MB" HorizontalAlignment="Right" FontFamily="Consolas"/>
                    </Border>
                    <Border Grid.Row="$rowIndex" Grid.Column="4" BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,0,0,1" Padding="8">
                        <TextBlock Text="$($percentUsed.ToString('N2'))%" HorizontalAlignment="Right" FontFamily="Consolas" Foreground="{DynamicResource AccentPrimary}"/>
                    </Border>
"@
        [void]$dataRows.Add($rowXaml)
        $rowIndex++
    }

    $totalFreeMB = $totalSizeMB - $totalUsedMB
    $totalPercentUsed = if ($totalSizeMB -gt 0) { ($totalUsedMB / $totalSizeMB) * 100 } else { 0 }

    Write-DzDebug "`t[DEBUG][DBSize] Totales: Size=$($totalSizeMB.ToString('N2')) MB, Used=$($totalUsedMB.ToString('N2')) MB, Free=$($totalFreeMB.ToString('N2')) MB, PercentUsed=$($totalPercentUsed.ToString('N2'))%"

    $safeDb = [Security.SecurityElement]::Escape($Database)
    $totalRowCount = $result.DataTable.Rows.Count + 2  # Headers + data + total

    # Generar definiciones de filas
    $rowDefs = ""
    for ($i = 0; $i -lt $totalRowCount; $i++) {
        $rowDefs += "                        <RowDefinition Height='Auto'/>`n"
    }

    # Unir todas las filas de datos
    $allDataRows = $dataRows -join "`n"
    $totalRowIndex = $result.DataTable.Rows.Count + 1

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Tamaño de Base de Datos"
        Height="550" Width="780"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True"
        Topmost="True">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Name="brdTitleBar" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="12" Margin="0,0,0,10">
            <DockPanel LastChildFill="True">
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="📊 Tamaño de Base de Datos" Foreground="{DynamicResource FormFg}"
                               FontSize="16" FontWeight="SemiBold"/>
                    <TextBlock Text="🗄️ $safeDb" Foreground="{DynamicResource AccentPrimary}"
                               Margin="0,2,0,0" FontSize="13"/>
                </StackPanel>
                <Button DockPanel.Dock="Right" Name="btnClose" Style="{StaticResource CloseButtonStyle}"/>
            </DockPanel>
        </Border>

        <Border Grid.Row="1" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="12" Margin="0,0,0,10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <Grid>
                    <Grid.RowDefinitions>
$rowDefs
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*"/>
                        <ColumnDefinition Width="1.2*"/>
                        <ColumnDefinition Width="1*"/>
                        <ColumnDefinition Width="1*"/>
                        <ColumnDefinition Width="0.8*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Headers -->
                    <Border Grid.Row="0" Grid.Column="0" Background="{DynamicResource AccentSecondary}"
                            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1,1,1,2" Padding="8">
                        <TextBlock Text="Archivo" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="0" Grid.Column="1" Background="{DynamicResource AccentSecondary}"
                            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,1,1,2" Padding="8">
                        <TextBlock Text="Tipo" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="0" Grid.Column="2" Background="{DynamicResource AccentSecondary}"
                            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,1,1,2" Padding="8">
                        <TextBlock Text="Tamaño Total" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="0" Grid.Column="3" Background="{DynamicResource AccentSecondary}"
                            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,1,1,2" Padding="8">
                        <TextBlock Text="Espacio Usado" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="0" Grid.Column="4" Background="{DynamicResource AccentSecondary}"
                            BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="0,1,1,2" Padding="8">
                        <TextBlock Text="% Usado" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>

                    <!-- Data rows -->
$allDataRows

                    <!-- Total row -->
                    <Border Grid.Row="$totalRowIndex" Grid.Column="0" Grid.ColumnSpan="2"
                            Background="{DynamicResource AccentMagenta}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1,2,1,1" Padding="8">
                        <TextBlock Text="📦 TOTAL" FontWeight="Bold" Foreground="{DynamicResource FormFg}" FontSize="13"/>
                    </Border>
                    <Border Grid.Row="$totalRowIndex" Grid.Column="2"
                            Background="{DynamicResource AccentMagenta}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="0,2,1,1" Padding="8">
                        <TextBlock Text="$($totalSizeMB.ToString('N2')) MB" HorizontalAlignment="Right"
                                   FontFamily="Consolas" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="$totalRowIndex" Grid.Column="3"
                            Background="{DynamicResource AccentMagenta}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="0,2,1,1" Padding="8">
                        <TextBlock Text="$($totalUsedMB.ToString('N2')) MB" HorizontalAlignment="Right"
                                   FontFamily="Consolas" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                    <Border Grid.Row="$totalRowIndex" Grid.Column="4"
                            Background="{DynamicResource AccentMagenta}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="0,2,1,1" Padding="8">
                        <TextBlock Text="$($totalPercentUsed.ToString('N2'))%" HorizontalAlignment="Right"
                                   FontFamily="Consolas" FontWeight="Bold" Foreground="{DynamicResource FormFg}"/>
                    </Border>
                </Grid>
            </ScrollViewer>
        </Border>

        <Border Grid.Row="2" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="10" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock FontSize="11" Foreground="{DynamicResource AccentMuted}" TextWrapping="Wrap">
                    ℹ️ Información:
                </TextBlock>
                <TextBlock FontSize="11" Foreground="{DynamicResource PanelFg}" TextWrapping="Wrap" Margin="0,4,0,0">
                    • ROWS = Archivos de datos (MDF/NDF)
                    • LOG = Archivos de registro de transacciones (LDF)
                    • Tamaño Total = Espacio asignado en disco
                    • Espacio Usado = Datos actualmente almacenados
                    • Espacio Libre = $($totalFreeMB.ToString('N2')) MB disponibles
                </TextBlock>
            </StackPanel>
        </Border>

        <Border Grid.Row="3" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Esc: Cerrar" VerticalAlignment="Center"/>
                <Button Grid.Column="1" Name="btnCerrar" Content="Cerrar" Width="100" Height="34"
                        Style="{StaticResource BaseButtonStyle}"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    try {
        Write-DzDebug "`t[DEBUG][DBSize] Creando ventana WPF..."
        $ui = New-WpfWindow -Xaml $xaml -PassThru
        $window = $ui.Window
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $window -Theme $theme
        try { Set-WpfDialogOwner -Dialog $window } catch {}

        $brdTitleBar = $window.FindName("brdTitleBar")
        if ($brdTitleBar) {
            $brdTitleBar.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) {
                        try { $window.DragMove() } catch {}
                    }
                })
        }

        $btnClose = $window.FindName("btnClose")
        $btnCerrar = $window.FindName("btnCerrar")

        $btnClose.Add_Click({ $window.Close() })
        $btnCerrar.Add_Click({ $window.Close() })

        $window.Add_PreviewKeyDown({
                param($sender, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                    $window.Close()
                }
            })

        Write-DzDebug "`t[DEBUG][DBSize] Mostrando diálogo..."
        $null = $window.ShowDialog()
        Write-DzDebug "`t[DEBUG][DBSize] Diálogo cerrado"
    } catch {
        Write-DzDebug "`t[DEBUG][DBSize] ERROR creando ventana: $($_.Exception.Message)"
        Ui-Error "Error al mostrar el diálogo: $($_.Exception.Message)" "Error" $null
    }
}

function Show-DatabaseRepairDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )

    $script:RepairRunning = $false
    $script:RepairDone = $false

    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) {
        Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null
    }

    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) {
        Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null
    }

    function Ui-Confirm([string]$m, [string]$t = "Confirmar", [System.Windows.Window]$o) {
        (Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "YesNo" -Icon "Question" -Owner $o) -eq [System.Windows.MessageBoxResult]::Yes
    }

    Write-DzDebug "`t[DEBUG][DBRepair] INICIO: Server='$Server' Database='$Database'"

    Add-Type -AssemblyName PresentationFramework

    $safeDb = [Security.SecurityElement]::Escape($Database)

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Reparación de Base de Datos"
        Height="720" Width="700"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="Transparent"
        AllowsTransparency="True"
        Topmost="True">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="{DynamicResource AccentSecondary}"/>
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
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
                    <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                    <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Name="brdTitleBar" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="12" Margin="0,0,0,10">
            <DockPanel LastChildFill="True">
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="⚠️ Reparación de Base de Datos" Foreground="{DynamicResource FormFg}"
                               FontSize="16" FontWeight="SemiBold"/>
                    <TextBlock Text="🗄️ $safeDb" Foreground="{DynamicResource AccentPrimary}"
                               Margin="0,2,0,0" FontSize="13"/>
                </StackPanel>
                <Button DockPanel.Dock="Right" Name="btnClose" Style="{StaticResource CloseButtonStyle}"/>
            </DockPanel>
        </Border>

        <Border Grid.Row="1" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="12" Margin="0,0,0,10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel>
                    <!-- Advertencia -->
                    <Border Background="#33FF0000" BorderBrush="#FFFF0000" BorderThickness="2"
                            CornerRadius="8" Padding="12" Margin="0,0,0,12">
                        <StackPanel>
                            <TextBlock Text="⚠️ ADVERTENCIA CRÍTICA" FontWeight="Bold" FontSize="14"
                                       Foreground="#FFFF3333" Margin="0,0,0,8"/>
                            <TextBlock TextWrapping="Wrap" Foreground="{DynamicResource PanelFg}">
                                Esta operación puede causar PÉRDIDA DE DATOS irreversible.
                                Solo continúa si entiendes completamente las consecuencias.
                                Se recomienda realizar un respaldo completo antes de proceder.
                            </TextBlock>
                        </StackPanel>
                    </Border>

                    <!-- Paso 1: Verificación -->
                    <TextBlock Text="Paso 1: Verificar Integridad" FontWeight="Bold" FontSize="13"
                               Margin="0,0,0,8"/>
                    <Border Background="{DynamicResource ControlBg}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,12">
                        <StackPanel>
                            <RadioButton x:Name="rbCheckOnly" IsChecked="True" GroupName="Action" Margin="0,0,0,8">
                                <TextBlock TextWrapping="Wrap">
                                    🔍 Solo verificar (DBCC CHECKDB sin reparación)
                                </TextBlock>
                            </RadioButton>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="{DynamicResource AccentMuted}"
                                       Margin="20,0,0,0">
                                Recomendado: Primero verifica si hay errores antes de intentar reparar.
                            </TextBlock>
                        </StackPanel>
                    </Border>

                    <!-- Paso 2: Reparación -->
                    <TextBlock Text="Paso 2: Reparación (Solo si hay errores)" FontWeight="Bold" FontSize="13"
                               Margin="0,0,0,8"/>
                    <Border Background="{DynamicResource ControlBg}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,12">
                        <StackPanel>
                            <RadioButton x:Name="rbRepairFast" GroupName="Action" Margin="0,0,0,8">
                                <TextBlock TextWrapping="Wrap">
                                    🔧 REPAIR_FAST (reparación rápida, sin pérdida de datos)
                                </TextBlock>
                            </RadioButton>
                            <RadioButton x:Name="rbRepairRebuild" GroupName="Action" Margin="0,0,0,8">
                                <TextBlock TextWrapping="Wrap">
                                    🔨 REPAIR_REBUILD (reconstruir índices, sin pérdida de datos)
                                </TextBlock>
                            </RadioButton>
                            <RadioButton x:Name="rbRepairAllowDataLoss" GroupName="Action" Margin="0,0,0,8">
                                <TextBlock TextWrapping="Wrap" Foreground="#FFFF6666">
                                    ⚠️ REPAIR_ALLOW_DATA_LOSS (puede causar PÉRDIDA DE DATOS)
                                </TextBlock>
                            </RadioButton>
                            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#FFFF6666"
                                       Margin="20,0,0,0">
                                PELIGRO: Esta opción eliminará datos corruptos. Úsala solo como último recurso.
                            </TextBlock>
                        </StackPanel>
                    </Border>

                    <!-- Opciones adicionales -->
                    <TextBlock Text="Opciones Adicionales" FontWeight="Bold" FontSize="13"
                               Margin="0,0,0,8"/>
                    <Border Background="{DynamicResource ControlBg}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,12">
                        <StackPanel>
                            <CheckBox x:Name="chkCloseConnections" IsChecked="True" Margin="0,0,0,8">
                                <TextBlock TextWrapping="Wrap">
                                    Cerrar conexiones activas (SINGLE_USER)
                                </TextBlock>
                            </CheckBox>
                            <CheckBox x:Name="chkEmergencyMode" IsChecked="True">
                                <TextBlock TextWrapping="Wrap">
                                    Poner en modo EMERGENCY antes de reparar
                                </TextBlock>
                            </CheckBox>
                        </StackPanel>
                    </Border>

                    <!-- Progress -->
                    <Border Background="{DynamicResource FormBg}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,0,0,12">
                        <StackPanel>
                            <TextBlock x:Name="txtProgress" Text="Esperando..." Margin="0,0,0,8" FontWeight="SemiBold"/>
                            <ProgressBar x:Name="pbProgress" Height="20" Minimum="0" Maximum="100" Value="0"/>
                        </StackPanel>
                    </Border>

                    <!-- Log -->
                    <TextBlock Text="Log de Operaciones" FontWeight="Bold" FontSize="13" Margin="0,0,0,8"/>
                    <Border Background="{DynamicResource ControlBg}" BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1" CornerRadius="8" Padding="8">
                        <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                                 Height="180" TextWrapping="Wrap" FontFamily="Consolas" FontSize="11"
                                 BorderThickness="0" Background="Transparent"/>
                    </Border>
                </StackPanel>
            </ScrollViewer>
        </Border>

        <Border Grid.Row="2" Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                CornerRadius="10" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="⚠️ Lee las advertencias antes de continuar"
                           VerticalAlignment="Center" Foreground="#FFFF6666"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btnCancelar" Content="Cancelar" Width="120" Height="34"
                            Margin="0,0,10,0" Style="{StaticResource OutlineButtonStyle}"/>
                    <Button x:Name="btnIniciar" Content="Iniciar" Width="140" Height="34"
                            Style="{StaticResource ActionButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    try {
        $ui = New-WpfWindow -Xaml $xaml -PassThru
        $window = $ui.Window
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $window -Theme $theme
        try { Set-WpfDialogOwner -Dialog $window } catch {}

        $brdTitleBar = $window.FindName("brdTitleBar")
        if ($brdTitleBar) {
            $brdTitleBar.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) {
                        try { $window.DragMove() } catch {}
                    }
                })
        }
    } catch {
        Write-DzDebug "`t[DEBUG][DBRepair] ERROR creando ventana: $($_.Exception.Message)"
        throw "No se pudo crear la ventana (XAML). $($_.Exception.Message)"
    }

    $rbCheckOnly = $window.FindName("rbCheckOnly")
    $rbRepairFast = $window.FindName("rbRepairFast")
    $rbRepairRebuild = $window.FindName("rbRepairRebuild")
    $rbRepairAllowDataLoss = $window.FindName("rbRepairAllowDataLoss")
    $chkCloseConnections = $window.FindName("chkCloseConnections")
    $chkEmergencyMode = $window.FindName("chkEmergencyMode")
    $pbProgress = $window.FindName("pbProgress")
    $txtProgress = $window.FindName("txtProgress")
    $txtLog = $window.FindName("txtLog")
    $btnIniciar = $window.FindName("btnIniciar")
    $btnCancelar = $window.FindName("btnCancelar")
    $btnClose = $window.FindName("btnClose")

    $logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $progressQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'

    function Paint-Progress {
        param([int]$Percent, [string]$Message)
        $pbProgress.Value = $Percent
        $txtProgress.Text = $Message
    }

    function Add-Log {
        param([string]$Message)
        $logQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message))
    }

    function Start-RepairWorkAsync {
        param(
            [string]$Server,
            [string]$Database,
            [string]$RepairOption,
            [bool]$CloseConnections,
            [bool]$EmergencyMode,
            [System.Management.Automation.PSCredential]$Credential,
            [System.Collections.Concurrent.ConcurrentQueue[string]]$LogQueue,
            [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$ProgressQueue
        )

        $worker = {
            param($Server, $Database, $RepairOption, $CloseConnections, $EmergencyMode, $Credential, $LogQueue, $ProgressQueue)

            function EnqLog([string]$m) { $LogQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $m)) }
            function EnqProg([int]$p, [string]$m) { $ProgressQueue.Enqueue(@{Percent = $p; Message = $m }) }

            function Invoke-SqlQueryLite {
                param([string]$Server, [string]$Database, [string]$Query, [System.Management.Automation.PSCredential]$Credential)

                $connection = $null
                $passwordBstr = [IntPtr]::Zero
                $plainPassword = $null

                try {
                    $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
                    $cs = "Server=$Server;Database=master;User Id=$($Credential.UserName);Password=$plainPassword;Connection Timeout=300"
                    $connection = New-Object System.Data.SqlClient.SqlConnection($cs)
                    $connection.Open()

                    $cmd = $connection.CreateCommand()
                    $cmd.CommandText = $Query
                    $cmd.CommandTimeout = 0

                    $reader = $cmd.ExecuteReader()
                    $messages = New-Object System.Collections.ArrayList

                    do {
                        while ($reader.Read()) {
                            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                                $val = $reader.GetValue($i)
                                if ($val -and $val -ne [DBNull]::Value) {
                                    [void]$messages.Add($val.ToString())
                                }
                            }
                        }
                    } while ($reader.NextResult())

                    $reader.Close()

                    @{ Success = $true; Messages = $messages }
                } catch {
                    @{ Success = $false; ErrorMessage = $_.Exception.Message }
                } finally {
                    if ($plainPassword) { $plainPassword = $null }
                    if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                    if ($connection) { try { $connection.Close(); $connection.Dispose() } catch { } }
                }
            }

            try {
                $safeName = $Database -replace ']', ']]'
                $steps = 0
                $currentStep = 0

                if ($EmergencyMode -and $RepairOption -ne "CHECK") { $steps++ }
                if ($CloseConnections) { $steps++ }
                $steps++ # La operación principal
                if ($CloseConnections) { $steps++ } # Restaurar MULTI_USER

                # Paso 1: Emergency Mode
                if ($EmergencyMode -and $RepairOption -ne "CHECK") {
                    $currentStep++
                    EnqProg ([int](($currentStep / $steps) * 90)) "Configurando modo EMERGENCY..."
                    EnqLog "🔧 Configurando base de datos en modo EMERGENCY"

                    $emergencyQuery = "ALTER DATABASE [$safeName] SET EMERGENCY"
                    $result = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $emergencyQuery -Credential $Credential

                    if (-not $result.Success) {
                        EnqProg 0 "Error"
                        EnqLog "❌ Error al configurar modo EMERGENCY: $($result.ErrorMessage)"
                        EnqLog "ERROR_RESULT|$($result.ErrorMessage)"
                        EnqLog "__DONE__"
                        return
                    }
                    EnqLog "✅ Modo EMERGENCY configurado"
                }

                # Paso 2: Close Connections
                if ($CloseConnections) {
                    $currentStep++
                    EnqProg ([int](($currentStep / $steps) * 90)) "Cerrando conexiones..."
                    EnqLog "🔒 Cerrando conexiones existentes (SINGLE_USER)"

                    $closeQuery = "ALTER DATABASE [$safeName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
                    $result = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $closeQuery -Credential $Credential

                    if (-not $result.Success) {
                        EnqProg 0 "Error"
                        EnqLog "❌ Error al cerrar conexiones: $($result.ErrorMessage)"
                        EnqLog "ERROR_RESULT|$($result.ErrorMessage)"
                        EnqLog "__DONE__"
                        return
                    }
                    EnqLog "✅ Conexiones cerradas"
                }

                # Paso 3: Ejecutar DBCC CHECKDB
                $currentStep++
                $actionText = if ($RepairOption -eq "CHECK") { "Verificando integridad..." } else { "Ejecutando reparación..." }
                EnqProg ([int](($currentStep / $steps) * 90)) $actionText
                EnqLog "🔍 Ejecutando DBCC CHECKDB ($RepairOption)"

                $dbccQuery = switch ($RepairOption) {
                    "CHECK" { "DBCC CHECKDB([$safeName]) WITH NO_INFOMSGS" }
                    "REPAIR_FAST" { "DBCC CHECKDB([$safeName], REPAIR_FAST) WITH NO_INFOMSGS" }
                    "REPAIR_REBUILD" { "DBCC CHECKDB([$safeName], REPAIR_REBUILD) WITH NO_INFOMSGS" }
                    "REPAIR_ALLOW_DATA_LOSS" { "DBCC CHECKDB([$safeName], REPAIR_ALLOW_DATA_LOSS) WITH NO_INFOMSGS" }
                }

                $result = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $dbccQuery -Credential $Credential

                if (-not $result.Success) {
                    EnqProg 0 "Error"
                    EnqLog "❌ Error ejecutando DBCC: $($result.ErrorMessage)"

                    # Intentar restaurar MULTI_USER si es necesario
                    if ($CloseConnections) {
                        try {
                            $restoreQuery = "ALTER DATABASE [$safeName] SET MULTI_USER"
                            Invoke-SqlQueryLite -Server $Server -Database "master" -Query $restoreQuery -Credential $Credential | Out-Null
                        } catch { }
                    }

                    EnqLog "ERROR_RESULT|$($result.ErrorMessage)"
                    EnqLog "__DONE__"
                    return
                }

                # Procesar mensajes del DBCC
                $hasErrors = $false
                foreach ($msg in $result.Messages) {
                    EnqLog "[DBCC] $msg"
                    if ($msg -match "(?i)(error|corruption|corrupt|dañ|inconsisten)") {
                        $hasErrors = $true
                    }
                }

                if ($result.Messages.Count -eq 0) {
                    EnqLog "✅ No se encontraron problemas de integridad"
                } elseif (-not $hasErrors) {
                    EnqLog "✅ Verificación completada sin errores críticos"
                } else {
                    EnqLog "⚠️ Se encontraron problemas de integridad"
                }

                # Paso 4: Restaurar MULTI_USER
                if ($CloseConnections) {
                    $currentStep++
                    EnqProg ([int](($currentStep / $steps) * 90)) "Restaurando acceso normal..."
                    EnqLog "🔓 Restaurando modo MULTI_USER"

                    $restoreQuery = "ALTER DATABASE [$safeName] SET MULTI_USER"
                    $result = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $restoreQuery -Credential $Credential

                    if (-not $result.Success) {
                        EnqLog "⚠️ Advertencia: No se pudo restaurar MULTI_USER: $($result.ErrorMessage)"
                    } else {
                        EnqLog "✅ Modo MULTI_USER restaurado"
                    }
                }

                EnqProg 100 "Operación completada"
                EnqLog "✅ Proceso completado exitosamente"
                EnqLog "SUCCESS_RESULT|Operación completada. Revisa el log para detalles."
                EnqLog "__DONE__"
            } catch {
                EnqProg 0 "Error"
                EnqLog "❌ Error inesperado: $($_.Exception.Message)"
                EnqLog "ERROR_RESULT|$($_.Exception.Message)"
                EnqLog "__DONE__"
            }
        }

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.ThreadOptions = 'ReuseThread'
        $rs.Open()

        $ps = [PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($worker).AddArgument($Server).AddArgument($Database).AddArgument($RepairOption).AddArgument($CloseConnections).AddArgument($EmergencyMode).AddArgument($Credential).AddArgument($LogQueue).AddArgument($ProgressQueue)

        $null = $ps.BeginInvoke()
    }

    $logTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $logTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $logTimer.Add_Tick({
            try {
                $count = 0
                $doneThisTick = $false
                $finalResult = $null

                while ($count -lt 50) {
                    $line = $null
                    if (-not $logQueue.TryDequeue([ref]$line)) { break }

                    if ($line -like "*SUCCESS_RESULT|*") {
                        $finalResult = @{ Success = $true; Message = $line -replace '^.*SUCCESS_RESULT\|', '' }
                    }
                    if ($line -like "*ERROR_RESULT|*") {
                        $finalResult = @{ Success = $false; Message = $line -replace '^.*ERROR_RESULT\|', '' }
                    }

                    if ($line -notmatch '(SUCCESS_RESULT|ERROR_RESULT)') {
                        $txtLog.AppendText("$line`n")
                        $txtLog.ScrollToEnd()
                    }

                    if ($line -like "*__DONE__*") {
                        $doneThisTick = $true
                        $script:RepairRunning = $false
                        $btnIniciar.IsEnabled = $true
                        $btnIniciar.Content = "Iniciar"

                        $tmp = $null
                        while ($progressQueue.TryDequeue([ref]$tmp)) { }
                        Paint-Progress -Percent 100 -Message "Completado"

                        $script:RepairDone = $true

                        if ($finalResult) {
                            $window.Dispatcher.Invoke([action] {
                                    if ($finalResult.Success) {
                                        Ui-Info "Operación completada.`n`n$($finalResult.Message)`n`nRevisa el log para más detalles." "✓ Completado" $window
                                    } else {
                                        Ui-Error "La operación falló:`n`n$($finalResult.Message)" "✗ Error" $window
                                    }
                                }, [System.Windows.Threading.DispatcherPriority]::Normal)
                        }
                    }
                    $count++
                }

                if (-not $doneThisTick) {
                    $last = $null
                    while ($true) {
                        $p = $null
                        if (-not $progressQueue.TryDequeue([ref]$p)) { break }
                        $last = $p
                    }
                    if ($last) { Paint-Progress -Percent $last.Percent -Message $last.Message }
                }
            } catch {
                Write-DzDebug "`t[DEBUG][UI][logTimer][repair] ERROR: $($_.Exception.Message)"
            }

            if ($script:RepairDone) {
                $tmpLine = $null
                $tmpProg = $null
                if (-not $logQueue.TryPeek([ref]$tmpLine) -and -not $progressQueue.TryPeek([ref]$tmpProg)) {
                    $logTimer.Stop()
                    $script:RepairDone = $false
                }
            }
        })

    $logTimer.Start()

    $btnClose.Add_Click({
            $window.Close()
        })

    $btnCancelar.Add_Click({
            $window.Close()
        })

    $window.Add_PreviewKeyDown({
            param($sender, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                $window.Close()
            }
        })

    $btnIniciar.Add_Click({
            if ($script:RepairRunning) { return }

            # Determinar opción seleccionada
            $repairOption = "CHECK"
            if ($rbRepairFast.IsChecked) { $repairOption = "REPAIR_FAST" }
            elseif ($rbRepairRebuild.IsChecked) { $repairOption = "REPAIR_REBUILD" }
            elseif ($rbRepairAllowDataLoss.IsChecked) { $repairOption = "REPAIR_ALLOW_DATA_LOSS" }

            # Confirmación especial para REPAIR_ALLOW_DATA_LOSS
            if ($repairOption -eq "REPAIR_ALLOW_DATA_LOSS") {
                $msg = @"
⚠️ ADVERTENCIA FINAL ⚠️

Estás a punto de ejecutar REPAIR_ALLOW_DATA_LOSS.

Esta operación:
• PUEDE ELIMINAR DATOS PERMANENTEMENTE
• Es IRREVERSIBLE
• Debe usarse solo como ÚLTIMO RECURSO
• Requiere que hayas respaldado la base de datos

¿Estás ABSOLUTAMENTE SEGURO de continuar?
"@
                if (-not (Ui-Confirm $msg "⚠️ Confirmar Reparación Destructiva" $window)) {
                    return
                }
            }

            $script:RepairDone = $false
            if (-not $logTimer.IsEnabled) { $logTimer.Start() }

            try {
                $btnIniciar.IsEnabled = $false
                $btnIniciar.Content = "Procesando..."
                $txtLog.Text = ""
                $pbProgress.Value = 0
                $txtProgress.Text = "Iniciando..."

                Add-Log "═══════════════════════════════════════"
                Add-Log "Iniciando operación de reparación"
                Add-Log "Base de datos: $Database"
                Add-Log "Servidor: $Server"
                Add-Log "Operación: $repairOption"
                Add-Log "Cerrar conexiones: $(if ($chkCloseConnections.IsChecked) { 'Sí' } else { 'No' })"
                Add-Log "Modo EMERGENCY: $(if ($chkEmergencyMode.IsChecked) { 'Sí' } else { 'No' })"
                Add-Log "═══════════════════════════════════════"

                Start-RepairWorkAsync -Server $Server -Database $Database -RepairOption $repairOption `
                    -CloseConnections ($chkCloseConnections.IsChecked -eq $true) `
                    -EmergencyMode ($chkEmergencyMode.IsChecked -eq $true) `
                    -Credential $Credential -LogQueue $logQueue -ProgressQueue $progressQueue

                $script:RepairRunning = $true
            } catch {
                Write-DzDebug "`t[DEBUG][UI] ERROR btnIniciar: $($_.Exception.Message)"
                Add-Log "❌ Error: $($_.Exception.Message)"
                $btnIniciar.IsEnabled = $true
                $btnIniciar.Content = "Iniciar"
                Paint-Progress -Percent 0 -Message "Error"
            }
        })

    $null = $window.ShowDialog()

    try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch { }
}
Export-ModuleMember -Function @('Show-RestoreDialog', 'Show-AttachDialog',
    'Reset-RestoreUI', 'Reset-AttachUI',
    'Show-BackupDialog', 'Reset-BackupUI',
    'Show-DetachDialog', 'Reset-DetachUI', 'Show-DatabaseSizeDialog',
    'Show-DatabaseRepairDialog', 'Reset-DatabaseRepairUI')