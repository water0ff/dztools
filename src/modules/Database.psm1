#requires -Version 5.0
function Process-SqlProgressMessage { param([string]$Message) if ($Message -match '(?i)(\d{1,3})\s*percent') { $percent = [int]$Matches[1]; Write-Output "Progreso: $percent%" } elseif ($Message -match 'BACKUP DATABASE successfully processed') { Write-Output "Backup completado exitosamente" } }
function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][scriptblock]$InfoMessageCallback
    )
    $connection = $null
    Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] INICIO"
    Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Server='$Server' Database='$Database' User='$($Credential.UserName)'"
    try {
        $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
        $connectionString = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Creando SqlConnection..."
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        if ($InfoMessageCallback) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Registrando InfoMessage handler..."
            $connection.add_InfoMessage({ param($sender, $e) try { & $InfoMessageCallback $e.Message } catch { Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Error en InfoMessageCallback: $_" } })
            $connection.FireInfoMessageEventOnUserErrors = $true
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Abriendo conexión..."
        $connection.Open()
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Estado conexión tras Open(): $($connection.State)"
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 0
        $returnsResultSet = $Query -match "(?si)^\s*(SELECT|WITH)" -or $Query -match "(?si)\bOUTPUT\b"
        if ($returnsResultSet) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Ejecutando consulta tipo SELECT/WITH"
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)
            return @{Success = $true; DataTable = $dataTable; Type = "Query" }
        } else {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Ejecutando consulta tipo NonQuery"
            $rowsAffected = $command.ExecuteNonQuery()
            return @{Success = $true; RowsAffected = $rowsAffected; Type = "NonQuery" }
        }
    } catch {
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] CATCH: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Tipo de excepción: $($_.Exception.GetType().FullName)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Stack: $($_.ScriptStackTrace)"
        return @{Success = $false; ErrorMessage = $_.Exception.Message; Type = "Error" }
    } finally {
        if ($plainPassword) { $plainPassword = $null }
        if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] FINALLY: connection = $($connection)"
        if ($null -ne $connection) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Estado antes de cerrar: $($connection.State)"
            if ($connection.State -eq [System.Data.ConnectionState]::Open) { Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Cerrando conexión..."; $connection.Close() }
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Disposing conexión..."
            $connection.Dispose()
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] FIN"
    }
}
function Invoke-SqlQueryMultiResultSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][scriptblock]$InfoMessageCallback
    )
    $connection = $null
    $reader = $null
    Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] INICIO"
    $script:HadSqlError = $false
    try {
        $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
        $connectionString = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $messages = New-Object System.Collections.Generic.List[string]
        $script:HadSqlError = $false
        $connection.add_InfoMessage({
                param($sender, $e)
                $msg = [string]$e.Message
                if (-not [string]::IsNullOrWhiteSpace($msg)) {
                    Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] InfoMessage: $msg" -Color DarkYellow
                    [void]$messages.Add($msg)
                }
                if ($e.Errors) {
                    foreach ($err in $e.Errors) {
                        # Evita imprimir basura vacía
                        $em = [string]$err.Message
                        if ([string]::IsNullOrWhiteSpace($em)) { continue }
                        # Solo severidad >= 11 es error “real”
                        if ($err.Class -ge 11) {
                            $script:HadSqlError = $true
                            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] SQL ERROR: $em (Class:$($err.Class), State:$($err.State))" -Color Red
                        } else {
                            # Es info / warning leve
                            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] SQL INFO: $em (Class:$($err.Class), State:$($err.State))" -Color Gray
                        }
                    }
                }
                if ($InfoMessageCallback -and -not [string]::IsNullOrWhiteSpace($msg)) {
                    try { & $InfoMessageCallback $msg } catch { }
                }
            })
        $connection.FireInfoMessageEventOnUserErrors = $true
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 0
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] QUERY:`n$Query"
        $reader = $command.ExecuteReader()
        $resultSets = New-Object System.Collections.Generic.List[object]
        while ($true) {
            if (-not $reader.IsClosed -and $reader.FieldCount -gt 0) {
                $dt = New-Object System.Data.DataTable
                for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    $colName = $reader.GetName($i)
                    if ([string]::IsNullOrWhiteSpace($colName)) { $colName = "Column$($i+1)" }
                    [void]$dt.Columns.Add($colName)
                }
                while ($reader.Read()) {
                    $row = $dt.NewRow()
                    for ($i = 0; $i -lt $reader.FieldCount; $i++) { $row[$i] = $reader.GetValue($i) }
                    [void]$dt.Rows.Add($row)
                }
                $resultSets.Add([PSCustomObject]@{ DataTable = $dt; RowCount = $dt.Rows.Count })
                Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RS#$($resultSets.Count) Filas: $($dt.Rows.Count) Columnas: $($dt.Columns.Count)"
            }
            if ($reader.IsClosed) { break }
            $hasNext = $false
            try { $hasNext = $reader.NextResult() } catch { break }
            if (-not $hasNext) { break }
        }
        $recordsAffected = -1
        try { if ($reader -and -not $reader.IsClosed) { $recordsAffected = $reader.RecordsAffected } } catch {}
        if ($reader -and -not $reader.IsClosed) { $reader.Close() }
        if ($resultSets.Count -eq 0 -and $recordsAffected -ge 0) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] ResultSets: $($resultSets.Count)"
            $idx = 0
            foreach ($rs in $resultSets) { $idx++; Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RS#$idx Filas: $($rs.RowCount)" }
            $allMsgs = ""
            try { $allMsgs = ($messages -join "`n") } catch {}
            $looksLikeError = $false
            if ($allMsgs -match '(?i)\b(no es válido|invalid object name|incorrect syntax|error\s+de\s+sql|msg\s+\d+|level\s+\d+)\b') { $looksLikeError = $true }
            if ($script:HadSqlError -or $looksLikeError) {
                return @{ Success = $false; ErrorMessage = $allMsgs; Type = "Error"; Messages = $messages; ResultSets = @() }
            }
            return @{ Success = $true; ResultSets = @(); RowsAffected = $recordsAffected; Messages = $messages }
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] ResultSets: $($resultSets.Count) TotalFilas: $((($resultSets | ForEach-Object { $_.RowCount }) | Measure-Object -Sum).Sum)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RESUMEN FINAL ANTES DE RETORNAR:"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] - resultSets.Count: $($resultSets.Count)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] - recordsAffected: $recordsAffected"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] - script:HadSqlError: $($script:HadSqlError)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] - messages count: $($messages.Count)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] - Últimos mensajes:"
        if ($messages.Count -gt 0) { for ($i = [math]::Max(0, $messages.Count - 3); $i -lt $messages.Count; $i++) { Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet]   [$i]: $($messages[$i])" -Color Gray } }
        if ($script:HadSqlError) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RETORNANDO ERROR (HadSqlError = true)" -Color Red
            $allMsgs = $messages -join "`n"
            return @{ Success = $false; ErrorMessage = $allMsgs; Type = "Error"; Messages = $messages; ResultSets = @() }
        }
        $allMsgs = $messages -join "`n"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] Revisando contenido de mensajes para errores..." -Color DarkGray
        if ($allMsgs -match '(?i)\b(no es válido|invalid object name|incorrect syntax|error|sintaxis incorrecta|no existe|does not exist|not found|no se pudo|cannot|no se puede)\b') {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] ERROR DETECTADO EN CONTENIDO DE MENSAJES" -Color Red
            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] Fragmento detectado: $($Matches[0])" -Color Red
            return @{ Success = $false; ErrorMessage = $allMsgs; Type = "Error"; Messages = $messages; ResultSets = @() }
        }
        if ($resultSets.Count -eq 0 -and $recordsAffected -ge 0) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RETORNANDO: Sin resultsets pero con recordsAffected" -Color Cyan
            return @{ Success = $true; ResultSets = @(); RowsAffected = $recordsAffected; Messages = $messages }
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] RETORNANDO: Con resultsets" -Color Green
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] ResultSets: $($resultSets.Count) TotalFilas: $((($resultSets | ForEach-Object { $_.RowCount }) | Measure-Object -Sum).Sum)"
        return @{ Success = $true; ResultSets = $resultSets.ToArray(); Messages = $messages }
    } catch {
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] CATCH: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] Tipo: $($_.Exception.GetType().FullName)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] Stack: $($_.ScriptStackTrace)"
        return @{ Success = $false; ErrorMessage = $_.Exception.Message; Type = "Error" }
    } finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($plainPassword) { $plainPassword = $null }
        if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
        if ($null -ne $connection) {
            try { if ($connection.State -eq [System.Data.ConnectionState]::Open) { $connection.Close() } } catch {}
            try { $connection.Dispose() } catch {}
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQueryMultiResultSet] FIN"
    }
}
function Remove-SqlComments { [CmdletBinding()] param([Parameter(Mandatory = $true)][string]$Query) $query = $Query -replace '(?s)/\*.*?\*/', ''; $query = $query -replace '(?m)^\s*--.*\n?', ''; $query = $query -replace '(?<!\w)--.*$', ''; $query.Trim() }
function Get-SqlDatabases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )
    $query = @"
SELECT name
FROM sys.databases
WHERE name NOT IN ('tempdb','model','msdb')
  AND state_desc = 'ONLINE'
ORDER BY CASE WHEN name = 'master' THEN 0 ELSE 1 END, name
"@
    $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $query -Credential $Credential
    if (-not $result.Success) { throw "Error obteniendo bases de datos: $($result.ErrorMessage)" }
    $databases = @()
    foreach ($row in $result.DataTable.Rows) { $databases += $row["name"] }
    $databases
}
function Backup-Database {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][string]$Database, [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential, [Parameter(Mandatory = $true)][string]$BackupPath, [Parameter(Mandatory = $false)][scriptblock]$ProgressCallback)
    try {
        Write-DzDebug "`t[DEBUG][Backup-Database] Iniciando backup de '$Database' a '$BackupPath'"
        $backupQuery = @"
BACKUP DATABASE [$Database]
TO DISK='$BackupPath'
WITH CHECKSUM, STATS = 1, FORMAT, INIT
"@
        $infoCallback = { param($message) if ($ProgressCallback) { & $ProgressCallback $message } }
        $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $backupQuery -Credential $Credential -InfoMessageCallback $infoCallback
        if ($result.Success) { Write-DzDebug "`t[DEBUG][Backup-Database] Backup completado exitosamente"; return @{Success = $true; BackupPath = $BackupPath } }
        Write-DzDebug "`t[DEBUG][Backup-Database] Error en backup: $($result.ErrorMessage)"
        @{Success = $false; ErrorMessage = $result.ErrorMessage }
    } catch {
        Write-DzDebug "`t[DEBUG][Backup-Database] Excepción: $($_.Exception.Message)"
        @{Success = $false; ErrorMessage = $_.Exception.Message }
    }
}
function Execute-SqlQuery {
    param([string]$server, [string]$database, [string]$query)
    $connection = $null
    try {
        $connectionString = "Server=$server;Database=$database;User Id=$global:user;Password=$global:password;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $infoMessages = New-Object System.Collections.ArrayList
        $connection.add_InfoMessage({ param($sender, $e) [void]$infoMessages.Add($e.Message) })
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $returnsResultSet = $query -match "(?si)^\s*(SELECT|WITH)" -or $query -match "(?si)\bOUTPUT\b"
        if ($returnsResultSet) {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)
            @{DataTable = $dataTable; Messages = $infoMessages }
        } else {
            $rowsAffected = $command.ExecuteNonQuery()
            @{RowsAffected = $rowsAffected; Messages = $infoMessages }
        }
    } catch { throw $_ } finally {
        if ($null -ne $connection) {
            try { if ($connection.State -ne [System.Data.ConnectionState]::Closed) { $connection.Close() } } catch {}
            try { $connection.Dispose() } catch {}
        }
    }
}
function Show-ResultsConsole {
    param([string]$query)
    try {
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query
        if ($results -is [hashtable] -and $results.ContainsKey('DataTable')) {
            $dt = $results.DataTable
            if ($dt.Rows.Count -gt 0) {
                $cols = @($dt.Columns | ForEach-Object { $_.ColumnName })
                $width = @{}
                foreach ($c in $cols) { $width[$c] = [Math]::Max($c.Length, 1) }
                foreach ($r in $dt.Rows) { foreach ($c in $cols) { $width[$c] = [Math]::Max($width[$c], ([string]$r[$c]).Length) } }
                $header = ($cols | ForEach-Object { $_.PadRight($width[$_] + 4) }) -join ''
                Write-Host ""
                Write-Host $header
                Write-Host ("-" * $header.Length)
                foreach ($r in $dt.Rows) { Write-Host (($cols | ForEach-Object { ([string]$r[$_]).PadRight($width[$_] + 4) }) -join '') }
            } else { Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow }
        } elseif ($results -is [hashtable] -and $results.ContainsKey('RowsAffected')) {
            Write-Host "`nFilas afectadas: $($results.RowsAffected)" -ForegroundColor Green
        } else {
            Write-Host "`nSin resultados." -ForegroundColor Yellow
        }
    } catch { Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red }
}
function Get-IniConnections {
    $connections = @()
    $pathsToCheck = @(
        @{Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
        @{Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
        @{Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
        @{Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
        @{Path = "C:\Program Files (x86)\NsBackOffice1.0"; INI = "DbConfig.ini"; Nombre = "NSBackOffice" },
        @{Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
        @{Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
    )
    function Get-IniValue { param([string]$FilePath, [string]$Key) if (Test-Path $FilePath) { $line = Get-Content $FilePath | Where-Object { $_ -match "^$Key\s*=" } | Select-Object -First 1; if ($line) { return $line.Split('=')[1].Trim() } } $null }
    foreach ($entry in $pathsToCheck) {
        $mainIni = Join-Path $entry.Path $entry.INI
        if (Test-Path $mainIni) { $dataSource = Get-IniValue -FilePath $mainIni -Key "DataSource"; if ($dataSource -and $dataSource -notin $connections) { $connections += $dataSource } }
        $inisFolder = Join-Path $entry.Path "INIS"
        if (Test-Path $inisFolder) {
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini" -ErrorAction SilentlyContinue
            foreach ($iniFile in $iniFiles) { $dataSource = Get-IniValue -FilePath $iniFile.FullName -Key "DataSource"; if ($dataSource -and $dataSource -notin $connections) { $connections += $dataSource } }
        }
    }
    $connections | Sort-Object
}
function Load-IniConnectionsToComboBox {
    param([Parameter(Mandatory = $true)]$Combo)
    $connections = Get-IniConnections
    $Combo.Items.Clear()
    if ($connections.Count -gt 0) { foreach ($c in $connections) { [void]$Combo.Items.Add($c) } } else { Write-Host "`tNo se encontraron conexiones en archivos INI" -ForegroundColor Yellow }
    $Combo.Text = ".\NationalSoft"
}
function ConvertTo-DataTable {
    param($InputObject)
    $dt = New-Object System.Data.DataTable
    if (-not $InputObject) { return $dt }
    $cols = $InputObject[0].Keys
    foreach ($c in $cols) { [void]$dt.Columns.Add($c) }
    foreach ($row in $InputObject) { $dr = $dt.NewRow(); foreach ($c in $cols) { $dr[$c] = $row[$c] }; [void]$dt.Rows.Add($dr) }
    $dt
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
                    if ($InfoMessageCallback) { $connection.add_InfoMessage({ param($sender, $e) try { & $InfoMessageCallback $e.Message } catch {} }); $connection.FireInfoMessageEventOnUserErrors = $true }
                    $connection.Open()
                    $cmd = $connection.CreateCommand()
                    $cmd.CommandText = $Query
                    $cmd.CommandTimeout = 0
                    [void]$cmd.ExecuteNonQuery()
                    @{Success = $true }
                } catch { @{Success = $false; ErrorMessage = $_.Exception.Message } } finally {
                    if ($plainPassword) { $plainPassword = $null }
                    if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                    if ($connection) { try { $connection.Close() } catch {}; try { $connection.Dispose() } catch {} }
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
                if (Test-Path $scriptBackupPath) {
                    $choice = Ui-Confirm "Ya existe un respaldo con ese nombre en:`n$scriptBackupPath`n`n¿Deseas sobrescribirlo?" "Archivo existente" $window
                    if (-not $choice) { $timestampsDefault = Get-Date -Format 'yyyyMMdd-HHmmss'; $txtNombre.Text = "$Database-$timestampsDefault.bak"; Add-Log "⚠️ Operación cancelada: el archivo ya existe. Se sugirió un nuevo nombre."; Reset-BackupUI -ProgressText "Cancelado (elige otro nombre y vuelve a intentar)"; return }
                }
                if ($sameHost) { if (-not (Test-Path $sqlBackupFolder)) { New-Item -ItemType Directory -Path $sqlBackupFolder -Force | Out-Null } } else { $uncFolder = "\\$machineName\C$\Temp\SQLBackups"; if (-not (Test-Path $uncFolder)) { Add-Log "⚠️ No pude validar la carpeta UNC: $uncFolder (puede ser permisos). SQL intentará escribir en $sqlBackupFolder en el servidor." } }
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
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch {}
            try { if ($timer -and $timer.IsEnabled) { $timer.Stop() } } catch {}
            try { if ($stopwatch) { $stopwatch.Stop() } } catch {}
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
function Show-RestoreDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][string]$User, [Parameter(Mandatory = $true)][string]$Password, [Parameter(Mandatory = $true)][string]$Database)
    $script:RestoreRunning = $false
    $script:RestoreDone = $false
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
                    # Estado compartido (reference type) para que el handler pueda modificarlo
                    $state = [pscustomobject]@{
                        HasError = $false
                        Errors   = New-Object System.Collections.Generic.List[string]
                    }

                    if ($InfoMessageCallback) {

                        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
                            param($sender, $e)

                            try {
                                $msg = [string]$e.Message
                                if ([string]::IsNullOrWhiteSpace($msg)) { return }

                                # Progreso: soporta "10 percent", "10% ", "10 por ciento"
                                $isProgressMsg = $msg -match '(?i)\b\d{1,3}\s*(%|percent|porcentaje|por\s+ciento)\b'

                                # Mensajes "ok" comunes
                                $isSuccessMsg = $msg -match '(?i)(successfully processed|procesad[oa]\s+correctamente|se proces[oó]\s+correctamente)'

                                # Errores críticos (sin confundir progreso)
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
                    @{Success = $true }
                } catch {
                    @{Success = $false; ErrorMessage = $_.Exception.Message; IsInfoMessageError = $false }
                } finally {
                    if ($plainPassword) { $plainPassword = $null }
                    if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
                    if ($connection) { try { $connection.Close() } catch {}; try { $connection.Dispose() } catch {} }
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
                    # 2) Éxito del paso
                    if ($m -match '(?i)\b(processed\s+\d{1,3}\s*percent\b|se proces[oó] correctamente|completad[oa])') {
                        EnqProg 100 "Restauración completada."
                        EnqLog "✅ Restauración finalizada"
                        return
                    }

                    # 3) Errores
                    $isCriticalError = $m -match '(?i)(abnormal termination|failed to|error|cannot restore|not compatible|no es compatible|imposible|terminado an[oó]malo)'

                    if ($IsError -or $isCriticalError) {
                        $script:hasErrorFlag = $true
                        $script:allErrors += $m
                        if (-not $script:errorDetails -or $m.Length -gt $script:errorDetails.Length) {
                            $script:errorDetails = $m
                        }
                        EnqLog ("[SQL ERROR] $m")
                        return
                    }

                    # 4) Info normal
                    EnqLog ("[SQL] $m")
                }
                $r = Invoke-SqlQueryLite -Server $Server -Database "master" -Query $RestoreQuery -Credential $Credential -InfoMessageCallback $progressCb

                # Verificar el resultado
                if (-not $r.Success) {
                    EnqProg 0 "❌ Error en restauración"
                    EnqLog ("❌ Error de SQL: {0}" -f $r.ErrorMessage)
                    EnqLog "ERROR_RESULT|$($r.ErrorMessage)"
                    EnqLog "__DONE__"
                    return
                }

                # Verificar si hubo errores en InfoMessages
                if ($hasErrorFlag) {
                    EnqProg 0 "❌ Restauración falló"
                    EnqLog "❌ La restauración no se completó correctamente"

                    # Construir mensaje de error detallado
                    if ($allErrors.Count -gt 0) {
                        # Usar todos los errores si son pocos, o los más importantes
                        if ($allErrors.Count -le 3) {
                            $finalErrorMsg = $allErrors -join "`n`n"
                        } else {
                            # Si hay muchos errores, usar solo los únicos más descriptivos
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

                # Si llegamos aquí, todo salió bien
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

                    # Capturar el resultado final
                    if ($line -like "*SUCCESS_RESULT|*") {
                        $finalResult = @{
                            Success = $true
                            Message = $line -replace '^.*SUCCESS_RESULT\|', ''
                        }
                    }
                    if ($line -like "*ERROR_RESULT|*") {
                        $finalResult = @{
                            Success = $false
                            Message = $line -replace '^.*ERROR_RESULT\|', ''
                        }
                    }

                    # Filtrar las líneas de resultado del log visual
                    if ($line -notmatch '(SUCCESS_RESULT|ERROR_RESULT)') {
                        $txtLog.Text = "$line`n" + $txtLog.Text
                    }

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

                        # Mostrar mensaje final al usuario
                        if ($finalResult) {
                            $window.Dispatcher.Invoke([action] {
                                    if ($finalResult.Success) {
                                        Ui-Info "Base de datos '$($txtDestino.Text)' restaurada con éxito.`n`n$($finalResult.Message)" "✓ Restauración Exitosa" $window
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
                    if ([string]::IsNullOrWhiteSpace($txtDestino.Text)) {
                        $txtDestino.Text = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
                    }
                }
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error btnBrowseBackup: $($_.Exception.Message)"
                Ui-Error "No se pudo abrir el selector de archivos: $($_.Exception.Message)" "Error" $window
            }
        })
    $txtDestino.Add_TextChanged({
            try {
                Update-RestorePaths -DatabaseName $txtDestino.Text
            } catch {
                Write-DzDebug "`t[DEBUG][UI] Error actualizando rutas: $($_.Exception.Message)"
            }
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
                # Reemplaza esta sección en tu función Show-RestoreDialog
                # Busca donde dice: Paint-Progress -Percent 5 -Message "Leyendo metadatos del backup..."

                Paint-Progress -Percent 5 -Message "Leyendo metadatos del backup..."
                $fileListQuery = "RESTORE FILELISTONLY FROM DISK = N'$escapedBackup'"

                # SOLUCIÓN: Crear una función inline que use ExecuteReader para FILELISTONLY
                function Get-BackupFileList {
                    param(
                        [string]$Server,
                        [string]$Query,
                        [System.Management.Automation.PSCredential]$Credential
                    )

                    $connection = $null
                    $passwordBstr = [IntPtr]::Zero
                    $plainPassword = $null

                    try {
                        # Extraer la contraseña de forma segura
                        $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
                        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)

                        # Crear conexión
                        $cs = "Server=$Server;Database=master;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
                        $connection = New-Object System.Data.SqlClient.SqlConnection($cs)
                        $connection.Open()

                        # Ejecutar query
                        $cmd = $connection.CreateCommand()
                        $cmd.CommandText = $Query
                        $cmd.CommandTimeout = 60

                        # CRÍTICO: Usar ExecuteReader para obtener resultados
                        $reader = $cmd.ExecuteReader()
                        $dataTable = New-Object System.Data.DataTable
                        $dataTable.Load($reader)
                        $reader.Close()

                        return @{
                            Success   = $true
                            DataTable = $dataTable
                        }

                    } catch {
                        return @{
                            Success      = $false
                            ErrorMessage = $_.Exception.Message
                            DataTable    = $null
                        }
                    } finally {
                        # Limpiar credenciales de memoria
                        if ($plainPassword) {
                            $plainPassword = $null
                        }
                        if ($passwordBstr -ne [IntPtr]::Zero) {
                            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
                        }
                        if ($connection) {
                            try { $connection.Close() } catch {}
                            try { $connection.Dispose() } catch {}
                        }
                    }
                }

                # Usar la nueva función
                $fileListResult = Get-BackupFileList -Server $Server -Query $fileListQuery -Credential $credential

                if (-not $fileListResult.Success) {
                    Write-DzDebug "`t[DEBUG][Show-RestoreDialog] Error FILELISTONLY: $($fileListResult.ErrorMessage)"
                    Ui-Error "No se pudo leer el contenido del backup: $($fileListResult.ErrorMessage)" "Error" $window
                    Reset-RestoreUI -ProgressText "Error leyendo backup"
                    return
                }

                # Validar que se obtuvieron resultados
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

                    if (-not $logicalData -and $type -eq "D") {
                        $logicalData = $logicalName
                        Add-Log "Encontrado archivo de datos: $logicalName"
                    }
                    if (-not $logicalLog -and $type -eq "L") {
                        $logicalLog = $logicalName
                        Add-Log "Encontrado archivo de log: $logicalName"
                    }
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
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch {}
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
Export-ModuleMember -Function @('Invoke-SqlQuery', 'Invoke-SqlQueryMultiResultSet', 'Remove-SqlComments', 'Get-SqlDatabases', 'Backup-Database', 'Execute-SqlQuery', 'Show-ResultsConsole', 'Get-IniConnections', 'Load-IniConnectionsToComboBox', 'ConvertTo-DataTable', 'Show-BackupDialog', 'Show-RestoreDialog', 'Reset-BackupUI', 'Reset-RestoreUI')
