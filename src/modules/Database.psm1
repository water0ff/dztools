if ($PSVersionTable.PSVersion.Major -lt 5) { throw "Se requiere PowerShell 5.0 o superior." }
$script:queryTabCounter = 1
if ([string]::IsNullOrWhiteSpace($global:DzSqlKeywords)) { $global:DzSqlKeywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE' }
function Process-SqlProgressMessage { param([string]$Message) if ($Message -match '(?i)(\d{1,3})\s*percent') { $percent = [int]$Matches[1]; Write-Output "Progreso: $percent%" } elseif ($Message -match 'BACKUP DATABASE successfully processed') { Write-Output "Backup completado exitosamente" } }
function New-DzSqlConnectionFromCredential {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )
    $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
    $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)
    $connectionString = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"
    @{
        Connection    = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        PlainPassword = $plainPassword
        PasswordBstr  = $passwordBstr
    }
}
function New-DzSqlConnectionFromPlain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$User,
        [Parameter(Mandatory = $true)][string]$Password
    )
    $connectionString = "Server=$Server;Database=$Database;User Id=$User;Password=$Password;MultipleActiveResultSets=True"
    New-Object System.Data.SqlClient.SqlConnection($connectionString)
}
function Resolve-DzSqlMessageSummary {
    [CmdletBinding()]
    param([Parameter()][System.Collections.Generic.List[string]]$Messages)
    if (-not $Messages) { return "" }
    $unique = @($Messages | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($unique.Count -eq 0) { return "" }
    $unique -join "`n"
}
function Write-DzSqlResultSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Result,
        [Parameter()][string]$Context
    )
    $prefix = if ([string]::IsNullOrWhiteSpace($Context)) { "" } else { "[$Context] " }
    if (-not $Result.Success) {
        Write-Host "$prefix Error SQL: $($Result.ErrorMessage)" -ForegroundColor Red
        return
    }
    if ($Result.ContainsKey('ResultSets') -and $Result.ResultSets -and $Result.ResultSets.Count -gt 0) {
        $totalRows = ($Result.ResultSets | Measure-Object -Property RowCount -Sum).Sum
        Write-Host "$prefix Resultado: $($Result.ResultSets.Count) conjunto(s), $totalRows fila(s)."
        return
    }
    if ($Result.ContainsKey('DataTable') -and $Result.DataTable) {
        Write-Host "$prefix Resultado: $($Result.DataTable.Rows.Count) fila(s)."
        return
    }
    if ($Result.ContainsKey('RowsAffected')) {
        Write-Host "$prefix Filas afectadas: $($Result.RowsAffected)"
        return
    }
    Write-Host "$prefix Sin resultados."
}
function Invoke-DzSqlCommandInternal {
    [CmdletBinding(DefaultParameterSetName = "Credential")]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(ParameterSetName = "Credential", Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(ParameterSetName = "Plain", Mandatory = $true)][string]$User,
        [Parameter(ParameterSetName = "Plain", Mandatory = $true)][string]$Password,
        [Parameter()][scriptblock]$InfoMessageCallback,
        [Parameter()][switch]$CollectMessages
    )
    $connection = $null
    $passwordBstr = [IntPtr]::Zero
    $plainPassword = $null
    $messages = New-Object System.Collections.Generic.List[string]
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] INICIO"
    Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] Server='$Server' Database='$Database'"
    try {
        if ($PSCmdlet.ParameterSetName -eq "Credential") {
            $connInfo = New-DzSqlConnectionFromCredential -Server $Server -Database $Database -Credential $Credential
            $connection = $connInfo.Connection
            $plainPassword = $connInfo.PlainPassword
            $passwordBstr = $connInfo.PasswordBstr
        } else {
            $connection = New-DzSqlConnectionFromPlain -Server $Server -Database $Database -User $User -Password $Password
        }
        if ($CollectMessages -or $InfoMessageCallback) {
            $connection.add_InfoMessage({
                    param($sender, $e)
                    $msg = [string]$e.Message
                    if (-not [string]::IsNullOrWhiteSpace($msg)) {
                        try { [void]$messages.Add($msg) } catch {}
                        if ($InfoMessageCallback) { try { & $InfoMessageCallback $msg } catch { Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] Error en InfoMessageCallback: $_" } }
                    }
                })
            $connection.FireInfoMessageEventOnUserErrors = $true
        }
        Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] Abriendo conexión..."
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 0
        $returnsResultSet = $Query -match "(?si)^\s*(SELECT|WITH)" -or $Query -match "(?si)\bOUTPUT\b"
        if ($returnsResultSet) {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)
            return @{
                Success    = $true
                DataTable  = $dataTable
                Type       = "Query"
                DurationMs = $stopwatch.ElapsedMilliseconds
                Messages   = $messages
            }
        }
        $rowsAffected = $command.ExecuteNonQuery()
        return @{
            Success      = $true
            RowsAffected = $rowsAffected
            Type         = "NonQuery"
            DurationMs   = $stopwatch.ElapsedMilliseconds
            Messages     = $messages
        }
    } catch {
        Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] CATCH: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] Tipo de excepción: $($_.Exception.GetType().FullName)"
        Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] Stack: $($_.ScriptStackTrace)"
        return @{
            Success      = $false
            ErrorMessage = $_.Exception.Message
            ErrorRecord  = $_
            Type         = "Error"
            DurationMs   = $stopwatch.ElapsedMilliseconds
            Messages     = $messages
        }
    } finally {
        if ($plainPassword) { $plainPassword = $null }
        if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
        if ($null -ne $connection) {
            try { if ($connection.State -eq [System.Data.ConnectionState]::Open) { $connection.Close() } } catch {}
            try { $connection.Dispose() } catch {}
        }
        try { $stopwatch.Stop() } catch {}
        Write-DzDebug "`t[DEBUG][Invoke-DzSqlCommandInternal] FIN"
    }
}
function Invoke-DzSqlBatchInternal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter()][scriptblock]$InfoMessageCallback
    )
    [string]$Query = $Query
    $connection = $null
    $passwordBstr = [IntPtr]::Zero
    $plainPassword = $null
    $messages = New-Object System.Collections.Generic.List[string]
    $debugLog = New-Object System.Collections.Generic.List[string]
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:HadSqlError = $false
    function Add-Debug([string]$m) {
        try { [void]$debugLog.Add(("{0} {1}" -f (Get-Date -Format "HH:mm:ss.fff"), $m)) } catch {}
        try { Write-DzDebug "`t[DEBUG][Invoke-DzSqlBatchInternal] $m" } catch {}
    }
    Add-Debug "INICIO"
    Add-Debug ("Server='{0}' DB='{1}' User='{2}' QueryLen={3}" -f $Server, $Database, $Credential.UserName, ($Query?.Length))
    try {
        $connInfo = New-DzSqlConnectionFromCredential -Server $Server -Database $Database -Credential $Credential
        $connection = $connInfo.Connection
        $plainPassword = $connInfo.PlainPassword
        $passwordBstr = $connInfo.PasswordBstr
        $connection.add_InfoMessage({
                param($sender, $e)
                $msg = [string]$e.Message
                if (-not [string]::IsNullOrWhiteSpace($msg)) {
                    try { [void]$messages.Add($msg) } catch {}
                    try { if ($InfoMessageCallback) { & $InfoMessageCallback $msg } } catch {}
                }
                if ($e.Errors) {
                    foreach ($err in $e.Errors) {
                        $em = [string]$err.Message
                        if ([string]::IsNullOrWhiteSpace($em)) { continue }
                        if ($err.Class -ge 11) { $script:HadSqlError = $true }
                    }
                }
            })
        $connection.FireInfoMessageEventOnUserErrors = $true
        Add-Debug "Abriendo conexión..."
        $connection.Open()
        Add-Debug ("Estado conexión: {0}" -f $connection.State)
        $batches = @([System.Text.RegularExpressions.Regex]::Split($Query, '(?im)^\s*GO\s*$') | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        Add-Debug ("Batches: {0}" -f $batches.Count)
        if ($batches.Count -gt 0) {
            $preview = $batches[0]
            if ($preview.Length -gt 300) { $preview = $preview.Substring(0, 300) + "..." }
            Add-Debug ("Batch[0] Preview: {0}" -f $preview.Replace("`r", " ").Replace("`n", " "))
        }
        $resultSets = New-Object System.Collections.Generic.List[object]
        $recordsAffected = $null
        $totalRowsAffected = 0
        $batchErrors = New-Object System.Collections.Generic.List[string]
        for ($batchIndex = 0; $batchIndex -lt $batches.Count; $batchIndex++) {
            $oneBatch = $batches[$batchIndex]
            $command = $connection.CreateCommand()
            $command.CommandTimeout = 0
            $command.CommandText = $oneBatch
            Add-Debug ("Ejecutando batch #{0} (len={1})..." -f ($batchIndex + 1), $oneBatch.Length)
            $isSelect = $oneBatch -match "(?si)^\s*(SELECT|WITH|EXEC|EXECUTE|DECLARE.*SELECT)" -or $oneBatch -match "(?si)\bOUTPUT\b"
            try {
                if ($isSelect) {
                    try {
                        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
                        $ds = New-Object System.Data.DataSet
                        $filledTables = $adapter.Fill($ds)
                        Add-Debug ("Adapter.Fill OK. ds.Tables={0} filledTables={1}" -f $ds.Tables.Count, $filledTables)
                        if ($ds.Tables.Count -gt 0) {
                            foreach ($t in $ds.Tables) {
                                for ($i = 0; $i -lt $t.Columns.Count; $i++) {
                                    if ([string]::IsNullOrWhiteSpace($t.Columns[$i].ColumnName)) {
                                        $t.Columns[$i].ColumnName = "Column$($i+1)"
                                    }
                                }
                                $resultSets.Add([PSCustomObject]@{ DataTable = $t; RowCount = $t.Rows.Count })
                                Add-Debug ("RS#{0} Filas={1} Cols={2}" -f $resultSets.Count, $t.Rows.Count, $t.Columns.Count)
                            }
                        }
                    } catch {
                        Add-Debug ("Adapter.Fill EX: {0}" -f $_.Exception.Message)
                        [void]$batchErrors.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                        [void]$messages.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                        break
                    }
                } else {
                    try {
                        $rows = $command.ExecuteNonQuery()
                        $totalRowsAffected += $rows
                        Add-Debug ("ExecuteNonQuery RowsAffected={0} (Total acumulado={1})" -f $rows, $totalRowsAffected)
                    } catch {
                        Add-Debug ("ExecuteNonQuery EX: {0}" -f $_.Exception.Message)
                        [void]$batchErrors.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                        [void]$messages.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                        break
                    }
                }
                $command.Dispose()
                $command = $null
            } catch {
                Add-Debug ("Error general en batch #{0}: {1}" -f ($batchIndex + 1), $_.Exception.Message)
                [void]$batchErrors.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                [void]$messages.Add("Batch #$($batchIndex + 1): $($_.Exception.Message)")
                break
            }
        }
        $allMsgs = Resolve-DzSqlMessageSummary -Messages $messages
        Add-Debug ("FIN loop. resultSets={0} totalRowsAffected={1} HadSqlError={2} messages={3} batchErrors={4}" -f $resultSets.Count, $totalRowsAffected, $script:HadSqlError, $messages.Count, $batchErrors.Count)
        if ($batchErrors.Count -gt 0) {
            $errorMessage = "Error en batch: " + ($batchErrors -join "; ")
            return @{
                Success      = $false
                Type         = "Error"
                ErrorMessage = $errorMessage
                Messages     = $messages
                ResultSets   = $resultSets.ToArray()  # Devolver los resultados que SÍ se obtuvieron
                DebugLog     = $debugLog
                DurationMs   = $stopwatch.ElapsedMilliseconds
            }
        }
        if ($script:HadSqlError) {
            if ([string]::IsNullOrWhiteSpace($allMsgs)) { $allMsgs = "Error SQL." }
            return @{
                Success      = $false
                Type         = "Error"
                ErrorMessage = $allMsgs
                Messages     = $messages
                ResultSets   = $resultSets.ToArray()
                DebugLog     = $debugLog
                DurationMs   = $stopwatch.ElapsedMilliseconds
            }
        }
        if ($resultSets.Count -gt 0) {
            return @{
                Success      = $true
                ResultSets   = $resultSets.ToArray()
                Messages     = $messages
                DebugLog     = $debugLog
                DurationMs   = $stopwatch.ElapsedMilliseconds
                RowsAffected = if ($totalRowsAffected -gt 0) { $totalRowsAffected } else { $null }
            }
        }
        return @{
            Success      = $true
            ResultSets   = @()
            RowsAffected = $totalRowsAffected
            Messages     = $messages
            DebugLog     = $debugLog
            DurationMs   = $stopwatch.ElapsedMilliseconds
        }
    } catch {
        Add-Debug ("CATCH: {0}" -f $_.Exception.Message)
        $summary = Resolve-DzSqlMessageSummary -Messages $messages
        $errorMessage = if (-not [string]::IsNullOrWhiteSpace($_.Exception.Message)) { $_.Exception.Message } elseif (-not [string]::IsNullOrWhiteSpace($summary)) { $summary } else { "Error ejecutando consulta." }
        return @{
            Success      = $false
            Type         = "Error"
            ErrorMessage = $errorMessage
            ErrorRecord  = $_
            Messages     = $messages
            ResultSets   = @()
            DebugLog     = $debugLog
            DurationMs   = $stopwatch.ElapsedMilliseconds
        }
    } finally {
        if ($plainPassword) { $plainPassword = $null }
        if ($passwordBstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr) }
        if ($connection) {
            try { if ($connection.State -eq [System.Data.ConnectionState]::Open) { $connection.Close() } } catch {}
            try { $connection.Dispose() } catch {}
        }
        try { $stopwatch.Stop() } catch {}
        try { Add-Debug "FIN" } catch {}
    }
}
function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Query,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][scriptblock]$InfoMessageCallback
    )
    Invoke-DzSqlCommandInternal -Server $Server -Database $Database -Query $Query -Credential $Credential -InfoMessageCallback $InfoMessageCallback
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
    Invoke-DzSqlBatchInternal -Server $Server -Database $Database -Query $Query -Credential $Credential -InfoMessageCallback $InfoMessageCallback
}
function Remove-SqlComments {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Query)
    [string]$query = $Query
    $query = $query -replace '(?s)/\*.*?\*/', ''
    $query = $query -replace '(?m)^\s*--.*\n?', ''
    $query = $query -replace '(?<!\w)--.*$', ''
    return [string]($query.Trim())
}
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

ORDER BY CASE WHEN name = 'master' THEN 0 ELSE 1 END, name
"@
    #AND state_desc = 'ONLINE' de momento mostrar todas!
    $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $query -Credential $Credential
    if (-not $result.Success) { throw "Error obteniendo bases de datos: $($result.ErrorMessage)" }
    $databases = @()
    foreach ($row in $result.DataTable.Rows) { $databases += $row["name"] }
    $databases
}
function Get-SqlDatabasesInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )
    $query = @"
SELECT
    name,
    state_desc,
    user_access_desc,
    CAST(is_read_only AS bit) AS is_read_only
FROM sys.databases
WHERE name NOT IN ('tempdb', 'model', 'msdb')
ORDER BY
    CASE WHEN name = 'master' THEN 0 ELSE 1 END,
    name
"@
    $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $query -Credential $Credential
    if (-not $result.Success) {
        throw "Error obteniendo bases de datos: $($result.ErrorMessage)"
    }
    $databases = @()
    foreach ($row in $result.DataTable.Rows) {
        $databases += [PSCustomObject]@{
            Name           = $row["name"]
            StateDesc      = $row["state_desc"]
            UserAccessDesc = $row["user_access_desc"]
            IsReadOnly     = [bool]$row["is_read_only"]
        }
    }
    Write-DzDebug "`t[DEBUG][Get-SqlDatabasesInfo] Retornando $($databases.Count) bases de datos"
    return $databases
}
function Backup-Database {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)][string]$BackupPath,
        [Parameter(Mandatory = $false)][scriptblock]$ProgressCallback
    )
    try {
        Write-DzDebug "`t[DEBUG][Backup-Database] Iniciando backup de '$Database' a '$BackupPath'"
        $backupQuery = @"
BACKUP DATABASE [$Database]
TO DISK='$BackupPath'
WITH CHECKSUM, STATS = 1, FORMAT, INIT
"@
        $infoCallback = { param($message) if ($ProgressCallback) { & $ProgressCallback $message } }
        $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $backupQuery -Credential $Credential -InfoMessageCallback $infoCallback
        if ($result.Success) { Write-DzDebug "`t[DEBUG][Backup-Database] Backup completado exitosamente"; return @{ Success = $true; BackupPath = $BackupPath } }
        Write-DzDebug "`t[DEBUG][Backup-Database] Error en backup: $($result.ErrorMessage)"
        @{ Success = $false; ErrorMessage = $result.ErrorMessage }
    } catch {
        Write-DzDebug "`t[DEBUG][Backup-Database] Excepción: $($_.Exception.Message)"
        @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}
function Execute-SqlQuery {
    param([string]$server, [string]$database, [string]$query)
    $result = Invoke-DzSqlCommandInternal -Server $server -Database $database -Query $query -User $global:user -Password $global:password -CollectMessages
    if (-not $result.Success) { throw $result.ErrorMessage }
    return $result
}
function Show-ResultsConsole {
    param([string]$query)
    try {
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query
        Write-DzSqlResultSummary -Result $results -Context "Consulta"
    } catch { Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red }
}
function Get-IniConnections {
    $connections = @()
    $pathsToCheck = @(
        @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
        @{ Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
        @{ Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
        @{ Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
        @{ Path = "C:\Program Files (x86)\NsBackOffice1.0"; INI = "DbConfig.ini"; Nombre = "NSBackOffice" },
        @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
        @{ Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
    )
    function Get-IniValue {
        param([string]$FilePath, [string]$Key)
        if (Test-Path $FilePath) {
            $line = Get-Content $FilePath | Where-Object { $_ -match "^$Key\s*=" } | Select-Object -First 1
            if ($line) { return $line.Split('=')[1].Trim() }
        }
        $null
    }
    foreach ($entry in $pathsToCheck) {
        $mainIni = Join-Path $entry.Path $entry.INI
        if (Test-Path $mainIni) {
            $dataSource = Get-IniValue -FilePath $mainIni -Key "DataSource"
            if ($dataSource -and $dataSource -notin $connections) { $connections += $dataSource }
        }
        $inisFolder = Join-Path $entry.Path "INIS"
        if (Test-Path $inisFolder) {
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini" -ErrorAction SilentlyContinue
            foreach ($iniFile in $iniFiles) {
                $dataSource = Get-IniValue -FilePath $iniFile.FullName -Key "DataSource"
                if ($dataSource -and $dataSource -notin $connections) { $connections += $dataSource }
            }
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
function bdd_RenameFromTree {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $false)][string]$DefaultValue = ""
    )
    $safeTitle = [Security.SecurityElement]::Escape($Title)
    $safePrompt = [Security.SecurityElement]::Escape($Prompt)
    $safeDefault = [Security.SecurityElement]::Escape($DefaultValue)
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$safeTitle"
        Height="260"
        Width="520"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        WindowStyle="None"
        Background="Transparent"
        AllowsTransparency="True"
        ShowInTaskbar="False"
        Topmost="True">
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
                <TextBlock DockPanel.Dock="Left"
                           Text="$safeTitle"
                           Foreground="{DynamicResource FormFg}"
                           FontSize="16"
                           FontWeight="SemiBold"
                           VerticalAlignment="Center"/>
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
                </Grid.RowDefinitions>
                <TextBlock Grid.Row="0"
                           Text="$safePrompt"
                           Margin="0,0,0,10"
                           FontWeight="SemiBold"/>
                <TextBox Grid.Row="1"
                         Name="txtInput"
                         Height="34"
                         Text="$safeDefault"
                         VerticalContentAlignment="Center"/>
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
                           Text="Enter: Aceptar   |   Esc: Cerrar"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1"
                            Orientation="Horizontal">
                    <Button Name="btnCancel"
                            Content="Cancelar"
                            Width="120"
                            Height="34"
                            Margin="0,0,10,0"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <Button Name="btnOk"
                            Content="Aceptar"
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
        $w = $ui.Window
        $c = $ui.Controls
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $w -Theme $theme
        try { Set-WpfDialogOwner -Dialog $w } catch {}
        $txtInput = $c['txtInput']
        $btnOk = $c['btnOk']
        $btnCancel = $c['btnCancel']
        $btnClose = $c['btnClose']
        $brdTitleBar = $w.FindName("brdTitleBar")
        if ($brdTitleBar) {
            $brdTitleBar.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) { try { $w.DragMove() } catch {} }
                })
        }
        $script:renameResult = $null
        $w.Add_Loaded({ $txtInput.Focus(); $txtInput.SelectAll() })
        $btnClose.Add_Click({ $script:renameResult = $null; $w.DialogResult = $false; $w.Close() })
        $btnCancel.Add_Click({ $script:renameResult = $null; $w.DialogResult = $false; $w.Close() })
        $w.Add_PreviewKeyDown({
                param($sender, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $script:renameResult = $null; $w.DialogResult = $false; $w.Close() }
            })
        $btnOk.Add_Click({ $script:renameResult = $txtInput.Text; $w.DialogResult = $true; $w.Close() })
        $result = $w.ShowDialog()
        if ($result -eq $true) { return $script:renameResult }
        return $null
    } catch {
        Write-Error "Error creando diálogo de renombrar: $($_.Exception.Message)"
        return $null
    }
}
function Get-ActiveQueryTab {
    param([Parameter(Mandatory = $true)]$TabControl)
    if (-not $TabControl) { return $null }
    $tab = $TabControl.SelectedItem
    if ($tab -and $tab.Tag -and $tab.Tag.Type -eq 'QueryTab') { return $tab }
    $null
}
function Get-ActiveQueryRichTextBox {
    param([Parameter(Mandatory = $true)]$TabControl)
    $tab = Get-ActiveQueryTab -TabControl $TabControl
    if ($tab -and $tab.Tag -and $tab.Tag.RichTextBox) { return $tab.Tag.RichTextBox }
    $null
}
function Set-QueryTextInActiveTab {
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)][string]$Text
    )
    $rtb = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $rtb) { return }
    $rtb.Document.Blocks.Clear()
    $paragraph = New-Object System.Windows.Documents.Paragraph
    $run = New-Object System.Windows.Documents.Run($Text)
    $paragraph.Inlines.Add($run)
    $rtb.Document.Blocks.Add($paragraph)
}
function Insert-TextIntoActiveQuery {
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)][string]$Text
    )
    $rtb = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $rtb) { return }
    $caret = $rtb.CaretPosition
    if ($caret) {
        try { $caret.InsertTextInRun($Text) } catch { $rtb.CaretPosition = $rtb.Document.ContentEnd; $rtb.CaretPosition.InsertTextInRun($Text) }
    }
    $rtb.Focus()
}
function Clear-ActiveQueryTab {
    param([Parameter(Mandatory = $true)]$TabControl)
    $rtb = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if ($rtb) {
        $rtb.Document.Blocks.Clear()
        $tab = Get-ActiveQueryTab -TabControl $TabControl
        if ($tab -and $tab.Tag) {
            $tab.Tag.IsDirty = $false
            $title = $tab.Tag.Title
            if ($tab.Tag.HeaderTextBlock) { $tab.Tag.HeaderTextBlock.Text = $title }
        }
    }
}
function Update-QueryTabHeader {
    param([Parameter(Mandatory = $true)]$TabItem)
    if (-not $TabItem.Tag) { return }
    $title = $TabItem.Tag.Title
    if ($TabItem.Tag.IsDirty) { $title = "*$title" }
    if ($TabItem.Tag.HeaderTextBlock) { $TabItem.Tag.HeaderTextBlock.Text = $title }
}
function Get-NextQueryNumber {
    param([Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl)
    $max = 0
    foreach ($item in $TabControl.Items) {
        if ($item -isnot [System.Windows.Controls.TabItem]) { continue }
        if (-not $item.Tag -or $item.Tag.Type -ne 'QueryTab') { continue }
        $title = [string]$item.Tag.Title
        if ($title -match 'Consulta\s+(\d+)') {
            $n = [int]$Matches[1]
            if ($n -gt $max) { $max = $n }
        }
    }
    return ($max + 1)
}
function New-QueryTab {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl)
    $tabNumber = Get-NextQueryNumber -TabControl $TabControl
    $tabTitle = "Consulta $tabNumber"
    $tabItem = New-Object System.Windows.Controls.TabItem
    $headerPanel = New-Object System.Windows.Controls.StackPanel
    $headerPanel.Orientation = "Horizontal"
    $headerText = New-Object System.Windows.Controls.TextBlock
    $headerText.Text = $tabTitle
    $headerText.VerticalAlignment = "Center"
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "×"
    $closeButton.Width = 20
    $closeButton.Height = 20
    $closeButton.Margin = "6,0,0,0"
    $closeButton.Padding = "0"
    $closeButton.FontSize = 14
    [void]$headerPanel.Children.Add($headerText)
    [void]$headerPanel.Children.Add($closeButton)
    $tabItem.Header = $headerPanel
    $grid = New-Object System.Windows.Controls.Grid
    $grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition))
    $rtb = New-Object System.Windows.Controls.RichTextBox
    $rtb.Margin = "0"
    $rtb.VerticalScrollBarVisibility = "Auto"
    $rtb.AcceptsReturn = $true
    $rtb.AcceptsTab = $true
    [void]$grid.Children.Add($rtb)
    $tabItem.Content = $grid
    $tabItem.Tag = [pscustomobject]@{ Type = "QueryTab"; RichTextBox = $rtb; Title = $tabTitle; HeaderTextBlock = $headerText; IsDirty = $false }
    $rtb.Add_TextChanged({
            if ($global:isHighlightingQuery) { return }
            $global:isHighlightingQuery = $true
            try {
                if ([string]::IsNullOrWhiteSpace($global:DzSqlKeywords)) { return }
                Set-WpfSqlHighlighting -RichTextBox $rtb -Keywords $global:DzSqlKeywords
                $tabItem.Tag.IsDirty = $true
                Update-QueryTabHeader -TabItem $tabItem
            } finally { $global:isHighlightingQuery = $false }
        }.GetNewClosure())
    $tcRef = $TabControl
    $closeButton.Add_Click({ Close-QueryTab -TabControl $tcRef -TabItem $tabItem }.GetNewClosure())
    $insertIndex = $TabControl.Items.Count
    for ($i = 0; $i -lt $TabControl.Items.Count; $i++) {
        $it = $TabControl.Items[$i]
        if ($it -is [System.Windows.Controls.TabItem] -and $it.Name -eq "tabAddQuery") { $insertIndex = $i; break }
    }
    [void]$TabControl.Items.Insert($insertIndex, $tabItem)
    $TabControl.SelectedItem = $tabItem
    return $tabItem
}
function Close-QueryTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)]$TabItem
    )
    if (-not $TabItem -or $TabItem.Tag.Type -ne 'QueryTab') { return }
    if ($TabItem.Tag.IsDirty) {
        $owner = [System.Windows.Window]::GetWindow($TabControl)
        if (Get-Command Show-WpfMessageBoxSafe -ErrorAction SilentlyContinue) {
            $result = Show-WpfMessageBoxSafe -Message "La consulta tiene cambios sin guardar. ¿Deseas cerrar la pestaña?" -Title "Confirmar" -Buttons "YesNo" -Icon "Warning" -Owner $owner
        } else {
            $result = [System.Windows.MessageBox]::Show("La consulta tiene cambios sin guardar. ¿Deseas cerrar la pestaña?", "Confirmar", "YesNo", "Warning")
        }
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }
    $removedIndex = $TabControl.Items.IndexOf($TabItem)
    $TabControl.Items.Remove($TabItem)
    if ($TabControl.Items.Count -lt 1) { return }
    $targetTab = $null
    for ($i = ($removedIndex - 1); $i -ge 0; $i--) {
        $candidate = $TabControl.Items[$i]
        if ($candidate -and $candidate.Tag -and $candidate.Tag.Type -eq 'QueryTab') { $targetTab = $candidate; break }
    }
    if (-not $targetTab) {
        for ($i = $removedIndex; $i -lt $TabControl.Items.Count; $i++) {
            $candidate = $TabControl.Items[$i]
            if ($candidate -and $candidate.Tag -and $candidate.Tag.Type -eq 'QueryTab') { $targetTab = $candidate; break }
        }
    }
    if ($targetTab) { $TabControl.SelectedItem = $targetTab }
}
function Execute-QueryInTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)]$ResultsTabControl,
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )
    $rtb = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $rtb) { throw "No hay una pestaña de consulta activa." }
    $rawQuery = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
    $cleanQuery = Remove-SqlComments -Query $rawQuery.Text
    if ([string]::IsNullOrWhiteSpace($cleanQuery)) { throw "La consulta está vacía." }
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] Ejecutando consulta en '$Database'"
    $result = Invoke-SqlQueryMultiResultSet -Server $Server -Database $Database -Query $cleanQuery -Credential $Credential
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] Resultado recibido:"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - Success: $($result.Success)"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - ErrorMessage: $($result.ErrorMessage)"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - ResultSets Count: $($result.ResultSets.Count)"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - Type: $($result.Type)"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - Tiene 'RowsAffected': $($result.ContainsKey('RowsAffected'))"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] - RowsAffected valor: $($result.RowsAffected)"
    Write-DzDebug "`t[DEBUG][Execute-QueryInTab] Entrando a las condiciones de resultado..."
    if (-not $result.Success) {
        $ResultsTabControl.Items.Clear()
        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = "Error"
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = $result.ErrorMessage
        $text.Margin = "10"
        $tab.Content = $text
        [void]$ResultsTabControl.Items.Add($tab)
        $ResultsTabControl.SelectedItem = $tab
        Write-DzSqlResultSummary -Result $result -Context "Consulta"
        return $result
    }
    if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
        Write-DzDebug "`t[DEBUG][Execute-QueryInTab] CONDICIÓN 1: Entrando a mostrar ResultSets (count: $($result.ResultSets.Count))"
        Show-MultipleResultSets -TabControl $ResultsTabControl -ResultSets $result.ResultSets
    } elseif ($result.ContainsKey('RowsAffected') -and $result.RowsAffected -ne $null) {
        Write-DzDebug "`t[DEBUG][Execute-QueryInTab] CONDICIÓN 2: Entrando a mostrar RowsAffected ($($result.RowsAffected))"
        $ResultsTabControl.Items.Clear()
        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = "Resultado"
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = "Filas afectadas: $($result.RowsAffected)"
        $text.Margin = "10"
        $tab.Content = $text
        [void]$ResultsTabControl.Items.Add($tab)
        $ResultsTabControl.SelectedItem = $tab
    } else {
        Write-DzDebug "`t[DEBUG][Execute-QueryInTab] CONDICIÓN 3: Entrando a mostrar vacío (sin resultados)"
        Show-MultipleResultSets -TabControl $ResultsTabControl -ResultSets @()
    }
    Write-DzSqlResultSummary -Result $result -Context "Consulta"
    return $result
}
function Show-MultipleResultSets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TabControl]$TabControl,
        [Parameter()][AllowEmptyCollection()][array]$ResultSets = @()
    )
    Write-DzDebug "`t[DEBUG][Show-MultipleResultSets] INICIO"
    Write-DzDebug "`t[DEBUG][Show-MultipleResultSets] ResultSets Count: $($ResultSets.Count)"
    $TabControl.Items.Clear()
    if (-not $ResultSets -or $ResultSets.Count -eq 0) {
        $tab = New-Object System.Windows.Controls.TabItem
        $ht = New-Object System.Windows.Controls.TextBlock
        $ht.Text = "Resultado"
        $ht.VerticalAlignment = "Center"
        $tab.Header = $ht
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = "La consulta no devolvió resultados."
        $text.Margin = "10"
        $tab.Content = $text
        [void]$TabControl.Items.Add($tab)
        $TabControl.SelectedIndex = 0
        if ($global:lblRowCount) {
            $global:lblRowCount.Text = "Filas: 0"
        }
        Write-DzDebug "`t[DEBUG][Show-MultipleResultSets] FIN (sin resultados)"
        return
    }
    $theme = $null
    try { $theme = Get-DzUiTheme } catch { $theme = $null }
    $isDark = $false
    try {
        if ($theme -and $theme.FormBackground) {
            $bg = [string]$theme.FormBackground
            if ($bg -match '^#') {
                $c = [System.Windows.Media.ColorConverter]::ConvertFromString($bg)
                $lum = (0.2126 * $c.R) + (0.7152 * $c.G) + (0.0722 * $c.B)
                if ($lum -lt 128) { $isDark = $true }
            } else {
                if ($bg -match '(?i)black|dark|#0|#1|#2') { $isDark = $true }
            }
        }
    } catch { $isDark = $false }
    $gridBg = $null
    $gridFg = $null
    $headerBg = $null
    $headerFg = $null
    $gridLine = $null
    $rowAlt = $null
    if ($isDark) {
        $gridBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#171717")
        $rowAlt = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1F1F1F")
        $gridFg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#E6E6E6")
        $headerBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#232323")
        $headerFg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFFFFF")
        $gridLine = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2E2E2E")
    } else {
        $gridBg = [System.Windows.Media.Brushes]::White
        $rowAlt = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F3F3F3")
        $gridFg = [System.Windows.Media.Brushes]::Black
        $headerBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#E9E9E9")
        $headerFg = [System.Windows.Media.Brushes]::Black
        $gridLine = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#D0D0D0")
    }
    $nullBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FDFBAC")
    $nullFg = [System.Windows.Media.Brushes]::Black
    $hdrStyle = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $headerBg)))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $headerFg)))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::HorizontalContentAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::VerticalContentAlignmentProperty, [System.Windows.VerticalAlignment]::Center)))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::PaddingProperty, (New-Object System.Windows.Thickness(8, 4, 8, 4)))))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderBrushProperty, $gridLine)))
    [void]$hdrStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderThicknessProperty, (New-Object System.Windows.Thickness(0, 0, 1, 1)))))
    $cellStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridCell])
    [void]$cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::PaddingProperty, (New-Object System.Windows.Thickness(8, 2, 8, 2)))))
    [void]$cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderBrushProperty, $gridLine)))
    [void]$cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderThicknessProperty, (New-Object System.Windows.Thickness(0, 0, 1, 1)))))
    [void]$cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::HorizontalContentAlignmentProperty, [System.Windows.HorizontalAlignment]::Stretch)))
    [void]$cellStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::VerticalContentAlignmentProperty, [System.Windows.VerticalAlignment]::Center)))
    $textStyleBase = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
    [void]$textStyleBase.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis)))
    [void]$textStyleBase.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::TextWrappingProperty, [System.Windows.TextWrapping]::NoWrap)))
    [void]$textStyleBase.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::VerticalAlignmentProperty, [System.Windows.VerticalAlignment]::Center)))
    $tNull = New-Object System.Windows.Trigger
    $tNull.Property = [System.Windows.Controls.TextBlock]::TextProperty
    $tNull.Value = "NULL"
    [void]$tNull.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::BackgroundProperty, $nullBrush)))
    [void]$tNull.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::ForegroundProperty, $nullFg)))
    [void]$textStyleBase.Triggers.Add($tNull)
    $i = 0
    foreach ($rs in $ResultSets) {
        $i++
        $tab = New-Object System.Windows.Controls.TabItem
        $rowCount = if ($rs.RowCount -ne $null) { $rs.RowCount } else { $rs.DataTable.Rows.Count }
        $ht = New-Object System.Windows.Controls.TextBlock
        $ht.Text = "Resultado $i ($rowCount filas)"
        $ht.VerticalAlignment = "Center"
        $tab.Header = $ht
        $dg = New-Object System.Windows.Controls.DataGrid
        $dg.AutoGenerateColumns = $true
        $dg.ItemsSource = $rs.DataTable.DefaultView
        $dg.IsReadOnly = $true
        $dg.CanUserAddRows = $false
        $dg.CanUserDeleteRows = $false
        $dg.SelectionMode = "Extended"
        $dg.HeadersVisibility = "Column"
        $dg.GridLinesVisibility = "All"
        $dg.HorizontalGridLinesBrush = $gridLine
        $dg.VerticalGridLinesBrush = $gridLine
        $dg.Background = $gridBg
        $dg.Foreground = $gridFg
        $dg.RowBackground = $gridBg
        $dg.AlternatingRowBackground = $rowAlt
        $dg.BorderBrush = $gridLine
        $dg.BorderThickness = "1"
        $dg.RowHeight = 26
        $dg.ColumnHeaderHeight = 28
        $dg.HorizontalScrollBarVisibility = "Auto"
        $dg.VerticalScrollBarVisibility = "Auto"
        $dg.CanUserResizeColumns = $true
        $dg.CanUserSortColumns = $true
        $dg.AlternationCount = 2
        $dg.ColumnHeaderStyle = $hdrStyle
        $dg.CellStyle = $cellStyle
        $dg.Add_AutoGeneratingColumn({
                param($s, $e)
                $prop = $e.PropertyName
                $hdr = $e.Column.Header
                if ($e.PropertyType -eq [datetime]) {
                    $col = New-Object System.Windows.Controls.DataGridTextColumn
                    $col.Header = $hdr
                    $b = New-Object System.Windows.Data.Binding($prop)
                    $b.StringFormat = "yyyy-MM-dd HH:mm:ss.fff"
                    $b.TargetNullValue = "NULL"
                    $col.Binding = $b
                    $ts = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
                    $ts.BasedOn = $textStyleBase
                    $col.ElementStyle = $ts
                    $e.Column = $col
                    return
                }
                if ($e.PropertyType -eq [bool]) {
                    if ($e.Column -is [System.Windows.Controls.DataGridCheckBoxColumn]) {
                        $e.Column.IsThreeState = $true
                        $cbBind = $e.Column.Binding
                        if ($cbBind -is [System.Windows.Data.Binding]) { $cbBind.TargetNullValue = $null }
                    }
                    return
                }
                if ($e.PropertyType -in @([int], [long], [decimal], [double], [single])) {
                    if ($e.Column -is [System.Windows.Controls.DataGridTextColumn]) {
                        $b = $e.Column.Binding
                        if (-not ($b -is [System.Windows.Data.Binding])) { $b = New-Object System.Windows.Data.Binding($prop) }
                        $b.TargetNullValue = "NULL"
                        $e.Column.Binding = $b
                        $ts = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
                        $ts.BasedOn = $textStyleBase
                        [void]$ts.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::TextAlignmentProperty, [System.Windows.TextAlignment]::Right)))
                        $e.Column.ElementStyle = $ts
                    }
                    return
                }
                if ($e.Column -is [System.Windows.Controls.DataGridTextColumn]) {
                    $b = $e.Column.Binding
                    if (-not ($b -is [System.Windows.Data.Binding])) { $b = New-Object System.Windows.Data.Binding($prop) }
                    $b.TargetNullValue = "NULL"
                    $e.Column.Binding = $b
                    $ts = New-Object System.Windows.Style([System.Windows.Controls.TextBlock])
                    $ts.BasedOn = $textStyleBase
                    $e.Column.ElementStyle = $ts
                }
            })
        $dg.Add_AutoGeneratedColumns({
                param($s, $e)
                try {
                    $min = 60
                    $max = 900
                    $pad = 14
                    $sampleMax = 60
                    $dv = $s.ItemsSource
                    $dt = $null
                    try { $dt = $dv.Table } catch { $dt = $null }
                    $dpi = 96.0
                    try {
                        $src = [System.Windows.PresentationSource]::FromVisual($s)
                        if ($src -and $src.CompositionTarget -and $src.CompositionTarget.TransformToDevice) { $dpi = 96.0 * $src.CompositionTarget.TransformToDevice.M11 }
                    } catch { $dpi = 96.0 }
                    $typeface = New-Object System.Windows.Media.Typeface($s.FontFamily, $s.FontStyle, $s.FontWeight, $s.FontStretch)
                    $fontSize = [double]$s.FontSize
                    foreach ($col in $s.Columns) { $col.MinWidth = $min; $col.MaxWidth = $max }
                    $s.Dispatcher.BeginInvoke([action] {
                            try {
                                $s.UpdateLayout()
                                foreach ($col in $s.Columns) {
                                    $prop = $null
                                    try { $prop = $col.SortMemberPath } catch { $prop = $null }
                                    if ([string]::IsNullOrWhiteSpace($prop)) { try { $prop = $col.Header.ToString() } catch { $prop = $null } }
                                    $best = 0.0
                                    $hText = ""
                                    try { $hText = [string]$col.Header } catch { $hText = "" }
                                    if (-not [string]::IsNullOrEmpty($hText)) {
                                        $ftH = New-Object System.Windows.Media.FormattedText($hText, [System.Globalization.CultureInfo]::CurrentCulture, [System.Windows.FlowDirection]::LeftToRight, $typeface, $fontSize, [System.Windows.Media.Brushes]::Black, $dpi)
                                        $best = [Math]::Max($best, $ftH.WidthIncludingTrailingWhitespace)
                                    }
                                    $count = 0
                                    if ($dt -and $prop -and $dt.Columns.Contains($prop)) {
                                        foreach ($row in $dt.Rows) {
                                            if ($count -ge $sampleMax) { break }
                                            $val = $row[$prop]
                                            $txt = $null
                                            if ($null -eq $val -or $val -is [System.DBNull]) { $txt = "NULL" } else { if ($val -is [datetime]) { $txt = ([datetime]$val).ToString("yyyy-MM-dd HH:mm:ss.fff") } else { $txt = [string]$val } }
                                            if (-not [string]::IsNullOrEmpty($txt)) {
                                                $ft = New-Object System.Windows.Media.FormattedText($txt, [System.Globalization.CultureInfo]::CurrentCulture, [System.Windows.FlowDirection]::LeftToRight, $typeface, $fontSize, [System.Windows.Media.Brushes]::Black, $dpi)
                                                if ($ft.WidthIncludingTrailingWhitespace -gt $best) { $best = $ft.WidthIncludingTrailingWhitespace }
                                            }
                                            $count++
                                        }
                                    } else { $best = [Math]::Max($best, 120.0) }
                                    $w = [Math]::Ceiling($best + $pad)
                                    if ($w -lt $col.MinWidth) { $w = $col.MinWidth }
                                    if ($w -gt $col.MaxWidth) { $w = $col.MaxWidth }
                                    $col.Width = $w
                                }
                            } catch { }
                        }, [System.Windows.Threading.DispatcherPriority]::Loaded) | Out-Null
                } catch { }
            })
        $tab.Content = $dg
        [void]$TabControl.Items.Add($tab)
        Write-Host "`tPestaña $i creada con $rowCount filas" -ForegroundColor Green
    }
    if ($global:lblRowCount) {
        $totalRows = ($ResultSets | Measure-Object -Property RowCount -Sum).Sum
        if ($ResultSets.Count -eq 1) {
            $global:lblRowCount.Text = "Filas: $totalRows"
        } else {
            $global:lblRowCount.Text = "Filas: $totalRows ($($ResultSets.Count) resultsets)"
        }
    }
    $TabControl.SelectedIndex = 0
    Write-DzDebug "`t[DEBUG][Show-MultipleResultSets] FIN"
}
function Export-ResultSetToCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$ResultSet,
        [Parameter(Mandatory = $true)][string]$Path
    )
    if (-not $ResultSet -or -not $ResultSet.DataTable) { throw "No hay datos para exportar." }
    $ResultSet.DataTable | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}
function Export-ResultSetToDelimitedText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$ResultSet,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter()][string]$Separator = "|"
    )
    if ([string]::IsNullOrWhiteSpace($Separator)) { $Separator = "|" }
    $dt = $null
    if ($ResultSet -is [System.Data.DataTable]) {
        $dt = $ResultSet
    } elseif ($ResultSet -and $ResultSet.DataTable) {
        $dt = $ResultSet.DataTable
    }
    if (-not $dt) { throw "No hay datos para exportar." }
    $encoding = New-Object System.Text.UTF8Encoding($true)
    $writer = New-Object System.IO.StreamWriter($Path, $false, $encoding)
    try {
        $escape = {
            param([string]$value, [string]$sep)
            if ($null -eq $value) { return "" }
            $text = [string]$value
            if ($text -match "[`r`n]" -or $text.Contains($sep) -or $text.Contains('"')) {
                $text = $text.Replace('"', '""')
                return '"' + $text + '"'
            }
            return $text
        }
        $headers = @()
        foreach ($col in $dt.Columns) {
            $headers += & $escape $col.ColumnName $Separator
        }
        $writer.WriteLine(($headers -join $Separator))
        foreach ($row in $dt.Rows) {
            $cells = @()
            foreach ($col in $dt.Columns) {
                $value = $row[$col]
                if ($value -is [System.DBNull]) { $value = $null }
                $cells += & $escape $value $Separator
            }
            $writer.WriteLine(($cells -join $Separator))
        }
    } finally {
        $writer.Dispose()
    }
}
function Get-TextPointerAtOffset {
    param(
        [Parameter(Mandatory)][System.Windows.Documents.TextPointer]$Start,
        [Parameter(Mandatory)][int]$Offset
    )
    $navigator = $Start
    $count = 0
    while ($navigator -ne $null) {
        if ($navigator.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward) -eq [System.Windows.Documents.TextPointerContext]::Text) {
            $run = $navigator.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
            $remaining = $Offset - $count
            if ($remaining -le $run.Length) { return $navigator.GetPositionAtOffset($remaining) }
            $count += $run.Length
        }
        $navigator = $navigator.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
    }
    return $Start
}
function Get-PredefinedQueries {
    return @{
        "Monitor de Servicios | Ventas a subir"           = @"
SELECT DISTINCT TOP (10)
    nofacturable,
    tablaventa.IDEMPRESA,
    codigo_unico_af AS ticket_cu,
    e.CLAVEUNICAEMPRESA AS empresa_id,
    numcheque AS ticket_folio,
    seriefolio AS ticket_serie,
    tablaventa.FOLIO AS ticket_folioSR,
    CONVERT(nvarchar(30), FECHA, 120) AS ticket_fecha,
    subtotal AS ticket_subtotal,
    total AS ticket_total,
    descuento AS ticket_descuento,
    totalconpropina AS ticket_totalconpropina,
    totalsindescuento AS ticket_totalsindescuento,
    (totalimpuestod1 + totalimpuestod2 + totalimpuestod3) AS ticket_totalimpuesto,
    totalotros AS ticket_totalotros,
    descuentoimporte AS ticket_totaldescuento,
    0 AS ticket_totaldescuento2,
    0 AS ticket_totaldescuento3,
    0 AS ticket_totaldescuento4,
    tablaventa.PROPINA AS ticket_propina,
    cancelado,
    CAST(numcheque AS VARCHAR) AS ticket_numcheque,
    numerotarjeta,
    puntosmonederogenerados,
    titulartarjetamonederodescuento AS titulartarjetamonedero,
    tarjetadescuento,
    descuentomonedero,
    e.idregimen_sat,
    tipopago.idformapago_SAT,
    tablaventa.idturno,
    nopersonas,
    tipodeservicio,
    idmesero,
    totalarticulos,
    LTRIM(RTRIM(estacion)) AS ticket_estacion,
    usuariodescuento,
    comentariodescuento,
    tablaventa.idtipodescuento,
    totalimpuestod2 AS TicketTotalIEPS,
    0 AS TicketTotalOtrosImpuestos,
    LEFT(CONVERT(VARCHAR, fecha + 15, 120), 10) + ' 23:59:59' AS ticket_fechavence,
    CONVERT(nvarchar(30), tablaventa.cierre, 120) AS ticket_fecha_cierre
FROM
    CHEQUES AS tablaventa
    INNER JOIN empresas AS e ON tablaventa.IDEMPRESA = e.IDEMPRESA
    LEFT JOIN chequespagos AS tablapago ON tablapago.folio = tablaventa.folio
    LEFT JOIN formasdepago AS tipopago ON tablapago.idformadepago = tipopago.idformadepago
WHERE
    fecha > (SELECT fecha_inicio_envio FROM configuracion_ws)
    AND (intentoEnvioAF < 20)
    AND (
        (enviado = 0) OR (enviado IS NULL)
    )
    AND (
        (pagado = 1 AND nofacturable = 0)
        OR cancelado = 1
    )
    AND codigo_unico_af IS NOT NULL
    AND codigo_unico_af <> ''
    AND tablaventa.IDEMPRESA = (SELECT TOP 1 idempresa FROM empresas)
ORDER BY
    numcheque;
"@
        "BackOffice Actualizar contraseña  administrador" = @"
UPDATE users
    SET Password = '08/Vqq0='
    OUTPUT inserted.UserName
    WHERE UserName = (SELECT TOP 1 UserName FROM users WHERE IsSuperAdmin = 1 and IsEnabled = 1);
"@
        "BackOffice Estaciones"                           = @"
SELECT
    t.Name,
    t.Ip,
    t.LastOnline,
    t.IsEnabled,
    u.UserName AS UltimoUsuario,
    t.AppVersion,
    t.IsMaximized,
    t.ForceAppUpdate,
    t.SkipDoPing
FROM Terminals t
LEFT JOIN Users u ON t.LastUserLogin = u.Id
ORDER BY t.IsEnabled DESC, t.Name;
"@
        "SR | Actualizar contraseña de administrador"     = @"
UPDATE usuarios
    SET contraseña = 'A9AE4E13D2A47998AC34'
    OUTPUT inserted.usuario
    WHERE usuario = (SELECT TOP 1 usuario FROM usuarios WHERE administrador = 1);
"@
        "SR | Revisar Pivot Table"                        = @"
SELECT app_id, field, COUNT(*)
    FROM app_settings
    GROUP BY app_id, field
    HAVING COUNT(*) > 1
"@
        "SR | Fecha Revisiones"                           = @"
WITH CTE AS (
        SELECT
            b.estacion,
            b.fecha       AS UltimoUso,
            ROW_NUMBER() OVER (PARTITION BY b.estacion ORDER BY b.fecha DESC) AS rn
        FROM bitacorasistema b
    )
    SELECT
        e.FECHAREV,
        c.estacion,
        c.UltimoUso
    FROM CTE c
    JOIN estaciones e
        ON c.estacion = e.idestacion
    WHERE c.rn = 1
    ORDER BY c.UltimoUso DESC;
"@
        "SR SYNC | nsplatformcontrol"                     = @"
BEGIN TRY
    BEGIN TRANSACTION;
    SELECT WorkspaceId, EntityType, OperationType, 0 AS IsSync, 0 AS Attempts, CreateDate
    INTO #tempcontroltaxes
    FROM nsplatformcontrol
    WHERE EntityType = 1;
    TRUNCATE TABLE nsplatformcontrol;
    INSERT INTO nsplatformcontrol (WorkspaceId, EntityType, OperationType, IsSync, Attempts, CreateDate)
    SELECT WorkspaceId, EntityType, OperationType, IsSync, Attempts, CreateDate
    FROM #tempcontroltaxes;
    DROP TABLE #tempcontroltaxes;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
"@
        "SR | Memoria Insuficiente"                       = @"
UPDATE empresas
        SET nombre='', razonsocial='',direccion='',sucursal='',
        rfc='',curp='', telefono='',giro='',contacto='',fax='',
        email='',idhardware='',web='',ciudad='',estado='',pais='',
        ciudadsucursal='',estadosucursal='',codigopostal='',a86ed5f9d02ec5b3='',
        codigopostalsucursal='',uid=newid();
GO
DELETE FROM registro_licencias;
"@
        "OTM | Eliminar Server en OTM"                    = @"
SELECT serie, ipserver, nombreservidor
    FROM configuracion;
"@
        "NSH | Eliminar Server en Hoteles"                = @"
SELECT serievalida, numserie, ipserver, nombreservidor, llave
    FROM configuracion;
"@
        "Restcard | Eliminar Server en Rest Card"         = @"
SELECT estacion, ipservidor FROM tabvariables;
"@
        "sql | Listar usuarios e idiomas"                 = @"
SELECT
    p.name AS Usuario,
    l.default_language_name AS Idioma
FROM
    sys.server_principals p
LEFT JOIN
    sys.sql_logins l ON p.principal_id = l.principal_id
WHERE
    p.type IN ('S', 'U')
"@
    }
}
function Get-DzThemeBrush {
    param(
        [Parameter(Mandatory = $true)][string]$Hex,
        [Parameter(Mandatory = $true)][System.Windows.Media.Brush]$Fallback
    )
    if ([string]::IsNullOrWhiteSpace($Hex)) { return $Fallback }
    try {
        $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
        if ($brush -is [System.Windows.Freezable] -and $brush.CanFreeze) { $brush.Freeze() }
        return $brush
    } catch {
        return $Fallback
    }
}
function Initialize-PredefinedQueries {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.ComboBox]$ComboQueries,
        [Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl,
        [Parameter(Mandatory = $true)][hashtable]$Queries
    )
    $ComboQueries.Items.Clear()
    [void]$ComboQueries.Items.Add("Selecciona una consulta predefinida")
    foreach ($key in ($Queries.Keys | Sort-Object)) { [void]$ComboQueries.Items.Add($key) }
    $ComboQueries.SelectedIndex = 0
    $ComboQueries.Tag = [pscustomobject]@{
        Queries     = $Queries
        TabControl  = $TabControl
        SqlKeywords = $global:DzSqlKeywords
    }
    $ComboQueries.Add_SelectionChanged({
            param($sender, $e)
            try {
                $selectedQuery = $sender.SelectedItem
                if (-not $selectedQuery -or $selectedQuery -eq "Selecciona una consulta predefinida") { return }
                $ctx = $sender.Tag
                if (-not $ctx -or -not $ctx.Queries.ContainsKey($selectedQuery)) { return }
                $rtb = Get-ActiveQueryRichTextBox -TabControl $ctx.TabControl
                if (-not $rtb) { return }
                $queryText = $ctx.Queries[$selectedQuery]
                $rtb.Document.Blocks.Clear()
                $p = New-Object System.Windows.Documents.Paragraph
                [void]$p.Inlines.Add((New-Object System.Windows.Documents.Run($queryText)))
                [void]$rtb.Document.Blocks.Add($p)
                if (-not [string]::IsNullOrWhiteSpace($ctx.SqlKeywords)) {
                    Set-WpfSqlHighlighting -RichTextBox $rtb -Keywords $ctx.SqlKeywords
                }
                $rtb.Focus()
            } catch {
                Write-DzDebug "`t[DEBUG] Error en SelectionChanged (queries): $($_.Exception.Message)" -Color Red
            }
        })
}
function Set-WpfSqlHighlighting {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.RichTextBox]$RichTextBox,
        [Parameter(Mandatory)][string]$Keywords
    )
    if ($null -eq $RichTextBox -or $null -eq $RichTextBox.Document) { return }
    if ([string]::IsNullOrWhiteSpace($Keywords)) { return }
    $theme = Get-DzUiTheme
    $defaultBrush = Get-DzThemeBrush -Hex $theme.ControlForeground -Fallback ([System.Windows.Media.Brushes]::Black)
    $commentBrush = Get-DzThemeBrush -Hex $theme.AccentMuted -Fallback ([System.Windows.Media.Brushes]::DarkGreen)
    $keywordBrush = Get-DzThemeBrush -Hex $theme.AccentPrimary -Fallback ([System.Windows.Media.Brushes]::Blue)
    $range = New-Object System.Windows.Documents.TextRange($RichTextBox.Document.ContentStart, $RichTextBox.Document.ContentEnd)
    $text = $range.Text
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    $range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $defaultBrush)
    $commentRanges = @()
    foreach ($c in [regex]::Matches($text, '--.*', [System.Text.RegularExpressions.RegexOptions]::Multiline)) {
        $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $c.Index
        $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($c.Index + $c.Length)
        if ($start -and $end) {
            (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $commentBrush)
            $commentRanges += [pscustomobject]@{ Start = $c.Index; End = $c.Index + $c.Length }
        }
    }
    foreach ($b in [regex]::Matches($text, '/\*[\s\S]*?\*/', [System.Text.RegularExpressions.RegexOptions]::Multiline)) {
        $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $b.Index
        $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($b.Index + $b.Length)
        if ($start -and $end) {
            (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $commentBrush)
            $commentRanges += [pscustomobject]@{ Start = $b.Index; End = $b.Index + $b.Length }
        }
    }
    $pattern = '\b(' + $Keywords + ')\b'
    $matches = [regex]::Matches($text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($m in $matches) {
        $inComment = $commentRanges | Where-Object { $m.Index -ge $_.Start -and $m.Index -lt $_.End }
        if ($inComment) { continue }
        $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $m.Index
        $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($m.Index + $m.Length)
        if ($start -and $end) {
            (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $keywordBrush)
        }
    }
}
function Get-TextPointerFromOffset {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.RichTextBox]$RichTextBox,
        [Parameter(Mandatory)][int]$Offset
    )
    if ($null -eq $RichTextBox.Document) { return $null }
    $pointer = $RichTextBox.Document.ContentStart
    $count = 0
    while ($pointer -ne $null) {
        $ctx = $pointer.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward)
        if ($ctx -eq [System.Windows.Documents.TextPointerContext]::Text) {
            $runText = $pointer.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
            $remaining = $Offset - $count
            if ($remaining -le $runText.Length) { return $pointer.GetPositionAtOffset($remaining) }
            $count += $runText.Length
            $pointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
            continue
        }
        if ($ctx -eq [System.Windows.Documents.TextPointerContext]::ElementStart) {
            $el = $pointer.GetAdjacentElement([System.Windows.Documents.LogicalDirection]::Forward)
            if ($el -is [System.Windows.Documents.LineBreak]) {
                if ($count + 2 -ge $Offset) { return $pointer }
                $count += 2
            }
            $pointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
            continue
        }
        if ($ctx -eq [System.Windows.Documents.TextPointerContext]::ElementEnd) {
            $parent = $pointer.Parent
            if ($parent -is [System.Windows.Documents.Paragraph]) {
                if ($count + 2 -ge $Offset) { return $pointer }
                $count += 2
            }
            $pointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
            continue
        }
        $pointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
    }
    return $RichTextBox.Document.ContentEnd
}
function Get-ResultTabHeaderText {
    param([System.Windows.Controls.TabItem]$TabItem)
    if (-not $TabItem) { return "Resultado" }
    $header = $TabItem.Header
    if ($header -is [System.Windows.Controls.TextBlock]) { return [string]$header.Text }
    if ($null -ne $header) { return [string]$header }
    return "Resultado"
}
function Get-ExportableResultTabs {
    param([System.Windows.Controls.TabControl]$TabControl)
    $exportable = @()
    if (-not $TabControl) { return $exportable }
    foreach ($item in $TabControl.Items) {
        if ($item -isnot [System.Windows.Controls.TabItem]) { continue }
        $dg = $item.Content
        if ($dg -isnot [System.Windows.Controls.DataGrid]) { continue }
        $dt = $null
        if ($dg.ItemsSource -is [System.Data.DataView]) {
            $dt = $dg.ItemsSource.Table
        } elseif ($dg.ItemsSource -is [System.Data.DataTable]) {
            $dt = $dg.ItemsSource
        } else {
            try { $dt = $dg.ItemsSource.Table } catch { $dt = $null }
        }
        if (-not $dt -or -not $dt.Rows -or $dt.Rows.Count -lt 1) { continue }
        $headerText = Get-ResultTabHeaderText -TabItem $item
        $exportable += [pscustomobject]@{
            Tab          = $item
            DataTable    = $dt
            RowCount     = $dt.Rows.Count
            Display      = "$headerText ($($dt.Rows.Count) filas)"
            DisplayShort = $headerText
        }
    }
    return $exportable
}
function Write-DataTableConsole {
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [int]$MaxRows = 50
    )
    if (-not $DataTable) { return }
    $rows = @($DataTable.Rows)
    $cols = @($DataTable.Columns | ForEach-Object { $_.ColumnName })
    Write-Host ""
    Write-Host ("Columnas: {0} | Filas: {1}" -f $cols.Count, $rows.Count) -ForegroundColor DarkGray
    $sample = $rows | Select-Object -First $MaxRows
    $width = @{}
    foreach ($c in $cols) { $width[$c] = [Math]::Max($c.Length, 4) }
    foreach ($r in $sample) {
        foreach ($c in $cols) {
            $v = $r[$c]
            if ($v -is [DBNull]) { $v = $null }
            $s = if ($null -eq $v) { "NULL" } else { [string]$v }
            if ($s.Length -gt 80) { $s = $s.Substring(0, 77) + "..." }
            $width[$c] = [Math]::Min(80, [Math]::Max($width[$c], $s.Length))
        }
    }
    $header = ($cols | ForEach-Object { $_.PadRight($width[$_] + 2) }) -join ""
    Write-Host $header -ForegroundColor Cyan
    Write-Host ("-" * $header.Length) -ForegroundColor DarkGray
    foreach ($r in $sample) {
        $line = ($cols | ForEach-Object {
                $v = $r[$_]
                if ($v -is [DBNull]) { $v = $null }
                $s = if ($null -eq $v) { "NULL" } else { [string]$v }
                if ($s.Length -gt 80) { $s = $s.Substring(0, 77) + "..." }
                $s.PadRight($width[$_] + 2)
            }) -join ""
        Write-Host $line
    }
    if ($rows.Count -gt $MaxRows) {
        Write-Host ("... mostrando {0} de {1} filas (limite MaxRows={0})" -f $MaxRows, $rows.Count) -ForegroundColor Yellow
    }
}
function Show-ErrorResultTab {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TabControl]$ResultsTabControl,
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][switch]$AddWithoutClear
    )
    if (-not $AddWithoutClear) {
        try { $ResultsTabControl.Items.Clear() } catch {}
    }
    $tab = New-Object System.Windows.Controls.TabItem
    $ht = New-Object System.Windows.Controls.TextBlock
    $ht.Text = "Error"
    $ht.VerticalAlignment = "Center"
    $tab.Header = $ht
    $text = New-Object System.Windows.Controls.TextBox
    $text.Text = $Message
    $text.Margin = "10"
    $text.IsReadOnly = $true
    $text.TextWrapping = "Wrap"
    $text.VerticalScrollBarVisibility = "Auto"
    $text.HorizontalScrollBarVisibility = "Auto"
    $tab.Content = $text
    [void]$ResultsTabControl.Items.Add($tab)
    if (-not $AddWithoutClear) {
        $ResultsTabControl.SelectedIndex = 0
    }
}
function Disconnect-DbCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Ctx
    )

    if ($Ctx.QueryRunning) {
        Write-DzDebug "`t[DEBUG][Disconnect] Cancelando query en ejecución..."
        try {
            if ($Ctx.CurrentQueryPowerShell) {
                $Ctx.CurrentQueryPowerShell.Stop()
                $Ctx.CurrentQueryPowerShell.Dispose()
            }
            if ($Ctx.CurrentQueryRunspace) {
                $Ctx.CurrentQueryRunspace.Close()
                $Ctx.CurrentQueryRunspace.Dispose()
            }
        } catch {
            Write-DzDebug "`t[DEBUG][Disconnect] Error cancelando query: $_"
        }
        $Ctx.CurrentQueryPowerShell = $null
        $Ctx.CurrentQueryRunspace = $null
        $Ctx.CurrentQueryAsync = $null
        $Ctx.QueryRunning = $false
        Write-DzDebug "`t[DEBUG][Disconnect] Query cancelada"
    }

    if ($Ctx.execUiTimer -and $Ctx.execUiTimer.IsEnabled) {
        Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo execUiTimer..."
        $Ctx.execUiTimer.Stop()
    }

    if ($Ctx.QueryDoneTimer -and $Ctx.QueryDoneTimer.IsEnabled) {
        Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo QueryDoneTimer..."
        $Ctx.QueryDoneTimer.Stop()
    }

    if ($Ctx.execStopwatch) {
        Write-DzDebug "`t[DEBUG][Disconnect] Deteniendo stopwatch..."
        $Ctx.execStopwatch.Stop()
        $Ctx.execStopwatch = $null
    }

    if ($Ctx.Connection) {
        Write-DzDebug "`t[DEBUG][Disconnect] Cerrando conexión SQL..."
        try {
            if ($Ctx.Connection.State -ne [System.Data.ConnectionState]::Closed) {
                $Ctx.Connection.Close()
            }
            $Ctx.Connection.Dispose()
        } catch {
            Write-DzDebug "`t[DEBUG][Disconnect] Error cerrando conexión: $_"
        }
        $Ctx.Connection = $null
    }

    Write-DzDebug "`t[DEBUG][Disconnect] Limpiando datos de conexión..."
    $Ctx.Server = $null
    $Ctx.User = $null
    $Ctx.Password = $null
    $Ctx.Database = $null
    $Ctx.DbCredential = $null

    if ($Ctx.tvDatabases) {
        Write-DzDebug "`t[DEBUG][Disconnect] Limpiando TreeView..."
        $Ctx.tvDatabases.Items.Clear()
    }

    if ($Ctx.cmbDatabases) {
        Write-DzDebug "`t[DEBUG][Disconnect] Limpiando ComboBox..."
        $Ctx.cmbDatabases.Items.Clear()
        $Ctx.cmbDatabases.IsEnabled = $false
    }

    if ($Ctx.lblConnectionStatus) {
        $Ctx.lblConnectionStatus.Text = "⚫ Desconectado"
    }

    Write-DzDebug "`t[DEBUG][Disconnect] Habilitando controles de conexión..."
    if ($Ctx.txtServer) { $Ctx.txtServer.IsEnabled = $true }
    if ($Ctx.txtUser) { $Ctx.txtUser.IsEnabled = $true }
    if ($Ctx.txtPassword) { $Ctx.txtPassword.IsEnabled = $true }
    if ($Ctx.btnConnectDb) { $Ctx.btnConnectDb.IsEnabled = $true }

    Write-DzDebug "`t[DEBUG][Disconnect] Deshabilitando controles de operaciones..."
    if ($Ctx.btnDisconnectDb) { $Ctx.btnDisconnectDb.IsEnabled = $false }
    if ($Ctx.btnExecute) { $Ctx.btnExecute.IsEnabled = $false }
    if ($Ctx.btnClearQuery) { $Ctx.btnClearQuery.IsEnabled = $false }
    if ($Ctx.btnExport) { $Ctx.btnExport.IsEnabled = $false }
    if ($Ctx.cmbQueries) { $Ctx.cmbQueries.IsEnabled = $false }
    if ($Ctx.tcQueries) { $Ctx.tcQueries.IsEnabled = $false }
    if ($Ctx.tcResults) { $Ctx.tcResults.IsEnabled = $false }
    if ($Ctx.rtbQueryEditor1) { $Ctx.rtbQueryEditor1.IsEnabled = $false }
    if ($Ctx.dgResults) { $Ctx.dgResults.IsEnabled = $false }
    if ($Ctx.txtMessages) { $Ctx.txtMessages.IsEnabled = $false }

    if ($Ctx.txtMessages) {
        $Ctx.txtMessages.Text = "Desconectado de la base de datos."
    }

    if ($Ctx.txtServer) {
        try { $Ctx.txtServer.Focus() | Out-Null } catch {}
    }

    Write-DzDebug "`t[DEBUG][Disconnect] Desconexión completada exitosamente"
}

function Get-DbUiContext {
    [CmdletBinding()]
    param()

    @{
        QueryRunning           = $script:QueryRunning
        CurrentQueryPowerShell = $script:CurrentQueryPowerShell
        CurrentQueryRunspace   = $script:CurrentQueryRunspace
        CurrentQueryAsync      = $script:CurrentQueryAsync
        execUiTimer            = $script:execUiTimer
        QueryDoneTimer         = $script:QueryDoneTimer
        execStopwatch          = $script:execStopwatch

        Connection             = $global:connection
        Server                 = $global:server
        User                   = $global:user
        Password               = $global:password
        Database               = $global:database
        DbCredential           = $global:dbCredential

        tvDatabases            = $global:tvDatabases
        cmbDatabases           = $global:cmbDatabases
        lblConnectionStatus    = $global:lblConnectionStatus

        txtServer              = $global:txtServer
        txtUser                = $global:txtUser
        txtPassword            = $global:txtPassword
        btnConnectDb           = $global:btnConnectDb

        btnDisconnectDb        = $global:btnDisconnectDb
        btnExecute             = $global:btnExecute
        btnClearQuery          = $global:btnClearQuery
        btnExport              = $global:btnExport
        cmbQueries             = $global:cmbQueries
        tcQueries              = $global:tcQueries
        tcResults              = $global:tcResults
        rtbQueryEditor1        = $global:rtbQueryEditor1
        dgResults              = $global:dgResults
        txtMessages            = $global:txtMessages

        MainWindow             = $global:MainWindow
    }
}

function Save-DbUiContext {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    $script:QueryRunning = [bool]$Ctx.QueryRunning
    $script:CurrentQueryPowerShell = $Ctx.CurrentQueryPowerShell
    $script:CurrentQueryRunspace = $Ctx.CurrentQueryRunspace
    $script:CurrentQueryAsync = $Ctx.CurrentQueryAsync
    $script:execUiTimer = $Ctx.execUiTimer
    $script:QueryDoneTimer = $Ctx.QueryDoneTimer
    $script:execStopwatch = $Ctx.execStopwatch

    $global:connection = $Ctx.Connection
    $global:server = $Ctx.Server
    $global:user = $Ctx.User
    $global:password = $Ctx.Password
    $global:database = $Ctx.Database
    $global:dbCredential = $Ctx.DbCredential
    $global:database = $Ctx.Database
    $global:dbCredential = $Ctx.DbCredential
    $global:server = $Ctx.Server
    $global:user = $Ctx.User
    $global:password = $Ctx.Password

}

function Disconnect-DbUiSafe {
    [CmdletBinding()]
    param()

    try {
        $ctx = Get-DbUiContext
        Disconnect-DbCore -Ctx $ctx
        Save-DbUiContext -Ctx $ctx
        Write-Host "✓ Desconectado exitosamente" -ForegroundColor Green
    } catch {
        Write-DzDebug "`t[DEBUG][Disconnect] ERROR: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Disconnect] Stack: $($_.ScriptStackTrace)"
        Write-Host "Error al desconectar: $($_.Exception.Message)" -ForegroundColor Red
        Ui-Error "Error al desconectar:`n`n$($_.Exception.Message)" $global:MainWindow
    }
}
function Connect-DbCore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    if ($null -eq $Ctx.txtServer -or $null -eq $Ctx.txtUser -or $null -eq $Ctx.txtPassword) {
        throw "Error interno: controles de conexión no inicializados."
    }

    $serverText = ([string]$Ctx.txtServer.Text).Trim()
    $userText = ([string]$Ctx.txtUser.Text).Trim()
    $passwordText = [string]$Ctx.txtPassword.Password

    Write-DzDebug "`t[DEBUG] | Server='$serverText' User='$userText' PasswordLen=$($passwordText.Length)"

    if ([string]::IsNullOrWhiteSpace($serverText) -or [string]::IsNullOrWhiteSpace($userText) -or [string]::IsNullOrWhiteSpace($passwordText)) {
        throw "Complete todos los campos de conexión"
    }

    $securePassword = (New-Object System.Net.NetworkCredential('', $passwordText)).SecurePassword
    $credential = New-Object System.Management.Automation.PSCredential($userText, $securePassword)

    $Ctx.Server = $serverText
    $Ctx.User = $userText
    $Ctx.Password = $passwordText
    $Ctx.DbCredential = $credential

    Write-DzDebug "`t[DEBUG] Obteniendo lista de bases de datos..."
    $databases = Get-SqlDatabases -Server $serverText -Credential $credential
    if (-not $databases -or $databases.Count -eq 0) {
        throw "Conexión correcta, pero no se encontraron bases de datos disponibles."
    }

    Write-DzDebug "`t[DEBUG] Se encontraron $($databases.Count) bases de datos"

    if ($Ctx.cmbDatabases) {
        $Ctx.cmbDatabases.Items.Clear()
        $Ctx.cmbDatabases.DisplayMemberPath = "Name"
        $Ctx.cmbDatabases.SelectedValuePath = "Name"

        foreach ($db in $databases) {
            $name = $null
            if ($db -is [string]) { $name = $db }
            elseif ($db.PSObject.Properties["Name"]) { $name = [string]$db.Name }
            elseif ($db.PSObject.Properties["Database"]) { $name = [string]$db.Database }
            else { $name = [string]$db }  # último recurso

            if (-not [string]::IsNullOrWhiteSpace($name)) {
                [void]$Ctx.cmbDatabases.Items.Add([pscustomobject]@{ Name = $name })
            }
        }

        $Ctx.cmbDatabases.IsEnabled = $true
        $Ctx.cmbDatabases.SelectedIndex = 0
        $Ctx.Database = [string]$Ctx.cmbDatabases.SelectedValue

    }

    if ($Ctx.lblConnectionStatus) {
        $Ctx.lblConnectionStatus.Text = "✓ Conectado a: $serverText | DB: $($Ctx.Database)"
    }

    if ($Ctx.txtServer) { $Ctx.txtServer.IsEnabled = $false }
    if ($Ctx.txtUser) { $Ctx.txtUser.IsEnabled = $false }
    if ($Ctx.txtPassword) { $Ctx.txtPassword.IsEnabled = $false }
    if ($Ctx.btnConnectDb) { $Ctx.btnConnectDb.IsEnabled = $false }

    if ($Ctx.btnDisconnectDb) { $Ctx.btnDisconnectDb.IsEnabled = $true }
    if ($Ctx.btnExecute) { $Ctx.btnExecute.IsEnabled = $true }
    if ($Ctx.btnClearQuery) { $Ctx.btnClearQuery.IsEnabled = $true }
    if ($Ctx.cmbQueries) { $Ctx.cmbQueries.IsEnabled = $true }
    if ($Ctx.btnExport) { $Ctx.btnExport.IsEnabled = $true }
    if ($Ctx.tcQueries) { $Ctx.tcQueries.IsEnabled = $true }
    if ($Ctx.tcResults) { $Ctx.tcResults.IsEnabled = $true }
    if ($Ctx.rtbQueryEditor1) { $Ctx.rtbQueryEditor1.IsEnabled = $true }
    if ($Ctx.dgResults) { $Ctx.dgResults.IsEnabled = $true }
    if ($Ctx.txtMessages) { $Ctx.txtMessages.IsEnabled = $true }

    if ($Ctx.rtbQueryEditor1) { try { $Ctx.rtbQueryEditor1.Focus() | Out-Null } catch {} }
    if ($Ctx.tvDatabases) {
        Write-DzDebug "`t[DEBUG] Inicializando TreeView..."

        Initialize-SqlTreeView `
            -TreeView $Ctx.tvDatabases `
            -Server $serverText `
            -Credential $credential `
            -User $userText `
            -Password $passwordText `
            -GetCurrentDatabase ({ $Ctx.Database }.GetNewClosure()) `
            -AutoExpand $true `
            -OnDatabaseSelected ({
                param($dbName)

                if (-not $Ctx.cmbDatabases) { return }

                $dbName = [string]$dbName

                try {
                    if ($Ctx.cmbDatabases.SelectedValuePath -and $Ctx.cmbDatabases.SelectedValuePath.Trim().Length -gt 0) {
                        $Ctx.cmbDatabases.SelectedValue = $dbName
                    } else {
                        $Ctx.cmbDatabases.SelectedItem = $dbName
                    }

                    if (-not $Ctx.cmbDatabases.SelectedItem -and -not $Ctx.cmbDatabases.SelectedValue) {
                        for ($i = 0; $i -lt $Ctx.cmbDatabases.Items.Count; $i++) {
                            $item = $Ctx.cmbDatabases.Items[$i]
                            $n = $null
                            if ($item -is [string]) { $n = $item }
                            elseif ($item.PSObject.Properties["Name"]) { $n = [string]$item.Name }
                            elseif ($item.PSObject.Properties["Database"]) { $n = [string]$item.Database }
                            else { $n = [string]$item }

                            if ($n -eq $dbName) {
                                $Ctx.cmbDatabases.SelectedIndex = $i
                                break
                            }
                        }
                    }

                    if ($Ctx.cmbDatabases.SelectedValuePath -and $Ctx.cmbDatabases.SelectedValuePath.Trim().Length -gt 0) {
                        $Ctx.Database = [string]$Ctx.cmbDatabases.SelectedValue
                    } else {
                        $sel = $Ctx.cmbDatabases.SelectedItem
                        if ($sel -is [string]) { $Ctx.Database = $sel }
                        elseif ($sel -and $sel.PSObject.Properties["Name"]) { $Ctx.Database = [string]$sel.Name }
                        elseif ($sel -and $sel.PSObject.Properties["Database"]) { $Ctx.Database = [string]$sel.Database }
                        else { $Ctx.Database = [string]$sel }
                    }

                    if ($Ctx.lblConnectionStatus) {
                        $Ctx.lblConnectionStatus.Text = "✓ Conectado a: $($Ctx.Server) | DB: $($Ctx.Database)"
                    }

                    Write-DzDebug "`t[DEBUG][TreeView] BD seleccionada: $($Ctx.Database) (dbName='$dbName')"
                } catch {
                    Write-DzDebug "`t[DEBUG][OnDatabaseSelected] ERROR: $($_.Exception.Message)" -Color Red
                }
            }).GetNewClosure() `
            -OnDatabasesRefreshed ({
                try {
                    if (-not $Ctx.Server -or -not $Ctx.DbCredential) { return }

                    $dbs = Get-SqlDatabases -Server $Ctx.Server -Credential $Ctx.DbCredential
                    if (-not $dbs -or $dbs.Count -lt 1 -or -not $Ctx.cmbDatabases) { return }

                    $Ctx.cmbDatabases.Items.Clear()

                    if ($Ctx.cmbDatabases.DisplayMemberPath -and $Ctx.cmbDatabases.DisplayMemberPath.Trim().Length -gt 0) {
                        foreach ($db in $dbs) {
                            $name = $null
                            if ($db -is [string]) { $name = $db }
                            elseif ($db.PSObject.Properties["Name"]) { $name = [string]$db.Name }
                            elseif ($db.PSObject.Properties["Database"]) { $name = [string]$db.Database }
                            else { $name = [string]$db }

                            if (-not [string]::IsNullOrWhiteSpace($name)) {
                                [void]$Ctx.cmbDatabases.Items.Add([pscustomobject]@{ Name = $name })
                            }
                        }
                    } else {
                        foreach ($db in $dbs) { [void]$Ctx.cmbDatabases.Items.Add($db) }
                    }

                    if ($Ctx.Database) {
                        if ($Ctx.cmbDatabases.SelectedValuePath -and $Ctx.cmbDatabases.SelectedValuePath.Trim().Length -gt 0) {
                            $Ctx.cmbDatabases.SelectedValue = [string]$Ctx.Database
                        } else {
                            $Ctx.cmbDatabases.SelectedItem = $Ctx.Database
                        }
                    }

                    if (-not $Ctx.cmbDatabases.SelectedItem -and $Ctx.cmbDatabases.Items.Count -gt 0) {
                        $Ctx.cmbDatabases.SelectedIndex = 0
                    }

                    if ($Ctx.cmbDatabases.SelectedValuePath -and $Ctx.cmbDatabases.SelectedValuePath.Trim().Length -gt 0) {
                        $Ctx.Database = [string]$Ctx.cmbDatabases.SelectedValue
                    } else {
                        $sel = $Ctx.cmbDatabases.SelectedItem
                        if ($sel -is [string]) { $Ctx.Database = $sel }
                        elseif ($sel -and $sel.PSObject.Properties["Name"]) { $Ctx.Database = [string]$sel.Name }
                        elseif ($sel -and $sel.PSObject.Properties["Database"]) { $Ctx.Database = [string]$sel.Database }
                        else { $Ctx.Database = [string]$sel }
                    }

                    if ($Ctx.lblConnectionStatus) {
                        $Ctx.lblConnectionStatus.Text = "✓ Conectado a: $($Ctx.Server) | DB: $($Ctx.Database)"
                    }

                    Write-DzDebug "`t[DEBUG][TreeView] ComboBox actualizado con $($dbs.Count) bases de datos"
                } catch {
                    Write-DzDebug "`t[DEBUG][OnDatabasesRefreshed] Error: $_"
                }
            }).GetNewClosure() `
            -InsertTextHandler ({
                param($text)
                if ($Ctx.tcQueries) {
                    Insert-TextIntoActiveQuery -TabControl $Ctx.tcQueries -Text $text
                }
            }).GetNewClosure()
    }

    Write-DzDebug "`t[DEBUG] Conexión establecida exitosamente"
}

function Connect-DbUiSafe {
    [CmdletBinding()]
    param()

    try {
        $ctx = Get-DbUiContext
        Connect-DbCore -Ctx $ctx
        Save-DbUiContext -Ctx $ctx
        Write-Host "✓ Conectado exitosamente a: $($ctx.Server)" -ForegroundColor Green
    } catch {
        Write-DzDebug "`t[DEBUG][Connect] CATCH: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Connect] Tipo: $($_.Exception.GetType().FullName)"
        Write-DzDebug "`t[DEBUG][Connect] Stack: $($_.ScriptStackTrace)"
        Ui-Error "Error de conexión: $($_.Exception.Message)" $global:MainWindow
        Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Export-ResultsCore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    if (-not $Ctx.tcResults) {
        Ui-Warn "No existe un panel de resultados para exportar." "Atención" $Ctx.MainWindow
        return
    }

    $resultTabs = Get-ExportableResultTabs -TabControl $Ctx.tcResults
    if (-not $resultTabs -or $resultTabs.Count -eq 0) {
        Ui-Warn "No existe pestaña con resultados para exportar." "Atención" $Ctx.MainWindow
        return
    }

    $target = $null
    if ($resultTabs.Count -gt 1) {
        $items = $resultTabs | ForEach-Object {
            [pscustomobject]@{
                Path         = $_
                Display      = $_.Display
                DisplayShort = $_.DisplayShort
            }
        }
        $selected = Show-WpfPathSelectionDialog -Title "Exportar resultados" -Prompt "Seleccione la pestaña de resultados a exportar:" -Items $items -ExecuteButtonText "Exportar"
        if (-not $selected) { return }
        $target = $selected.Path
    } else {
        $target = $resultTabs[0]
    }

    $rowCount = [int]$target.RowCount
    $headerText = [string]$target.DisplayShort

    if (-not (Ui-Confirm "Se exportarán $rowCount filas de '$headerText'. ¿Deseas continuar?" "Confirmar exportación" $Ctx.MainWindow)) { return }

    $safeName = ($headerText -replace '[\\/:*?"<>|]', '-')
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "resultado" }

    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter = "CSV (*.csv)|*.csv|Texto delimitado (*.txt)|*.txt"
    $saveDialog.FileName = "$safeName.csv"
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($saveDialog.ShowDialog() -ne $true) { return }

    $filePath = $saveDialog.FileName
    $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()

    if ($extension -eq ".txt" -or $saveDialog.FilterIndex -eq 2) {
        $separator = New-WpfInputDialog -Title "Separador de exportación" -Prompt "Ingrese el separador para el archivo de texto:" -DefaultValue "|"
        if ($null -eq $separator) { return }
        Export-ResultSetToDelimitedText -ResultSet $target.DataTable -Path $filePath -Separator $separator
    } else {
        Export-ResultSetToCsv -ResultSet ([pscustomobject]@{ DataTable = $target.DataTable }) -Path $filePath
    }

    Ui-Info "Exportación completada en:`n$filePath" "Exportación" $Ctx.MainWindow
}

function Export-ResultsUiSafe {
    [CmdletBinding()]
    param()

    try {
        $ctx = Get-DbUiContext
        Export-ResultsCore -Ctx $ctx
    } catch {
        Ui-Error "Error al exportar resultados:`n$($_.Exception.Message)" "Error" $global:MainWindow
        Write-DzDebug "`t[DEBUG][Export] CATCH: $($_.Exception.Message)" -Color Red
    }
}
# Database.psm1

function Get-DbUiContext {
    [CmdletBinding()]
    param()

    @{
        QueryRunning           = $script:QueryRunning
        CurrentQueryPowerShell = $script:CurrentQueryPowerShell
        CurrentQueryRunspace   = $script:CurrentQueryRunspace
        CurrentQueryAsync      = $script:CurrentQueryAsync
        execUiTimer            = $script:execUiTimer
        QueryDoneTimer         = $script:QueryDoneTimer
        execStopwatch          = $script:execStopwatch

        Connection             = $global:connection
        Server                 = $global:server
        User                   = $global:user
        Password               = $global:password
        Database               = $global:database
        DbCredential           = $global:dbCredential

        tvDatabases            = $global:tvDatabases
        cmbDatabases           = $global:cmbDatabases
        lblConnectionStatus    = $global:lblConnectionStatus

        txtServer              = $global:txtServer
        txtUser                = $global:txtUser
        txtPassword            = $global:txtPassword
        btnConnectDb           = $global:btnConnectDb

        btnDisconnectDb        = $global:btnDisconnectDb
        btnExecute             = $global:btnExecute
        btnClearQuery          = $global:btnClearQuery
        btnExport              = $global:btnExport
        cmbQueries             = $global:cmbQueries
        tcQueries              = $global:tcQueries
        tcResults              = $global:tcResults
        rtbQueryEditor1        = $global:rtbQueryEditor1
        dgResults              = $global:dgResults
        txtMessages            = $global:txtMessages
        lblRowCount            = $global:lblRowCount
        lblExecutionTimer      = $global:lblExecutionTimer

        MainWindow             = $global:MainWindow
    }
}

function Save-DbUiContext {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    $script:QueryRunning = [bool]$Ctx.QueryRunning
    $script:CurrentQueryPowerShell = $Ctx.CurrentQueryPowerShell
    $script:CurrentQueryRunspace = $Ctx.CurrentQueryRunspace
    $script:CurrentQueryAsync = $Ctx.CurrentQueryAsync
    $script:execUiTimer = $Ctx.execUiTimer
    $script:QueryDoneTimer = $Ctx.QueryDoneTimer
    $script:execStopwatch = $Ctx.execStopwatch

    $global:connection = $Ctx.Connection
    $global:server = $Ctx.Server
    $global:user = $Ctx.User
    $global:password = $Ctx.Password
    $global:database = $Ctx.Database
    $global:dbCredential = $Ctx.DbCredential
}

function Get-DbNameFromComboSelection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ComboBox)

    $selected = $ComboBox.SelectedItem
    if (-not $selected) { return $null }

    $db = $null
    if ($ComboBox.SelectedValuePath -and $ComboBox.SelectedValuePath.Trim().Length -gt 0) {
        $db = [string]$ComboBox.SelectedValue
    } else {
        if ($selected -is [string]) { $db = [string]$selected }
        elseif ($selected.PSObject.Properties["Name"]) { $db = [string]$selected.Name }
        elseif ($selected.PSObject.Properties["Database"]) { $db = [string]$selected.Database }
        else { $db = [string]$selected }
    }

    if ([string]::IsNullOrWhiteSpace($db)) { return $null }
    return $db
}

function Execute-QueryCore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    if ($Ctx.QueryRunning) {
        Write-DzDebug "`t[DEBUG] Query ya en ejecución, ignorando click"
        return
    }

    $Ctx.QueryRunning = $true

    $activeRtb = Get-ActiveQueryRichTextBox -TabControl $Ctx.tcQueries
    if (-not $activeRtb) { throw "No hay una pestaña de consulta activa." }

    $textRange = New-Object System.Windows.Documents.TextRange(
        $activeRtb.Document.ContentStart,
        $activeRtb.Document.ContentEnd
    )

    $rawQuery = $textRange.Text
    if ([string]::IsNullOrWhiteSpace($rawQuery)) { throw "La consulta está vacía." }

    $query = Remove-SqlComments -Query $rawQuery
    if ([string]::IsNullOrWhiteSpace($query)) { throw "La consulta está vacía después de limpiar comentarios." }

    if (-not $Ctx.cmbDatabases) { throw "Control de bases de datos no inicializado." }
    $db = Get-DbNameFromComboSelection -ComboBox $Ctx.cmbDatabases
    if ([string]::IsNullOrWhiteSpace($db)) { throw "Selecciona una base de datos válida." }

    $server = [string]$Ctx.Server
    $userText = [string]$Ctx.User
    $passwordTxt = [string]$Ctx.Password
    if ([string]::IsNullOrWhiteSpace($server) -or [string]::IsNullOrWhiteSpace($userText) -or [string]::IsNullOrWhiteSpace($passwordTxt)) {
        throw "Faltan datos de conexión."
    }

    if ($Ctx.tcResults) { $Ctx.tcResults.Items.Clear() }
    if ($Ctx.dgResults) { $Ctx.dgResults.ItemsSource = $null }
    if ($Ctx.txtMessages) { $Ctx.txtMessages.Text = "" }
    if ($Ctx.lblRowCount) { $Ctx.lblRowCount.Text = "Filas: --" }
    if ($Ctx.lblExecutionTimer) { $Ctx.lblExecutionTimer.Text = "Tiempo: 00:00.0" }

    if (-not $Ctx.execStopwatch) { $Ctx.execStopwatch = [System.Diagnostics.Stopwatch]::new() }

    if ($Ctx.execUiTimer) {
        try { if ($Ctx.execUiTimer.IsEnabled) { $Ctx.execUiTimer.Stop() } } catch {}
    }

    $Ctx.execUiTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Ctx.execUiTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $Ctx.execUiTimer.Add_Tick({
            try {
                if ($global:lblExecutionTimer -and $script:execStopwatch) {
                    $t = $script:execStopwatch.Elapsed
                    $global:lblExecutionTimer.Text = ("Tiempo: {0:mm\:ss\.f}" -f $t)
                }
            } catch {
                Write-DzDebug "`t[DEBUG][Timer] Error actualizando: $_"
            }
        })

    $Ctx.execStopwatch.Restart()
    $Ctx.execUiTimer.Start()

    if ($Ctx.btnExecute) { $Ctx.btnExecute.IsEnabled = $false }

    Write-Host "Query:" -ForegroundColor Cyan
    foreach ($line in ($query -split "`r?`n")) { Write-Host "`t$line" -ForegroundColor Green }
    Write-Host ""
    Write-Host "Ejecutando consulta en '$db'..." -ForegroundColor Cyan
    # CORRECTO
    $modulesPath = $PSScriptRoot
    if (-not (Test-Path $modulesPath)) {
        throw "No se encuentra la carpeta de módulos en: $modulesPath"
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs

    $Ctx.CurrentQueryRunspace = $rs
    $Ctx.CurrentQueryPowerShell = $ps

    $worker = {
        param($Server, $Database, $Query, $User, $Password, $ModulesPath)
        try {
            # Los paths ya vienen correctos desde el main
            $utilPath = Join-Path $ModulesPath "Utilities.psm1"
            $dbPath = Join-Path $ModulesPath "Database.psm1"

            # Verificar que existan antes de importar
            if (-not (Test-Path $utilPath)) {
                throw "No se encuentra Utilities.psm1 en: $utilPath"
            }
            if (-not (Test-Path $dbPath)) {
                throw "No se encuentra Database.psm1 en: $dbPath"
            }

            Import-Module $utilPath -Force -DisableNameChecking -ErrorAction Stop
            Import-Module $dbPath -Force -DisableNameChecking -ErrorAction Stop

            $secure = New-Object System.Security.SecureString
            foreach ($ch in $Password.ToCharArray()) { $secure.AppendChar($ch) }
            $secure.MakeReadOnly()
            $cred = New-Object System.Management.Automation.PSCredential($User, $secure)

            $r = Invoke-SqlQueryMultiResultSet -Server $Server -Database $Database -Query $Query -Credential $cred

            if ($null -eq $r) {
                return @{
                    Success      = $false
                    ErrorMessage = "La ejecución devolvió NULL."
                    ResultSets   = @()
                    Messages     = @()
                    Type         = "Error"
                }
            }

            return $r
        } catch {
            return @{
                Success      = $false
                ErrorMessage = $_.Exception.Message
                ResultSets   = @()
                Messages     = @()
                Type         = "Error"
                InnerError   = $_.Exception.InnerException.Message
                StackTrace   = $_.ScriptStackTrace
            }
        }
    }

    [void]$ps.AddScript($worker).
    AddArgument($server).
    AddArgument($db).
    AddArgument($query).
    AddArgument($userText).
    AddArgument($passwordTxt).
    AddArgument($modulesPath)

    $Ctx.CurrentQueryAsync = $ps.BeginInvoke()

    if ($Ctx.QueryDoneTimer) {
        try { if ($Ctx.QueryDoneTimer.IsEnabled) { $Ctx.QueryDoneTimer.Stop() } } catch {}
    }

    $Ctx.QueryDoneTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $Ctx.QueryDoneTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $Ctx.QueryDoneTimer.Add_Tick({
            try {
                Write-DzDebug "`t[DEBUG][TICK] Verificando query..."
                if (-not $script:CurrentQueryAsync) { Write-DzDebug "`t[DEBUG][TICK] No hay async"; return }
                if (-not $script:CurrentQueryAsync.IsCompleted) { Write-DzDebug "`t[DEBUG] Aún en ejecución..."; return }

                $script:QueryDoneTimer.Stop()
                Write-DzDebug "`t[DEBUG][TICK] Query completada, procesando..."

                $result = $null
                try {
                    $result = $script:CurrentQueryPowerShell.EndInvoke($script:CurrentQueryAsync)
                } catch {
                    $result = @{
                        Success      = $false
                        ErrorMessage = $_.Exception.Message
                        ResultSets   = @()
                        Messages     = @()
                        Type         = "Error"
                    }
                }

                if ($result -is [System.Management.Automation.PSDataCollection[psobject]]) {
                    if ($result.Count -gt 0) { $result = $result[0] } else { $result = $null }
                } elseif ($result -is [System.Array]) {
                    if ($result.Count -gt 0) { $result = $result[0] } else { $result = $null }
                }

                try { if ($script:CurrentQueryPowerShell) { $script:CurrentQueryPowerShell.Dispose() } } catch {}
                try { if ($script:CurrentQueryRunspace) { $script:CurrentQueryRunspace.Close(); $script:CurrentQueryRunspace.Dispose() } } catch {}

                $script:CurrentQueryPowerShell = $null
                $script:CurrentQueryRunspace = $null
                $script:CurrentQueryAsync = $null

                try {
                    if ($script:execStopwatch) { $script:execStopwatch.Stop() }
                    if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) { $script:execUiTimer.Stop() }
                    if ($global:lblExecutionTimer -and $script:execStopwatch) {
                        $t = $script:execStopwatch.Elapsed
                        $global:lblExecutionTimer.Text = ("Tiempo: {0:mm\:ss\.fff}" -f $t)
                    }
                } catch {}

                try { if ($global:btnExecute) { $global:btnExecute.IsEnabled = $true } } catch {}
                $script:QueryRunning = $false

                if (-not $result -or -not $result.Success) {
                    $msg = ""
                    try { $msg = [string]$result.ErrorMessage } catch {}

                    # Agregar detalles adicionales del error si existen
                    if ($result.InnerError) { $msg += "`n`nError interno: $($result.InnerError)" }
                    if ($result.StackTrace) { $msg += "`n`nStack trace:`n$($result.StackTrace)" }

                    if ([string]::IsNullOrWhiteSpace($msg) -and $result -and $result.Messages) {
                        try { $msg = ($result.Messages -join "`n") } catch {}
                    }
                    if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "Error desconocido al ejecutar la consulta." }

                    if ($result -and $result.ResultSets -and $result.ResultSets.Count -gt 0) {
                        try {
                            Show-MultipleResultSets -TabControl $global:tcResults -ResultSets $result.ResultSets
                            Show-ErrorResultTab -ResultsTabControl $global:tcResults -Message $msg -AddWithoutClear
                            if ($global:txtMessages) {
                                $currentText = $global:txtMessages.Text
                                $global:txtMessages.Text = "ERROR: $msg`n`n$currentText"
                            }
                            if ($global:lblRowCount) {
                                $totalRows = ($result.ResultSets | Measure-Object -Property RowCount -Sum).Sum
                                if ($result.ResultSets.Count -eq 1) { $global:lblRowCount.Text = "Filas: $totalRows (con error)" }
                                else { $global:lblRowCount.Text = "Filas: $totalRows ($($result.ResultSets.Count) resultsets, con error)" }
                            }
                        } catch {
                            if ($global:txtMessages) { $global:txtMessages.Text = "Error mostrando resultados: $($_.Exception.Message)" }
                        }
                    } else {
                        if ($global:txtMessages) { $global:txtMessages.Text = $msg }
                        if ($global:lblRowCount) { $global:lblRowCount.Text = "Filas: --" }
                        if ($global:tcResults) { try { Show-ErrorResultTab -ResultsTabControl $global:tcResults -Message $msg } catch {} }
                    }
                    return
                }

                if ($result.DebugLog -and $global:txtMessages) {
                    try {
                        $dbg = ($result.DebugLog -join "`n")
                        if (-not [string]::IsNullOrWhiteSpace($dbg)) { $global:txtMessages.Text = $dbg + "`n`n" + $global:txtMessages.Text }
                    } catch {}
                }

                if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
                    try { Show-MultipleResultSets -TabControl $global:tcResults -ResultSets $result.ResultSets } catch {
                        if ($global:txtMessages) { $global:txtMessages.Text = "Error mostrando resultados: $($_.Exception.Message)" }
                    }
                    return
                }

                if ($result -and ($result -is [hashtable]) -and $result.ContainsKey('RowsAffected') -and $result.RowsAffected -ne $null) {
                    if ($global:tcResults) {
                        $global:tcResults.Items.Clear()
                        $tab = New-Object System.Windows.Controls.TabItem
                        $tab.Header = "Resultado"
                        $text = New-Object System.Windows.Controls.TextBlock
                        $text.Text = "Filas afectadas: $($result.RowsAffected)"
                        $text.Margin = "10"
                        $text.FontSize = 14
                        $text.FontWeight = "Bold"
                        $tab.Content = $text
                        [void]$global:tcResults.Items.Add($tab)
                        $global:tcResults.SelectedItem = $tab
                    }
                    if ($global:txtMessages) { $global:txtMessages.Text = "Filas afectadas: $($result.RowsAffected)" }
                    if ($global:lblRowCount) { $global:lblRowCount.Text = "Filas afectadas: $($result.RowsAffected)" }
                    return
                }

                Show-MultipleResultSets -TabControl $global:tcResults -ResultSets @()
                if ($global:lblRowCount) { $global:lblRowCount.Text = "Filas: 0" }
            } catch {
                $err = "[UI][QueryDoneTimer ERROR] $($_.Exception.Message)`n$($_.ScriptStackTrace)"
                if ($global:txtMessages) { $global:txtMessages.Text = $err }
                Write-Host $err -ForegroundColor Red
                try { if ($global:btnExecute) { $global:btnExecute.IsEnabled = $true } } catch {}
                $script:QueryRunning = $false
                try {
                    if ($script:execStopwatch) { $script:execStopwatch.Stop() }
                    if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) { $script:execUiTimer.Stop() }
                } catch {}
            }
        })

    $Ctx.QueryDoneTimer.Start()
}
function Execute-QueryUiSafe {
    [CmdletBinding()]
    param()

    try {
        $ctx = Get-DbUiContext
        Execute-QueryCore -Ctx $ctx
        Save-DbUiContext -Ctx $ctx
    } catch {
        $msg = $_.Exception.Message
        if ($global:txtMessages) { $global:txtMessages.Text = $msg }
        Write-Host "`n[ERROR Execute-QueryUiSafe] $msg" -ForegroundColor Red
        try { if ($global:btnExecute) { $global:btnExecute.IsEnabled = $true } } catch {}
        $script:QueryRunning = $false
        try {
            if ($script:execStopwatch) { $script:execStopwatch.Stop() }
            if ($script:execUiTimer -and $script:execUiTimer.IsEnabled) { $script:execUiTimer.Stop() }
        } catch {}
    }
}
Export-ModuleMember -Function @(
    'Invoke-SqlQuery', 'Invoke-SqlQueryMultiResultSet', 'Remove-SqlComments', 'Get-SqlDatabases', 'Get-SqlDatabasesInfo', 'Backup-Database',
    'Execute-SqlQuery', 'Show-ResultsConsole', 'Get-IniConnections', 'Load-IniConnectionsToComboBox', 'ConvertTo-DataTable', 'New-QueryTab',
    'Close-QueryTab', 'Execute-QueryInTab', 'Show-MultipleResultSets', 'Export-ResultSetToCsv', 'Export-ResultSetToDelimitedText', 'Get-ActiveQueryRichTextBox',
    'Set-QueryTextInActiveTab', 'Insert-TextIntoActiveQuery', 'Clear-ActiveQueryTab', 'Update-QueryTabHeader', 'Get-ActiveQueryTab', 'Get-TextPointerAtOffset',
    'Get-PredefinedQueries', 'Initialize-PredefinedQueries', 'Remove-SqlComments', 'Set-WpfSqlHighlighting', 'Get-TextPointerFromOffset', 'Get-ResultTabHeaderText',
    'Get-ExportableResultTabs', 'Write-DataTableConsole', 'Show-ErrorResultTab', 'Get-UseDatabaseFromQuery', 'Disconnect-DbUiSafe', 'Connect-DbUiSafe',
    'Export-ResultsUiSafe', 'Execute-QueryUiSafe'
)