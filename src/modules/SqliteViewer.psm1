#requires -Version 5.0
Add-Type -AssemblyName System.Data

$script:SqliteProvider = $null

function Get-SqliteProvider {
    if (-not [string]::IsNullOrWhiteSpace($script:SqliteProvider)) { return $script:SqliteProvider }
    try {
        Add-Type -AssemblyName Microsoft.Data.Sqlite -ErrorAction Stop
        $script:SqliteProvider = "Microsoft.Data.Sqlite"
        return $script:SqliteProvider
    } catch {
        # continue
    }
    try {
        Add-Type -AssemblyName System.Data.SQLite -ErrorAction Stop
        $script:SqliteProvider = "System.Data.SQLite"
        return $script:SqliteProvider
    } catch {
        # continue
    }
    return $null
}

function Open-SqliteConnection {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{ Success = $false; ErrorMessage = "No se encontró el archivo: $Path" }
    }

    $provider = Get-SqliteProvider
    if (-not $provider) {
        return @{ Success = $false; ErrorMessage = "No se encontró un proveedor SQLite. Instala Microsoft.Data.Sqlite o System.Data.SQLite." }
    }

    try {
        switch ($provider) {
            "Microsoft.Data.Sqlite" {
                $connection = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$Path")
            }
            "System.Data.SQLite" {
                $connection = [System.Data.SQLite.SQLiteConnection]::new("Data Source=$Path;Version=3;")
            }
        }
        $connection.Open()
        return @{ Success = $true; Connection = $connection; Provider = $provider }
    } catch {
        return @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}

function Close-SqliteConnection {
    param([Parameter(Mandatory = $true)]$Connection)
    try {
        if ($Connection -and $Connection.State -ne [System.Data.ConnectionState]::Closed) {
            $Connection.Close()
        }
        if ($Connection) { $Connection.Dispose() }
    } catch {
        Write-DzDebug "`t[DEBUG][SQLite] Error cerrando conexión: $($_.Exception.Message)" -Color DarkGray
    }
}

function Get-SqliteTables {
    param([Parameter(Mandatory = $true)]$Connection)

    $query = @"
SELECT name
FROM sqlite_master
WHERE type = 'table'
  AND name NOT LIKE 'sqlite_%'
ORDER BY name;
"@
    $result = Invoke-SqliteQuery -Connection $Connection -Query $query
    if (-not $result.Success -or -not $result.DataTable) { return @() }
    $tables = @()
    foreach ($row in $result.DataTable.Rows) {
        $tables += [string]$row["name"]
    }
    $tables
}

function Invoke-SqliteQuery {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$Query
    )

    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 30

        $shouldRead = $Query -match '^\s*(SELECT|PRAGMA|WITH|EXPLAIN)\b'
        if ($shouldRead) {
            $reader = $command.ExecuteReader()
            $table = New-Object System.Data.DataTable
            $table.Load($reader)
            $reader.Close()
            return @{ Success = $true; DataTable = $table; RowsAffected = $table.Rows.Count }
        }

        $rows = $command.ExecuteNonQuery()
        return @{ Success = $true; RowsAffected = $rows }
    } catch {
        return @{ Success = $false; ErrorMessage = $_.Exception.Message }
    }
}

function Initialize-SqliteTreeView {
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.TreeView]$TreeView,
        [Parameter(Mandatory = $true)][string[]]$Tables,
        [Parameter(Mandatory = $false)][scriptblock]$OnTableSelected
    )

    $TreeView.Items.Clear()
    foreach ($table in $Tables) {
        $item = New-Object System.Windows.Controls.TreeViewItem
        $item.Header = "📄 $table"
        $item.Tag = $table
        $item.Add_MouseDoubleClick({
                param($s, $e)
                if ($OnTableSelected) { & $OnTableSelected ([string]$s.Tag) }
            }.GetNewClosure())
        [void]$TreeView.Items.Add($item)
    }
}

Export-ModuleMember -Function @(
    'Get-SqliteProvider',
    'Open-SqliteConnection',
    'Close-SqliteConnection',
    'Get-SqliteTables',
    'Invoke-SqliteQuery',
    'Initialize-SqliteTreeView'
)
