#requires -Version 5.0

function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [string]$Query,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password
    )

    try {
        $connectionString = "Server=$Server;Database=$Database;User Id=$Username;Password=$Password;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)

        $connection.Open()

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 0

        # Detectar tipo de consulta
        if ($Query -match "(?si)^\s*(SELECT|WITH)") {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)

            return @{
                Success = $true
                DataTable = $dataTable
                Type = "Query"
            }
        }
        else {
            $rowsAffected = $command.ExecuteNonQuery()

            return @{
                Success = $true
                RowsAffected = $rowsAffected
                Type = "NonQuery"
            }
        }
    }
    catch {
        return @{
            Success = $false
            ErrorMessage = $_.Exception.Message
            Type = "Error"
        }
    }
    finally {
        if ($connection.State -eq [System.Data.ConnectionState]::Open) {
            $connection.Close()
        }
    }
}
function Remove-SqlComments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query
    )

    # Eliminar comentarios multi-línea /* */
    $query = $Query -replace '(?s)/\*.*?\*/', ''

    # Eliminar comentarios de línea --
    $query = $Query -replace '(?m)^\s*--.*\n?', ''
    $query = $Query -replace '(?<!\w)--.*$', ''

    return $query.Trim()
}
function Get-SqlDatabases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password
    )

    $query = @"
SELECT name
FROM sys.databases
WHERE name NOT IN ('tempdb','model','msdb')
  AND state_desc = 'ONLINE'
ORDER BY CASE WHEN name = 'master' THEN 0 ELSE 1 END, name
"@

    $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $query -Username $Username -Password $Password

    if (-not $result.Success) {
        Write-Error "Error obteniendo bases de datos: $($result.ErrorMessage)"
        return @()
    }

    $databases = @()
    foreach ($row in $result.DataTable.Rows) {
        $databases += $row["name"]
    }

    return $databases
}
function Backup-Database {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password,

        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    try {
        $backupQuery = "BACKUP DATABASE [$Database] TO DISK='$BackupPath' WITH CHECKSUM"

        $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $backupQuery -Username $Username -Password $Password

        if ($result.Success) {
            return @{
                Success = $true
                BackupPath = $BackupPath
            }
        }
        else {
            return @{
                Success = $false
                ErrorMessage = $result.ErrorMessage
            }
        }
    }
    catch {
        return @{
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}
function Execute-SqlQuery {
    param (
        [string]$server,
        [string]$database,
        [string]$query
    )
    try {
        $connectionString = "Server=$server;Database=$database;User Id=$global:user;Password=$global:password;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $infoMessages = New-Object System.Collections.ArrayList
        $connection.add_InfoMessage({
                param($sender, $e)
                $infoMessages.Add($e.Message) | Out-Null
            })
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        if ($query -match "(?si)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|RESTORE)") {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable) | Out-Null
            $command.ExecuteNonQuery() | Out-Null
            return @{
                DataTable = $dataTable
                Messages  = $infoMessages
            }
        } else {
            $rowsAffected = $command.ExecuteNonQuery()
            return @{
                RowsAffected = $rowsAffected
                Messages     = $infoMessages
            }
        }
    } catch {
        Write-Host "Error en consulta: $($_.Exception.Message)" -ForegroundColor Red
        throw $_
    } finally {
        $connection.Close()
    }
}
function Show-ResultsConsole {
    param (
        [string]$query
    )
    try {
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query

        if ($results.GetType().Name -eq 'Hashtable') {
            $consoleData = $results.ConsoleData
            if ($consoleData.Count -gt 0) {
                $columns = $consoleData[0].Keys
                $columnWidths = @{}
                foreach ($col in $columns) {
                    $columnWidths[$col] = $col.Length
                }

                Write-Host ""
                $header = ""
                foreach ($col in $columns) {
                    $header += $col.PadRight($columnWidths[$col] + 4)
                }
                Write-Host $header
                Write-Host ("-" * $header.Length)

                foreach ($row in $consoleData) {
                    $rowText = ""
                    foreach ($col in $columns) {
                        $rowText += ($row[$col].ToString()).PadRight($columnWidths[$col] + 4)
                    }
                    Write-Host $rowText
                }
            } else {
                Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow
            }
        } else {
            Write-Host "`nFilas afectadas: $results" -ForegroundColor Green
        }
    } catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
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
            $line = Get-Content $FilePath | Where-Object { $_ -match "^$Key\s*=" }
            if ($line) {
                return $line.Split('=')[1].Trim()
            }
        }
        return $null
    }
    foreach ($entry in $pathsToCheck) {
        $mainIni = Join-Path $entry.Path $entry.INI
        if (Test-Path $mainIni) {
            $dataSource = Get-IniValue -FilePath $mainIni -Key "DataSource"
            if ($dataSource -and $dataSource -notin $connections) {
                $connections += $dataSource
            }
        }
        $inisFolder = Join-Path $entry.Path "INIS"
        if (Test-Path $inisFolder) {
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
            foreach ($iniFile in $iniFiles) {
                $dataSource = Get-IniValue -FilePath $iniFile.FullName -Key "DataSource"
                if ($dataSource -and $dataSource -notin $connections) {
                    $connections += $dataSource
                }
            }
        }
    }

    return $connections | Sort-Object
}
Export-ModuleMember -Function Invoke-SqlQuery, Remove-SqlComments, Get-SqlDatabases, Backup-Database, Execute-SqlQuery, Show-ResultsConsole, Get-IniConnections