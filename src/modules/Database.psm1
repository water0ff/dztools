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

Export-ModuleMember -Function Invoke-SqlQuery, Remove-SqlComments, Get-SqlDatabases, Backup-Database
