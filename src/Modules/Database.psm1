function Remove-SqlComments {
    <#
        .SYNOPSIS
            Limpia comentarios de consultas SQL.
        .PARAMETER Query
            Texto SQL a limpiar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query
    )

    $cleanedQuery = $Query -replace '(?s)/\*.*?\*/', ''
    $cleanedQuery = $cleanedQuery -replace '(?m)^\s*--.*\n?', ''
    $cleanedQuery = $cleanedQuery -replace '(?<!\w)--.*$', ''
    return $cleanedQuery.Trim()
}

function ConvertTo-DataTable {
    <#
        .SYNOPSIS
            Convierte objetos de PowerShell en un DataTable .NET.
        .PARAMETER InputObject
            Colección o elemento a transformar.
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    $dt = New-Object System.Data.DataTable
    process {
        if (-not $_) { return }
        if (-not $dt.Columns.Count) {
            $_.PSObject.Properties | ForEach-Object { $null = $dt.Columns.Add($_.Name, $_.Value.GetType()) }
        }
        $row = $dt.NewRow()
        $_.PSObject.Properties | ForEach-Object { $row[$_.Name] = $_.Value }
        $dt.Rows.Add($row) | Out-Null
    }
    end { return $dt }
}

function Execute-SqlQuery {
    <#
        .SYNOPSIS
            Ejecuta una consulta SQL y devuelve resultados o mensajes.
        .PARAMETER Server
            Instancia de SQL Server.
        .PARAMETER Database
            Base de datos a utilizar.
        .PARAMETER Query
            Consulta a ejecutar (se limpia de comentarios antes de ejecutarse).
        .PARAMETER Username
            Usuario SQL Server.
        .PARAMETER Password
            Contraseña del usuario SQL.
    #>
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

    $cleanQuery = Remove-SqlComments -Query $Query
    $connectionString = "Server=$Server;Database=$Database;User Id=$Username;Password=$Password;MultipleActiveResultSets=True"
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $infoMessages = New-Object System.Collections.ArrayList

    $connection.add_InfoMessage({
        param($sender, $e)
        $infoMessages.Add($e.Message) | Out-Null
    })

    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $cleanQuery

        if ($cleanQuery -match "(?si)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|RESTORE)") {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable) | Out-Null
            $command.ExecuteNonQuery() | Out-Null
            return @{ DataTable = $dataTable; Messages = $infoMessages }
        }
        else {
            $rowsAffected = $command.ExecuteNonQuery()
            return @{ RowsAffected = $rowsAffected; Messages = $infoMessages }
        }
    }
    catch {
        Write-Error "Error en consulta: $($_.Exception.Message)"
        throw $_
    }
    finally {
        if ($connection.State -eq [System.Data.ConnectionState]::Open) {
            $connection.Close()
        }
    }
}

function Show-ResultsConsole {
    <#
        .SYNOPSIS
            Muestra en consola los resultados devueltos por Execute-SqlQuery.
        .PARAMETER Query
            Consulta a ejecutar y mostrar.
        .PARAMETER Server
            Instancia SQL.
        .PARAMETER Database
            Base de datos seleccionada.
        .PARAMETER Username
            Usuario SQL.
        .PARAMETER Password
            Contraseña SQL.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query,

        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [string]$Database,

        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        [string]$Password
    )

    try {
        $results = Execute-SqlQuery -Server $Server -Database $Database -Query $Query -Username $Username -Password $Password

        if ($results.ContainsKey('DataTable')) {
            $consoleData = $results.DataTable | ConvertTo-DataTable
            if ($consoleData.Rows.Count -gt 0) {
                $columns = $consoleData.Columns | ForEach-Object { $_.ColumnName }
                $columnWidths = @{}
                foreach ($col in $columns) { $columnWidths[$col] = $col.Length }

                Write-Host ""
                $header = ""
                foreach ($col in $columns) { $header += $col.PadRight($columnWidths[$col] + 4) }
                Write-Host $header
                Write-Host ("-" * $header.Length)

                foreach ($row in $consoleData.Rows) {
                    $rowText = ""
                    foreach ($col in $columns) { $rowText += ($row[$col].ToString()).PadRight($columnWidths[$col] + 4) }
                    Write-Host $rowText
                }
            }
            else {
                Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "`nFilas afectadas: $($results.RowsAffected)" -ForegroundColor Green
        }

        if ($results.Messages.Count -gt 0) {
            Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
            $results.Messages | ForEach-Object { Write-Host $_ -ForegroundColor DarkGray }
        }
    }
    catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
}

Export-ModuleMember -Function Remove-SqlComments, ConvertTo-DataTable, Execute-SqlQuery, Show-ResultsConsole
