#requires -Version 5.0

function New-SqlTreeNode {
    param(
        [Parameter(Mandatory = $true)][string]$Header,
        [Parameter(Mandatory = $true)][hashtable]$Tag,
        [Parameter(Mandatory = $false)][bool]$HasPlaceholder
    )

    $node = New-Object System.Windows.Controls.TreeViewItem
    $node.Header = $Header
    $node.Tag = $Tag

    if ($HasPlaceholder) {
        $placeholder = New-Object System.Windows.Controls.TreeViewItem
        $placeholder.Header = "Cargando..."
        $placeholder.Tag = @{ Type = "Placeholder" }
        [void]$node.Items.Add($placeholder)
    }

    return $node
}

function Initialize-SqlTreeView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TreeView,
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][scriptblock]$InsertTextHandler
    )

    $TreeView.Items.Clear()

    $serverTag = @{ Type = "Server"; Server = $Server; Credential = $Credential; InsertTextHandler = $InsertTextHandler }
    $serverNode = New-SqlTreeNode -Header "📊 $Server" -Tag $serverTag -HasPlaceholder $true

    $serverNode.Add_Expanded({
            param($sender, $e)
            if (-not $sender.Tag.Loaded) {
                $sender.Tag.Loaded = $true
                Load-DatabasesIntoTree -ServerNode $sender
            }
        })

    [void]$TreeView.Items.Add($serverNode)
}

function Load-DatabasesIntoTree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$ServerNode
    )

    $ServerNode.Items.Clear()
    $server = $ServerNode.Tag.Server
    $credential = $ServerNode.Tag.Credential

    try {
        $databases = Get-SqlDatabases -Server $server -Credential $credential
    } catch {
        $errorNode = New-SqlTreeNode -Header "Error cargando bases" -Tag @{ Type = "Error" } -HasPlaceholder $false
        [void]$ServerNode.Items.Add($errorNode)
        return
    }

    foreach ($db in $databases) {
        $dbTag = @{ Type = "Database"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler }
        $dbNode = New-SqlTreeNode -Header "🗄️ $db" -Tag $dbTag -HasPlaceholder $false

        $tablesTag = @{ Type = "TablesRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler }
        $viewsTag = @{ Type = "ViewsRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler }
        $procsTag = @{ Type = "ProceduresRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler }

        $tablesNode = New-SqlTreeNode -Header "📋 Tablas" -Tag $tablesTag -HasPlaceholder $true
        $viewsNode = New-SqlTreeNode -Header "👁️ Vistas" -Tag $viewsTag -HasPlaceholder $true
        $procsNode = New-SqlTreeNode -Header "⚙️ Procedimientos" -Tag $procsTag -HasPlaceholder $true

        $tablesNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })
        $viewsNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })
        $procsNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })

        [void]$dbNode.Items.Add($tablesNode)
        [void]$dbNode.Items.Add($viewsNode)
        [void]$dbNode.Items.Add($procsNode)

        [void]$ServerNode.Items.Add($dbNode)
    }
}

function Load-TablesIntoNode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$RootNode
    )

    if ($RootNode.Tag.Loaded) { return }
    $RootNode.Tag.Loaded = $true
    $RootNode.Items.Clear()

    $server = $RootNode.Tag.Server
    $database = $RootNode.Tag.Database
    $credential = $RootNode.Tag.Credential
    $nodeType = $RootNode.Tag.Type

    switch ($nodeType) {
        "TablesRoot" {
            $query = @"
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
ORDER BY TABLE_SCHEMA, TABLE_NAME
"@
            $result = Invoke-SqlQuery -Server $server -Database $database -Query $query -Credential $credential
            if (-not $result.Success) { return }

            foreach ($row in $result.DataTable.Rows) {
                $schema = $row.TABLE_SCHEMA
                $table = $row.TABLE_NAME
                $tag = @{ Type = "Table"; Database = $database; Schema = $schema; Table = $table; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler }
                $tableNode = New-SqlTreeNode -Header "📋 [$schema].[$table]" -Tag $tag -HasPlaceholder $true
                $tableNode.Add_Expanded({ param($s, $e) Load-ColumnsIntoTableNode -TableNode $s })
                $tableNode.Add_MouseDoubleClick({
                        param($s, $e)
                        $queryText = "SELECT TOP 100 * FROM [$($s.Tag.Schema)].[$($s.Tag.Table)]"
                        if ($s.Tag.InsertTextHandler) { & $s.Tag.InsertTextHandler $queryText }
                        $e.Handled = $true
                    })
                Add-TreeNodeContextMenu -TableNode $tableNode
                [void]$RootNode.Items.Add($tableNode)
            }
        }
        "ViewsRoot" {
            $query = @"
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.VIEWS
ORDER BY TABLE_SCHEMA, TABLE_NAME
"@
            $result = Invoke-SqlQuery -Server $server -Database $database -Query $query -Credential $credential
            if (-not $result.Success) { return }

            foreach ($row in $result.DataTable.Rows) {
                $schema = $row.TABLE_SCHEMA
                $view = $row.TABLE_NAME
                $tag = @{ Type = "View"; Database = $database; Schema = $schema; View = $view; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler }
                $viewNode = New-SqlTreeNode -Header "👁️ [$schema].[$view]" -Tag $tag -HasPlaceholder $false
                $viewNode.Add_MouseDoubleClick({
                        param($s, $e)
                        $queryText = "SELECT TOP 100 * FROM [$($s.Tag.Schema)].[$($s.Tag.View)]"
                        if ($s.Tag.InsertTextHandler) { & $s.Tag.InsertTextHandler $queryText }
                        $e.Handled = $true
                    })
                [void]$RootNode.Items.Add($viewNode)
            }
        }
        "ProceduresRoot" {
            $query = @"
SELECT SPECIFIC_SCHEMA, SPECIFIC_NAME
FROM INFORMATION_SCHEMA.ROUTINES
WHERE ROUTINE_TYPE = 'PROCEDURE'
ORDER BY SPECIFIC_SCHEMA, SPECIFIC_NAME
"@
            $result = Invoke-SqlQuery -Server $server -Database $database -Query $query -Credential $credential
            if (-not $result.Success) { return }

            foreach ($row in $result.DataTable.Rows) {
                $schema = $row.SPECIFIC_SCHEMA
                $proc = $row.SPECIFIC_NAME
                $tag = @{ Type = "Procedure"; Database = $database; Schema = $schema; Procedure = $proc; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler }
                $procNode = New-SqlTreeNode -Header "⚙️ [$schema].[$proc]" -Tag $tag -HasPlaceholder $false
                $procNode.Add_MouseDoubleClick({
                        param($s, $e)
                        $queryText = "EXEC [$($s.Tag.Schema)].[$($s.Tag.Procedure)]"
                        if ($s.Tag.InsertTextHandler) { & $s.Tag.InsertTextHandler $queryText }
                        $e.Handled = $true
                    })
                [void]$RootNode.Items.Add($procNode)
            }
        }
    }
}

function Load-ColumnsIntoTableNode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TableNode
    )

    if ($TableNode.Tag.Loaded) { return }
    $TableNode.Tag.Loaded = $true
    $TableNode.Items.Clear()

    $server = $TableNode.Tag.Server
    $database = $TableNode.Tag.Database
    $schema = $TableNode.Tag.Schema
    $table = $TableNode.Tag.Table
    $credential = $TableNode.Tag.Credential

    $query = @"
SELECT
    c.COLUMN_NAME,
    c.DATA_TYPE,
    c.CHARACTER_MAXIMUM_LENGTH,
    c.NUMERIC_PRECISION,
    c.NUMERIC_SCALE,
    c.IS_NULLABLE,
    CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PRIMARY_KEY
FROM INFORMATION_SCHEMA.COLUMNS c
LEFT JOIN (
    SELECT ku.TABLE_SCHEMA, ku.TABLE_NAME, ku.COLUMN_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
    JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
        ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
    WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
) pk ON c.TABLE_SCHEMA = pk.TABLE_SCHEMA AND c.TABLE_NAME = pk.TABLE_NAME AND c.COLUMN_NAME = pk.COLUMN_NAME
WHERE c.TABLE_SCHEMA = '$schema' AND c.TABLE_NAME = '$table'
ORDER BY c.ORDINAL_POSITION
"@

    $result = Invoke-SqlQuery -Server $server -Database $database -Query $query -Credential $credential
    if (-not $result.Success) { return }

    foreach ($row in $result.DataTable.Rows) {
        $name = $row.COLUMN_NAME
        $type = $row.DATA_TYPE
        $length = $row.CHARACTER_MAXIMUM_LENGTH
        $isPk = [int]$row.IS_PRIMARY_KEY -eq 1
        $label = if ($isPk) { "🔑 $name ($type)" } else { "• $name ($type)" }
        $tag = @{ Type = "Column"; Column = $name; InsertTextHandler = $TableNode.Tag.InsertTextHandler }
        $colNode = New-SqlTreeNode -Header $label -Tag $tag -HasPlaceholder $false
        $colNode.Add_MouseDoubleClick({
                param($s, $e)
                if ($s.Tag.InsertTextHandler) { & $s.Tag.InsertTextHandler "[$($s.Tag.Column)]" }
                $e.Handled = $true
            })
        [void]$TableNode.Items.Add($colNode)
    }
}

function Get-CreateTableScript {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][string]$Schema,
        [Parameter(Mandatory = $true)][string]$Table,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential
    )

    $query = @"
SELECT
    c.COLUMN_NAME,
    c.DATA_TYPE,
    c.CHARACTER_MAXIMUM_LENGTH,
    c.NUMERIC_PRECISION,
    c.NUMERIC_SCALE,
    c.IS_NULLABLE,
    CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PRIMARY_KEY
FROM INFORMATION_SCHEMA.COLUMNS c
LEFT JOIN (
    SELECT ku.TABLE_SCHEMA, ku.TABLE_NAME, ku.COLUMN_NAME
    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
    JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
        ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
    WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
) pk ON c.TABLE_SCHEMA = pk.TABLE_SCHEMA AND c.TABLE_NAME = pk.TABLE_NAME AND c.COLUMN_NAME = pk.COLUMN_NAME
WHERE c.TABLE_SCHEMA = '$Schema' AND c.TABLE_NAME = '$Table'
ORDER BY c.ORDINAL_POSITION
"@

    $result = Invoke-SqlQuery -Server $Server -Database $Database -Query $query -Credential $Credential
    if (-not $result.Success) { return $null }

    $lines = New-Object System.Collections.Generic.List[string]
    $pkColumns = New-Object System.Collections.Generic.List[string]

    foreach ($row in $result.DataTable.Rows) {
        $colName = $row.COLUMN_NAME
        $dataType = $row.DATA_TYPE
        $len = $row.CHARACTER_MAXIMUM_LENGTH
        $precision = $row.NUMERIC_PRECISION
        $scale = $row.NUMERIC_SCALE
        $nullable = $row.IS_NULLABLE -eq 'YES'

        $typeSuffix = ""
        if ($dataType -in @('char', 'varchar', 'nchar', 'nvarchar', 'binary', 'varbinary')) {
            if ($len -eq -1) { $typeSuffix = "(MAX)" } else { $typeSuffix = "($len)" }
        } elseif ($dataType -in @('decimal', 'numeric')) {
            $typeSuffix = "($precision,$scale)"
        }

        $nullText = if ($nullable) { "NULL" } else { "NOT NULL" }
        $lines.Add("    [$colName] $dataType$typeSuffix $nullText")

        if ([int]$row.IS_PRIMARY_KEY -eq 1) {
            $pkColumns.Add("[$colName]")
        }
    }

    if ($pkColumns.Count -gt 0) {
        $pkLine = "    CONSTRAINT [PK_$Table] PRIMARY KEY (" + ($pkColumns -join ', ') + ")"
        $lines.Add($pkLine)
    }

    $script = @()
    $script += "CREATE TABLE [$Schema].[$Table] ("
    $script += ($lines -join ",`n")
    $script += ")"

    return ($script -join "`n")
}

function Add-TreeNodeContextMenu {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TableNode
    )

    $menu = New-Object System.Windows.Controls.ContextMenu

    $menuSelectTop = New-Object System.Windows.Controls.MenuItem
    $menuSelectTop.Header = "SELECT TOP 100 *"
    $menuSelectTop.Add_Click({
            $queryText = "SELECT TOP 100 * FROM [$($TableNode.Tag.Schema)].[$($TableNode.Tag.Table)]"
            if ($TableNode.Tag.InsertTextHandler) { & $TableNode.Tag.InsertTextHandler $queryText }
        })

    $menuSelectAll = New-Object System.Windows.Controls.MenuItem
    $menuSelectAll.Header = "SELECT *"
    $menuSelectAll.Add_Click({
            $queryText = "SELECT * FROM [$($TableNode.Tag.Schema)].[$($TableNode.Tag.Table)]"
            if ($TableNode.Tag.InsertTextHandler) { & $TableNode.Tag.InsertTextHandler $queryText }
        })

    $menuScript = New-Object System.Windows.Controls.MenuItem
    $menuScript.Header = "Script CREATE TABLE"
    $menuScript.Add_Click({
            $scriptText = Get-CreateTableScript -Server $TableNode.Tag.Server -Database $TableNode.Tag.Database -Schema $TableNode.Tag.Schema -Table $TableNode.Tag.Table -Credential $TableNode.Tag.Credential
            if ($scriptText -and $TableNode.Tag.InsertTextHandler) { & $TableNode.Tag.InsertTextHandler $scriptText }
        })

    $menuCopy = New-Object System.Windows.Controls.MenuItem
    $menuCopy.Header = "Copiar nombre"
    $menuCopy.Add_Click({
            $name = "[$($TableNode.Tag.Schema)].[$($TableNode.Tag.Table)]"
            [System.Windows.Clipboard]::SetText($name)
        })

    [void]$menu.Items.Add($menuSelectTop)
    [void]$menu.Items.Add($menuSelectAll)
    [void]$menu.Items.Add($menuScript)
    [void]$menu.Items.Add($menuCopy)

    $TableNode.ContextMenu = $menu
}

Export-ModuleMember -Function @(
    'Initialize-SqlTreeView',
    'Load-DatabasesIntoTree',
    'Load-TablesIntoNode',
    'Load-ColumnsIntoTableNode',
    'Add-TreeNodeContextMenu'
)
