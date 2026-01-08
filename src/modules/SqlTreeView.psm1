#requires -Version 5.0
function New-SqlTreeNode {
    param([Parameter(Mandatory = $true)][string]$Header, [Parameter(Mandatory = $true)][hashtable]$Tag, [Parameter(Mandatory = $false)][bool]$HasPlaceholder)
    $node = New-Object System.Windows.Controls.TreeViewItem
    $node.Header = $Header
    $node.Tag = $Tag
    if ($HasPlaceholder) {
        $placeholder = New-Object System.Windows.Controls.TreeViewItem
        $placeholder.Header = "Cargando..."
        $placeholder.Tag = @{Type = "Placeholder" }
        [void]$node.Items.Add($placeholder)
    }
    $node
}
function Initialize-SqlTreeView {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$TreeView, [Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential, [Parameter(Mandatory = $false)][scriptblock]$InsertTextHandler, [Parameter(Mandatory = $false)][scriptblock]$OnDatabaseSelected, [Parameter(Mandatory = $false)][bool]$AutoExpand = $true)
    $TreeView.Items.Clear()
    $serverTag = @{Type = "Server"; Server = $Server; Credential = $Credential; InsertTextHandler = $InsertTextHandler; OnDatabaseSelected = $OnDatabaseSelected; Loaded = $false }
    $serverNode = New-SqlTreeNode -Header "📊 $Server" -Tag $serverTag -HasPlaceholder $true
    $serverNode.Add_Expanded({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG][TreeView] Expand Server: $($sender.Tag.Server) Loaded=$($sender.Tag.Loaded)"
            if (-not $sender.Tag.Loaded) {
                $sender.Tag.Loaded = $true
                Load-DatabasesIntoTree -ServerNode $sender
            }
        })
    [void]$TreeView.Items.Add($serverNode)
    if ($AutoExpand) {
        Write-DzDebug "`t[DEBUG][TreeView] AutoExpand Server: $Server"
        $serverNode.Tag.Loaded = $true
        Load-DatabasesIntoTree -ServerNode $serverNode
        $serverNode.IsExpanded = $true
    }
}
function Load-DatabasesIntoTree {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$ServerNode)
    $ServerNode.Items.Clear()
    $server = [string]$ServerNode.Tag.Server
    $credential = $ServerNode.Tag.Credential
    Write-DzDebug "`t[DEBUG][TreeView] Load DBs: Server='$server'"
    try {
        $databases = Get-SqlDatabases -Server $server -Credential $credential
    } catch {
        $errorNode = New-SqlTreeNode -Header "Error cargando bases" -Tag @{Type = "Error" } -HasPlaceholder $false
        [void]$ServerNode.Items.Add($errorNode)
        return
    }
    foreach ($db in $databases) {
        $dbTag = @{Type = "Database"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler; OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected }
        $dbNode = New-SqlTreeNode -Header "🗄️ $db" -Tag $dbTag -HasPlaceholder $false
        $dbNode.Add_MouseDoubleClick({
                param($s, $e)
                Write-DzDebug "`t[DEBUG][TreeView] Doble clic DB: $($s.Tag.Database) | HasHandler=$([bool]$s.Tag.OnDatabaseSelected)"
                if ($s.Tag.OnDatabaseSelected) { & $s.Tag.OnDatabaseSelected $s.Tag.Database }
                $e.Handled = $true
            })
        Add-DatabaseContextMenu -DatabaseNode $dbNode
        $tablesTag = @{Type = "TablesRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler; OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected; Loaded = $false }
        $viewsTag = @{Type = "ViewsRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler; OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected; Loaded = $false }
        $procsTag = @{Type = "ProceduresRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $ServerNode.Tag.InsertTextHandler; OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected; Loaded = $false }
        $tablesNode = New-SqlTreeNode -Header "📋 Tablas" -Tag $tablesTag -HasPlaceholder $true
        $viewsNode = New-SqlTreeNode -Header "👁️ Vistas" -Tag $viewsTag -HasPlaceholder $true
        $procsNode = New-SqlTreeNode -Header "⚙️ Procedimientos" -Tag $procsTag -HasPlaceholder $true
        $tablesNode.Add_Expanded({ param($s, $e)Write-DzDebug "`t[DEBUG][TreeView] Expand TablasRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
        $viewsNode.Add_Expanded({ param($s, $e)Write-DzDebug "`t[DEBUG][TreeView] Expand VistasRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
        $procsNode.Add_Expanded({ param($s, $e)Write-DzDebug "`t[DEBUG][TreeView] Expand ProcsRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
        [void]$dbNode.Items.Add($tablesNode)
        [void]$dbNode.Items.Add($viewsNode)
        [void]$dbNode.Items.Add($procsNode)
        [void]$ServerNode.Items.Add($dbNode)
    }
}
function Load-TablesIntoNode {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$RootNode)
    if ($RootNode.Tag.Loaded) { return }
    $RootNode.Tag.Loaded = $true
    $RootNode.Items.Clear()
    $server = [string]$RootNode.Tag.Server
    $database = [string]$RootNode.Tag.Database
    $credential = $RootNode.Tag.Credential
    $nodeType = [string]$RootNode.Tag.Type
    Write-DzDebug "`t[DEBUG][TreeView] Load NodeType=$nodeType Server='$server' DB='$database'"
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
                $tag = @{Type = "Table"; Database = $database; Schema = $schema; Table = $table; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler; OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected; Loaded = $false }
                $tableNode = New-SqlTreeNode -Header "📋 [$schema].[$table]" -Tag $tag -HasPlaceholder $true
                $tableNode.Add_Expanded({ param($s, $e)Write-DzDebug "`t[DEBUG][TreeView] Expand Table: $($s.Tag.Database) [$($s.Tag.Schema)].[$($s.Tag.Table)]"; Load-ColumnsIntoTableNode -TableNode $s })
                $tableNode.Add_MouseDoubleClick({
                        param($s, $e)
                        Write-DzDebug "`t[DEBUG][TreeView] Doble clic Table: DB=$($s.Tag.Database) [$($s.Tag.Schema)].[$($s.Tag.Table)]"
                        if ($s.Tag.OnDatabaseSelected) { & $s.Tag.OnDatabaseSelected $s.Tag.Database }
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
                $tag = @{Type = "View"; Database = $database; Schema = $schema; View = $view; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler; OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected }
                $viewNode = New-SqlTreeNode -Header "👁️ [$schema].[$view]" -Tag $tag -HasPlaceholder $false
                $viewNode.Add_MouseDoubleClick({
                        param($s, $e)
                        Write-DzDebug "`t[DEBUG][TreeView] Doble clic View: DB=$($s.Tag.Database) [$($s.Tag.Schema)].[$($s.Tag.View)]"
                        if ($s.Tag.OnDatabaseSelected) { & $s.Tag.OnDatabaseSelected $s.Tag.Database }
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
                $tag = @{Type = "Procedure"; Database = $database; Schema = $schema; Procedure = $proc; Server = $server; Credential = $credential; InsertTextHandler = $RootNode.Tag.InsertTextHandler; OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected }
                $procNode = New-SqlTreeNode -Header "⚙️ [$schema].[$proc]" -Tag $tag -HasPlaceholder $false
                $procNode.Add_MouseDoubleClick({
                        param($s, $e)
                        Write-DzDebug "`t[DEBUG][TreeView] Doble clic Proc: DB=$($s.Tag.Database) [$($s.Tag.Schema)].[$($s.Tag.Procedure)]"
                        if ($s.Tag.OnDatabaseSelected) { & $s.Tag.OnDatabaseSelected $s.Tag.Database }
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
    param([Parameter(Mandatory = $true)]$TableNode)
    if ($TableNode.Tag.Loaded) { return }
    $TableNode.Tag.Loaded = $true
    $TableNode.Items.Clear()
    $server = [string]$TableNode.Tag.Server
    $database = [string]$TableNode.Tag.Database
    $schema = [string]$TableNode.Tag.Schema
    $table = [string]$TableNode.Tag.Table
    $credential = $TableNode.Tag.Credential
    Write-DzDebug "`t[DEBUG][TreeView] Load Columns: Server='$server' DB='$database' [$schema].[$table]"
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
        $isPk = ([int]$row.IS_PRIMARY_KEY -eq 1)
        $label = if ($isPk) { "🔑 $name ($type)" }else { "• $name ($type)" }
        $tag = @{Type = "Column"; Column = $name; InsertTextHandler = $TableNode.Tag.InsertTextHandler }
        $colNode = New-SqlTreeNode -Header $label -Tag $tag -HasPlaceholder $false
        $colNode.Add_MouseDoubleClick({
                param($s, $e)
                Write-DzDebug "`t[DEBUG][TreeView] Doble clic Column: [$($s.Tag.Column)]"
                if ($s.Tag.InsertTextHandler) { & $s.Tag.InsertTextHandler "[$($s.Tag.Column)]" }
                $e.Handled = $true
            })
        [void]$TableNode.Items.Add($colNode)
    }
}
function Get-CreateTableScript {
    param([Parameter(Mandatory = $true)][string]$Server, [Parameter(Mandatory = $true)][string]$Database, [Parameter(Mandatory = $true)][string]$Schema, [Parameter(Mandatory = $true)][string]$Table, [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential)
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
        $nullable = ($row.IS_NULLABLE -eq "YES")
        $typeSuffix = ""
        if ($dataType -in @("char", "varchar", "nchar", "nvarchar", "binary", "varbinary")) {
            if ($len -eq -1) { $typeSuffix = "(MAX)" }else { $typeSuffix = "($len)" }
        } elseif ($dataType -in @("decimal", "numeric")) {
            $typeSuffix = "($precision,$scale)"
        }
        $nullText = if ($nullable) { "NULL" }else { "NOT NULL" }
        $lines.Add("    [$colName] $dataType$typeSuffix $nullText")
        if ([int]$row.IS_PRIMARY_KEY -eq 1) { $pkColumns.Add("[$colName]") }
    }
    if ($pkColumns.Count -gt 0) {
        $pkLine = "    CONSTRAINT [PK_$Table] PRIMARY KEY (" + ($pkColumns -join ", ") + ")"
        $lines.Add($pkLine)
    }
    $script = @()
    $script += "CREATE TABLE [$Schema].[$Table] ("
    $script += ($lines -join ",`n")
    $script += ")"
    ($script -join "`n")
}
function Add-TreeNodeContextMenu {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$TableNode)
    $menu = New-Object System.Windows.Controls.ContextMenu
    $menuSelectTop = New-Object System.Windows.Controls.MenuItem
    $menuSelectTop.Header = "SELECT TOP 100 *"
    $menuSelectTop.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $db = [string]$node.Tag.Database
            $schema = [string]$node.Tag.Schema
            $table = [string]$node.Tag.Table
            Write-DzDebug "`t[DEBUG][TreeView] Context SELECT TOP: DB=$db [$schema].[$table]"
            if ($node.Tag.OnDatabaseSelected) { & $node.Tag.OnDatabaseSelected $db }
            $queryText = "SELECT TOP 100 * FROM [$schema].[$table]"
            if ($node.Tag.InsertTextHandler) { & $node.Tag.InsertTextHandler $queryText }
        })
    $menuSelectAll = New-Object System.Windows.Controls.MenuItem
    $menuSelectAll.Header = "SELECT *"
    $menuSelectAll.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $db = [string]$node.Tag.Database
            $schema = [string]$node.Tag.Schema
            $table = [string]$node.Tag.Table
            Write-DzDebug "`t[DEBUG][TreeView] Context SELECT *: DB=$db [$schema].[$table]"
            if ($node.Tag.OnDatabaseSelected) { & $node.Tag.OnDatabaseSelected $db }
            $queryText = "SELECT * FROM [$schema].[$table]"
            if ($node.Tag.InsertTextHandler) { & $node.Tag.InsertTextHandler $queryText }
        })
    $menuScript = New-Object System.Windows.Controls.MenuItem
    $menuScript.Header = "Script CREATE TABLE"
    $menuScript.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $srv = [string]$node.Tag.Server
            $db = [string]$node.Tag.Database
            $schema = [string]$node.Tag.Schema
            $table = [string]$node.Tag.Table
            Write-DzDebug "`t[DEBUG][TreeView] Context CREATE: Server='$srv' DB='$db' [$schema].[$table]"
            if ([string]::IsNullOrWhiteSpace($srv)) { Ui-Error "TreeView: Server vacío en el Tag. Reconecta a la BD para recargar el explorador." $global:MainWindow; return }
            if ([string]::IsNullOrWhiteSpace($db)) { Ui-Error "TreeView: Database vacía en el Tag." $global:MainWindow; return }
            if ($node.Tag.OnDatabaseSelected) { & $node.Tag.OnDatabaseSelected $db }
            $scriptText = Get-CreateTableScript -Server $srv -Database $db -Schema $schema -Table $table -Credential $node.Tag.Credential
            if ($scriptText -and $node.Tag.InsertTextHandler) { & $node.Tag.InsertTextHandler $scriptText }
        })
    $menuCopy = New-Object System.Windows.Controls.MenuItem
    $menuCopy.Header = "Copiar nombre"
    $menuCopy.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $schema = [string]$node.Tag.Schema
            $table = [string]$node.Tag.Table
            Write-DzDebug "`t[DEBUG][TreeView] Context COPY: [$schema].[$table]"
            $name = "[$schema].[$table]"
            [System.Windows.Clipboard]::SetText($name)
        })
    [void]$menu.Items.Add($menuSelectTop)
    [void]$menu.Items.Add($menuSelectAll)
    [void]$menu.Items.Add($menuScript)
    [void]$menu.Items.Add($menuCopy)
    $TableNode.ContextMenu = $menu
}

# AGREGAR esta función al módulo SqlTreeView.psm1

function Add-DatabaseContextMenu {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$DatabaseNode)

    $menu = New-Object System.Windows.Controls.ContextMenu

    # Opción: Eliminar base de datos
    $menuDelete = New-Object System.Windows.Controls.MenuItem
    $menuDelete.Header = "🗑️ Eliminar base de datos..."
    $menuDelete.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }

            $dbName = [string]$node.Tag.Database
            $server = [string]$node.Tag.Server
            $credential = $node.Tag.Credential

            Write-DzDebug "`t[DEBUG][TreeView] Context DELETE DB: Server='$server' DB='$dbName'"

            # Mostrar diálogo de confirmación con opciones
            Show-DeleteDatabaseDialog -Server $server -Database $dbName -Credential $credential -ParentNode $node
        })

    # Separador
    $separator = New-Object System.Windows.Controls.Separator

    # Opción: Refresh
    $menuRefresh = New-Object System.Windows.Controls.MenuItem
    $menuRefresh.Header = "🔄 Actualizar"
    $menuRefresh.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }

            Write-DzDebug "`t[DEBUG][TreeView] Context REFRESH DB: $($node.Tag.Database)"

            # Limpiar y recargar los hijos
            $node.Items.Clear()

            $db = $node.Tag.Database
            $server = $node.Tag.Server
            $credential = $node.Tag.Credential
            $insertHandler = $node.Tag.InsertTextHandler
            $dbSelectedHandler = $node.Tag.OnDatabaseSelected

            # Recrear nodos de tablas, vistas, procedimientos
            $tablesTag = @{Type = "TablesRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }
            $viewsTag = @{Type = "ViewsRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }
            $procsTag = @{Type = "ProceduresRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }

            $tablesNode = New-SqlTreeNode -Header "📋 Tablas" -Tag $tablesTag -HasPlaceholder $true
            $viewsNode = New-SqlTreeNode -Header "👁️ Vistas" -Tag $viewsTag -HasPlaceholder $true
            $procsNode = New-SqlTreeNode -Header "⚙️ Procedimientos" -Tag $procsTag -HasPlaceholder $true

            $tablesNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })
            $viewsNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })
            $procsNode.Add_Expanded({ param($s, $e) Load-TablesIntoNode -RootNode $s })

            [void]$node.Items.Add($tablesNode)
            [void]$node.Items.Add($viewsNode)
            [void]$node.Items.Add($procsNode)
        })

    [void]$menu.Items.Add($menuDelete)
    [void]$menu.Items.Add($separator)
    [void]$menu.Items.Add($menuRefresh)

    $DatabaseNode.ContextMenu = $menu
}

# REEMPLAZAR la función Show-DeleteDatabaseDialog completa

function Show-DeleteDatabaseDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)]$ParentNode
    )

    function Ui-Info([string]$m, [string]$t = "Información", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Information" -Owner $o | Out-Null }
    function Ui-Error([string]$m, [string]$t = "Error", [System.Windows.Window]$o) { Show-WpfMessageBoxSafe -Message $m -Title $t -Buttons "OK" -Icon "Error" -Owner $o | Out-Null }

    Write-DzDebug "`t[DEBUG][DeleteDB] INICIO: Server='$Server' Database='$Database'"

    Add-Type -AssemblyName PresentationFramework

    $theme = Get-DzUiTheme

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Eliminar Base de Datos"
        Height="280"
        Width="500"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" FontWeight="Bold" FontSize="14" Margin="0,0,0,15">
            ⚠️ Advertencia: Eliminar Base de Datos
        </TextBlock>

        <TextBlock Grid.Row="1" TextWrapping="Wrap" Margin="0,0,0,15">
            Estás a punto de eliminar permanentemente la base de datos:
        </TextBlock>

        <TextBlock Grid.Row="2" FontWeight="Bold" FontSize="13" Margin="10,0,0,20"
                   Foreground="$($theme.AccentPrimary)">
            🗄️ $Database
        </TextBlock>

        <StackPanel Grid.Row="3" Margin="0,0,0,15">
            <CheckBox x:Name="chkDeleteBackupHistory" IsChecked="True" Margin="0,0,0,8">
                <TextBlock Text="Eliminar historial de backup y restore" TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="chkCloseConnections" IsChecked="True">
                <TextBlock Text="Cerrar conexiones existentes" TextWrapping="Wrap"/>
            </CheckBox>
        </StackPanel>

        <TextBlock Grid.Row="4" FontSize="11" Foreground="Gray" TextWrapping="Wrap" VerticalAlignment="Center">
            Esta acción es irreversible. Asegúrate de tener un respaldo antes de continuar.
        </TextBlock>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button x:Name="btnEliminar" Content="Eliminar" Width="100" Height="30" Margin="5,0"
                    Style="{StaticResource SystemButtonStyle}"/>
            <Button x:Name="btnCancelar" Content="Cancelar" Width="100" Height="30" Margin="5,0"
                    Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    if (-not $window) {
        Write-DzDebug "`t[DEBUG][DeleteDB] ERROR: window=NULL"
        throw "No se pudo crear la ventana (XAML)."
    }

    $chkDeleteBackupHistory = $window.FindName("chkDeleteBackupHistory")
    $chkCloseConnections = $window.FindName("chkCloseConnections")
    $btnEliminar = $window.FindName("btnEliminar")
    $btnCancelar = $window.FindName("btnCancelar")

    # Botón Eliminar
    $btnEliminar.Add_Click({
            Write-DzDebug "`t[DEBUG][DeleteDB] btnEliminar Click"

            $deleteBackupHistory = $chkDeleteBackupHistory.IsChecked -eq $true
            $closeConnections = $chkCloseConnections.IsChecked -eq $true

            try {
                # SOLUCIÓN: Ejecutar comandos en LOTES SEPARADOS (sin GO)
                # Los comandos se ejecutan uno por uno

                $escapedDb = $Database -replace "'", "''"
                $safeName = $Database -replace ']', ']]'

                # 1. Cerrar conexiones existentes si está marcado
                if ($closeConnections) {
                    Write-DzDebug "`t[DEBUG][DeleteDB] Paso 1: Cerrando conexiones"
                    $closeQuery = "ALTER DATABASE [$safeName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
                    $result1 = Invoke-SqlQuery -Server $Server -Database "master" -Query $closeQuery -Credential $Credential

                    if (-not $result1.Success) {
                        Write-DzDebug "`t[DEBUG][DeleteDB] Error cerrando conexiones: $($result1.ErrorMessage)"
                        Ui-Error "Error al cerrar conexiones existentes:`n`n$($result1.ErrorMessage)" "Error" $window
                        return
                    }
                    Write-DzDebug "`t[DEBUG][DeleteDB] Conexiones cerradas OK"
                }

                # 2. Eliminar la base de datos
                Write-DzDebug "`t[DEBUG][DeleteDB] Paso 2: Eliminando base de datos"
                $dropQuery = "DROP DATABASE [$safeName]"
                $result2 = Invoke-SqlQuery -Server $Server -Database "master" -Query $dropQuery -Credential $Credential

                if (-not $result2.Success) {
                    Write-DzDebug "`t[DEBUG][DeleteDB] Error eliminando DB: $($result2.ErrorMessage)"
                    Ui-Error "Error al eliminar la base de datos:`n`n$($result2.ErrorMessage)" "Error" $window
                    return
                }
                Write-DzDebug "`t[DEBUG][DeleteDB] Base de datos eliminada OK"

                # 3. Eliminar historial de backup/restore si está marcado
                if ($deleteBackupHistory) {
                    Write-DzDebug "`t[DEBUG][DeleteDB] Paso 3: Eliminando historial de backup"
                    $historyQuery = "EXEC msdb.dbo.sp_delete_database_backuphistory @database_name = N'$escapedDb'"
                    $result3 = Invoke-SqlQuery -Server $Server -Database "master" -Query $historyQuery -Credential $Credential

                    # El historial puede no existir, así que no es crítico si falla
                    if ($result3.Success) {
                        Write-DzDebug "`t[DEBUG][DeleteDB] Historial eliminado OK"
                    } else {
                        Write-DzDebug "`t[DEBUG][DeleteDB] Historial no eliminado (puede no existir): $($result3.ErrorMessage)"
                    }
                }

                # Todo salió bien
                Write-DzDebug "`t[DEBUG][DeleteDB] Base de datos eliminada exitosamente"
                Ui-Info "La base de datos '$Database' ha sido eliminada exitosamente." "✓ Éxito" $window

                # Remover el nodo del TreeView
                if ($ParentNode.Parent -is [System.Windows.Controls.ItemsControl]) {
                    $window.Dispatcher.Invoke([action] {
                            try {
                                [void]$ParentNode.Parent.Items.Remove($ParentNode)
                                Write-DzDebug "`t[DEBUG][DeleteDB] Nodo removido del TreeView"
                            } catch {
                                Write-DzDebug "`t[DEBUG][DeleteDB] Error removiendo nodo: $($_.Exception.Message)"
                            }
                        })
                }

                $window.DialogResult = $true
                $window.Close()

            } catch {
                Write-DzDebug "`t[DEBUG][DeleteDB] Excepción: $($_.Exception.Message)"
                Ui-Error "Error inesperado al eliminar la base de datos:`n`n$($_.Exception.Message)" "Error" $window
            }
        })

    # Botón Cancelar
    $btnCancelar.Add_Click({
            Write-DzDebug "`t[DEBUG][DeleteDB] btnCancelar Click"
            $window.DialogResult = $false
            $window.Close()
        })

    Write-DzDebug "`t[DEBUG][DeleteDB] Mostrando ventana"
    $null = $window.ShowDialog()
}

Export-ModuleMember -Function @("Initialize-SqlTreeView", "Load-DatabasesIntoTree", "Load-TablesIntoNode", "Load-ColumnsIntoTableNode", "Add-TreeNodeContextMenu", "Add-DatabaseContextMenu", "Show-DeleteDatabaseDialog")