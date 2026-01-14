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
    $node
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
                    if ($e.ButtonState -eq [System.Windows.Input.MouseButtonState]::Pressed) {
                        try { $w.DragMove() } catch {}
                    }
                })
        }
        $script:renameResult = $null
        $w.Add_Loaded({
                $txtInput.Focus()
                $txtInput.SelectAll()
            })
        $btnClose.Add_Click({
                $script:renameResult = $null
                $w.DialogResult = $false
                $w.Close()
            })
        $btnCancel.Add_Click({
                $script:renameResult = $null
                $w.DialogResult = $false
                $w.Close()
            })
        $w.Add_PreviewKeyDown({
                param($sender, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                    $script:renameResult = $null
                    $w.DialogResult = $false
                    $w.Close()
                }
            })
        $btnOk.Add_Click({
                $script:renameResult = $txtInput.Text
                $w.DialogResult = $true
                $w.Close()
            })
        $result = $w.ShowDialog()
        if ($result -eq $true) { return $script:renameResult }
        return $null
    } catch {
        Write-Error "Error creando diálogo de renombrar: $($_.Exception.Message)"
        return $null
    }
}
function Initialize-SqlTreeView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TreeView,
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password,
        [Parameter(Mandatory = $false)][scriptblock]$InsertTextHandler,
        [Parameter(Mandatory = $false)][scriptblock]$OnDatabaseSelected,
        [Parameter(Mandatory = $false)][scriptblock]$GetCurrentDatabase,
        [Parameter(Mandatory = $false)][bool]$AutoExpand = $true
    )
    $TreeView.Items.Clear()
    $serverTag = @{
        Type               = "Server"
        Server             = $Server
        Credential         = $Credential
        User               = $User
        Password           = $Password
        InsertTextHandler  = $InsertTextHandler
        OnDatabaseSelected = $OnDatabaseSelected
        GetCurrentDatabase = $GetCurrentDatabase
        Loaded             = $false
    }
    $serverNode = New-SqlTreeNode -Header "📊 $Server" -Tag $serverTag -HasPlaceholder $true
    Add-ServerContextMenu -ServerNode $serverNode
    $serverNode.Add_Expanded({
            param($sender, $e)
            Write-DzDebug "`t[DEBUG][TreeView] Expand Server: $($sender.Tag.Server) Loaded=$($sender.Tag.Loaded)"
            if (-not $sender.Tag.Loaded) {
                $sender.Tag.Loaded = $true
                Load-DatabasesIntoTree -ServerNode $sender
            }
        })
    [void]$TreeView.Items.Add($serverNode)
    if (-not $TreeView.Tag -or -not $TreeView.Tag.TreeViewKeyHandlerAdded) {
        $TreeView.Add_KeyDown({
                param($sender, $e)
                if ($e.Key -eq 'F5') {
                    $selected = $sender.SelectedItem
                    if (-not $selected -or -not $selected.Tag -or $selected.Tag.Type -ne 'Server') { return }
                    Write-DzDebug "`t[DEBUG][TreeView] F5 Refresh Server: $($selected.Tag.Server)"
                    Refresh-SqlTreeServerNode -ServerNode $selected
                    $e.Handled = $true
                    return
                }
                if ($e.Key -eq 'F2') {
                    $selected = $sender.SelectedItem
                    if (-not $selected -or -not $selected.Tag -or $selected.Tag.Type -ne 'Database') { return }
                    Write-DzDebug "`t[DEBUG][TreeView] F2 Rename Database: $($selected.Tag.Database)"
                    $dbName = [string]$selected.Tag.Database
                    $server = [string]$selected.Tag.Server
                    $credential = $selected.Tag.Credential
                    $newName = bdd_RenameFromTree -Title "Renombrar base de datos" `
                        -Prompt "Nuevo nombre para la base de datos:" `
                        -DefaultValue $dbName
                    if ($null -eq $newName) {
                        Write-DzDebug "`t[DEBUG][TreeView] Rename cancelado"
                        $e.Handled = $true
                        return
                    }
                    $newName = $newName.Trim()
                    if ([string]::IsNullOrWhiteSpace($newName)) {
                        Ui-Error "El nombre no puede estar vacío." $global:MainWindow
                        $e.Handled = $true
                        return
                    }
                    if ($newName -eq $dbName) {
                        Write-DzDebug "`t[DEBUG][TreeView] Rename sin cambios"
                        $e.Handled = $true
                        return
                    }
                    Write-DzDebug "`t[DEBUG][TreeView] F2 RENAME DB: Server='$server' Old='$dbName' New='$newName'"
                    $safeOld = $dbName -replace ']', ']]'
                    $safeNew = $newName -replace ']', ']]'
                    $query = @"
ALTER DATABASE [$safeOld] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
ALTER DATABASE [$safeOld] MODIFY NAME = [$safeNew];
ALTER DATABASE [$safeNew] SET MULTI_USER;
"@
                    $result = Invoke-SqlQuery -Server $server -Database "master" -Query $query -Credential $credential
                    if (-not $result.Success) {
                        Ui-Error "Error al renombrar la base de datos:`n`n$($result.ErrorMessage)" $global:MainWindow
                        $e.Handled = $true
                        return
                    }
                    Refresh-SqlTreeView -TreeView $global:tvDatabases -Server $server
                    $e.Handled = $true
                    return
                }
            })
        $TreeView.Tag = [pscustomobject]@{ TreeViewKeyHandlerAdded = $true }
    }
    if ($AutoExpand) {
        Write-DzDebug "`t[DEBUG][TreeView] AutoExpand Server: $Server"
        $serverNode.Tag.Loaded = $true
        Load-DatabasesIntoTree -ServerNode $serverNode
        $serverNode.IsExpanded = $true
    }
}
function Refresh-SqlTreeServerNode {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$ServerNode)
    if (-not $ServerNode -or -not $ServerNode.Tag) { return }
    Write-DzDebug "`t[DEBUG][TreeView] Refresh Server: $($ServerNode.Tag.Server)"
    $ServerNode.Tag.Databases = $null
    $ServerNode.Tag.Loaded = $true
    $ServerNode.Items.Clear()
    Load-DatabasesIntoTree -ServerNode $ServerNode
    $ServerNode.IsExpanded = $true
}
function Refresh-SqlTreeView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TreeView,
        [Parameter(Mandatory = $true)][string]$Server
    )
    if (-not $TreeView) { return }
    Write-DzDebug "`t[DEBUG][TreeView] Refresh TreeView Server='$Server'"
    $targetNode = $null
    foreach ($item in $TreeView.Items) {
        if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Tag -and $item.Tag.Type -eq "Server" -and $item.Tag.Server -eq $Server) {
            $targetNode = $item
            break
        }
    }
    if ($targetNode) { Refresh-SqlTreeServerNode -ServerNode $targetNode }
}
function Load-DatabasesIntoTree {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$ServerNode)
    $ServerNode.Items.Clear()
    $server = [string]$ServerNode.Tag.Server
    $credential = $ServerNode.Tag.Credential
    Write-DzDebug "`t[DEBUG][TreeView] Load DBs: Server='$server'"
    try {
        $databases = $ServerNode.Tag.Databases
        if (-not $databases -or $databases.Count -eq 0) {
            Write-DzDebug "`t[DEBUG][TreeView] No hay Databases precargadas, consultando a SQL..."
            try {
                $databases = Get-SqlDatabases -Server $server -Credential $credential
                $ServerNode.Tag.Databases = $databases
            } catch {
                $errorNode = New-SqlTreeNode -Header "Error cargando bases" -Tag @{ Type = "Error" } -HasPlaceholder $false
                [void]$ServerNode.Items.Add($errorNode)
                return
            }
        } else {
            Write-DzDebug "`t[DEBUG][TreeView] Usando Databases precargadas: $($databases.Count)"
        }
    } catch {
        $errorNode = New-SqlTreeNode -Header "Error cargando bases" -Tag @{ Type = "Error" } -HasPlaceholder $false
        [void]$ServerNode.Items.Add($errorNode)
        return
    }
    foreach ($db in $databases) {
        $dbTag = @{
            Type               = "Database"
            Database           = $db
            Server             = $server
            Credential         = $credential
            User               = $ServerNode.Tag.User
            Password           = $ServerNode.Tag.Password
            InsertTextHandler  = $ServerNode.Tag.InsertTextHandler
            OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected
        }
        $dbNode = New-SqlTreeNode -Header "🗄️ $db" -Tag $dbTag -HasPlaceholder $false
        $dbNode.Add_MouseDoubleClick({
                param($s, $e)
                Write-DzDebug "`t[DEBUG][TreeView] Doble clic DB: $($s.Tag.Database) | HasHandler=$([bool]$s.Tag.OnDatabaseSelected)"
                if ($s.Tag.OnDatabaseSelected) { & $s.Tag.OnDatabaseSelected $s.Tag.Database }
                $e.Handled = $true
            })
        Add-DatabaseContextMenu -DatabaseNode $dbNode
        $tablesTag = @{
            Type               = "TablesRoot"
            Database           = $db
            Server             = $server
            Credential         = $credential
            InsertTextHandler  = $ServerNode.Tag.InsertTextHandler
            OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected
            Loaded             = $false
        }
        $viewsTag = @{
            Type               = "ViewsRoot"
            Database           = $db
            Server             = $server
            Credential         = $credential
            InsertTextHandler  = $ServerNode.Tag.InsertTextHandler
            OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected
            Loaded             = $false
        }
        $procsTag = @{
            Type               = "ProceduresRoot"
            Database           = $db
            Server             = $server
            Credential         = $credential
            InsertTextHandler  = $ServerNode.Tag.InsertTextHandler
            OnDatabaseSelected = $ServerNode.Tag.OnDatabaseSelected
            Loaded             = $false
        }
        $tablesNode = New-SqlTreeNode -Header "📋 Tablas" -Tag $tablesTag -HasPlaceholder $true
        $viewsNode = New-SqlTreeNode -Header "👁️ Vistas" -Tag $viewsTag -HasPlaceholder $true
        $procsNode = New-SqlTreeNode -Header "⚙️ Procedimientos" -Tag $procsTag -HasPlaceholder $true
        $tablesNode.Add_Expanded({ param($s, $e) Write-DzDebug "`t[DEBUG][TreeView] Expand TablasRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
        $viewsNode.Add_Expanded({ param($s, $e) Write-DzDebug "`t[DEBUG][TreeView] Expand VistasRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
        $procsNode.Add_Expanded({ param($s, $e) Write-DzDebug "`t[DEBUG][TreeView] Expand ProcsRoot DB: $($s.Tag.Database)"; Load-TablesIntoNode -RootNode $s })
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
                $tag = @{
                    Type               = "Table"
                    Database           = $database
                    Schema             = $schema
                    Table              = $table
                    Server             = $server
                    Credential         = $credential
                    InsertTextHandler  = $RootNode.Tag.InsertTextHandler
                    OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected
                    Loaded             = $false
                }
                $tableNode = New-SqlTreeNode -Header "📋 [$schema].[$table]" -Tag $tag -HasPlaceholder $true
                $tableNode.Add_Expanded({ param($s, $e) Write-DzDebug "`t[DEBUG][TreeView] Expand Table: $($s.Tag.Database) [$($s.Tag.Schema)].[$($s.Tag.Table)]"; Load-ColumnsIntoTableNode -TableNode $s })
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
                $tag = @{
                    Type               = "View"
                    Database           = $database
                    Schema             = $schema
                    View               = $view
                    Server             = $server
                    Credential         = $credential
                    InsertTextHandler  = $RootNode.Tag.InsertTextHandler
                    OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected
                }
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
                $tag = @{
                    Type               = "Procedure"
                    Database           = $database
                    Schema             = $schema
                    Procedure          = $proc
                    Server             = $server
                    Credential         = $credential
                    InsertTextHandler  = $RootNode.Tag.InsertTextHandler
                    OnDatabaseSelected = $RootNode.Tag.OnDatabaseSelected
                }
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
        $label = if ($isPk) { "🔑 $name ($type)" } else { "• $name ($type)" }
        $tag = @{ Type = "Column"; Column = $name; InsertTextHandler = $TableNode.Tag.InsertTextHandler }
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
        $nullable = ($row.IS_NULLABLE -eq "YES")
        $typeSuffix = ""
        if ($dataType -in @("char", "varchar", "nchar", "nvarchar", "binary", "varbinary")) {
            if ($len -eq -1) { $typeSuffix = "(MAX)" } else { $typeSuffix = "($len)" }
        } elseif ($dataType -in @("decimal", "numeric")) {
            $typeSuffix = "($precision,$scale)"
        }
        $nullText = if ($nullable) { "NULL" } else { "NOT NULL" }
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
function Add-DatabaseContextMenu {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$DatabaseNode)
    if ($DatabaseNode -and $DatabaseNode.Tag) {
        Write-DzDebug "`t[DEBUG][TreeView] Init DB ContextMenu: $($DatabaseNode.Tag.Database)"
    }
    $menu = New-Object System.Windows.Controls.ContextMenu
    $menuBackup = New-Object System.Windows.Controls.MenuItem
    $menuBackup.Header = "💾 Respaldar..."
    $menuBackup.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $dbName = [string]$node.Tag.Database
            $server = [string]$node.Tag.Server
            $user = [string]$node.Tag.User
            $password = [string]$node.Tag.Password
            Write-DzDebug "`t[DEBUG][TreeView] Context BACKUP DB: Server='$server' DB='$dbName'"
            if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($password)) {
                Ui-Error "No se encontraron credenciales para respaldar la base de datos." $global:MainWindow
                return
            }
            Show-BackupDialog -Server $server -User $user -Password $password -Database $dbName
        })
    $menuRename = New-Object System.Windows.Controls.MenuItem
    $menuRename.Header = "✏️ Renombrar..."
    $menuRename.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $dbName = [string]$node.Tag.Database
            $server = [string]$node.Tag.Server
            $credential = $node.Tag.Credential
            $newName = bdd_RenameFromTree -Title "Renombrar base de datos" `
                -Prompt "Nuevo nombre para la base de datos:" `
                -DefaultValue $dbName
            if ($null -eq $newName) {
                Write-DzDebug "`t[DEBUG][TreeView] Rename cancelado"
                return
            }
            $newName = $newName.Trim()
            if ([string]::IsNullOrWhiteSpace($newName)) {
                Ui-Error "El nombre no puede estar vacío." $global:MainWindow
                return
            }
            if ($newName -eq $dbName) {
                Write-DzDebug "`t[DEBUG][TreeView] Rename sin cambios"
                return
            }
            Write-DzDebug "`t[DEBUG][TreeView] Context RENAME DB: Server='$server' Old='$dbName' New='$newName'"
            $safeOld = $dbName -replace ']', ']]'
            $safeNew = $newName -replace ']', ']]'
            $query = @"
ALTER DATABASE [$safeOld] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
ALTER DATABASE [$safeOld] MODIFY NAME = [$safeNew];
ALTER DATABASE [$safeNew] SET MULTI_USER;
"@
            $result = Invoke-SqlQuery -Server $server -Database "master" -Query $query -Credential $credential
            if (-not $result.Success) {
                Ui-Error "Error al renombrar la base de datos:`n`n$($result.ErrorMessage)" $global:MainWindow
                return
            }
            Refresh-SqlTreeView -TreeView $global:tvDatabases -Server $server
        })
    $menuNewQuery = New-Object System.Windows.Controls.MenuItem
    $menuNewQuery.Header = "🧾 Nuevo Query"
    $menuNewQuery.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $dbName = [string]$node.Tag.Database
            Write-DzDebug "`t[DEBUG][TreeView] Context NEW QUERY: DB='$dbName'"
            if ($node.Tag.OnDatabaseSelected) { & $node.Tag.OnDatabaseSelected $dbName }
            if (-not $global:tcQueries) { return }
            $tab = New-QueryTab -TabControl $global:tcQueries
            $safeDb = $dbName -replace ']', ']]'
            Set-QueryTextInActiveTab -TabControl $global:tcQueries -Text "USE [$safeDb]`n`n"
        })
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
            Show-DeleteDatabaseDialog -Server $server -Database $dbName -Credential $credential -ParentNode $node
        })
    $separator1 = New-Object System.Windows.Controls.Separator
    $separator2 = New-Object System.Windows.Controls.Separator
    $menuRefresh = New-Object System.Windows.Controls.MenuItem
    $menuRefresh.Header = "🔄 Actualizar"
    $menuRefresh.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            Write-DzDebug "`t[DEBUG][TreeView] Context REFRESH DB: $($node.Tag.Database)"
            $node.Items.Clear()
            $db = $node.Tag.Database
            $server = $node.Tag.Server
            $credential = $node.Tag.Credential
            $insertHandler = $node.Tag.InsertTextHandler
            $dbSelectedHandler = $node.Tag.OnDatabaseSelected
            $tablesTag = @{ Type = "TablesRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }
            $viewsTag = @{ Type = "ViewsRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }
            $procsTag = @{ Type = "ProceduresRoot"; Database = $db; Server = $server; Credential = $credential; InsertTextHandler = $insertHandler; OnDatabaseSelected = $dbSelectedHandler; Loaded = $false }
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
    [void]$menu.Items.Add($menuBackup)
    [void]$menu.Items.Add($menuRename)
    [void]$menu.Items.Add($menuNewQuery)
    [void]$menu.Items.Add($separator1)
    [void]$menu.Items.Add($menuDelete)
    [void]$menu.Items.Add($separator2)
    [void]$menu.Items.Add($menuRefresh)
    $DatabaseNode.ContextMenu = $menu
}
function Add-ServerContextMenu {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$ServerNode)
    if ($ServerNode -and $ServerNode.Tag) { Write-DzDebug "`t[DEBUG][TreeView] Init Server ContextMenu: $($ServerNode.Tag.Server)" }
    $menu = New-Object System.Windows.Controls.ContextMenu
    $menuRefresh = New-Object System.Windows.Controls.MenuItem
    $menuRefresh.Header = "🔄 Actualizar"
    $menuRefresh.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            Write-DzDebug "`t[DEBUG][TreeView] Context REFRESH Server: $($node.Tag.Server)"
            Refresh-SqlTreeServerNode -ServerNode $node
        })
    $menuRestore = New-Object System.Windows.Controls.MenuItem
    $menuRestore.Header = "♻️ Restaurar..."
    $serverNodeRef = $ServerNode
    $menuRestore.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $server = [string]$node.Tag.Server
            $user = [string]$node.Tag.User
            $password = [string]$node.Tag.Password
            $defaultDb = "master"
            if ($node.Tag.GetCurrentDatabase) { $defaultDb = & $node.Tag.GetCurrentDatabase }
            Write-DzDebug "`t[DEBUG][TreeView] Context RESTORE Server: $server DefaultDB='$defaultDb'"
            if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($password)) {
                Ui-Error "No se encontraron credenciales para restaurar." $global:MainWindow
                return
            }
            Show-RestoreDialog -Server $server -User $user -Password $password -Database $defaultDb -OnRestoreCompleted {
                param($dbName)
                Write-DzDebug "`t[DEBUG][TreeView] Restore completed. Refresh Server: $server"
                Refresh-SqlTreeServerNode -ServerNode $serverNodeRef
            }
        }.GetNewClosure())
    $menuAttach = New-Object System.Windows.Controls.MenuItem
    $menuAttach.Header = "📎 Adjuntar..."
    $menuAttach.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $server = [string]$node.Tag.Server
            $user = [string]$node.Tag.User
            $password = [string]$node.Tag.Password
            if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($password)) {
                Ui-Error "No se encontraron credenciales para adjuntar." $global:MainWindow
                return
            }
            $modulesPath = Join-Path $PSScriptRoot "modules"
            Show-AttachDialog -Server $server -User $user -Password $password -Database "master" `
                -ModulesPath $modulesPath `
                -OnAttachCompleted {
                param($dbName)
                Write-DzDebug "`t[DEBUG][TreeView] Attach completed. Refresh Server: $server"
                Refresh-SqlTreeServerNode -ServerNode $serverNodeRef
            }
        }.GetNewClosure())
    $menuCreateDb = New-Object System.Windows.Controls.MenuItem
    $menuCreateDb.Header = "🆕 Crear base de datos..."
    $menuCreateDb.Add_Click({
            $cm = $this.Parent
            $node = $null
            if ($cm -is [System.Windows.Controls.ContextMenu]) { $node = $cm.PlacementTarget }
            if ($null -eq $node -or $null -eq $node.Tag) { return }
            $server = [string]$node.Tag.Server
            $credential = $node.Tag.Credential
            $dbName = New-WpfInputDialog -Title "Crear base de datos" -Prompt "Nombre de la nueva base de datos:" -DefaultValue ""
            if ($null -eq $dbName) { Write-DzDebug "`t[DEBUG][TreeView] Create DB cancelado"; return }
            $dbName = $dbName.Trim()
            if ([string]::IsNullOrWhiteSpace($dbName)) { Ui-Error "El nombre no puede estar vacío." $global:MainWindow; return }
            Write-DzDebug "`t[DEBUG][TreeView] Context CREATE DB: Server='$server' DB='$dbName'"
            $safeName = $dbName -replace ']', ']]'
            $query = "CREATE DATABASE [$safeName]"
            $result = Invoke-SqlQuery -Server $server -Database "master" -Query $query -Credential $credential
            if (-not $result.Success) {
                Ui-Error "Error al crear la base de datos:`n`n$($result.ErrorMessage)" $global:MainWindow
                return
            }
            Refresh-SqlTreeServerNode -ServerNode $node
        })
    [void]$menu.Items.Add($menuRefresh)
    [void]$menu.Items.Add($menuRestore)
    [void]$menu.Items.Add($menuAttach)
    [void]$menu.Items.Add($menuCreateDb)
    $ServerNode.ContextMenu = $menu
}
function Show-DeleteDatabaseDialog {
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
    Write-DzDebug "`t[DEBUG][DeleteDB] INICIO: Server='$Server' Database='$Database'"
    Add-Type -AssemblyName PresentationFramework
    $safeDb = [Security.SecurityElement]::Escape($Database)
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Eliminar Base de Datos"
        Height="360" Width="560"
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
                    <TextBlock Text="⚠️ Advertencia: Eliminar Base de Datos"
                               Foreground="{DynamicResource FormFg}"
                               FontSize="16"
                               FontWeight="SemiBold"/>
                    <TextBlock Text="Esta acción es irreversible. Verifica antes de continuar."
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
                           Text="Estás a punto de eliminar permanentemente la base de datos:"
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
                    <CheckBox x:Name="chkDeleteBackupHistory" IsChecked="True" Margin="0,0,0,8">
                        <TextBlock Text="Eliminar historial de backup y restore" TextWrapping="Wrap"/>
                    </CheckBox>
                    <CheckBox x:Name="chkCloseConnections" IsChecked="True">
                        <TextBlock Text="Cerrar conexiones existentes (SINGLE_USER + ROLLBACK IMMEDIATE)" TextWrapping="Wrap"/>
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
                        Recomendación: realiza un respaldo antes de continuar. Si la base de datos está en uso,
                        se forzará el cierre de sesiones para poder eliminarla.
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
                           Text="Enter: Eliminar   |   Esc: Cerrar"
                           VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button x:Name="btnCancelar"
                            Content="Cancelar"
                            Width="120"
                            Height="34"
                            Margin="0,0,10,0"
                            IsCancel="True"
                            Style="{StaticResource OutlineButtonStyle}"/>
                    <Button x:Name="btnEliminar"
                            Content="Eliminar"
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
        Write-DzDebug "`t[DEBUG][DeleteDB] ERROR creando ventana: $($_.Exception.Message)" -Color Red
        throw "No se pudo crear la ventana (XAML). $($_.Exception.Message)"
    }
    $chkDeleteBackupHistory = $window.FindName("chkDeleteBackupHistory")
    $chkCloseConnections = $window.FindName("chkCloseConnections")
    $btnEliminar = $window.FindName("btnEliminar")
    $btnCancelar = $window.FindName("btnCancelar")
    $btnClose = $window.FindName("btnClose")
    $btnClose.Add_Click({
            Write-DzDebug "`t[DEBUG][DeleteDB] btnClose Click"
            $window.DialogResult = $false
            $window.Close()
        })
    $btnCancelar.Add_Click({
            Write-DzDebug "`t[DEBUG][DeleteDB] btnCancelar Click"
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
    $btnEliminar.Add_Click({
            Write-DzDebug "`t[DEBUG][DeleteDB] btnEliminar Click"
            $deleteBackupHistory = $chkDeleteBackupHistory.IsChecked -eq $true
            $closeConnections = $chkCloseConnections.IsChecked -eq $true
            try {
                $escapedDb = $Database -replace "'", "''"
                $safeName = $Database -replace ']', ']]'
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
                Write-DzDebug "`t[DEBUG][DeleteDB] Paso 2: Eliminando base de datos"
                $dropQuery = "DROP DATABASE [$safeName]"
                $result2 = Invoke-SqlQuery -Server $Server -Database "master" -Query $dropQuery -Credential $Credential
                if (-not $result2.Success) {
                    Write-DzDebug "`t[DEBUG][DeleteDB] Error eliminando DB: $($result2.ErrorMessage)"
                    Ui-Error "Error al eliminar la base de datos:`n`n$($result2.ErrorMessage)" "Error" $window
                    return
                }
                Write-DzDebug "`t[DEBUG][DeleteDB] Base de datos eliminada OK"
                if ($deleteBackupHistory) {
                    Write-DzDebug "`t[DEBUG][DeleteDB] Paso 3: Eliminando historial de backup"
                    $historyQuery = "EXEC msdb.dbo.sp_delete_database_backuphistory @database_name = N'$escapedDb'"
                    $result3 = Invoke-SqlQuery -Server $Server -Database "master" -Query $historyQuery -Credential $Credential
                    if ($result3.Success) {
                        Write-DzDebug "`t[DEBUG][DeleteDB] Historial eliminado OK"
                    } else {
                        Write-DzDebug "`t[DEBUG][DeleteDB] Historial no eliminado (puede no existir): $($result3.ErrorMessage)"
                    }
                }
                Ui-Info "La base de datos '$Database' ha sido eliminada exitosamente." "✓ Éxito" $window
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
    Write-DzDebug "`t[DEBUG][DeleteDB] Mostrando ventana"
    $null = $window.ShowDialog()
}
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

    if (-not $txtMdfPath -or -not $btnBrowseMdf -or -not $txtLdfPath -or -not $btnBrowseLdf -or -not $chkRebuildLog -or -not $chkReadOnly -or -not $txtDbName -or -not $txtOwner -or -not $pbAttach -or -not $txtProgress -or -not $txtLog -or -not $btnAttach -or -not $btnClose) {
        Write-DzDebug "`t[DEBUG][Show-AttachDialog] ERROR: controles NULL"
        throw "Controles WPF incompletos (FindName devolvió NULL)."
    }

    $logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $progressQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[hashtable]'

    function Paint-Progress { param([int]$Percent, [string]$Message) $pbAttach.Value = $Percent; $txtProgress.Text = $Message }
    function Add-Log { param([string]$Message) $logQueue.Enqueue(("{0} {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message)) }
    function New-SafeCredential { param([string]$Username, [string]$PlainPassword) $secure = New-Object System.Security.SecureString; foreach ($ch in $PlainPassword.ToCharArray()) { $secure.AppendChar($ch) }; $secure.MakeReadOnly(); New-Object System.Management.Automation.PSCredential($Username, $secure) }

    function Set-LogControlsState {
        param([bool]$Enabled)
        $txtLdfPath.IsEnabled = $Enabled
        $btnBrowseLdf.IsEnabled = $Enabled
    }

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
                if (-not (Test-Path -LiteralPath $DatabaseModulePath)) {
                    throw "No se encontró el módulo requerido: $DatabaseModulePath"
                }
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

        [void]$ps.AddScript($worker).
        AddArgument($Server).
        AddArgument($AttachQuery).
        AddArgument($Credential).
        AddArgument($LogQueue).
        AddArgument($ProgressQueue).
        AddArgument($UtilitiesModulePath).
        AddArgument($DatabaseModulePath)

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

                    if ($line -notmatch '(SUCCESS_RESULT|ERROR_RESULT)') {
                        $txtLog.Text = "$line`n" + $txtLog.Text
                    }

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
                                        if ($OnAttachCompleted) {
                                            try { & $OnAttachCompleted $txtDbName.Text } catch { Write-DzDebug "`t[DEBUG][Show-AttachDialog] Error OnAttachCompleted: $($_.Exception.Message)" }
                                        }
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

                    if ([string]::IsNullOrWhiteSpace($txtDbName.Text)) {
                        $txtDbName.Text = [System.IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
                    }

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

                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $txtLdfPath.Text = $dlg.FileName
                }
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
                # Construir ruta del módulo de forma segura
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

                Start-AttachWorkAsync -Server $Server -AttachQuery $query -Credential $credential `
                    -LogQueue $logQueue -ProgressQueue $progressQueue `
                    -DatabaseModulePath $dbModulePath -UtilitiesModulePath $utilModulePath

                $script:AttachRunning = $true
            } catch {
                Write-DzDebug "`t[DEBUG][UI] ERROR btnAttach: $($_.Exception.Message)"
                Add-Log "❌ Error: $($_.Exception.Message)"
                Reset-AttachUI -ProgressText "Error inesperado"
            }
        })

    $btnClose.Add_Click({
            try { if ($logTimer -and $logTimer.IsEnabled) { $logTimer.Stop() } } catch {}
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
Export-ModuleMember -Function @(
    "bdd_RenameFromTree",
    "Initialize-SqlTreeView",
    "Refresh-SqlTreeServerNode",
    "Refresh-SqlTreeView",
    "Load-DatabasesIntoTree",
    "Load-TablesIntoNode",
    "Load-ColumnsIntoTableNode",
    "Add-TreeNodeContextMenu",
    "Add-DatabaseContextMenu",
    "Add-ServerContextMenu",
    "Show-DeleteDatabaseDialog",
    'Show-RestoreDialog',
    'Show-AttachDialog',
    'Reset-RestoreUI',
    'Reset-AttachUI')