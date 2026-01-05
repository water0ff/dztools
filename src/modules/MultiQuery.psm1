#requires -Version 5.0

$script:queryTabCounter = 1

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
        try {
            $caret.InsertTextInRun($Text)
        } catch {
            $rtb.CaretPosition = $rtb.Document.ContentEnd
            $rtb.CaretPosition.InsertTextInRun($Text)
        }
    }
    $rtb.Focus()
}

function Clear-ActiveQueryTab {
    param([Parameter(Mandatory = $true)]$TabControl)
    $rtb = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if ($rtb) {
        $rtb.Document.Blocks.Clear()
        $tab = Get-ActiveQueryTab -TabControl $TabControl
        if ($tab -and $tab.Tag) { $tab.Tag.IsDirty = $false; Update-QueryTabHeader -TabItem $tab }
    }
}

function Update-QueryTabHeader {
    param([Parameter(Mandatory = $true)]$TabItem)
    if (-not $TabItem.Tag) { return }
    $title = $TabItem.Tag.Title
    if ($TabItem.Tag.IsDirty) { $title = "*$title" }
    if ($TabItem.Tag.HeaderTextBlock) { $TabItem.Tag.HeaderTextBlock.Text = $title }
}

function New-QueryTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TabControl
    )

    $tabNumber = $script:queryTabCounter
    $script:queryTabCounter++
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

    $tabItem.Tag = @{ Type = "QueryTab"; RichTextBox = $rtb; Title = $tabTitle; HeaderTextBlock = $headerText; IsDirty = $false }

    $rtb.Add_TextChanged({
            $tabItem.Tag.IsDirty = $true
            Update-QueryTabHeader -TabItem $tabItem
        })

    $closeButton.Add_Click({
            Close-QueryTab -TabControl $TabControl -TabItem $tabItem
        })

    $insertIndex = $TabControl.Items.Count
    for ($i = 0; $i -lt $TabControl.Items.Count; $i++) {
        $item = $TabControl.Items[$i]
        if ($item -is [System.Windows.Controls.TabItem] -and $item.Name -eq "tabAddQuery") {
            $insertIndex = $i
            break
        }
    }

    $TabControl.Items.Insert($insertIndex, $tabItem) | Out-Null
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

    $TabControl.Items.Remove($TabItem)
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

    $result = Invoke-SqlQueryMultiResultSet -Server $Server -Database $Database -Query $cleanQuery -Credential $Credential
    if (-not $result.Success) { throw $result.ErrorMessage }

    if ($result.ResultSets -and $result.ResultSets.Count -gt 0) {
        Show-MultipleResultSets -TabControl $ResultsTabControl -ResultSets $result.ResultSets
    } elseif ($result.ContainsKey('RowsAffected') -and $result.RowsAffected -ne $null) {
        $ResultsTabControl.Items.Clear()
        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = \"Resultado\"
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = \"Filas afectadas: $($result.RowsAffected)\"
        $text.Margin = \"10\"
        $tab.Content = $text
        [void]$ResultsTabControl.Items.Add($tab)
    } else {
        Show-MultipleResultSets -TabControl $ResultsTabControl -ResultSets @()
    }

    if ($result.Messages -and $result.Messages.Count -gt 0) {
        Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
        $result.Messages | ForEach-Object { Write-Host $_ }
    }

    return $result
}

function Show-MultipleResultSets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)][array]$ResultSets
    )

    $TabControl.Items.Clear()

    if (-not $ResultSets -or $ResultSets.Count -eq 0) {
        $tab = New-Object System.Windows.Controls.TabItem
        $tab.Header = "Resultado"
        $text = New-Object System.Windows.Controls.TextBlock
        $text.Text = "La consulta no devolvió resultados."
        $text.Margin = "10"
        $tab.Content = $text
        [void]$TabControl.Items.Add($tab)
        return
    }

    $i = 0
    foreach ($rs in $ResultSets) {
        $i++
        $tab = New-Object System.Windows.Controls.TabItem
        $rowCount = if ($rs.RowCount -ne $null) { $rs.RowCount } else { $rs.DataTable.Rows.Count }
        $tab.Header = "Resultado $i ($rowCount filas)"

        $dg = New-Object System.Windows.Controls.DataGrid
        $dg.ItemsSource = $rs.DataTable.DefaultView
        $dg.IsReadOnly = $true
        $dg.AutoGenerateColumns = $true
        $dg.CanUserAddRows = $false
        $dg.CanUserDeleteRows = $false
        $dg.SelectionMode = "Extended"

        $tab.Content = $dg
        [void]$TabControl.Items.Add($tab)
    }
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

Export-ModuleMember -Function @(
    'New-QueryTab',
    'Close-QueryTab',
    'Execute-QueryInTab',
    'Show-MultipleResultSets',
    'Export-ResultSetToCsv',
    'Get-ActiveQueryRichTextBox',
    'Set-QueryTextInActiveTab',
    'Insert-TextIntoActiveQuery',
    'Clear-ActiveQueryTab'
)
