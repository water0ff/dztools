$script:QueriesConfigPath = "C:\Temp\dztools\Queries.ini"
$script:MaxHistoryItems = 100
function Initialize-QueriesConfig {
    [CmdletBinding()]
    param()
    $configDir = Split-Path $script:QueriesConfigPath -Parent
    if (-not (Test-Path -LiteralPath $configDir)) {
        try {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
            Write-DzDebug "`t[DEBUG][Queries] Directorio creado: $configDir"
        } catch {
            Write-DzDebug "`t[DEBUG][Queries] Error creando directorio: $_" Red
            return $false
        }
    }
    if (-not (Test-Path -LiteralPath $script:QueriesConfigPath)) {
        try {
            $initialContent = @"
[QueriesHistory]
; Historial de queries ejecutadas (máximo $script:MaxHistoryItems)
; Formato: timestamp|database|query_hash|query_preview|success

[QueriesOpen]
; Estado de pestañas abiertas al cerrar
; Formato: tab_index|tab_title|query_content|is_dirty
"@
            Set-Content -LiteralPath $script:QueriesConfigPath -Value $initialContent -Encoding UTF8
            Write-DzDebug "`t[DEBUG][Queries] Archivo de configuración creado"
        } catch {
            Write-DzDebug "`t[DEBUG][Queries] Error creando archivo: $_" Red
            return $false
        }
    }
    return $true
}
function Add-QueryToHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Query,
        [Parameter(Mandatory)]
        [string]$Database,
        [Parameter()]
        [bool]$Success = $true,
        [Parameter()]
        [int]$RowsAffected = 0,
        [Parameter()]
        [string]$ErrorMessage = ""
    )
    try {
        if (-not (Initialize-QueriesConfig)) {
            Write-DzDebug "`t[DEBUG][Queries] No se pudo inicializar configuración" Yellow
            return
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $cleanQuery = ($Query -replace '\s+', ' ').Trim()
        $preview = if ($cleanQuery.Length -gt 100) { $cleanQuery.Substring(0, 97) + "..." } else { $cleanQuery }
        $hash = Get-StringHash -InputString $Query
        $resultInfo = if ($Success) { if ($RowsAffected -gt 0) { "OK ($RowsAffected filas)" } else { "OK" } } else { "ERROR: $($ErrorMessage.Substring(0, [Math]::Min(50, $ErrorMessage.Length)))" }
        $queryBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Query))
        $historyLine = "$timestamp|$Database|$hash|$preview|$resultInfo|$queryBase64"
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $lines = New-Object System.Collections.Generic.List[string]
        $historyLines = New-Object System.Collections.Generic.List[string]
        $inHistorySection = $false
        $inOpenSection = $false
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesHistory]") {
                $inHistorySection = $true
                $inOpenSection = $false
                $lines.Add($line)
                continue
            }
            if ($trimmed -eq "[QueriesOpen]") {
                if ($inHistorySection) {
                    $historyLines.Add($historyLine)
                    $startIndex = [Math]::Max(0, $historyLines.Count - $script:MaxHistoryItems)
                    for ($i = $startIndex; $i -lt $historyLines.Count; $i++) { $lines.Add($historyLines[$i]) }
                }
                $inHistorySection = $false
                $inOpenSection = $true
                $lines.Add($line)
                continue
            }
            if ($inHistorySection -and -not ($trimmed -match '^\s*;')) {
                if (-not [string]::IsNullOrWhiteSpace($trimmed)) { $historyLines.Add($line) }
                continue
            }
            $lines.Add($line)
        }
        if ($inHistorySection) {
            $historyLines.Add($historyLine)
            $startIndex = [Math]::Max(0, $historyLines.Count - $script:MaxHistoryItems)
            for ($i = $startIndex; $i -lt $historyLines.Count; $i++) { $lines.Add($historyLines[$i]) }
        }
        Set-Content -LiteralPath $script:QueriesConfigPath -Value $lines -Encoding UTF8
        Write-DzDebug "`t[DEBUG][Queries] Query agregado al historial: $preview"
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error agregando al historial: $_" Red
    }
}
function Get-QueryHistory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$MaxItems = 50
    )
    try {
        if (-not (Test-Path -LiteralPath $script:QueriesConfigPath)) {
            Write-DzDebug "`t[DEBUG][Queries] Archivo de historial no existe"
            return @()
        }
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $history = New-Object System.Collections.Generic.List[object]
        $inHistorySection = $false
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesHistory]") { $inHistorySection = $true; continue }
            if ($trimmed -match '^\[') { $inHistorySection = $false; continue }
            if ($inHistorySection -and -not ($trimmed -match '^\s*;') -and -not [string]::IsNullOrWhiteSpace($trimmed)) {
                $parts = $trimmed -split '\|', 6
                if ($parts.Count -ge 6) {
                    try {
                        $queryBytes = [Convert]::FromBase64String($parts[5])
                        $fullQuery = [System.Text.Encoding]::UTF8.GetString($queryBytes)
                        $history.Add([PSCustomObject]@{
                                Timestamp = $parts[0]
                                Database  = $parts[1]
                                Hash      = $parts[2]
                                Preview   = $parts[3]
                                Result    = $parts[4]
                                FullQuery = $fullQuery
                                Success   = $parts[4] -match '^OK'
                            })
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error parseando línea de historial: $_" Yellow
                    }
                }
            }
        }
        $history = @([Linq.Enumerable]::Reverse($history) | Select-Object -First $MaxItems)
        Write-DzDebug "`t[DEBUG][Queries] Historial cargado: $($history.Count) items"

        # CORRECCIÓN: Solo esto
        return $history

    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error leyendo historial: $_" Red
        return @()
    }
}
function Clear-QueryHistory {
    [CmdletBinding()]
    param()
    try {
        if (-not (Test-Path -LiteralPath $script:QueriesConfigPath)) { return $true }
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $lines = New-Object System.Collections.Generic.List[string]
        $inHistorySection = $false
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesHistory]") { $inHistorySection = $true; $lines.Add($line); continue }
            if ($trimmed -match '^\[') { $inHistorySection = $false; $lines.Add($line); continue }
            if (-not $inHistorySection) { $lines.Add($line) } elseif ($trimmed -match '^\s*;') { $lines.Add($line) }
        }
        Set-Content -LiteralPath $script:QueriesConfigPath -Value $lines -Encoding UTF8
        Write-DzDebug "`t[DEBUG][Queries] Historial limpiado"
        return $true
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error limpiando historial: $_" Red
        return $false
    }
}
function Save-OpenQueryTabs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.TabControl]$TabControl
    )
    try {
        if (-not (Initialize-QueriesConfig)) {
            Write-DzDebug "`t[DEBUG][Queries] No se pudo inicializar configuración" Yellow
            return $false
        }
        Write-DzDebug "`t[DEBUG][Queries] Guardando pestañas abiertas..."
        $tabsData = New-Object System.Collections.Generic.List[string]
        $tabIndex = 0
        $selectedIndex = $TabControl.SelectedIndex
        foreach ($item in $TabControl.Items) {
            if ($item -isnot [System.Windows.Controls.TabItem]) { continue }
            if (-not $item.Tag -or $item.Tag.Type -ne 'QueryTab') { continue }
            $editor = $item.Tag.Editor
            if (-not $editor) { continue }
            $queryText = Get-SqlEditorText -Editor $editor
            if ([string]::IsNullOrWhiteSpace($queryText)) { continue }
            $title = if ($item.Tag.Title) { $item.Tag.Title } else { "Consulta $($tabIndex + 1)" }
            $isDirty = if ($item.Tag.IsDirty) { "1" } else { "0" }
            $queryBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($queryText))
            $tabLine = "$tabIndex|$title|$isDirty|$queryBase64"
            $tabsData.Add($tabLine)
            $tabIndex++
        }
        Write-DzDebug "`t[DEBUG][Queries] Tabs a guardar: $($tabsData.Count)"
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $lines = New-Object System.Collections.Generic.List[string]
        $inOpenSection = $false
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesOpen]") {
                $inOpenSection = $true
                $lines.Add($line)
                $lines.Add("; Estado de pestañas abiertas al cerrar")
                $lines.Add("; Formato: tab_index|tab_title|is_dirty|query_content_base64")
                foreach ($tabData in $tabsData) { $lines.Add($tabData) }
                continue
            }
            if ($inOpenSection) {
                if ($trimmed -match '^\[') { $inOpenSection = $false; $lines.Add($line) }
                continue
            }
            $lines.Add($line)
        }
        Set-Content -LiteralPath $script:QueriesConfigPath -Value $lines -Encoding UTF8
        Write-DzDebug "`t[DEBUG][Queries] Pestañas guardadas exitosamente"
        return $true
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error guardando pestañas: $_" Red
        return $false
    }
}
function Restore-OpenQueryTabs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.TabControl]$TabControl
    )
    try {
        if (-not (Test-Path -LiteralPath $script:QueriesConfigPath)) {
            Write-DzDebug "`t[DEBUG][Queries] No hay archivo de configuración para restaurar"
            return 0
        }
        Write-DzDebug "`t[DEBUG][Queries] Restaurando pestañas..."
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $inOpenSection = $false
        $restoredCount = 0
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesOpen]") { $inOpenSection = $true; continue }
            if ($trimmed -match '^\[') { $inOpenSection = $false; continue }
            if ($inOpenSection -and -not ($trimmed -match '^\s*;') -and -not [string]::IsNullOrWhiteSpace($trimmed)) {
                $parts = $trimmed -split '\|', 4
                if ($parts.Count -ge 4) {
                    try {
                        $queryBytes = [Convert]::FromBase64String($parts[3])
                        $queryText = [System.Text.Encoding]::UTF8.GetString($queryBytes)
                        $newTab = New-QueryTab -TabControl $TabControl
                        if ($newTab -and $newTab.Tag -and $newTab.Tag.Editor) {
                            $editor = $newTab.Tag.Editor
                            Set-SqlEditorText -Editor $editor -Text $queryText
                            $savedTitle = $parts[1]
                            if (-not [string]::IsNullOrWhiteSpace($savedTitle) -and $savedTitle -ne $newTab.Tag.Title) {
                                $newTab.Tag.Title = $savedTitle
                                if ($newTab.Tag.HeaderTextBlock) { $newTab.Tag.HeaderTextBlock.Text = $savedTitle }
                            }
                            $newTab.Tag.IsDirty = ($parts[2] -eq "1")
                            Update-QueryTabHeader -TabItem $newTab
                            $restoredCount++
                        }
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error restaurando tab: $_" Yellow
                    }
                }
            }
        }
        Write-DzDebug "`t[DEBUG][Queries] Pestañas restauradas: $restoredCount"
        if ($restoredCount -gt 0) {
            $lastQueryTabIndex = -1
            for ($i = 0; $i -lt $TabControl.Items.Count; $i++) {
                $it = $TabControl.Items[$i]
                if ($it -is [System.Windows.Controls.TabItem] -and $it.Tag -and $it.Tag.Type -eq 'QueryTab') { $lastQueryTabIndex = $i }
            }
            if ($lastQueryTabIndex -ge 0) { $TabControl.SelectedIndex = $lastQueryTabIndex }
        }
        return $restoredCount
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error restaurando pestañas: $_" Red
        return 0
    }
}
function Add-QueryHistoryTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Windows.Controls.TabControl]$TabControl
    )
    try {
        Write-DzDebug "`t[DEBUG][Queries] Creando pestaña de historial en TabControl"
        foreach ($item in $TabControl.Items) {
            if ($item -is [System.Windows.Controls.TabItem] -and $item.Tag -and $item.Tag.Type -eq 'HistoryTab') {
                Write-DzDebug "`t[DEBUG][Queries] Pestaña de historial ya existe"
                return $item
            }
        }
        $tabItem = New-Object System.Windows.Controls.TabItem
        $headerPanel = New-Object System.Windows.Controls.StackPanel
        $headerPanel.Orientation = "Horizontal"
        $iconText = New-Object System.Windows.Controls.TextBlock
        $iconText.Text = "📋"
        $iconText.Margin = "0,0,6,0"
        $iconText.VerticalAlignment = "Center"
        $headerText = New-Object System.Windows.Controls.TextBlock
        $headerText.Text = "Historial"
        $headerText.VerticalAlignment = "Center"
        [void]$headerPanel.Children.Add($iconText)
        [void]$headerPanel.Children.Add($headerText)
        $tabItem.Header = $headerPanel
        $grid = New-Object System.Windows.Controls.Grid
        $row1 = New-Object System.Windows.Controls.RowDefinition
        $row1.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $row2 = New-Object System.Windows.Controls.RowDefinition
        $row2.Height = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $grid.RowDefinitions.Add($row1)
        $grid.RowDefinitions.Add($row2)
        $dataGrid = New-Object System.Windows.Controls.DataGrid
        $dataGrid.IsReadOnly = $true
        $dataGrid.AutoGenerateColumns = $false
        $dataGrid.CanUserAddRows = $false
        $dataGrid.CanUserDeleteRows = $false
        $dataGrid.SelectionMode = "Single"
        $dataGrid.HeadersVisibility = "Column"
        $dataGrid.Margin = "5"
        $col1 = New-Object System.Windows.Controls.DataGridTextColumn
        $col1.Header = "Fecha/Hora"
        $col1.Binding = New-Object System.Windows.Data.Binding("Timestamp")
        $col1.Width = 140
        $col2 = New-Object System.Windows.Controls.DataGridTextColumn
        $col2.Header = "Base de Datos"
        $col2.Binding = New-Object System.Windows.Data.Binding("Database")
        $col2.Width = 150
        $col3 = New-Object System.Windows.Controls.DataGridTextColumn
        $col3.Header = "Query"
        $col3.Binding = New-Object System.Windows.Data.Binding("Preview")
        $col3.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star)
        $col4 = New-Object System.Windows.Controls.DataGridTextColumn
        $col4.Header = "Resultado"
        $col4.Binding = New-Object System.Windows.Data.Binding("Result")
        $col4.Width = 180
        [void]$dataGrid.Columns.Add($col1)
        [void]$dataGrid.Columns.Add($col2)
        [void]$dataGrid.Columns.Add($col3)
        [void]$dataGrid.Columns.Add($col4)
        [System.Windows.Controls.Grid]::SetRow($dataGrid, 0)
        [void]$grid.Children.Add($dataGrid)
        $buttonPanel = New-Object System.Windows.Controls.StackPanel
        $buttonPanel.Orientation = "Horizontal"
        $buttonPanel.HorizontalAlignment = "Right"
        $buttonPanel.Margin = "5"
        [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
        $btnRefresh = New-Object System.Windows.Controls.Button
        $btnRefresh.Content = "🔄 Actualizar"
        $btnRefresh.Width = 100
        $btnRefresh.Height = 28
        $btnRefresh.Margin = "0,0,5,0"
        $btnCopy = New-Object System.Windows.Controls.Button
        $btnCopy.Content = "📋 Copiar"
        $btnCopy.Width = 90
        $btnCopy.Height = 28
        $btnCopy.Margin = "0,0,5,0"
        $btnLoad = New-Object System.Windows.Controls.Button
        $btnLoad.Content = "📥 Cargar"
        $btnLoad.Width = 90
        $btnLoad.Height = 28
        $btnLoad.Margin = "0,0,5,0"
        $btnClear = New-Object System.Windows.Controls.Button
        $btnClear.Content = "🗑️ Limpiar"
        $btnClear.Width = 100
        $btnClear.Height = 28
        [void]$buttonPanel.Children.Add($btnRefresh)
        [void]$buttonPanel.Children.Add($btnCopy)
        [void]$buttonPanel.Children.Add($btnLoad)
        [void]$buttonPanel.Children.Add($btnClear)
        [void]$grid.Children.Add($buttonPanel)
        $tabItem.Content = $grid
        $tabItem.Tag = [pscustomobject]@{
            Type     = "HistoryTab"
            DataGrid = $dataGrid
        }
        $loadHistory = {
            try {
                $history = @(Get-QueryHistory -MaxItems 100)
                $dataGrid.ItemsSource = $history
                Write-DzDebug "`t[DEBUG][Queries] Historial cargado: $($history.Count) items"
            } catch {
                Write-DzDebug "`t[DEBUG][Queries] Error cargando historial: $_" Red
            }
        }.GetNewClosure()
        $btnRefresh.Add_Click({
                & $loadHistory
            }.GetNewClosure())
        $btnCopy.Add_Click({
                try {
                    $selected = $dataGrid.SelectedItem
                    if (-not $selected) {
                        Ui-Warn "Selecciona un query del historial" "Atención"
                        return
                    }
                    if (Get-Command Set-ClipboardTextSafe -ErrorAction SilentlyContinue) {
                        Set-ClipboardTextSafe -Text $selected.FullQuery | Out-Null
                    } else {
                        [System.Windows.Clipboard]::SetText($selected.FullQuery)
                    }
                    Write-Host "`n✓ Query copiado al portapapeles" -ForegroundColor Green
                } catch {
                    Write-Host "`n✗ Error copiando query: $_" -ForegroundColor Red
                }
            }.GetNewClosure())
        $tcRef = $TabControl
        $btnLoad.Add_Click({
                try {
                    $selected = $dataGrid.SelectedItem
                    if (-not $selected) {
                        Ui-Warn "Selecciona un query del historial" "Atención"
                        return
                    }
                    $activeTab = Get-ActiveQueryTab -TabControl $tcRef
                    if (-not $activeTab) { $activeTab = New-QueryTab -TabControl $tcRef }
                    if ($activeTab -and $activeTab.Tag -and $activeTab.Tag.Editor) {
                        $editor = $activeTab.Tag.Editor
                        Set-SqlEditorText -Editor $editor -Text $selected.FullQuery
                        $tcRef.SelectedItem = $activeTab
                        Write-Host "`n✓ Query cargado en el editor" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "`n✗ Error cargando query: $_" -ForegroundColor Red
                    Ui-Error "Error cargando query: $($_.Exception.Message)"
                }
            }.GetNewClosure())
        $btnClear.Add_Click({
                $confirm = Ui-Confirm "¿Estás seguro de limpiar todo el historial de queries?" "Confirmar"
                if ($confirm) {
                    if (Clear-QueryHistory) {
                        Write-Host "`n✓ Historial limpiado" -ForegroundColor Green
                        & $loadHistory
                    } else {
                        Ui-Error "Error limpiando historial"
                    }
                }
            }.GetNewClosure())
        & $loadHistory
        $insertIndex = $TabControl.Items.Count
        for ($i = $TabControl.Items.Count - 1; $i -ge 0; $i--) {
            $it = $TabControl.Items[$i]
            if ($it -is [System.Windows.Controls.TabItem] -and $it.Name -eq "tabAddQuery") { $insertIndex = $i; break }
        }
        [void]$TabControl.Items.Insert($insertIndex, $tabItem)
        Write-DzDebug "`t[DEBUG][Queries] Pestaña de historial creada exitosamente en posición $insertIndex"
        return $tabItem
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error creando pestaña de historial: $_" Red
        return $null
    }
}
function Get-StringHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$InputString
    )
    try {
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
        $hashBytes = $md5.ComputeHash($bytes)
        $hash = [System.BitConverter]::ToString($hashBytes).Replace('-', '').Substring(0, 16)
        return $hash.ToLower()
    } catch {
        return [Guid]::NewGuid().ToString("N").Substring(0, 16)
    }
}
function Execute-QueryCore {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Ctx)

    if ($Ctx.QueryRunning) {
        Write-DzDebug "`t[DEBUG] Query ya en ejecución, ignorando click"
        return
    }

    $Ctx.QueryRunning = $true

    $activeEditor = Get-ActiveQueryRichTextBox -TabControl $Ctx.tcQueries
    if (-not $activeEditor) { throw "No hay una pestaña de consulta activa." }

    $rawQuery = $null
    $rawQuery = Get-SqlEditorSelectedText -Editor $activeEditor
    if ([string]::IsNullOrWhiteSpace($rawQuery)) {
        $rawQuery = Get-SqlEditorText -Editor $activeEditor
    }
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
    $script:rawQueryToSave = $rawQuery
    $script:dbToSave = $db
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
                $rows = 0
                try {
                    $result = $script:CurrentQueryPowerShell.EndInvoke($script:CurrentQueryAsync)
                    if ($result -is [System.Management.Automation.PSDataCollection[psobject]]) {
                        if ($result.Count -gt 0) { $result = $result[0] } else { $result = $null }
                    } elseif ($result -is [System.Array]) {
                        if ($result.Count -gt 0) { $result = $result[0] } else { $result = $null }
                    }
                    try {
                        if ($result -and $result.ResultSets -and $result.ResultSets.Count -gt 0) {
                            $rows = ($result.ResultSets | Measure-Object -Property RowCount -Sum).Sum
                            if ($rows -eq $null -or $rows -lt 0) { $rows = 0 }
                        } elseif ($result -and ($result -is [hashtable]) -and $result.ContainsKey('RowsAffected') -and $result.RowsAffected -ne $null) {
                            $rows = [int]$result.RowsAffected
                        }
                    } catch { $rows = 0 }
                    try {
                        $ok = $false
                        try { $ok = [bool]$result.Success } catch { $ok = $false }

                        $errMsg = ""
                        if (-not $ok) {
                            try { $errMsg = [string]$result.ErrorMessage } catch {}
                            if ($result.InnerError) { $errMsg += " | Inner: $($result.InnerError)" }
                        }
                        Add-QueryToHistory -Query $rawQueryToSave -Database $dbToSave -Success $ok -RowsAffected $rows -ErrorMessage $errMsg
                        Write-DzDebug "`t[DEBUG][Queries] Historial guardado (DB='$db' Success=$ok Rows=$rows)"
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error guardando historial: $($_.Exception.Message)" Yellow
                    }
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
                    # --- HISTORIAL: registrar ejecución (éxito o error) ---
                    try {
                        $ok = $false
                        try { $ok = [bool]$result.Success } catch { $ok = $false }

                        $errMsg = ""
                        if (-not $ok) {
                            try { $errMsg = [string]$result.ErrorMessage } catch {}
                            if ($result.InnerError) { $errMsg += " | Inner: $($result.InnerError)" }
                        }

                        Add-QueryToHistory -Query $rawQuery -Database $db -Success $ok -RowsAffected $rows -ErrorMessage $errMsg
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error registrando historial desde Execute-QueryCore: $_" Yellow
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
        sqlEditor1             = $global:sqlEditor1
        dgResults              = $global:dgResults
        txtMessages            = $global:txtMessages
        lblRowCount            = $global:lblRowCount
        lblExecutionTimer      = $global:lblExecutionTimer

        MainWindow             = $global:MainWindow
    }
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
    $border = New-Object System.Windows.Controls.Border
    $border.BorderThickness = "1"
    $border.CornerRadius = "4"
    $border.Margin = "5"
    $border.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, "BorderBrushColor")
    $border.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "ControlBg")
    $editor = New-SqlEditor -Container $border -FontFamily "Consolas" -FontSize 12
    $tabItem.Content = $border
    $tabItem.Tag = [pscustomobject]@{ Type = "QueryTab"; Editor = $editor; Title = $tabTitle; HeaderTextBlock = $headerText; IsDirty = $false }
    $editor.Add_TextChanged({
            $tabItem.Tag.IsDirty = $true
            Update-QueryTabHeader -TabItem $tabItem
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
    if ($tab -and $tab.Tag -and $tab.Tag.Editor) { return $tab.Tag.Editor }
    $null
}
function Set-QueryTextInActiveTab {
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)][string]$Text
    )
    $editor = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $editor) { return }
    Set-SqlEditorText -Editor $editor -Text $Text
}
function Insert-TextIntoActiveQuery {
    param(
        [Parameter(Mandatory = $true)]$TabControl,
        [Parameter(Mandatory = $true)][string]$Text
    )
    $editor = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $editor) { return }
    Insert-SqlEditorText -Editor $editor -Text $Text
    $editor.Focus()
}
function Clear-ActiveQueryTab {
    param([Parameter(Mandatory = $true)]$TabControl)
    $editor = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if ($editor) {
        Clear-SqlEditorText -Editor $editor
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
    $editor = Get-ActiveQueryRichTextBox -TabControl $TabControl
    if (-not $editor) { throw "No hay una pestaña de consulta activa." }
    $rawQuery = Get-SqlEditorText -Editor $editor
    $cleanQuery = Remove-SqlComments -Query $rawQuery
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
        try {
            if (Get-Command Add-QueryToHistory -ErrorAction SilentlyContinue) {
                $dbName = Get-DbNameFromComboSelection -ComboBox $global:cmbDatabases
                if (-not [string]::IsNullOrWhiteSpace($dbName)) {
                    $totalRows = ($result.ResultSets | Measure-Object -Property RowCount -Sum).Sum
                    Add-QueryToHistory -Query $cleanQuery -Database $dbName -Success $true -RowsAffected $totalRows
                    Write-DzDebug "`t[DEBUG] Query exitoso registrado en historial (ResultSets)"
                }
            }
        } catch {
            Write-DzDebug "`t[DEBUG] Error registrando query exitoso: $_" Yellow
        }
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
        try {
            if (Get-Command Add-QueryToHistory -ErrorAction SilentlyContinue) {
                $dbName = Get-DbNameFromComboSelection -ComboBox $global:cmbDatabases
                if (-not [string]::IsNullOrWhiteSpace($dbName)) {
                    Add-QueryToHistory -Query $cleanQuery -Database $dbName -Success $true -RowsAffected $result.RowsAffected
                    Write-DzDebug "`t[DEBUG] Query exitoso registrado en historial (RowsAffected)"
                }
            }
        } catch {
            Write-DzDebug "`t[DEBUG] Error registrando query exitoso: $_" Yellow
        }
    }
    Write-DzSqlResultSummary -Result $result -Context "Consulta"
    return $result
}
$script:SqlEditorAssemblyLoaded = $false
$script:SqlEditorHighlighting = $null
function Get-SqlEditorPaths {
    $moduleRoot = Split-Path -Parent $PSScriptRoot

    $assemblyPath = Join-Path (Join-Path $moduleRoot "lib") "AvalonEdit.dll"
    $highlightingPath = Join-Path (Join-Path $moduleRoot "resources") "SQL.xshd"

    return [pscustomobject]@{
        AssemblyPath     = $assemblyPath
        HighlightingPath = $highlightingPath
    }
}

function Import-AvalonEditAssembly {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AssemblyPath
    )
    if ($script:SqlEditorAssemblyLoaded) { return }
    if (-not (Test-Path -LiteralPath $AssemblyPath)) {
        throw "No se encontró AvalonEdit.dll en '$AssemblyPath'."
    }
    Add-Type -Path $AssemblyPath
    $script:SqlEditorAssemblyLoaded = $true
}
function Get-SqlEditorHighlighting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HighlightingPath
    )

    if ($script:SqlEditorHighlighting) { return $script:SqlEditorHighlighting }
    if (-not (Test-Path -LiteralPath $HighlightingPath)) { return $null }

    try {
        $reader = [System.Xml.XmlReader]::Create($HighlightingPath)
        try {
            $script:SqlEditorHighlighting = [ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader]::Load(
                $reader,
                [ICSharpCode.AvalonEdit.Highlighting.HighlightingManager]::Instance
            )
        } finally {
            $reader.Close()
        }
        return $script:SqlEditorHighlighting
    } catch {
        Write-Host "⚠ Highlighting inválido ($HighlightingPath). Se iniciará sin resaltado. Detalle: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function New-SqlEditor {
    [CmdletBinding()]
    param(
        [Parameter()][System.Windows.Controls.Border]$Container,
        [string]$FontFamily = "Consolas",
        [int]$FontSize = 12
    )
    $paths = Get-SqlEditorPaths
    Import-AvalonEditAssembly -AssemblyPath $paths.AssemblyPath
    $editor = New-Object ICSharpCode.AvalonEdit.TextEditor
    $editor.ShowLineNumbers = $true
    $editor.FontFamily = $FontFamily
    $editor.FontSize = $FontSize
    $editor.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.Options.ConvertTabsToSpaces = $false
    $editor.SyntaxHighlighting = Get-SqlEditorHighlighting -HighlightingPath $paths.HighlightingPath
    if ($Container) {
        $Container.Child = $editor
    }
    return $editor
}
function Set-SqlEditorText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Editor,
        [Parameter(Mandatory)][string]$Text
    )
    if (-not $Editor) { return }
    $Editor.Text = $Text
}
function Get-SqlEditorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return "" }
    return [string]$Editor.Text
}
function Clear-SqlEditorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return }
    $Editor.Clear()
}
function Insert-SqlEditorText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Editor,
        [Parameter(Mandatory)][string]$Text
    )
    if (-not $Editor) { return }
    $offset = $Editor.CaretOffset
    $Editor.Document.Insert($offset, $Text)
    $Editor.CaretOffset = $offset + $Text.Length
}
function Get-SqlEditorSelectedText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return "" }
    return [string]$Editor.SelectedText
}
Export-ModuleMember -Function @(
    'Initialize-QueriesConfig',
    'Add-QueryToHistory',
    'Get-QueryHistory',
    'Clear-QueryHistory',
    'Save-OpenQueryTabs',
    'Restore-OpenQueryTabs',
    'Show-QueryHistoryWindow',
    'Execute-QueryUiSafe',
    'Export-ResultsUiSafe', 'Connect-DbUiSafe', 'Disconnect-DbUiSafe', 'Add-QueryHistoryTab', 'New-QueryTab',
    'Get-ActiveQueryTab', 'Get-ActiveQueryRichTextBox', 'Set-QueryTextInActiveTab', 'Insert-TextIntoActiveQuery',
    'Clear-ActiveQueryTab', 'Update-QueryTabHeader', 'Close-QueryTab', 'Execute-QueryInTab',
    'Get-SqlEditorPaths',
    'Import-AvalonEditAssembly',
    'Get-SqlEditorHighlighting',
    'New-SqlEditor',
    'Set-SqlEditorText',
    'Get-SqlEditorText',
    'Clear-SqlEditorText',
    'Insert-SqlEditorText',
    'Get-SqlEditorSelectedText'
)
