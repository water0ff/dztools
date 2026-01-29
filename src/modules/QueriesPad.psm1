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
        [string]$Server = "",
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
        $historyLine = "$timestamp|$Server|$Database|$hash|$preview|$resultInfo|$queryBase64"
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
            if ($trimmed -eq "[QueriesHistory]") {
                $inHistorySection = $true
                continue
            }
            if ($trimmed -match '^\[') {
                $inHistorySection = $false
                continue
            }
            if ($inHistorySection -and -not ($trimmed -match '^\s*;') -and -not [string]::IsNullOrWhiteSpace($trimmed)) {
                $parts = $trimmed -split '\|', 7
                if ($parts.Count -eq 6) {
                    try {
                        $queryBytes = [Convert]::FromBase64String($parts[5])
                        $fullQuery = [System.Text.Encoding]::UTF8.GetString($queryBytes)
                        $history.Add([PSCustomObject]@{
                                Timestamp = $parts[0]
                                Server    = ""
                                Database  = $parts[1]
                                Hash      = $parts[2]
                                Preview   = $parts[3]
                                Result    = $parts[4]
                                FullQuery = $fullQuery
                                Success   = $parts[4] -match '^OK'
                            })
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error parseando línea antigua: $_" Yellow
                    }
                } elseif ($parts.Count -ge 7) {
                    try {
                        $queryBytes = [Convert]::FromBase64String($parts[6])
                        $fullQuery = [System.Text.Encoding]::UTF8.GetString($queryBytes)
                        $history.Add([PSCustomObject]@{
                                Timestamp = $parts[0]
                                Server    = $parts[1]
                                Database  = $parts[2]
                                Hash      = $parts[3]
                                Preview   = $parts[4]
                                Result    = $parts[5]
                                FullQuery = $fullQuery
                                Success   = $parts[5] -match '^OK'
                            })
                    } catch {
                        Write-DzDebug "`t[DEBUG][Queries] Error parseando línea nueva: $_" Yellow
                    }
                }
            }
        }
        $history = @([Linq.Enumerable]::Reverse($history) | Select-Object -First $MaxItems)
        Write-DzDebug "`t[DEBUG][Queries] Historial cargado: $($history.Count) items"
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
function Remove-QueriesFromHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Array]$Items
    )
    try {
        if (-not (Test-Path -LiteralPath $script:QueriesConfigPath)) {
            return $false
        }
        $content = Get-Content -LiteralPath $script:QueriesConfigPath -ErrorAction Stop
        $lines = New-Object System.Collections.Generic.List[string]
        $inHistorySection = $false
        $hashesToRemove = @{}
        foreach ($item in $Items) {
            $key = "$($item.Timestamp)|$($item.Hash)"
            $hashesToRemove[$key] = $true
        }
        foreach ($line in $content) {
            $trimmed = $line.Trim()
            if ($trimmed -eq "[QueriesHistory]") {
                $inHistorySection = $true
                $lines.Add($line)
                continue
            }
            if ($trimmed -match '^\[') {
                $inHistorySection = $false
                $lines.Add($line)
                continue
            }
            if ($inHistorySection -and -not ($trimmed -match '^\s*;') -and -not [string]::IsNullOrWhiteSpace($trimmed)) {
                $parts = $trimmed -split '\|'
                $timestamp = $parts[0]
                $hash = if ($parts.Count -ge 7) { $parts[3] } else { $parts[2] }
                $key = "$timestamp|$hash"
                if (-not $hashesToRemove.ContainsKey($key)) {
                    $lines.Add($line)
                }
            } elseif (-not $inHistorySection -or ($trimmed -match '^\s*;')) {
                $lines.Add($line)
            }
        }
        Set-Content -LiteralPath $script:QueriesConfigPath -Value $lines -Encoding UTF8
        Write-DzDebug "`t[DEBUG][Queries] Eliminados $($Items.Count) items del historial"
        return $true
    } catch {
        Write-DzDebug "`t[DEBUG][Queries] Error eliminando items: $_" Red
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
            $title = if ($item.Tag.Title) { $item.Tag.Title } else { "Query$($tabIndex + 1)" }
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
                            $parsedNumber = Get-QueryTabNumberFromTitle -Title ([string]$newTab.Tag.Title)
                            if ($parsedNumber) { $newTab.Tag.Number = $parsedNumber }
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
            $utilPath = Join-Path $ModulesPath "Utilities.psm1"
            $dbPath = Join-Path $ModulesPath "Database.psm1"
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
    [void]$ps.AddScript($worker).AddArgument($server).AddArgument($db).AddArgument($query).AddArgument($userText).AddArgument($passwordTxt).AddArgument($modulesPath)
    $Ctx.CurrentQueryAsync = $ps.BeginInvoke()
    if ($Ctx.QueryDoneTimer) {
        try { if ($Ctx.QueryDoneTimer.IsEnabled) { $Ctx.QueryDoneTimer.Stop() } } catch {}
    }
    $script:rawQueryToSave = $rawQuery
    $script:dbToSave = $db
    $script:serverToSave = $server
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
                        Add-QueryToHistory -Query $rawQueryToSave -Database $dbToSave -Server $serverToSave -Success $ok -RowsAffected $rows -ErrorMessage $errMsg
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

function Get-QueryTabTitle {
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter()][string]$Database,
        [Parameter()][switch]$Short
    )
    if ([string]::IsNullOrWhiteSpace($Database)) {
        return "Query$Number"
    }
    if ($Short) {
        if ($Database.Length -gt 7) {
            $shortDb = "..." + $Database.Substring($Database.Length - 7)
            return "Query$Number ($shortDb)"
        }
    }
    return "Query$Number ($Database)"
}

function Get-QueryTabNumberFromTitle {
    param([Parameter(Mandatory = $true)][string]$Title)
    if ($Title -match 'Query\s*(\d+)') { return [int]$Matches[1] }
    if ($Title -match 'Consulta\s+(\d+)') { return [int]$Matches[1] }
    return $null
}

function Set-QueryTabsDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl,
        [Parameter()][string]$Database
    )
    Write-DzDebug "`t[DEBUG] Actualizando todas las pestañas con DB: '$Database'"
    foreach ($item in $TabControl.Items) {
        if ($item -isnot [System.Windows.Controls.TabItem]) { continue }
        if (-not $item.Tag -or $item.Tag.Type -ne 'QueryTab') { continue }
        if (-not $item.Tag.Number) {
            $parsedNumber = Get-QueryTabNumberFromTitle -Title ([string]$item.Tag.Title)
            if ($parsedNumber) {
                $item.Tag.Number = $parsedNumber
            }
        }
        $item.Tag.Database = $Database
        if ($item.Tag.Number) {
            $fullTitle = Get-QueryTabTitle -Number $item.Tag.Number -Database $Database
            $shortTitle = Get-QueryTabTitle -Number $item.Tag.Number -Database $Database -Short
            $item.Tag.Title = $fullTitle
            $item.Tag.TitleShort = $shortTitle
            if (-not [string]::IsNullOrWhiteSpace($Database)) {
                $item.ToolTip = $fullTitle
            } else {
                $item.ToolTip = $null
            }
            Write-DzDebug "`t[DEBUG]   Query$($item.Tag.Number): '$shortTitle' (tooltip: '$fullTitle')"
        }
        Update-QueryTabHeader -TabItem $item
    }
}

function Close-OtherQueryTabs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl,
        [Parameter(Mandatory = $true)][System.Windows.Controls.TabItem]$TabItem
    )
    $toClose = @()
    foreach ($item in $TabControl.Items) {
        if ($item -isnot [System.Windows.Controls.TabItem]) { continue }
        if (-not $item.Tag -or $item.Tag.Type -ne 'QueryTab') { continue }
        if ($item -eq $TabItem) { continue }
        $toClose += $item
    }
    foreach ($item in $toClose) {
        Close-QueryTab -TabControl $TabControl -TabItem $item
    }
}
function New-QueryTab {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Windows.Controls.TabControl]$TabControl)
    $tabNumber = Get-NextQueryNumber -TabControl $TabControl
    $dbName = $null
    if ($global:cmbDatabases -and (Get-Command Get-DbNameFromComboSelection -ErrorAction SilentlyContinue)) {
        $dbName = Get-DbNameFromComboSelection -ComboBox $global:cmbDatabases
    }
    $tabTitleShort = Get-QueryTabTitle -Number $tabNumber -Database $dbName -Short
    $tabTitleFull = Get-QueryTabTitle -Number $tabNumber -Database $dbName
    $tabItem = New-Object System.Windows.Controls.TabItem
    $headerPanel = New-Object System.Windows.Controls.StackPanel
    $headerPanel.Orientation = "Horizontal"
    $headerText = New-Object System.Windows.Controls.TextBlock
    $headerText.Text = $tabTitleShort
    $headerText.VerticalAlignment = "Center"
    $headerText.FontSize = 10
    $closeButton = New-Object System.Windows.Controls.Button
    $closeButton.Content = "×"
    $closeButton.Width = 16
    $closeButton.Height = 16
    $closeButton.Margin = "4,0,0,0"
    $closeButton.Padding = "0"
    $closeButton.FontSize = 10
    [void]$headerPanel.Children.Add($headerText)
    [void]$headerPanel.Children.Add($closeButton)
    $tabItem.Header = $headerPanel
    if (-not [string]::IsNullOrWhiteSpace($dbName)) {
        $tabItem.ToolTip = $tabTitleFull
    }
    $border = New-Object System.Windows.Controls.Border
    $border.BorderThickness = "1"
    $border.CornerRadius = "4"
    $border.Margin = "5"
    $border.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, "BorderBrushColor")
    $border.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, "ControlBg")
    $editor = New-SqlEditor -Container $border -FontFamily "Consolas" -FontSize 11
    $tabItem.Content = $border
    $tabItem.Tag = [pscustomobject]@{
        Type            = "QueryTab"
        Editor          = $editor
        Title           = $tabTitleFull
        TitleShort      = $tabTitleShort
        HeaderTextBlock = $headerText
        IsDirty         = $false
        Number          = $tabNumber
        Database        = $dbName
    }
    $editor.Add_TextChanged({
            $tabItem.Tag.IsDirty = $true
            Update-QueryTabHeader -TabItem $tabItem
        }.GetNewClosure())
    $tcRef = $TabControl
    $closeButton.Add_Click({ Close-QueryTab -TabControl $tcRef -TabItem $tabItem }.GetNewClosure())
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $menuClose = New-Object System.Windows.Controls.MenuItem
    $menuClose.Header = "Cerrar esta pestaña"
    $menuClose.Add_Click({ Close-QueryTab -TabControl $tcRef -TabItem $tabItem }.GetNewClosure())
    $menuCloseOthers = New-Object System.Windows.Controls.MenuItem
    $menuCloseOthers.Header = "Cerrar otras pestañas"
    $menuCloseOthers.Add_Click({ Close-OtherQueryTabs -TabControl $tcRef -TabItem $tabItem }.GetNewClosure())
    [void]$contextMenu.Items.Add($menuClose)
    [void]$contextMenu.Items.Add($menuCloseOthers)
    $tabItem.ContextMenu = $contextMenu
    $insertIndex = $TabControl.Items.Count
    for ($i = 0; $i -lt $TabControl.Items.Count; $i++) {
        $it = $TabControl.Items[$i]
        if ($it -is [System.Windows.Controls.TabItem] -and $it.Name -eq "tabAddQuery") {
            $insertIndex = $i
            break
        }
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
        if ($item.Tag.Number) {
            $n = [int]$item.Tag.Number
            if ($n -gt $max) { $max = $n }
            continue
        }
        $title = [string]$item.Tag.Title
        $parsed = Get-QueryTabNumberFromTitle -Title $title
        if ($parsed -and $parsed -gt $max) { $max = $parsed }
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
    $displayTitle = if ($TabItem.Tag.TitleShort) {
        $TabItem.Tag.TitleShort
    } else {
        $TabItem.Tag.Title
    }
    if ($TabItem.Tag.IsDirty) {
        $displayTitle = "*$displayTitle"
    }
    if ($TabItem.Tag.HeaderTextBlock) {
        $TabItem.Tag.HeaderTextBlock.Text = $displayTitle
    }
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
function Show-QueryHistoryWindow {
    [CmdletBinding()]
    param(
        [System.Windows.Window]$Owner
    )
    try {
        $theme = Get-DzUiTheme
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Historial de Consultas SQL"
        SizeToContent="Manual"
        Height="400" Width="800"
        WindowStartupLocation="CenterOwner"
        WindowStyle="None"
        ResizeMode="CanResize"
        ShowInTaskbar="False"
        Background="#66000000"
        AllowsTransparency="True"
        Topmost="True"
        MinHeight="400" MinWidth="800">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style x:Key="BaseButtonStyle" TargetType="Button">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.6"/>
                    <Setter Property="Cursor" Value="Arrow"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DatabaseButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentDatabase}"/>
            <Setter Property="Foreground" Value="#111111"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentDatabaseHover}"/>
                    <Setter Property="Foreground" Value="#111111"/>
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
            </Style.Triggers>
        </Style>
        <Style x:Key="DangerButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentRed}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentRedHover}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="OutlineButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                    <Setter Property="BorderThickness" Value="0"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="CloseButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Width" Value="28"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Content" Value="×"/>
        </Style>
    </Window.Resources>
    <Grid Background="Transparent">
        <Border Background="{DynamicResource FormBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="10"
                Margin="9"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Stretch">
            <Border.Effect>
                <DropShadowEffect BlurRadius="2"
                                  ShadowDepth="0"
                                  Opacity="0.45"/>
            </Border.Effect>

            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                    <Border Grid.Row="0" Name="brdTitleBar"
                            Background="{DynamicResource PanelBg}"
                            Cursor="SizeAll"
                            BorderBrush="{DynamicResource BorderBrushColor}"
                            BorderThickness="1"
                            CornerRadius="8"
                            Padding="8"
                            Margin="0,0,0,6">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" VerticalAlignment="Center">
                            <TextBlock Text="📜 Historial de Consultas SQL"
                                       FontSize="13" FontWeight="SemiBold"
                                       Foreground="{DynamicResource AccentPrimary}"/>
                            <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                <TextBlock Name="lblHistoryCount" Text="0 consultas"
                                           FontSize="10" Foreground="{DynamicResource AccentMuted}"/>
                                <TextBlock Text=" • " Margin="4,0" FontSize="10"
                                           Foreground="{DynamicResource AccentMuted}"/>
                                <TextBlock Name="lblSelectedCount" Text="0 seleccionadas"
                                           FontSize="10" Foreground="{DynamicResource AccentMuted}"/>
                            </StackPanel>
                        </StackPanel>
                        <Button Grid.Column="1" Name="btnClose"
                                Content="Cerrar" Width="60" Height="26"
                                Style="{StaticResource DatabaseButtonStyle}"
                                Margin="8,0,0,0"
                                HorizontalAlignment="Right"
                                VerticalAlignment="Center"/>
                    </Grid>
                </Border>

                <Border Grid.Row="1" Background="{DynamicResource PanelBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                        CornerRadius="8" Padding="6" Margin="0,0,0,6">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="80"/>
                            <ColumnDefinition Width="110"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="txtSearch" Grid.Column="0"
                                 Height="24" Padding="6,2"
                                 Margin="0,0,6,0"
                                 MinWidth="120"/>
                        <Button Name="btnClearSearch" Grid.Column="1"
                                Content="✖ Limpiar"
                                Height="24"
                                Style="{StaticResource DatabaseButtonStyle}"
                                Margin="0,0,4,0"/>
                        <Button Name="btnRefresh" Grid.Column="2"
                                Content="🔄 Actualizar"
                                Height="24"
                                ToolTip="Actualizar"
                                Style="{StaticResource DatabaseButtonStyle}"/>
                    </Grid>
                </Border>

                <Border Grid.Row="2" Background="{DynamicResource PanelBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                        CornerRadius="8" Margin="0,0,0,6">
                    <DataGrid Name="dgHistory"
                              IsReadOnly="True"
                              AutoGenerateColumns="False"
                              CanUserAddRows="False"
                              CanUserDeleteRows="False"
                              SelectionMode="Extended"
                              SelectionUnit="FullRow"
                              HeadersVisibility="Column"
                              GridLinesVisibility="Horizontal"
                              AlternatingRowBackground="{DynamicResource ControlBg}"
                              RowHeight="28"
                            FontFamily="Consolas"
                            FontSize="11"
                            Padding="4"
                            ScrollViewer.VerticalScrollBarVisibility="Auto"
                            ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="📅 Fecha" Binding="{Binding Timestamp}" Width="120">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Setter Property="Padding" Value="4,0"/>
                                        <Setter Property="FontFamily" Value="Consolas"/>
                                        <Setter Property="FontSize" Value="10"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="🖥️ Servidor" Binding="{Binding Server}" Width="110">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Setter Property="Padding" Value="4,0"/>
                                        <Setter Property="FontFamily" Value="Consolas"/>
                                        <Setter Property="FontSize" Value="10"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="🗄️ DB" Binding="{Binding Database}" Width="120">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Setter Property="Padding" Value="4,0"/>
                                        <Setter Property="FontFamily" Value="Consolas"/>
                                        <Setter Property="FontSize" Value="10"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="📝 Query" Binding="{Binding Preview}" Width="*">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Setter Property="Padding" Value="4,0"/>
                                        <Setter Property="TextWrapping" Value="NoWrap"/>
                                        <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                                        <Setter Property="FontFamily" Value="Consolas"/>
                                        <Setter Property="FontSize" Value="10"/>
                                        <Setter Property="ToolTipService.ShowDuration" Value="60000"/>
                                        <Setter Property="ToolTip">
                                            <Setter.Value>
                                                <ToolTip MaxWidth="900">
                                                    <TextBox Text="{Binding FullQuery}"
                                                            IsReadOnly="True"
                                                            TextWrapping="Wrap"
                                                            BorderThickness="0"
                                                            Background="Transparent"
                                                            FontFamily="Consolas"
                                                            FontSize="11"/>
                                                </ToolTip>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="✅ Estado" Binding="{Binding Result}" Width="140">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Setter Property="Padding" Value="4,0"/>
                                        <Setter Property="FontFamily" Value="Consolas"/>
                                        <Setter Property="FontSize" Value="10"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </Border>

                <Border Grid.Row="3" Background="{DynamicResource PanelBg}"
                        BorderBrush="{DynamicResource BorderBrushColor}" BorderThickness="1"
                        CornerRadius="8" Padding="6">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Orientation="Horizontal">
                            <Button Name="btnDeleteSelected" Content="🗑️ Eliminar"
                                    Width="90" Height="26"
                                    Style="{StaticResource DangerButtonStyle}"
                                    Margin="0,0,4,0"/>
                            <Button Name="btnClearAll" Content="🗑️ Todo"
                                    Width="70" Height="26"
                                    Style="{StaticResource DangerButtonStyle}"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button Name="btnCopy" Content="📋 Copiar"
                                    Width="80" Height="26"
                                    Style="{StaticResource DatabaseButtonStyle}"
                                    Margin="0,0,4,0"/>
                            <Button Name="btnLoadNew" Content="📥 Nueva Tab"
                                    Width="100" Height="26"
                                    Style="{StaticResource DatabaseButtonStyle}"/>
                        </StackPanel>
                    </Grid>
                </Border>

            </Grid>
        </Border>
    </Grid>
</Window>
"@
        $result = New-WpfWindow -Xaml $xaml -PassThru
        $window = $result.Window
        $controls = $result.Controls
        Set-DzWpfThemeResources -Window $window -Theme $theme
        $titleBar = $window.FindName("brdTitleBar")
        if ($titleBar) {
            $titleBar.Add_PreviewMouseLeftButtonDown({
                    param($sender, $e)
                    try {
                        $src = $e.OriginalSource
                        $dep = [System.Windows.DependencyObject]$src
                        while ($dep -ne $null) {
                            if ($dep -is [System.Windows.Controls.Button] -or
                                $dep -is [System.Windows.Controls.Primitives.ButtonBase] -or
                                $dep -is [System.Windows.Controls.TextBox] -or
                                $dep -is [System.Windows.Controls.Primitives.TextBoxBase] -or
                                $dep -is [System.Windows.Controls.ComboBox] -or
                                $dep -is [System.Windows.Controls.DataGrid]) {
                                return   # deja que el control maneje su click normal
                            }
                            $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
                        }
                        if ($e.ClickCount -eq 2) {
                            $window.WindowState = if ($window.WindowState -eq 'Maximized') { 'Normal' } else { 'Maximized' }
                            $e.Handled = $true
                            return
                        }
                        $window.DragMove()
                        $e.Handled = $true
                    } catch {}
                }.GetNewClosure())
        }
        $dataGrid = $controls['dgHistory']
        $txtSearch = $controls['txtSearch']
        $lblHistoryCount = $controls['lblHistoryCount']
        $lblSelectedCount = $controls['lblSelectedCount']
        $btnRefresh = $controls['btnRefresh']
        $btnClearSearch = $controls['btnClearSearch']
        $btnCopy = $controls['btnCopy']
        $btnLoadNew = $controls['btnLoadNew']
        $btnDeleteSelected = $controls['btnDeleteSelected']
        $btnClearAll = $controls['btnClearAll']
        $btnClose = $controls['btnClose']
        $fullHistory = New-Object System.Collections.ArrayList
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($fullHistory)
        $dataGrid.ItemsSource = $view
        $loadHistory = {
            try {
                $newHistory = @(Get-QueryHistory -MaxItems 100)
                $fullHistory.Clear()
                foreach ($item in $newHistory) {
                    $fullHistory.Add($item) | Out-Null
                }
                $lblHistoryCount.Text = "$($fullHistory.Count) consultas guardadas"
                $view.Refresh()
                Write-DzDebug "`t[DEBUG][Historial] Cargado: $($fullHistory.Count) items"
            } catch {
                Write-DzDebug "`t[DEBUG][Historial] Error cargando: $_" Red
            }
        }.GetNewClosure()
        $filterHistory = {
            param([string]$searchText)
            $search = if ([string]::IsNullOrWhiteSpace($searchText)) { $null } else { $searchText.ToLowerInvariant() }
            $view.Filter = {
                param($item)
                if (-not $search) { return $true }
                $query = ($item.FullQuery  | ForEach-Object { "$_" }).ToLowerInvariant()
                $preview = ($item.Preview    | ForEach-Object { "$_" }).ToLowerInvariant()
                $database = ($item.Database   | ForEach-Object { "$_" }).ToLowerInvariant()
                $server = ($item.Server     | ForEach-Object { "$_" }).ToLowerInvariant()
                return ($query.Contains($search) -or
                    $preview.Contains($search) -or
                    $database.Contains($search) -or
                    $server.Contains($search))
            }
            $view.Refresh()
        }.GetNewClosure()
        $txtSearch.Add_TextChanged({ & $filterHistory -searchText $txtSearch.Text }.GetNewClosure())
        $btnClearSearch.Add_Click({
                Write-DzDebug "`t[DEBUG][Historial] Limpiar búsqueda"
                $txtSearch.Text = ""
            })
        $btnRefresh.Add_Click({
                & $loadHistory
                $txtSearch.Text = ""
            }.GetNewClosure())
        $btnCopy.Add_Click({
                try {
                    $selected = $dataGrid.SelectedItem
                    if (-not $selected) {
                        Ui-Warn "Selecciona un query del historial" "Atención" $window
                        return
                    }
                    if (Get-Command Set-ClipboardTextSafe -ErrorAction SilentlyContinue) {
                        Set-ClipboardTextSafe -Text $selected.FullQuery -Owner $window | Out-Null
                    } else {
                        [System.Windows.Clipboard]::SetText($selected.FullQuery)
                    }
                    Write-Host "`n✓ Query copiado al portapapeles" -ForegroundColor Green
                    Write-Host "  Query: $($selected.FullQuery)" -ForegroundColor DarkGray
                } catch {
                    Write-Host "`n✗ Error copiando query: $_" -ForegroundColor Red
                    Ui-Error "Error copiando query: $($_.Exception.Message)" "Error" $window
                }
            }.GetNewClosure())
        $btnLoadNew.Add_Click({
                try {
                    $selected = $dataGrid.SelectedItem
                    if (-not $selected) {
                        Ui-Warn "Selecciona un query del historial" "Atención" $window
                        return
                    }
                    if (-not $global:tcQueries) {
                        Ui-Error "TabControl de queries no disponible" "Error" $window
                        return
                    }
                    $newTab = New-QueryTab -TabControl $global:tcQueries
                    if ($newTab -and $newTab.Tag -and $newTab.Tag.Editor) {
                        $editor = $newTab.Tag.Editor
                        Set-SqlEditorText -Editor $editor -Text $selected.FullQuery
                        $global:tcQueries.SelectedItem = $newTab
                        Write-Host "`n✓ Query cargado en nueva pestaña" -ForegroundColor Green
                        $window.Close()
                    }
                } catch {
                    Write-Host "`n✗ Error cargando query: $_" -ForegroundColor Red
                    Ui-Error "Error cargando query: $($_.Exception.Message)" "Error" $window
                }
            }.GetNewClosure())
        $btnDeleteSelected.Add_Click({
                try {
                    $selectedItems = @($dataGrid.SelectedItems)
                    if ($selectedItems.Count -eq 0) {
                        Ui-Warn "Selecciona uno o más queries para eliminar" "Atención" $window
                        return
                    }
                    $confirm = Ui-Confirm "¿Estás seguro de eliminar $($selectedItems.Count) consulta$(if ($selectedItems.Count -ne 1) { 's' } else { '' }) del historial?" "Confirmar eliminación" $window
                    if (-not $confirm) { return }
                    if (Remove-QueriesFromHistory -Items $selectedItems) {
                        & $loadHistory
                        Write-Host "`n✓ $($selectedItems.Count) consulta$(if ($selectedItems.Count -ne 1) { 's eliminadas' } else { ' eliminada' })" -ForegroundColor Green
                        & $filterHistory -searchText $txtSearch.Text
                        $dataGrid.UnselectAll()
                        & $updateSelectionCount
                    } else {
                        Ui-Error "Error eliminando queries del historial" "Error" $window
                    }
                } catch {
                    Write-Host "`n✗ Error eliminando queries: $_" -ForegroundColor Red
                    Ui-Error "Error: $($_.Exception.Message)" "Error" $window
                }
            }.GetNewClosure())
        $btnClearAll.Add_Click({
                $confirm = Ui-Confirm "¿Estás seguro de limpiar TODO el historial de queries?" "Confirmar" $window
                if ($confirm) {
                    if (Clear-QueryHistory) {
                        Write-Host "`n✓ Historial limpiado completamente" -ForegroundColor Green
                        & $loadHistory
                    } else {
                        Ui-Error "Error limpiando historial" "Error" $window
                    }
                }
            }.GetNewClosure())
        $btnClose.Add_Click({ $window.Close() })
        $dataGrid.Add_MouseDoubleClick({
                if ($dataGrid.SelectedItem) {
                    $btnLoadNew.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
                }
            })
        & $loadHistory
        $null = $window.ShowDialog()
    } catch {
        Write-Host "`n✗ Error mostrando ventana de historial: $_" -ForegroundColor Red
        Write-DzDebug "`t[DEBUG][Historial] Error: $($_.Exception.Message)" Red
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
    # Configuración básica
    $editor.ShowLineNumbers = $true
    $editor.FontFamily = $FontFamily
    $editor.FontSize = $FontSize
    $editor.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.Options.ConvertTabsToSpaces = $false
    $editor.SyntaxHighlighting = Get-SqlEditorHighlighting -HighlightingPath $paths.HighlightingPath
    # ===== FONDO GRIS UNIVERSAL =====
    $grayBackground = "#F5F5F5"  # Gris oscuro (estilo VS Code)#2D2D30
    $editor.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($grayBackground)
    $editor.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#000000")  # Texto claro
    # Color de selección
    $editor.TextArea.SelectionBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#264F78")
    $editor.TextArea.SelectionForeground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFFFFF")
    $editor.Options.IndentationSize = 4
    $editor.Options.ConvertTabsToSpaces = $true  # Espacios en lugar de tabs
    $editor.Options.EnableRectangularSelection = $true  # Selección rectangular (Alt+Drag)
    $editor.Options.EnableTextDragDrop = $true  # Arrastrar y soltar texto
    $editor.Options.ShowSpaces = $false  # Cambia a $true para ver espacios
    $editor.Options.ShowTabs = $true    # Cambia a $true para ver tabs
    $editor.Options.ShowEndOfLine = $false  # Cambia a $true para ver saltos de línea
    $editor.WordWrap = $true  # Cambia a $true si quieres que las líneas se ajusten
    $editor.Options.ColumnRulerPosition = 80
    $editor.Options.ShowColumnRuler = $false  # Cambia a $true para activar
    # FALTA IMPLEMENTAR
    # Esto requiere configuración adicional con BracketHighlightRenderer
    #$renderer = New-Object BracketHighlightRenderer
    #$editor.TextArea.TextView.BackgroundRenderers.Add($renderer)
    $editor.Options.HighlightCurrentLine = $true
    $editor.TextArea.TextView.CurrentLineBackground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#edecec")
    $editor.TextArea.TextView.CurrentLineBorder = $null
    $editor.Padding = [System.Windows.Thickness]::new(5, 2, 5, 2)
    $editor.Options.EnableHyperlinks = $false
    $editor.Options.EnableEmailHyperlinks = $false
    $editor.TextArea.Caret.CaretBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#AEAFAD")
    # FALTA IMPLEMENTAR
    # Requiere FoldingManager - lo configuramos después
    #$foldingManager = [ICSharpCode.AvalonEdit.Folding.FoldingManager]::Install($editor.TextArea)
    #$foldingStrategy = New-Object ICSharpCode.AvalonEdit.Folding.XmlFoldingStrategy
    #$foldingStrategy.UpdateFoldings($foldingManager, $editor.Document)

    if ($Container) {
        $Container.Child = $editor
    }
    # FALTA IMPLEMENTAR EL AUTOCOMPLETE
    #$completionWindow = $null
    #$editor.TextArea.Add_TextEntering({
    #        param($s, $e)
    #        if ($completionWindow -and $e.Text -eq " ") {
    #            $completionWindow.CompletionList.RequestInsertion($e)
    #        }
    #    })
    try {
        $searchPanel = [ICSharpCode.AvalonEdit.Search.SearchPanel]::Install($editor)
        # Personalizar colores del panel
        $searchPanel.MarkerBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FFD700")  # Amarillo
        $searchPanel.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($grayBackground)
        $searchPanel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#000000")
        Write-DzDebug "`t[DEBUG] Panel de búsqueda instalado correctamente" -Color Green
    } catch {
        Write-DzDebug "`t[DEBUG] Error instalando panel de búsqueda: $_" -Color Yellow
    }

    $editor.Add_KeyDown({
            param($sender, $e)

            if ($e.Key -eq [System.Windows.Input.Key]::F -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
                try {
                    if ($searchPanel) {
                        $searchPanel.IsReplaceMode = $false
                        $searchPanel.Open()
                        if (-not [string]::IsNullOrWhiteSpace($sender.SelectedText)) {
                            $searchPanel.SearchPattern = $sender.SelectedText
                        }
                    }
                } catch {
                    Write-DzDebug "`t[DEBUG] Error abriendo panel de búsqueda: $_" -Color Red
                }
                $e.Handled = $true
            }

            if ($e.Key -eq [System.Windows.Input.Key]::H -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control) {
                try {
                    if ($searchPanel) {
                        $searchPanel.IsReplaceMode = $true
                        $searchPanel.Open()
                        if (-not [string]::IsNullOrWhiteSpace($sender.SelectedText)) {
                            $searchPanel.SearchPattern = $sender.SelectedText
                        }
                    }
                } catch {
                    Write-DzDebug "`t[DEBUG] Error abriendo panel de reemplazo: $_" -Color Red
                }
                $e.Handled = $true
            }

            if ($e.Key -eq [System.Windows.Input.Key]::F3 -and [System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Shift) {
                try {
                    if ($searchPanel) { $searchPanel.FindPrevious() }
                } catch {}
                $e.Handled = $true
            } elseif ($e.Key -eq [System.Windows.Input.Key]::F3) {
                try {
                    if ($searchPanel) { $searchPanel.FindNext() }
                } catch {}
                $e.Handled = $true
            }

            if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                try {
                    if ($searchPanel -and $searchPanel.IsClosed -eq $false) {
                        $searchPanel.Close()
                        $e.Handled = $true
                    }
                } catch {}
            }
        }.GetNewClosure())
    return $editor
}
Export-ModuleMember -Function @(
    'Initialize-QueriesConfig',
    'Add-QueryToHistory',
    'Save-DbUiContext',
    'Get-QueryHistory',
    'Clear-QueryHistory',
    'Remove-QueriesFromHistory',
    'Save-OpenQueryTabs',
    'Restore-OpenQueryTabs',
    'Show-QueryHistoryWindow',
    'Execute-QueryUiSafe',
    'Export-ResultsUiSafe',
    'New-QueryTab',
    'Set-QueryTabsDatabase',
    'Close-OtherQueryTabs',
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
    'Get-SqlEditorSelectedText',
    'Show-QueryHistoryWindow'
)
