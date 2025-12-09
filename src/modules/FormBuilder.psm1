#requires -Version 5.0

class UiState {
    [hashtable]$Resources
    [hashtable]$Controls

    UiState([hashtable]$resources) {
        if ($resources) {
            $this.Resources = $resources
        } else {
            $this.Resources = @{}
        }
        $this.Controls = @{}
    }

    [void]AddControl([string]$name, [object]$control) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw [System.ArgumentException]::new("El nombre del control no puede estar vacío.", "name")
        }
        $this.Controls[$name] = $control
    }

    [object]GetControl([string]$name) {
        if ($this.Controls.ContainsKey($name)) {
            return $this.Controls[$name]
        }
        return $null
    }

    [bool]TryGetControl([string]$name, [ref]$control) {
        if ($this.Controls.ContainsKey($name)) {
            $control.Value = $this.Controls[$name]
            return $true
        }
        $control.Value = $null
        return $false
    }
}

function New-FormState {
    [CmdletBinding()]
    param(
        [hashtable]$Resources = @{}
    )

    $state = [UiState]::new($Resources)
    return $state
}

function Build-MainTabs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )

    # Crear ToolTip compartido
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $State.AddControl('ToolTip', $toolTip)

    # Crear el control de pestañas
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(990, 515)
    $tabControl.Location = New-Object System.Drawing.Point(5, 5)

    # Crear las pestañas
    $tabAplicaciones = New-Object System.Windows.Forms.TabPage
    $tabAplicaciones.Text = "Aplicaciones"

    $tabProSql = New-Object System.Windows.Forms.TabPage
    $tabProSql.Text = "Base de datos"
    $tabProSql.AutoScroll = $true

    # Construir el contenido de cada pestaña
    $State = Build-ApplicationsTab -State $State -TabPage $tabAplicaciones
    $State = Build-DatabaseTab -State $State -TabPage $tabProSql

    # Agregar las pestañas al control
    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)

    # Devolver un objeto con el control, las tabs y el estado actualizado
    return @{
        TabControl = $tabControl
        Tabs       = @{
            Aplicaciones = $tabAplicaciones
            BaseDatos    = $tabProSql
        }
        State      = $State
        ToolTip    = $toolTip
    }
}

function Build-DatabaseTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TabPage]$TabPage
    )

    # Obtener fuentes predeterminadas (asumiendo que están definidas en el scope global)
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

    $controls = @{}

    # === LABELS ===
    $controls.LblServer = Create-Label -Text "Instancia SQL:" `
        -Location (New-Object System.Drawing.Point(10, 10))

    $controls.LblUser = Create-Label -Text "Usuario:" `
        -Location (New-Object System.Drawing.Point(10, 50))

    $controls.LblPassword = Create-Label -Text "Contraseña:" `
        -Location (New-Object System.Drawing.Point(10, 90))

    $controls.LblDatabase = Create-Label -Text "Base de datos" `
        -Location (New-Object System.Drawing.Point(10, 130))

    $controls.LblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" `
        -Location (New-Object System.Drawing.Point(10, 400)) `
        -Size (New-Object System.Drawing.Size(180, 80)) `
        -ForeColor ([System.Drawing.Color]::Red)

    $controls.LblQueries = Create-Label -Text "Consultas SQL" `
        -Location (New-Object System.Drawing.Point(220, 0)) `
        -Size (New-Object System.Drawing.Size(200, 20))

    # === COMBOBOXES ===
    $controls.ComboServer = Create-ComboBox `
        -Location (New-Object System.Drawing.Point(10, 20)) `
        -Size (New-Object System.Drawing.Size(180, 20)) `
        -DropDownStyle "DropDown"
    $controls.ComboServer.Text = ".\NationalSoft"

    $controls.ComboDatabases = Create-ComboBox `
        -Location (New-Object System.Drawing.Point(10, 140)) `
        -Size (New-Object System.Drawing.Size(180, 20)) `
        -DropDownStyle DropDownList
    $controls.ComboDatabases.Enabled = $false

    $controls.ComboQueries = Create-ComboBox `
        -Location (New-Object System.Drawing.Point(220, 20)) `
        -Size (New-Object System.Drawing.Size(320, 25)) `
        -DropDownStyle DropDownList `
        -DefaultText "Selecciona una consulta predefinida"
    $controls.ComboQueries.Enabled = $false

    # === TEXTBOXES ===
    $controls.TxtUser = Create-TextBox `
        -Location (New-Object System.Drawing.Point(10, 60)) `
        -Size (New-Object System.Drawing.Size(180, 20))
    $controls.TxtUser.Text = "sa"

    $controls.TxtPassword = Create-TextBox `
        -Location (New-Object System.Drawing.Point(10, 100)) `
        -Size (New-Object System.Drawing.Size(180, 20)) `
        -UseSystemPasswordChar $true

    # === RICHTEXTBOX para queries ===
    $controls.RichTextQuery = New-Object System.Windows.Forms.RichTextBox
    $controls.RichTextQuery.Location = New-Object System.Drawing.Point(220, 50)
    $controls.RichTextQuery.Size = New-Object System.Drawing.Size(740, 200)
    $controls.RichTextQuery.Font = $defaultFont
    $controls.RichTextQuery.WordWrap = $false
    $controls.RichTextQuery.Enabled = $false

    # === DATAGRIDVIEW ===
    $controls.DataGridResults = New-Object System.Windows.Forms.DataGridView
    $controls.DataGridResults.Location = New-Object System.Drawing.Point(220, 260)
    $controls.DataGridResults.Size = New-Object System.Drawing.Size(740, 220)
    $controls.DataGridResults.ReadOnly = $true
    $controls.DataGridResults.AllowUserToAddRows = $false
    $controls.DataGridResults.AllowUserToDeleteRows = $false
    $controls.DataGridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $controls.DataGridResults.Enabled = $false

    # === CONTEXT MENU (no es un Control, no se agrega al TabPage) ===
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $copyMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyMenuItem.Text = "Copiar celda"
    $contextMenu.Items.Add($copyMenuItem) | Out-Null
    $controls.DataGridResults.ContextMenuStrip = $contextMenu

    # Guardar referencia al menú contextual en el State (pero NO en la lista de controles a agregar)
    $State.AddControl('ContextMenu', $contextMenu)
    $State.AddControl('CopyMenuItem', $copyMenuItem)

    # === BUTTONS ===
    $controls.BtnExecute = Create-Button -Text "Ejecutar" `
        -Location (New-Object System.Drawing.Point(340, 20)) `
        -Size (New-Object System.Drawing.Size(100, 30))
    $controls.BtnExecute.Enabled = $false

    $controls.BtnConnectDb = Create-Button -Text "Conectar a BDD" `
        -Location (New-Object System.Drawing.Point(10, 220)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))

    $controls.BtnDisconnectDb = Create-Button -Text "Desconectar de BDD" `
        -Location (New-Object System.Drawing.Point(10, 260)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 180, 180))
    $controls.BtnDisconnectDb.Enabled = $false

    $controls.BtnBackup = Create-Button -Text "Backup de BDD" `
        -Location (New-Object System.Drawing.Point(10, 300)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 200))
    $controls.BtnBackup.Enabled = $false

    # === AGREGAR TODOS LOS CONTROLES AL STATE ===
    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    # === AGREGAR SOLO LOS CONTROLES VISUALES AL TABPAGE ===
    # Filtrar explícitamente solo los controles que heredan de System.Windows.Forms.Control
    $controlsToAdd = @()
    foreach ($ctrl in $controls.Values) {
        if ($ctrl -is [System.Windows.Forms.Control]) {
            $controlsToAdd += $ctrl
        }
    }

    if ($controlsToAdd.Count -gt 0) {
        $TabPage.Controls.AddRange($controlsToAdd)
    }

    return $State
}

function Build-ApplicationsTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TabPage]$TabPage
    )

    $controls = @{}

    $controls.LblAppsHeader = Create-Label -Text "Aplicaciones disponibles" `
        -Location (New-Object System.Drawing.Point(10, 10)) `
        -Size (New-Object System.Drawing.Size(300, 25))

    # Agregar controles al State
    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    # Agregar controles al TabPage
    $controlsToAdd = @()
    foreach ($ctrl in $controls.Values) {
        if ($ctrl -is [System.Windows.Forms.Control]) {
            $controlsToAdd += $ctrl
        }
    }

    if ($controlsToAdd.Count -gt 0) {
        $TabPage.Controls.AddRange($controlsToAdd)
    }

    return $State
}

function Initialize-FormControls {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )

    # Cargar conexiones desde INI
    $serverCombo = $State.GetControl('ComboServer')
    if ($serverCombo -and (Get-Command Load-IniConnectionsToComboBox -ErrorAction SilentlyContinue)) {
        Load-IniConnectionsToComboBox -Combo $serverCombo
    }

    # Inicializar queries predefinidas
    $comboQueries = $State.GetControl('ComboQueries')
    $rtbQuery = $State.GetControl('RichTextQuery')
    if ($comboQueries -and $rtbQuery -and (Get-Command Initialize-PredefinedQueries -ErrorAction SilentlyContinue)) {
        $predefinedQueries = Get-PredefinedQueries
        Initialize-PredefinedQueries -ComboQueries $comboQueries -RichTextBox $rtbQuery -Queries $predefinedQueries
    }

    # Inicializar eventos básicos
    if (Get-Command Initialize-BasicEvents -ErrorAction SilentlyContinue) {
        Initialize-BasicEvents -State $State
    }

    return $State
}

function Initialize-BasicEvents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )

    $connectButton = $State.GetControl('BtnConnectDb')
    $disconnectButton = $State.GetControl('BtnDisconnectDb')
    $statusLabel = $State.GetControl('LblConnectionStatus')

    if ($connectButton -and $disconnectButton -and $statusLabel) {
        $connectButton.Add_Click({
                $statusLabel.Text = "Intentando conectar..."
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            })

        $disconnectButton.Add_Click({
                $statusLabel.Text = "Conectado a BDD: Ninguna"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            })
    }

    return $State
}

Export-ModuleMember -Function New-FormState, Build-DatabaseTab, Build-ApplicationsTab, Initialize-FormControls, Initialize-BasicEvents, Build-MainTabs