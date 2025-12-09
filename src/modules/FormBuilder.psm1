#requires -Version 5.0

class UiState {
    [hashtable]$Resources
    [hashtable]$Controls

    UiState([hashtable]$initialResources = $null) {
        $this.Resources = [hashtable]::new()

        if ($null -ne $initialResources) {
            if (-not ($initialResources -is [hashtable])) {
                throw [System.ArgumentException]::new("initialResources debe ser un hashtable.", "initialResources")
            }

            foreach ($key in $initialResources.Keys) {
                $this.Resources[$key] = $initialResources[$key]
            }
        }

        $this.Controls = [hashtable]::new()
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

function Build-DatabaseTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TabPage]$TabPage
    )

    $controls = @{}

    # Labels
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
        -Size (New-Object System.Drawing.Size(180, 80))

    # Comboboxes
    $controls.ComboServer = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) `
        -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle "DropDown"
    $controls.ComboServer.Text = ".\\NationalSoft"

    $controls.ComboDatabases = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 140)) `
        -Size (New-Object System.Drawing.Size(180, 20)) `
        -DropDownStyle DropDownList
    $controls.ComboDatabases.Enabled = $false

    # Textboxes
    $controls.TxtUser = Create-TextBox -Location (New-Object System.Drawing.Point(10, 60)) `
        -Size (New-Object System.Drawing.Size(180, 20))

    $controls.TxtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(10, 100)) `
        -Size (New-Object System.Drawing.Size(180, 20)) -UseSystemPasswordChar $true

    # Buttons
    $controls.BtnExecute = Create-Button -Text "Ejecutar" -Location (New-Object System.Drawing.Point(220, 20))
    $controls.BtnExecute.Size = New-Object System.Drawing.Size(100, 30)
    $controls.BtnExecute.Enabled = $false

    $controls.BtnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 220)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))

    $controls.BtnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(10, 260)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 180, 180))

    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    $TabPage.Controls.AddRange($controls.Values)

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

    $controls.BtnInstallSelected = Create-Button -Text "Instalar selección" `
        -Location (New-Object System.Drawing.Point(10, 45)) `
        -Size (New-Object System.Drawing.Size(180, 30))

    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    $TabPage.Controls.AddRange($controls.Values)

    return $State
}

function Initialize-FormControls {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )

    $serverCombo = $State.GetControl('ComboServer')
    if ($serverCombo) {
        Load-IniConnectionsToComboBox -Combo $serverCombo
    }

    $comboQueries = $State.GetControl('ComboQueries')
    $rtbQuery = $State.GetControl('RichTextQuery')
    if ($comboQueries -and $rtbQuery) {
        Initialize-PredefinedQueries -ComboQueries $comboQueries -RichTextBox $rtbQuery
    }

    if (Get-Command Initialize-BasicEvents -ErrorAction SilentlyContinue) {
        Initialize-BasicEvents -State $State
    }

    return $State
}

function Build-MainTabs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )

    $toolTip = New-Object System.Windows.Forms.ToolTip

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(990, 515)
    $tabControl.Location = New-Object System.Drawing.Point(5, 5)

    $tabAplicaciones = New-Object System.Windows.Forms.TabPage
    $tabAplicaciones.Text = "Aplicaciones"

    $tabProSql = New-Object System.Windows.Forms.TabPage
    $tabProSql.Text = "Base de datos"
    $tabProSql.AutoScroll = $true

    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)

    $lblServer = Create-Label -Text "Instancia SQL:" `
        -Location (New-Object System.Drawing.Point(10, 10)) `
        -Size (New-Object System.Drawing.Size(100, 10))

    $txtServer = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) `
        -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle "DropDown"
    $txtServer.Text = ".\\NationalSoft"
    Load-IniConnectionsToComboBox -Combo $txtServer

    $lblUser = Create-Label -Text "Usuario:" `
        -Location (New-Object System.Drawing.Point(10, 50)) `
        -Size (New-Object System.Drawing.Size(100, 10))

    $txtUser = Create-TextBox -Location (New-Object System.Drawing.Point(10, 60)) `
        -Size (New-Object System.Drawing.Size(180, 20))

    $lblPassword = Create-Label -Text "Contraseña:" `
        -Location (New-Object System.Drawing.Point(10, 90)) `
        -Size (New-Object System.Drawing.Size(100, 10))

    $txtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(10, 100)) `
        -Size (New-Object System.Drawing.Size(180, 20)) -UseSystemPasswordChar $true
    $txtUser.Text = "sa"

    $lblbdd_cmb = Create-Label -Text "Base de datos" `
        -Location (New-Object System.Drawing.Point(10, 130)) `
        -Size (New-Object System.Drawing.Size(100, 10))

    $cmbDatabases = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 140)) `
        -Size (New-Object System.Drawing.Size(180, 20)) `
        -DropDownStyle DropDownList
    $cmbDatabases.Enabled = $false

    $lblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" `
        -Location (New-Object System.Drawing.Point(10, 400)) `
        -Size (New-Object System.Drawing.Size(180, 80))

    $btnExecute = Create-Button -Text "Ejecutar" -Location (New-Object System.Drawing.Point(220, 20))
    $btnExecute.Size = New-Object System.Drawing.Size(100, 30)
    $btnExecute.Enabled = $false

    $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 220)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))

    $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(10, 260)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) `
        -Enabled $false

    $btnBackup = Create-Button -Text "Backup BDD" -Location (New-Object System.Drawing.Point(10, 300)) `
        -Size (New-Object System.Drawing.Size(180, 30)) `
        -BackColor ([System.Drawing.Color]::FromArgb(0, 192, 0)) `
        -ToolTip "Realizar backup de la base de datos seleccionada"

    $cmbQueries = New-Object System.Windows.Forms.ComboBox
    $cmbQueries.Location = New-Object System.Drawing.Point(330, 25)
    $cmbQueries.Size = New-Object System.Drawing.Size(350, 20)
    $cmbQueries.Enabled = $false

    $rtbQuery = New-Object System.Windows.Forms.RichTextBox
    $rtbQuery.Location = New-Object System.Drawing.Point(220, 60)
    $rtbQuery.Size = New-Object System.Drawing.Size(740, 140)
    $rtbQuery.Multiline = $true
    $rtbQuery.ScrollBars = 'Vertical'
    $rtbQuery.WordWrap = $true

    $dgvResults = New-Object System.Windows.Forms.DataGridView
    $dgvResults.Location = New-Object System.Drawing.Point(220, 205)
    $dgvResults.Size = New-Object System.Drawing.Size(740, 280)
    $dgvResults.ReadOnly = $true
    $dgvResults.AllowUserToAddRows = $false
    $dgvResults.AllowUserToDeleteRows = $false
    $dgvResults.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditProgrammatically
    $dgvResults.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $dgvResults.MultiSelect = $true
    $dgvResults.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
    $dgvResults.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightBlue
    $dgvResults.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $copyMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyMenu.Text = "Copiar selección"
    $contextMenu.Items.Add($copyMenu) | Out-Null
    $dgvResults.ContextMenuStrip = $contextMenu

    $panelGrid = New-Object System.Windows.Forms.Panel
    $panelGrid.Location = $dgvResults.Location
    $panelGrid.Size = $dgvResults.Size
    $panelGrid.AutoScroll = $true
    $dgvResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelGrid.Controls.Add($dgvResults)

    $btnConnectDb.Enabled = $True
    $btnBackup.Enabled = $false
    $btnDisconnectDb.Enabled = $false
    $btnExecute.Enabled = $false
    $rtbQuery.Enabled = $false
    $txtServer.Enabled = $true
    $txtUser.Enabled = $true
    $txtPassword.Enabled = $true
    $cmbQueries.Enabled = $false

    $tabProSql.Controls.AddRange(@(
            $btnConnectDb,
            $btnDisconnectDb,
            $cmbDatabases,
            $lblConnectionStatus,
            $btnExecute,
            $cmbQueries,
            $rtbQuery,
            $lblServer,
            $txtServer,
            $lblUser,
            $txtUser,
            $lblPassword,
            $txtPassword,
            $lblbdd_cmb,
            $panelGrid,
            $btnBackup
        ))

    $State.AddControl('ToolTip', $toolTip)
    $State.AddControl('TabControl', $tabControl)
    $State.AddControl('TabAplicaciones', $tabAplicaciones)
    $State.AddControl('TabBaseDatos', $tabProSql)
    $State.AddControl('LblServer', $lblServer)
    $State.AddControl('ComboServer', $txtServer)
    $State.AddControl('LblUser', $lblUser)
    $State.AddControl('TxtUser', $txtUser)
    $State.AddControl('LblPassword', $lblPassword)
    $State.AddControl('TxtPassword', $txtPassword)
    $State.AddControl('LblDatabase', $lblbdd_cmb)
    $State.AddControl('ComboDatabases', $cmbDatabases)
    $State.AddControl('LblConnectionStatus', $lblConnectionStatus)
    $State.AddControl('BtnExecute', $btnExecute)
    $State.AddControl('BtnConnectDb', $btnConnectDb)
    $State.AddControl('BtnDisconnectDb', $btnDisconnectDb)
    $State.AddControl('BtnBackup', $btnBackup)
    $State.AddControl('ComboQueries', $cmbQueries)
    $State.AddControl('RichTextQuery', $rtbQuery)
    $State.AddControl('DataGridResults', $dgvResults)
    $State.AddControl('PanelGrid', $panelGrid)
    $State.AddControl('CopyMenuItem', $copyMenu)
    $State.AddControl('ContextMenu', $contextMenu)

    return [PSCustomObject]@{
        State        = $State
        TabControl   = $tabControl
        Tabs         = @{ Aplicaciones = $tabAplicaciones; BaseDatos = $tabProSql }
        ToolTip      = $toolTip
        DataGridView = $dgvResults
    }
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
            })

        $disconnectButton.Add_Click({
                $statusLabel.Text = "Conectado a BDD: Ninguna"
            })
    }

    return $State
}

Export-ModuleMember -Function New-FormState, Build-DatabaseTab, Build-ApplicationsTab, Initialize-FormControls, Initialize-BasicEvents, Build-MainTabs
