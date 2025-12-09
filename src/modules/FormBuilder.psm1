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

    [void]AddControl([string]$name, [System.Windows.Forms.Control]$control) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw [System.ArgumentException]::new("El nombre del control no puede estar vacío.", "name")
        }
        $this.Controls[$name] = $control
    }

    [System.Windows.Forms.Control]GetControl([string]$name) {
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

    # Construir el contenido de cada pestaña usando las funciones existentes
    $State = Build-ApplicationsTab -State $State -TabPage $tabAplicaciones
    $State = Build-DatabaseTab -State $State -TabPage $tabProSql

    # Agregar las pestañas al control
    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)

    # Devolver un objeto con el control y el estado actualizado
    return @{
        TabControl = $tabControl
        State      = $State
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

    $controls.ComboQueries = Create-ComboBox -Location (New-Object System.Drawing.Point(220, 20)) `
        -Size (New-Object System.Drawing.Size(320, 25)) -DropDownStyle DropDownList `
        -DefaultText "Selecciona una consulta predefinida"

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

    # Query browser controls
    $controls.LblQueries = Create-Label -Text "Consultas SQL" `
        -Location (New-Object System.Drawing.Point(220, 0)) `
        -Size (New-Object System.Drawing.Size(200, 20))

    $controls.RichTextQuery = New-Object System.Windows.Forms.RichTextBox
    $controls.RichTextQuery.Location = New-Object System.Drawing.Point(220, 50)
    $controls.RichTextQuery.Size = New-Object System.Drawing.Size(740, 200)
    $controls.RichTextQuery.Font = $defaultFont
    $controls.RichTextQuery.WordWrap = $false

    $controls.DataGridResults = New-Object System.Windows.Forms.DataGridView
    $controls.DataGridResults.Location = New-Object System.Drawing.Point(220, 260)
    $controls.DataGridResults.Size = New-Object System.Drawing.Size(740, 220)
    $controls.DataGridResults.ReadOnly = $true
    $controls.DataGridResults.AllowUserToAddRows = $false
    $controls.DataGridResults.AllowUserToDeleteRows = $false
    $controls.DataGridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

    $controls.ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $controls.CopyMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $controls.CopyMenuItem.Text = "Copiar celda"
    $controls.ContextMenu.Items.Add($controls.CopyMenuItem) | Out-Null
    $controls.DataGridResults.ContextMenuStrip = $controls.ContextMenu

    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    $controlsToAdd = $controls.Values | Where-Object { $_ -is [System.Windows.Forms.Control] }
    $TabPage.Controls.AddRange($controlsToAdd)

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