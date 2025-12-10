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
    $toolTip = $State.GetControl('ToolTip')
    if ($null -eq $toolTip) {
        Write-DzDebug "`t[DEBUG] ToolTip no encontrado en el State, creando uno nuevo"
        $toolTip = New-Object System.Windows.Forms.ToolTip
        $State.AddControl('ToolTip', $toolTip)
    }
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(990, 515)
    $tabControl.Location = New-Object System.Drawing.Point(5, 5)
    $tabAplicaciones = New-Object System.Windows.Forms.TabPage
    $tabAplicaciones.Text = "Aplicaciones"
    $tabProSql = New-Object System.Windows.Forms.TabPage
    $tabProSql.Text = "Base de datos"
    $tabProSql.AutoScroll = $true
    $State = Build-ApplicationsTab -State $State -TabPage $tabAplicaciones
    $State = Build-DatabaseTab -State $State -TabPage $tabProSql
    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)
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
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $controls = @{}
    # Crear todos los controles
    $controls.LblServer = Create-Label -Text "Instancia SQL:" -Location (New-Object System.Drawing.Point(10, 10))
    $controls.LblUser = Create-Label -Text "Usuario:" -Location (New-Object System.Drawing.Point(10, 50))
    $controls.LblPassword = Create-Label -Text "Contraseña:" -Location (New-Object System.Drawing.Point(10, 90))
    $controls.LblDatabase = Create-Label -Text "Base de datos" -Location (New-Object System.Drawing.Point(10, 130))
    $controls.LblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" -Location (New-Object System.Drawing.Point(10, 400)) -Size (New-Object System.Drawing.Size(180, 80)) -ForeColor ([System.Drawing.Color]::Red)
    $controls.LblQueries = Create-Label -Text "Consultas SQL" -Location (New-Object System.Drawing.Point(220, 0)) -Size (New-Object System.Drawing.Size(200, 20))
    $controls.ComboServer = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle "DropDown"
    $controls.ComboServer.Text = ".\NationalSoft"
    $controls.ComboDatabases = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 140)) -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle DropDownList
    $controls.ComboDatabases.Enabled = $false
    $controls.ComboQueries = Create-ComboBox -Location (New-Object System.Drawing.Point(220, 20)) -Size (New-Object System.Drawing.Size(320, 25)) -DropDownStyle DropDownList -DefaultText "Selecciona una consulta predefinida"
    $controls.ComboQueries.Enabled = $false
    $controls.TxtUser = Create-TextBox -Location (New-Object System.Drawing.Point(10, 60)) -Size (New-Object System.Drawing.Size(180, 20))
    $controls.TxtUser.Text = "sa"
    $controls.TxtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(10, 100)) -Size (New-Object System.Drawing.Size(180, 20)) -UseSystemPasswordChar $true
    $controls.RichTextQuery = New-Object System.Windows.Forms.RichTextBox
    $controls.RichTextQuery.Location = New-Object System.Drawing.Point(220, 50)
    $controls.RichTextQuery.Size = New-Object System.Drawing.Size(740, 200)
    $controls.RichTextQuery.Font = $defaultFont
    $controls.RichTextQuery.WordWrap = $false
    $controls.RichTextQuery.Enabled = $false
    $controls.DataGridResults = New-Object System.Windows.Forms.DataGridView
    $controls.DataGridResults.Location = New-Object System.Drawing.Point(220, 260)
    $controls.DataGridResults.Size = New-Object System.Drawing.Size(740, 220)
    $controls.DataGridResults.ReadOnly = $true
    $controls.DataGridResults.AllowUserToAddRows = $false
    $controls.DataGridResults.AllowUserToDeleteRows = $false
    $controls.DataGridResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $controls.DataGridResults.Enabled = $false
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $copyMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $copyMenuItem.Text = "Copiar celda"
    $contextMenu.Items.Add($copyMenuItem) | Out-Null
    $controls.DataGridResults.ContextMenuStrip = $contextMenu
    $controls.ContextMenu = $contextMenu
    $controls.CopyMenuItem = $copyMenuItem
    $controls.BtnExecute = Create-Button -Text "Ejecutar" -Location (New-Object System.Drawing.Point(340, 20)) -Size (New-Object System.Drawing.Size(100, 30))
    $controls.BtnExecute.Enabled = $false
    $controls.BtnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 220)) -Size (New-Object System.Drawing.Size(180, 30)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))

    $controls.BtnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(10, 260)) -Size (New-Object System.Drawing.Size(180, 30)) -BackColor ([System.Drawing.Color]::FromArgb(255, 180, 180))
    $controls.BtnDisconnectDb.Enabled = $false

    $controls.BtnBackup = Create-Button -Text "Backup de BDD" -Location (New-Object System.Drawing.Point(10, 300)) -Size (New-Object System.Drawing.Size(180, 30)) -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 200))
    $controls.BtnBackup.Enabled = $false

    # Agregar controles al State
    foreach ($name in $controls.Keys) {
        $State.AddControl($name, $controls[$name])
    }

    # Agregar controles al TabPage
    $controlsToAdd = $controls.Values | Where-Object {
        $_ -is [System.Windows.Forms.Control] -and
        $_.GetType() -ne [System.Windows.Forms.ContextMenuStrip]
    }

    if ($controlsToAdd.Count -gt 0) {
        $TabPage.Controls.AddRange($controlsToAdd)
    }

    # Hacer todos los controles VISIBLES
    foreach ($control in $TabPage.Controls) {
        if ($control -is [System.Windows.Forms.Control]) {
            $control.Visible = $true
            Write-DzDebug "`t[DEBUG] Haciendo visible control: $($control.GetType().Name) - '$($control.Text)'"
        } else {
            Write-DzDebug "`t[DEBUG] Control ignorado al hacer visible (tipo: $($control.GetType().FullName))"
        }
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
    return $State
}

function Initialize-FormControls {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UiState]$State
    )
    $serverCombo = $State.GetControl('ComboServer')
    if ($serverCombo -and (Get-Command Load-IniConnectionsToComboBox -ErrorAction SilentlyContinue)) {
        Load-IniConnectionsToComboBox -Combo $serverCombo
    }
    $comboQueries = $State.GetControl('ComboQueries')
    $rtbQuery = $State.GetControl('RichTextQuery')
    if ($comboQueries -and $rtbQuery -and (Get-Command Initialize-PredefinedQueries -ErrorAction SilentlyContinue)) {
        $predefinedQueries = Get-PredefinedQueries
        Initialize-PredefinedQueries -ComboQueries $comboQueries -RichTextBox $rtbQuery -Queries $predefinedQueries
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
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            })
        $disconnectButton.Add_Click({
                $statusLabel.Text = "Conectado a BDD: Ninguna"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
            })
    }
    return $State
}

function Make-AllControlsVisible {
    param(
        [System.Windows.Forms.TabPage]$TabPage
    )
    foreach ($control in $TabPage.Controls) {
        if ($control -is [System.Windows.Forms.Control]) {
            $control.Visible = $true
            Write-DzDebug "`t[DEBUG] Haciendo visible: $($control.GetType().Name) - Text: $($control.Text)"
            if ($control.HasChildren) {
                foreach ($child in $control.Controls) {
                    if ($child -is [System.Windows.Forms.Control]) {
                        $child.Visible = $true
                    }
                }
            }
        } else {
            Write-DzDebug "`t[DEBUG] Elemento sin propiedad Visible ignorado: $($control.GetType().FullName)"
        }
    }
    $TabPage.Visible = $true
}

function Add-ApplicationControls {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TabPage]$TabPage,
        [UiState]$State
    )
    Write-DzDebug "`t[DEBUG] En Add-ApplicationControls - State recibido: $($State -ne $null)"
    if ($null -eq $State) {
        Write-DzDebug "`t[DEBUG] ERROR: State es null en Add-ApplicationControls"
        return
    }
    $toolTip = $State.GetControl('ToolTip')
    Write-DzDebug "`t[DEBUG] ToolTip obtenido del State: $($toolTip -ne $null)"
    if ($null -eq $toolTip) {
        Write-DzDebug "`t[DEBUG] ERROR: ToolTip no encontrado en el State"
        $toolTip = New-Object System.Windows.Forms.ToolTip
        $State.AddControl('ToolTip', $toolTip)
        Write-DzDebug "`t[DEBUG] ToolTip creado de emergencia"
    }
    $lblHostname = Create-Label -Text ([System.Net.Dns]::GetHostName()) -Location (New-Object System.Drawing.Point(10, 1)) -Size (New-Object System.Drawing.Size(220, 40)) -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Hostname al portapapeles."
    $btnInstalarHerramientas = Create-Button -Text "Instalar Herramientas" -Location (New-Object System.Drawing.Point(10, 50)) -ToolTip "Abrir el menú de instaladores de Chocolatey."
    $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 90)) -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 130)) -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 170)) -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "De momento solo si es SQL 2014."
    $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 210)) -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Busca SQL Management en tu equipo y te confirma la versión previo a ejecutarlo."
    $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 250)) -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Herramienta de Star con funciones multiples para impresoras POS."
    $btnLectorDPicacls = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(10, 290)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
    $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Instalador LZMA" -Location (New-Object System.Drawing.Point(10, 330)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
    $btnConfigurarIPs = Create-Button -Text "Agregar IPs" -Location (New-Object System.Drawing.Point(10, 370)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
    $btnAddUser = Create-Button -Text "Agregar usuario de Windows" -Location (New-Object System.Drawing.Point(10, 410)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Crear nuevo usuario local en Windows"
    $btnForzarActualizacion = Create-Button -Text "Actualizar datos del sistema" -Location (New-Object System.Drawing.Point(10, 450)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Actualiza información de hardware del sistema"
    $lblPort = Create-Label -Text "Puerto: No disponible" -Location (New-Object System.Drawing.Point(250, 1)) -Size (New-Object System.Drawing.Size(220, 40)) -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Puerto al portapapeles."
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(250, 50)) -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(250, 90)) -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(250, 130)) -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
    $txt_IpAdress = Create-TextBox -Location (New-Object System.Drawing.Point(490, 1)) -Size (New-Object System.Drawing.Size(220, 40)) -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -ScrollBars 'Vertical' -Multiline $true -ReadOnly $true
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(490, 50)) -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
    $btnCambiarOTM = Create-Button -Text "Cambiar OTM a SQL/DBF" -Location (New-Object System.Drawing.Point(490, 90)) -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Cambiar la configuración entre SQL y DBF para On The Minute."
    $btnCheckPermissions = Create-Button -Text "Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(490, 130)) -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
    $btnCreateAPK = Create-Button -Text "Creación de SRM APK" -Location (New-Object System.Drawing.Point(490, 170)) -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Generar archivo APK para Comandero Móvil"
    $txt_AdapterStatus = Create-TextBox -Location (New-Object System.Drawing.Point(730, 1)) -Size(New-Object System.Drawing.Size(220, 40)) -BackColor([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -ScrollBars 'Vertical' -Multiline $true -ReadOnly  $true
    $global:txt_AdapterStatus = $txt_AdapterStatus
    if ($null -ne $toolTip) {
        $toolTip.SetToolTip($txt_AdapterStatus, "Lista de adaptadores y su estado. Haga clic en 'Actualizar adaptadores' para refrescar.")
        Write-DzDebug "`t[DEBUG] ToolTip asignado correctamente a txt_AdapterStatus"
    } else {
        Write-DzDebug "`t[DEBUG] ERROR: ToolTip sigue siendo nulo incluso después de intentar crearlo"
    }
    $State.AddControl('LblHostname', $lblHostname)
    $State.AddControl('LblPort', $lblPort)
    $State.AddControl('TxtIpAddress', $txt_IpAdress)
    $State.AddControl('TxtAdapterStatus', $txt_AdapterStatus)
    $State.AddControl('BtnInstalarHerramientas', $btnInstalarHerramientas)
    $State.AddControl('BtnProfiler', $btnProfiler)
    $State.AddControl('BtnDatabase', $btnDatabase)
    $State.AddControl('BtnSQLManager', $btnSQLManager)
    $State.AddControl('BtnSQLManagement', $btnSQLManagement)
    $State.AddControl('BtnPrinterTool', $btnPrinterTool)
    $State.AddControl('BtnLectorDPicacls', $btnLectorDPicacls)
    $State.AddControl('LZMAbtnBuscarCarpeta', $LZMAbtnBuscarCarpeta)
    $State.AddControl('BtnConfigurarIPs', $btnConfigurarIPs)
    $State.AddControl('BtnAddUser', $btnAddUser)
    $State.AddControl('BtnForzarActualizacion', $btnForzarActualizacion)
    $State.AddControl('BtnClearAnyDesk', $btnClearAnyDesk)
    $State.AddControl('BtnShowPrinters', $btnShowPrinters)
    $State.AddControl('BtnClearPrintJobs', $btnClearPrintJobs)
    $State.AddControl('BtnAplicacionesNS', $btnAplicacionesNS)
    $State.AddControl('BtnCheckPermissions', $btnCheckPermissions)
    $State.AddControl('BtnCambiarOTM', $btnCambiarOTM)
    $State.AddControl('BtnCreateAPK', $btnCreateAPK)
    $TabPage.Controls.AddRange(@(
            $lblHostname,
            $btnInstalarHerramientas,
            $btnProfiler,
            $btnDatabase,
            $btnSQLManager,
            $btnSQLManagement,
            $btnClearPrintJobs,
            $btnClearAnyDesk,
            $btnShowPrinters,
            $btnPrinterTool,
            $btnAplicacionesNS,
            $btnCheckPermissions,
            $btnCambiarOTM,
            $btnConfigurarIPs,
            $btnAddUser,
            $btnForzarActualizacion,
            $LZMAbtnBuscarCarpeta,
            $btnLectorDPicacls,
            $lblPort,
            $txt_IpAdress,
            $txt_AdapterStatus
        ))
    foreach ($control in $TabPage.Controls) {
        # Solo establecer Visible = $true en controles que tengan esa propiedad
        if ($control -is [System.Windows.Forms.Control] -and $control.GetType().Name -notmatch '(ContextMenu|Menu)') {
            $control.Visible = $true
        }
    }
}

Export-ModuleMember -Function New-FormState, Build-DatabaseTab, Build-ApplicationsTab,
Initialize-FormControls, Initialize-BasicEvents, Build-MainTabs,
Make-AllControlsVisible, Add-ApplicationControls