# Crear la carpeta 'C:\Temp' si no existe
if (!(Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    Write-Host "Carpeta 'C:\Temp' creada correctamente."
}
# Mostrar advertencia ALFA y solicitar confirmación
Write-Host "`n==============================================" -ForegroundColor Red
Write-Host "           ADVERTENCIA DE VERSIÓN ALFA          " -ForegroundColor Red
Write-Host "==============================================" -ForegroundColor Red
Write-Host "Esta aplicación se encuentra en fase de desarrollo ALFA.`n" -ForegroundColor Yellow
Write-Host "Algunas funciones pueden realizar cambios irreversibles en: `n"
Write-Host " - Su equipo" -ForegroundColor Red
Write-Host " - Bases de datos" -ForegroundColor Red
Write-Host " - Configuraciones del sistema`n" -ForegroundColor Red
Write-Host "¿Acepta ejecutar esta aplicación bajo su propia responsabilidad? (Y/N)" -ForegroundColor Yellow
$response = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
while ($response.Character -notin 'Y','N') {
    $response = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
if ($response.Character -ne 'Y') {
    Write-Host "`nEjecución cancelada por el usuario.`n" -ForegroundColor Red
    exit
}
Clear-Host
Write-Host "El usuario aceptó los riesgos. Corriendo programa..." -ForegroundColor Green
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
# Crear el formulario
    $formPrincipal = New-Object System.Windows.Forms.Form
    $formPrincipal.Size = New-Object System.Drawing.Size(1000, 600)  # Aumentado de 720x400
    $formPrincipal.StartPosition = "CenterScreen"
    $formPrincipal.BackColor = [System.Drawing.Color]::White
    $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formPrincipal.MaximizeBox = $false
    $formPrincipal.MinimizeBox = $false
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                                                                                                        $version = "Alfa 250521.1413"  # Valor predeterminado para la versión
    $formPrincipal.Text = "Daniel Tools v$version"
    Write-Host "`n=============================================" -ForegroundColor DarkCyan
    Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
    Write-Host "              Versión: v$($version)               " -ForegroundColor Green
    Write-Host "=============================================" -ForegroundColor DarkCyan
    Write-Host "`nTodos los derechos reservados para Daniel Tools." -ForegroundColor Cyan
    Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan
# Creación maestra de botones
    $toolTip = New-Object System.Windows.Forms.ToolTip
function Create-Button {
                param (
                    [string]$Text,
                    [System.Drawing.Point]$Location,
                    [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,
                    [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
                    [string]$ToolTipText = $null,
                    [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35)),
[System.Drawing.Font]$Font = $defaultFont, # Agregar parámetro Font con valor predeterminado
                    [bool]$Enabled = $true
                )
                $buttonStyle = @{
                    FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
                    Font      = $defaultFont
                }
                $button_MouseEnter = {
                    $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)  # Cambia el color al pasar el mouse
                    $this.Font = $boldFont
                }
                $button_MouseLeave = {
                    $this.BackColor = $this.Tag  # Restaura el color original almacenado en Tag
                    $this.Font = $defaultFont
                }
                $button = New-Object System.Windows.Forms.Button
                $button.Text = $Text
                $button.Size = $Size  # Usar el tamaño personalizado
                $button.Location = $Location
                $button.BackColor = $BackColor
                $button.ForeColor = $ForeColor
$button.Font = $Font # Usar el parámetro Font
                $button.FlatStyle = $buttonStyle.FlatStyle
                $button.Tag = $BackColor  # Almacena el color original en Tag
                $button.Add_MouseEnter($button_MouseEnter)
                $button.Add_MouseLeave($button_MouseLeave)
                $button.Enabled = $Enabled
                if ($ToolTipText) {
                    $toolTip.SetToolTip($button, $ToolTipText)
                }
                return $button
            }
#Lo mismo pero para las labels
function Create-Label {
                    param (
                        [string]$Text,
                        [System.Drawing.Point]$Location,
                        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::Transparent,
                        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
                        [string]$ToolTipText = $null,
                        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
                        [System.Drawing.Font]$Font = $defaultFont,
                        [System.Windows.Forms.BorderStyle]$BorderStyle = [System.Windows.Forms.BorderStyle]::None,
                        [System.Drawing.ContentAlignment]$TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    )
                    $label = New-Object System.Windows.Forms.Label
                    $label.Text = $Text
                    $label.Size = $Size
                    $label.Location = $Location
                    $label.BackColor = $BackColor
                    $label.ForeColor = $ForeColor
                    $label.Font = $Font
                    $label.BorderStyle = $BorderStyle
                    $label.TextAlign = $TextAlign
                    if ($ToolTipText) { $toolTip.SetToolTip($label, $ToolTipText) }
                    return $label
                }
    function Create-Form {
                param (
                    [string]$Title,
                    [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(350, 200)),
                    [System.Windows.Forms.FormStartPosition]$StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen,
                    [System.Windows.Forms.FormBorderStyle]$FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog,
                    [bool]$MaximizeBox = $false,
                    [bool]$MinimizeBox = $false
                )
                $form = New-Object System.Windows.Forms.Form
                $form.Text = $Title
                $form.Size = $Size
                $form.StartPosition = $StartPosition
                $form.FormBorderStyle = $FormBorderStyle
                $form.MaximizeBox = $MaximizeBox
                $form.MinimizeBox = $MinimizeBox
                return $form
    }
function Create-ComboBox {
            param (
                [System.Drawing.Point]$Location,
                [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
                [System.Windows.Forms.ComboBoxStyle]$DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList,
                [System.Drawing.Font]$Font = $defaultFont,
                [string[]]$Items = @(),
                [int]$SelectedIndex = -1,
                [string]$DefaultText = $null
            )
            $comboBox = New-Object System.Windows.Forms.ComboBox
            $comboBox.Location = $Location
            $comboBox.Size = $Size
            $comboBox.DropDownStyle = $DropDownStyle
            $comboBox.Font = $Font
            if ($Items.Count -gt 0) {
                $comboBox.Items.AddRange($Items)
                $comboBox.SelectedIndex = $SelectedIndex
            }
            if ($DefaultText) {
                $comboBox.Text = $DefaultText
            }
            return $comboBox
}
function Create-TextBox {
    param (
        [System.Drawing.Point]$Location,
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
        [System.Drawing.Font]$Font = $defaultFont,
        [string]$Text = "",
        [bool]$Multiline = $false,
        [System.Windows.Forms.ScrollBars]$ScrollBars = [System.Windows.Forms.ScrollBars]::None,
        [bool]$ReadOnly = $false,
        [bool]$UseSystemPasswordChar = $false
    )
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = $Location
    $textBox.Size = $Size
    $textBox.BackColor = $BackColor
    $textBox.ForeColor = $ForeColor
    $textBox.Font = $Font
    $textBox.Text = $Text
    $textBox.Multiline = $Multiline
    $textBox.ScrollBars = $ScrollBars
    $textBox.ReadOnly = $ReadOnly
    $textBox.WordWrap = $false
    if ($UseSystemPasswordChar) {
        $textBox.UseSystemPasswordChar = $true
    }
    return $textBox
}
#COMIENZA EL PROGRAMA ----------------------------------
$global:lastReportedPct = -1  # Añadir al inicio del script
$tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(990, 515)    # Original: 710x315
    $tabControl.Location = New-Object System.Drawing.Point(5, 5)   # Pequeño margen
    $tabAplicaciones = New-Object System.Windows.Forms.TabPage
    $tabAplicaciones.Text = "Aplicaciones"
    $tabProSql = New-Object System.Windows.Forms.TabPage
    $tabProSql.Text = "Base de datos"
    $tabProSql.AutoScroll = $true  # Habilitar scrollbar si el contenido excede el área
    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)
# Panel izquierdo (controles manualmente posicionados)
$lblServer = Create-Label -Text "Instancia SQL:" `
    -Location (New-Object System.Drawing.Point(10, 10)) `
    -Size (New-Object System.Drawing.Size(100, 10))
    $txtServer = Create-TextBox -Location (New-Object System.Drawing.Point(10, 20)) `
        -Size (New-Object System.Drawing.Size(180, 20))
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
        $txtServer.Text = ".\NationalSoft"
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
# Crear RichTextBox
$rtbQuery = New-Object System.Windows.Forms.RichTextBox
    $rtbQuery.Location   = New-Object System.Drawing.Point(220, 60)
    $rtbQuery.Size = New-Object System.Drawing.Size(740, 140)     # Mayor ancho
    $rtbQuery.Multiline  = $true
    $rtbQuery.ScrollBars = 'Vertical'
    $rtbQuery.WordWrap   = $true
    $keywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE'
    $rtbQuery.Add_TextChanged({
        $pos = $rtbQuery.SelectionStart
        $rtbQuery.SuspendLayout()
    
        # 1. Restablecer todo a negro
        $rtbQuery.SelectAll()
        $rtbQuery.SelectionColor = [System.Drawing.Color]::Black
    
        # 2. Encontrar y resaltar comentarios de línea (--) y almacenar rangos
        $commentRanges = @()
        foreach ($c in [regex]::Matches($rtbQuery.Text, '--.*', 'Multiline')) {
            $rtbQuery.Select($c.Index, $c.Length)
            $rtbQuery.SelectionColor = [System.Drawing.Color]::Green
            $commentRanges += [PSCustomObject]@{ Start = $c.Index; End = $c.Index + $c.Length }
        }
    
        # 3. Encontrar y resaltar comentarios de bloque (/* ... */) y agregar rangos
        foreach ($b in [regex]::Matches($rtbQuery.Text, '/\*[\s\S]*?\*/', 'Multiline')) {
            $rtbQuery.Select($b.Index, $b.Length)
            $rtbQuery.SelectionColor = [System.Drawing.Color]::Green
            $commentRanges += [PSCustomObject]@{ Start = $b.Index; End = $b.Index + $b.Length }
        }
    
        # 4. Resaltar keywords en azul sólo fuera de comentarios
        foreach ($m in [regex]::Matches($rtbQuery.Text, "\b($keywords)\b", 'IgnoreCase')) {
            $inComment = $commentRanges | Where-Object { $m.Index -ge $_.Start -and $m.Index -lt $_.End }
            if (-not $inComment) {
                $rtbQuery.Select($m.Index, $m.Length)
                $rtbQuery.SelectionColor = [System.Drawing.Color]::Blue
            }
        }
    
        # 5. Restaurar posición del cursor
        $rtbQuery.Select($pos, 0)
        $rtbQuery.ResumeLayout()
    })

# Configuración del DataGridView
# Configuración del DataGridView
$dgvResults = New-Object System.Windows.Forms.DataGridView
$dgvResults.Location                   = New-Object System.Drawing.Point(220, 205)
$dgvResults.Size                       = New-Object System.Drawing.Size(740, 280)
$dgvResults.ReadOnly                   = $true
$dgvResults.AllowUserToAddRows         = $false
$dgvResults.AllowUserToDeleteRows      = $false
$dgvResults.EditMode                   = [System.Windows.Forms.DataGridViewEditMode]::EditProgrammatically
$dgvResults.SelectionMode              = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
$dgvResults.MultiSelect                = $true
$dgvResults.ClipboardCopyMode          = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$dgvResults.DefaultCellStyle.SelectionBackColor  = [System.Drawing.Color]::LightBlue
$dgvResults.DefaultCellStyle.SelectionForeColor  = [System.Drawing.Color]::Black

# Capturar estilos originales
$script:originalForeColor = $dgvResults.DefaultCellStyle.ForeColor
$script:originalHeaderBackColor = $dgvResults.ColumnHeadersDefaultCellStyle.BackColor
$script:originalAutoSizeMode = $dgvResults.AutoSizeColumnsMode

# 1. Crear menú contextual
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$copyMenu    = New-Object System.Windows.Forms.ToolStripMenuItem
$copyMenu.Text = "Copiar selección"
$contextMenu.Items.Add($copyMenu) | Out-Null
$dgvResults.ContextMenuStrip = $contextMenu

# 2. Evento Click en “Copiar selección”
$copyMenu.Add_Click({
    if ($dgvResults.GetCellCount("Selected") -gt 0) {
        # Obtener texto tabulado de todas las celdas seleccionadas
        $dataObj = $dgvResults.GetClipboardContent()
        if ($dataObj) {
            [Windows.Forms.Clipboard]::SetText($dataObj.GetText())
        }
    }
})

# 3. Habilitar Ctrl+C para copiar selección
$dgvResults.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        $dataObj = $dgvResults.GetClipboardContent()
        if ($dataObj) {
            [Windows.Forms.Clipboard]::SetText($dataObj.GetText())
        }
        $e.Handled = $true
    }
})

# 4. Selección con clic derecho puntual
$dgvResults.Add_MouseDown({
    param($sender, $args)
    if ($args.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hit = $dgvResults.HitTest($args.X, $args.Y)
        if ($hit.RowIndex -ge 0 -and $hit.ColumnIndex -ge 0) {
            $dgvResults.CurrentCell = $dgvResults.Rows[$hit.RowIndex].Cells[$hit.ColumnIndex]
            # Retener otras celdas si Ctrl está presionado
            if (-not $args.Modifiers.HasFlag([System.Windows.Forms.Keys]::Control)) {
                $dgvResults.ClearSelection()
            }
            $dgvResults.Rows[$hit.RowIndex].Cells[$hit.ColumnIndex].Selected = $true
        }
    }
})
#PANEL
$panelGrid = New-Object System.Windows.Forms.Panel
    $panelGrid.Location = $dgvResults.Location
    $panelGrid.Size = $dgvResults.Size
    $panelGrid.AutoScroll = $true
    $dgvResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panelGrid.Controls.Add($dgvResults)
#DISABLE:
        $btnConnectDb.Enabled    = $True
        $btnBackup.Enabled       = $false
        $btnDisconnectDb.Enabled = $false
        $btnExecute.Enabled      = $false
        $rtbQuery.Enabled        = $false
        $txtServer.Enabled = $true
        $txtUser.Enabled = $true
        $txtPassword.Enabled = $true
        $btnExecute.Enabled = $false
        $cmbQueries.Enabled = $false

$tabProSql.Controls.AddRange(@(
    $btnConnectDb,
    $btnDisconnectDb,
    $cmbDatabases,  # <-- Aquí el ComboBox reemplaza al ListBox
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
# Hashtable con consultas predefinidas usando here-strings
$script:predefinedQueries = @{
"Monitor de Servicios | Ventas a subir" = @"
SELECT DISTINCT TOP (10)
    nofacturable,
    tablaventa.IDEMPRESA,
    codigo_unico_af AS ticket_cu,
    e.CLAVEUNICAEMPRESA AS empresa_id,
    numcheque AS ticket_folio,
    seriefolio AS ticket_serie,
    tablaventa.FOLIO AS ticket_folioSR,
    CONVERT(nvarchar(30), FECHA, 120) AS ticket_fecha,
    subtotal AS ticket_subtotal,
    total AS ticket_total,
    descuento AS ticket_descuento,
    totalconpropina AS ticket_totalconpropina,
    totalsindescuento AS ticket_totalsindescuento,
    (totalimpuestod1 + totalimpuestod2 + totalimpuestod3) AS ticket_totalimpuesto,
    totalotros AS ticket_totalotros,
    descuentoimporte AS ticket_totaldescuento,
    0 AS ticket_totaldescuento2,
    0 AS ticket_totaldescuento3,
    0 AS ticket_totaldescuento4,
    tablaventa.PROPINA AS ticket_propina,
    cancelado,
    CAST(numcheque AS VARCHAR) AS ticket_numcheque,
    numerotarjeta,
    puntosmonederogenerados,
    titulartarjetamonederodescuento AS titulartarjetamonedero,
    tarjetadescuento,
    descuentomonedero,
    e.idregimen_sat,
    tipopago.idformapago_SAT,
    tablaventa.idturno,
    nopersonas,
    tipodeservicio,
    idmesero,
    totalarticulos,
    LTRIM(RTRIM(estacion)) AS ticket_estacion,
    usuariodescuento,
    comentariodescuento,
    tablaventa.idtipodescuento,
    totalimpuestod2 AS TicketTotalIEPS,
    0 AS TicketTotalOtrosImpuestos,
    LEFT(CONVERT(VARCHAR, fecha + 15, 120), 10) + ' 23:59:59' AS ticket_fechavence,
    CONVERT(nvarchar(30), tablaventa.cierre, 120) AS ticket_fecha_cierre
FROM
    CHEQUES AS tablaventa
    INNER JOIN empresas AS e ON tablaventa.IDEMPRESA = e.IDEMPRESA
    LEFT JOIN chequespagos AS tablapago ON tablapago.folio = tablaventa.folio
    LEFT JOIN formasdepago AS tipopago ON tablapago.idformadepago = tipopago.idformadepago
WHERE
    fecha > (SELECT fecha_inicio_envio FROM configuracion_ws)
    AND (intentoEnvioAF < 20)
    AND (
        (enviado = 0) OR (enviado IS NULL)
    )
    AND (
        (pagado = 1 AND nofacturable = 0)
        OR cancelado = 1
    )
    AND codigo_unico_af IS NOT NULL
    AND codigo_unico_af <> ''
    AND tablaventa.IDEMPRESA = (SELECT TOP 1 idempresa FROM empresas)
ORDER BY
    numcheque;
"@
"BackOffice Actualizar contraseña  administrador" = @"
    -- Actualiza la contraseña del primer UserName con rol administrador y retorna el UserName actualizado
UPDATE users
    SET Password = '08/Vqq0='
    OUTPUT inserted.UserName 
    WHERE UserName = (SELECT TOP 1 UserName FROM users WHERE IsSuperAdmin = 1 and IsEnabled = 1);
"@
"BackOffice Estaciones" = @"
SELECT 
    t.Name, 
    t.Ip, 
    t.LastOnline, 
    t.IsEnabled, 
    u.UserName AS UltimoUsuario, 
    t.AppVersion, 
    t.IsMaximized, 
    t.ForceAppUpdate, 
    t.SkipDoPing
FROM Terminals t
LEFT JOIN Users u ON t.LastUserLogin = u.Id
--WHERE t.IsEnabled = 1.0000 
ORDER BY t.IsEnabled DESC, t.Name;
"@
"SR | Actualizar contraseña de administrador" = @"
    -- Actualiza la contraseña del primer usuario con rol administrador y retorna el usuario actualizado
    UPDATE usuarios 
    SET contraseña = 'A9AE4E13D2A47998AC34' 
    OUTPUT inserted.usuario 
    WHERE usuario = (SELECT TOP 1 usuario FROM usuarios WHERE administrador = 1);
"@
"SR | Revisar Pivot Table" = @"
    SELECT app_id, field, COUNT(*) 
    FROM app_settings 
    GROUP BY app_id, field 
    HAVING COUNT(*) > 1
/* Consulta SQL para eliminar duplicados
        BEGIN TRANSACTION;
                                                    WITH CTE AS (
                                                        SELECT id, app_id, field,
                                                               ROW_NUMBER() OVER (PARTITION BY app_id, field ORDER BY id DESC) AS rn
                                                        FROM app_settings
                                                    )
                                                    DELETE FROM app_settings
                                                    WHERE id IN (
                                                        SELECT id FROM CTE WHERE rn > 1
                                                    );
                                                    COMMIT TRANSACTION;
*/
"@
"SR | Fecha Revisiones" = @"
    WITH CTE AS (
        SELECT 
            b.estacion, 
            b.fecha       AS UltimoUso, 
            ROW_NUMBER() OVER (PARTITION BY b.estacion ORDER BY b.fecha DESC) AS rn 
        FROM bitacorasistema b
    )
    SELECT 
        e.FECHAREV, 
        c.estacion, 
        c.UltimoUso 
    FROM CTE c 
    JOIN estaciones e 
        ON c.estacion = e.idestacion 
    WHERE c.rn = 1 
    ORDER BY c.UltimoUso DESC;
"@
"OTM | Eliminar Server en OTM" = @"
    SELECT serie, ipserver, nombreservidor 
    FROM configuracion;  
    -- UPDATE configuracion 
    --   SET serie='', ipserver='', nombreservidor=''
"@
"NSH | Eliminar Server en Hoteles" = @"
    SELECT serievalida, numserie, ipserver, nombreservidor, llave 
    FROM configuracion; 
    -- UPDATE configuracion 
    --   SET serievalida='', numserie='', ipserver='', nombreservidor='', llave=''
"@
"Restcard | Eliminar Server en Rest Card" = @"
    -- update tabvariables 
    --   SET estacion='', ipservidor='';
"@
"sql | Listar usuarios e idiomas" = @"
    -- Lista los usuarios del sistema y su idioma configurado
SELECT 
    p.name AS Usuario, 
    l.default_language_name AS Idioma
FROM 
    sys.server_principals p
LEFT JOIN 
    sys.sql_logins l ON p.principal_id = l.principal_id
WHERE 
    p.type IN ('S', 'U') -- Usuarios SQL y Windows
"@
}
#Agregando los Queries en orden:
    $sortedKeys = $script:predefinedQueries.Keys | Sort-Object
    $cmbQueries.Items.Clear()
    foreach ($key in $sortedKeys) {
        $cmbQueries.Items.Add($key) | Out-Null
    }
$cmbQueries.Add_SelectedIndexChanged({
    $rtbQuery.Text = $script:predefinedQueries[$cmbQueries.SelectedItem]
})
# COLUMNA 1 | INSTALADORES EJECUTABLES | X:10
    $lblHostname = Create-Label -Text ([System.Net.Dns]::GetHostName()) -Location (New-Object System.Drawing.Point(10, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Hostname al portapapeles."
    $btnInstalarHerramientas = Create-Button -Text "Instalar Herramientas" -Location (New-Object System.Drawing.Point(10, 50)) `
        -ToolTip "Abrir el menú de instaladores de Chocolatey."
    $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 90)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "De momento solo si es SQL 2014."
    $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 210)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Busca SQL Management en tu equipo y te confirma la versión previo a ejecutarlo."
    $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 250)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(224, 224, 224)) -ToolTip "Herramienta de Star con funciones multiples para impresoras POS."
#COLUMNA 2 | FUNCIONES Y SERVICIOS DE WINDOWS | X: 250
    $lblPort = Create-Label -Text "Puerto: No disponible" -Location (New-Object System.Drawing.Point(250, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Puerto al portapapeles."
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(250, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(250, 90)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(250, 130)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
#COLUMNA 3 | FUNCIONES NATIONAL SOFT | X: 490
    $txt_IpAdress = Create-TextBox -Location (New-Object System.Drawing.Point(490, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) `
        -ScrollBars 'Vertical' -Multiline $true -ReadOnly $true
    $btnLectorDPicacls = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(490, 90)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
    $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Instalador LZMA" -Location (New-Object System.Drawing.Point(490, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
    $btnConfigurarIPs = Create-Button -Text "Agregar IPs" -Location (New-Object System.Drawing.Point(490, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
    $btnAddUser = Create-Button -Text "Agregar usuario de Windows" -Location (New-Object System.Drawing.Point(490, 210)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Crear nuevo usuario local en Windows"
    $btnForzarActualizacion = Create-Button -Text "Actualizar datos del sistema" -Location (New-Object System.Drawing.Point(490, 250)) `
        -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Actualiza información de hardware del sistema"
# Columna 4: | FIXES Y NOVEDADES |  X:730
$txt_AdapterStatus = Create-TextBox -Location (New-Object System.Drawing.Point(730, 1)) -Size(New-Object System.Drawing.Size(220, 40)) `
    -BackColor([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) `
    -ScrollBars 'Vertical' -Multiline $true -ReadOnly  $true
    $toolTip.SetToolTip($txt_AdapterStatus, "Lista de adaptadores y su estado. Haga clic en 'Actualizar adaptadores' para refrescar.")
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(730, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
    $btnCambiarOTM = Create-Button -Text "Cambiar OTM a SQL/DBF" -Location (New-Object System.Drawing.Point(730, 90)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Cambiar la configuración entre SQL y DBF para On The Minute."
    $btnCheckPermissions = Create-Button -Text "Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(730, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
    $btnCreateAPK = Create-Button -Text "Creación de SRM APK" -Location (New-Object System.Drawing.Point(730, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Generar archivo APK para Comandero Móvil"

#FUERA DEL TAB
    $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(350, 525)) `
                                -Size (New-Object System.Drawing.Size(500, 30)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
# Agregar botones a la pestaña de aplicaciones
$tabAplicaciones.Controls.AddRange(@(
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
            $lblHostname,
            $lblPort,
            $txt_IpAdress,
            $txt_AdapterStatus,
            $btnCreateAPK
))
# Función para manejar MouseEnter y cambiar el color
$changeColorOnHover = {
    param($sender, $eventArgs)
    $sender.BackColor = [System.Drawing.Color]::Orange
}
# Función para manejar MouseLeave y restaurar el color
$restoreColorOnLeave = {
    param($sender, $eventArgs)
    $sender.BackColor = [System.Drawing.Color]::Black
}
        $lblHostname.Add_MouseEnter($changeColorOnHover)
        $lblHostname.Add_MouseLeave($restoreColorOnLeave)

#FUNCIONES:
# Función para revisar permisos y agregar Full Control a "Everyone" si es necesario
    function Check-Permissions {
                $folderPath = "C:\NationalSoft"
                $acl = Get-Acl -Path $folderPath
                $permissions = @()
                # Obtener el SID universal de "Everyone" (independiente del idioma del sistema)
                $everyoneSid = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::WorldSid, $null)
                $everyonePermissions = @()
                $everyoneHasFullControl = $false
                foreach ($access in $acl.Access) {
                    # Obtener el SID del usuario en la ACL
                    $userSid = (New-Object System.Security.Principal.NTAccount($access.IdentityReference)).Translate([System.Security.Principal.SecurityIdentifier])
                    # Almacenar los permisos de todos los usuarios
                    $permissions += [PSCustomObject]@{
                        Usuario = $access.IdentityReference
                        Permiso = $access.FileSystemRights
                        Tipo    = $access.AccessControlType
                    }
                    # Comparar usando el SID universal de "Everyone"
                    if ($userSid -eq $everyoneSid) {
                        $everyonePermissions += $access.FileSystemRights
                        if ($access.FileSystemRights -match "FullControl") {
                            $everyoneHasFullControl = $true
                        }
                    }
                }
                # Mostrar los permisos en la consola
                $permissions | ForEach-Object { 
                    Write-Host "`t$($_.Usuario) - $($_.Tipo) - " -NoNewline
                    Write-Host "` $($_.Permiso)" -ForegroundColor Green
                }
                # Mostrar los permisos de "Everyone" de forma consolidada
                if ($everyonePermissions.Count -gt 0) {
                    Write-Host "`tEveryone tiene los siguientes permisos:"  -NoNewline -ForegroundColor Yellow
                    Write-Host "` $($everyonePermissions -join ', ')" -ForegroundColor Green
                } else {
                    Write-Host "`tNo hay permisos para 'Everyone'" -ForegroundColor Red
                }
                # Si "Everyone" no tiene Full Control, preguntar si se desea concederlo
                if (-not $everyoneHasFullControl) {
                    $message = "El usuario 'Everyone' no tiene permisos de 'Full Control'. ¿Deseas concederlo?"
                    $title = "Permisos 'Everyone'"
                    $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
                    $icon = [System.Windows.Forms.MessageBoxIcon]::Question
                    $result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Agregar Full Control a "Everyone" en la carpeta
                        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($everyoneSid, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                        $acl.AddAccessRule($accessRule)
                        # Forzar la herencia para subcarpetas y archivos
                        $acl.SetAccessRuleProtection($false, $true)
                        Set-Acl -Path $folderPath -AclObject $acl
                        Write-Host "Se ha concedido 'Full Control' a 'Everyone'." -ForegroundColor Green
                    }
                }
    }
# Obtener las direcciones IP y los adaptadores
    $ipsWithAdapters = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
        Where-Object { $_.OperationalStatus -eq 'Up' } |
        ForEach-Object {
            $interface = $_
            $interface.GetIPProperties().UnicastAddresses |
            Where-Object { 
                $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '127.0.0.1' 
            } |
            ForEach-Object {
                @{
                    AdapterName = $interface.Name
                    IPAddress = $_.Address.ToString()
                }
            }
        }
# Construir el texto para mostrar todas las IPs y adaptadores
    if ($ipsWithAdapters.Count -gt 0) {
        $ipsTextForClipboard = ($ipsWithAdapters | ForEach-Object {
            $_.IPAddress
        }) -join ", "
        $ipsTextForLabel = $ipsWithAdapters | ForEach-Object {
            "- $($_.AdapterName) - IP: $($_.IPAddress)"
        } | Out-String
        $txt_IpAdress.Text = $ipsTextForLabel
    } else {
        $txt_IpAdress.Text = "No se encontraron direcciones IP"
    }
# Agregar los controles al formulario
            $formPrincipal.Controls.Add($tabControl)
            $formPrincipal.Controls.Add($btnExit)
# Obtener el puerto de SQL Server desde el registro
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\NATIONALSOFT\MSSQLServer\SuperSocketNetLib\Tcp"
        $tcpPort = Get-ItemProperty -Path $regKeyPath -Name "TcpPort" -ErrorAction SilentlyContinue
        if ($tcpPort -and $tcpPort.TcpPort) {
            $lblPort.Text = "Puerto SQL \NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $lblPort.Text = "No se encontró puerto o instancia."
        }
# Función para recargar y mostrar estados en el TextBox
function Refresh-AdapterStatus {
    $statuses = Get-NetworkAdapterStatus
    if ($statuses.Count -gt 0) {
        $lines = $statuses | ForEach-Object {
            "- $($_.AdapterName) - $($_.NetworkCategory)"
        }
        $txt_AdapterStatus.Text = $lines -join "`r`n"
    } else {
        $txt_AdapterStatus.Text = "No se encontraron adaptadores activos."
    }
}
function Get-NetworkAdapterStatus {
            $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
            $profiles = Get-NetConnectionProfile
            $adapterStatus = @()
            foreach ($adapter in $adapters) {
                $profile = $profiles | Where-Object { $_.InterfaceIndex -eq $adapter.ifIndex }
                $networkCategory = if ($profile) { $profile.NetworkCategory } else { "Desconocido" }
                $adapterStatus += [PSCustomObject]@{
                    AdapterName     = $adapter.Name
                    NetworkCategory = $networkCategory
                    InterfaceIndex  = $adapter.ifIndex  # Guardar el InterfaceIndex para identificar el adaptador
                }
            }
            return $adapterStatus
    }
#------------------------ download&run 1.0
function DownloadAndRun($url, $zipPath, $extractPath, $exeName, $validationPath) {
    # Validar si el archivo o aplicación ya existe
    if (!(Test-Path -Path $validationPath)) {
        $response = [System.Windows.Forms.MessageBox]::Show(
            "El archivo o aplicación no se encontró en '$validationPath'. ¿Desea descargarlo?",
            "Archivo no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        # Si el usuario selecciona "No", salir de la función
        if ($response -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
            return
        }
    }
    # Verificar si el archivo ZIP ya existe
    if (Test-Path -Path $zipPath) {
        $response = [System.Windows.Forms.MessageBox]::Show(
            "Archivo encontrado. ¿Lo desea eliminar y volver a descargar?",
            "Archivo ya descargado",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
        )
        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-Item -Path $zipPath -Force
            Remove-Item -Path $extractPath -Recurse -Force
            Write-Host "`tEliminando archivos anteriores..."
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::No) {
            # Si selecciona "No", abrir el programa sin eliminar archivos
            $exePath = Join-Path -Path $extractPath -ChildPath $exeName
            if (Test-Path -Path $exePath) {
                Write-Host "`tEjecutando el archivo ya descargado..."
                Start-Process -FilePath $exePath #-Wait   # Se quitó para ver si se usaban múltiples apps.
                Write-Host "`t$exeName se está ejecutando."
                return
            } else {
                Write-Host "`tNo se pudo encontrar el archivo ejecutable."  -ForegroundColor Red
                return
            }
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # Si selecciona "Cancelar", no hacer nada y decir que el usuario canceló
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
            return  # Aquí se termina la ejecución si el usuario cancela
        }
    }
    # Proceder con la descarga si no fue cancelada
    Write-Host "`tDescargando desde: $url"
    # Obtener el tamaño total del archivo antes de la descarga
    $response = Invoke-WebRequest -Uri $url -Method Head
    $totalSize = $response.Headers["Content-Length"]
    $totalSizeKB = [math]::round($totalSize / 1KB, 2)
    Write-Host "`tTamaño total: $totalSizeKB KB" -ForegroundColor Yellow
    # Descargar el archivo con barra de progreso
    $downloaded = 0
    $request = Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
    foreach ($chunk in $request.Content) {
        $downloaded += $chunk.Length
        $downloadedKB = [math]::round($downloaded / 1KB, 2)
        $progress = [math]::round(($downloaded / $totalSize) * 100, 2)
        Write-Progress -PercentComplete $progress -Status "Descargando..." -Activity "Progreso de la descarga" -CurrentOperation "$downloadedKB KB de $totalSizeKB KB descargados"
    }
    Write-Host "`tDescarga completada."  -ForegroundColor Green
    # Crear directorio de extracción si no existe
    if (!(Test-Path -Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }
    Write-Host "`tExtrayendo archivos..."
    try {
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Host "`tArchivos extraídos correctamente."  -ForegroundColor Green
    } catch {
        Write-Host "`tError al descomprimir el archivo: $_"   -ForegroundColor Red
    }
    $exePath = Join-Path -Path $extractPath -ChildPath $exeName
    if (Test-Path -Path $exePath) {
        Write-Host "`tEjecutando $exeName..."
        Start-Process -FilePath $exePath #-Wait
        Write-Host "`n$exeName se está ejecutando."
    } else {
        Write-Host "`nNo se pudo encontrar el archivo ejecutable."  -ForegroundColor Red
    }
}
# Funcion para agregar nuevas ips
function Show-NewIpForm {
        $formIpAssign = Create-Form -Title "Agregar IP Adicional" -Size (New-Object System.Drawing.Size(350, 150)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
    
        $lblipAssignER = Create-Label -Text "Ingrese la nueva dirección IP:" -Location (New-Object System.Drawing.Point(10, 20))
        $lblipAssignER.AutoSize = $true
        $formIpAssign.Controls.Add($lblipAssignER)
    
        $ipAssignTextBox1 = Create-TextBox -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(50, 20))
        $ipAssignTextBox1.MaxLength = 3
        $ipAssignTextBox1.Add_KeyPress({
            if (-not [char]::IsDigit($_.KeyChar) -and $_.KeyChar -ne 8 -and $_.KeyChar -ne '.') { $_.Handled = $true }
            if ($_.KeyChar -eq '.') {
                $ipAssignTextBox2.Focus()
                $_.Handled = $true
            }
        })
        $ipAssignTextBox1.Add_TextChanged({
            if ($ipAssignTextBox1.Text.Length -eq 3) { $ipAssignTextBox2.Focus() }
        })
        $formIpAssign.Controls.Add($ipAssignTextBox1)
    
        $lblipAssignERDot1 = Create-Label -Text "." -Location (New-Object System.Drawing.Point(65, 53))
        $lblipAssignERDot1.AutoSize = $true
        $formIpAssign.Controls.Add($lblipAssignERDot1)    
        $ipAssignTextBox2 = Create-TextBox -Location (New-Object System.Drawing.Point(80, 50)) -Size (New-Object System.Drawing.Size(50, 20))
        $ipAssignTextBox2.MaxLength = 3
        $ipAssignTextBox2.Add_KeyPress({
            if (-not [char]::IsDigit($_.KeyChar) -and $_.KeyChar -ne 8 -and $_.KeyChar -ne '.') { $_.Handled = $true }
            if ($_.KeyChar -eq '.') {
                $ipAssignTextBox3.Focus()
                $_.Handled = $true
            }
        })
        $ipAssignTextBox2.Add_TextChanged({
            if ($ipAssignTextBox2.Text.Length -eq 3) { $ipAssignTextBox3.Focus() }
        })
        $formIpAssign.Controls.Add($ipAssignTextBox2)
    
        $lblipAssignERDot2 = Create-Label -Text "." -Location (New-Object System.Drawing.Point(135, 53))
        $lblipAssignERDot2.AutoSize = $true
        $formIpAssign.Controls.Add($lblipAssignERDot2)
    
        $ipAssignTextBox3 = Create-TextBox -Location (New-Object System.Drawing.Point(150, 50)) -Size (New-Object System.Drawing.Size(50, 20))
        $ipAssignTextBox3.MaxLength = 3
        $ipAssignTextBox3.Add_KeyPress({
            if (-not [char]::IsDigit($_.KeyChar) -and $_.KeyChar -ne 8 -and $_.KeyChar -ne '.') { $_.Handled = $true }
            if ($_.KeyChar -eq '.') {
                $ipAssignTextBox4.Focus()
                $_.Handled = $true
            }
        })
        $ipAssignTextBox3.Add_TextChanged({
            if ($ipAssignTextBox3.Text.Length -eq 3) { $ipAssignTextBox4.Focus() }
        })
        $formIpAssign.Controls.Add($ipAssignTextBox3)
    
        $lblipAssignERDot3 = Create-Label -Text "." -Location (New-Object System.Drawing.Point(205, 53))
        $lblipAssignERDot3.AutoSize = $true
        $formIpAssign.Controls.Add($lblipAssignERDot3)
    
        $ipAssignTextBox4 = Create-TextBox -Location (New-Object System.Drawing.Point(220, 50)) -Size (New-Object System.Drawing.Size(50, 20))
        $ipAssignTextBox4.MaxLength = 3
        $ipAssignTextBox4.Add_KeyPress({
            if (-not [char]::IsDigit($_.KeyChar) -and $_.KeyChar -ne 8) { $_.Handled = $true }
        })
        $formIpAssign.Controls.Add($ipAssignTextBox4)
        $bntipAssign = Create-Button -Text "Aceptar" -Location (New-Object System.Drawing.Point(100, 80))  -Size (New-Object System.Drawing.Size(140, 30))
        $bntipAssign.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $formIpAssign.AcceptButton = $bntipAssign
        $formIpAssign.Controls.Add($bntipAssign)
        $result = $formIpAssign.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $octet1 = [int]$ipAssignTextBox1.Text
            $octet2 = [int]$ipAssignTextBox2.Text
            $octet3 = [int]$ipAssignTextBox3.Text
            $octet4 = [int]$ipAssignTextBox4.Text
    
            if ($octet1 -ge 0 -and $octet1 -le 255 -and
                $octet2 -ge 0 -and $octet2 -le 255 -and
                $octet3 -ge 0 -and $octet3 -le 255 -and
                $octet4 -ge 0 -and $octet4 -le 255) {
                $newIp = "$octet1.$octet2.$octet3.$octet4"
    
                if ($newIp -eq "0.0.0.0") {
                    Write-Host "La dirección IP no puede ser 0.0.0.0." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("La dirección IP no puede ser 0.0.0.0.", "Error")
                    return $null
                } else {
                    Write-Host "Nueva IP ingresada: $newIp" -ForegroundColor Green
                    return $newIp
                }
            } else {
                Write-Host "Uno o más octetos están fuera del rango válido (0-255)." -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("Uno o más octetos están fuera del rango válido (0-255).", "Error")
                return $null
            }
        } else {
            return $null
        }
}
# Función para verificar e instalar Chocolatey
function Check-Chocolatey {
            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                $response = [System.Windows.Forms.MessageBox]::Show(
                    "Chocolatey no está instalado. ¿Desea instalarlo ahora?",
                    "Chocolatey no encontrado",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
        
                if ($response -eq [System.Windows.Forms.DialogResult]::No) {
                    Write-Host "`nEl usuario canceló la instalación de Chocolatey." -ForegroundColor Red
                    return $false  # Retorna falso si el usuario cancela
                }
        
                Write-Host "`nInstalando Chocolatey..." -ForegroundColor Cyan
                try {
                    Set-ExecutionPolicy Bypass -Scope Process -Force
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
                    Write-Host "`nChocolatey se instaló correctamente." -ForegroundColor Green
        
                    # Configurar cacheLocation
                    Write-Host "`nConfigurando Chocolatey..." -ForegroundColor Yellow
                    choco config set cacheLocation C:\Choco\cache
        
                    [System.Windows.Forms.MessageBox]::Show(
                        "Chocolatey se instaló correctamente y ha sido configurado. Por favor, reinicie PowerShell antes de continuar.",
                        "Reinicio requerido",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
        
                    # Cerrar el programa automáticamente
                    Write-Host "`nCerrando la aplicación para permitir reinicio de PowerShell..." -ForegroundColor Red
                    Stop-Process -Id $PID -Force
                    return $false # Retorna falso para indicar que se debe reiniciar
                } catch {
                    Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error al instalar Chocolatey. Por favor, inténtelo manualmente.",
                        "Error de instalación",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                    return $false # Retorna falso en caso de error
                }
            } else {
                Write-Host "`tChocolatey ya está instalado." -ForegroundColor Green
                return $true # Retorna verdadero si Chocolatey ya está instalado
            }
}
# Función para detectar el nombre del grupo de administradores
function Get-AdminGroupName {
    $groups = net localgroup | Where-Object { $_ -match "Administrador|Administrators" }
    
    # Buscar coincidencia exacta
    if ($groups -match "\bAdministradores\b") {
        return "Administradores"
    } elseif ($groups -match "\bAdministrators\b") {
        return "Administrators"
    }
    # Si no se encuentra, intentar con otro método
    try {
        $adminGroup = Get-LocalGroup | Where-Object { $_.SID -like "S-1-5-32-544" }
        return $adminGroup.Name
    } catch {
        return "Administrators" # Valor por defecto
    }
}
#componentes
function Clear-TemporaryFiles {
    param([string]$folderPath)
    
    try {
        $items = Get-ChildItem -Path $folderPath -Recurse -Force -ErrorAction Stop
        $count = $items.Count
        Remove-Item -Path "$folderPath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "Eliminados $count archivos en $folderPath" -ForegroundColor Green
        return $count
    }
    catch {
        Write-Host "`n`tError limpiando $folderPath : $($_.Exception.Message)" -ForegroundColor Red
        return 0
    }
}
function Invoke-DiskCleanup {
    try {
        Write-Host "`nEjecutando Liberador de espacio en disco..." -ForegroundColor Cyan
        # Configurar parámetros de limpieza
        $cleanmgr = "$env:SystemDrive\Windows\System32\cleanmgr.exe"
        $sagerun = "9999"
        
        # Crear registro para limpieza completa
        Start-Process $cleanmgr -ArgumentList "/sageset:$sagerun" -Wait
        Start-Process $cleanmgr -ArgumentList "/sagerun:$sagerun" -Wait
        
        Write-Host "Limpieza de disco completada correctamente" -ForegroundColor Green
    }
    catch {
        Write-Host "Error en limpieza de disco: $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Show-SystemComponents {
    $criticalError = $false
    
    Write-Host "`n=== Componentes del sistema detectados ===" -ForegroundColor Cyan
    
    # Versión de Windows (componente crítico)
    try {
        $os = Get-CimInstance -ClassName CIM_OperatingSystem -ErrorAction Stop
        Write-Host "`n[Windows]" -ForegroundColor Yellow
        Write-Host "Versión: $($os.Caption) (Build $($os.Version))" -ForegroundColor White
    }
    catch {
        $criticalError = $true
        Write-Host "`n[Windows]" -ForegroundColor Yellow
        Write-Host "ERROR CRÍTICO: $($_.Exception.Message)" -ForegroundColor Red
        throw "No se pudo obtener información crítica del sistema"
    }

    # Resto de componentes (no críticos)
    if (-not $criticalError) {
        # Procesador
        try {
            $procesador = Get-CimInstance -ClassName CIM_Processor -ErrorAction Stop
            Write-Host "`n[Procesador]" -ForegroundColor Yellow
            Write-Host "Modelo: $($procesador.Name)" -ForegroundColor White
            Write-Host "Núcleos: $($procesador.NumberOfCores)" -ForegroundColor White
        }
        catch {
            Write-Host "`n[Procesador]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Memoria RAM
        try {
            $memoria = Get-CimInstance -ClassName CIM_PhysicalMemory -ErrorAction Stop
            Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
            $memoria | ForEach-Object {
                Write-Host "Módulo: $([math]::Round($_.Capacity/1GB, 2)) GB $($_.Manufacturer) ($($_.Speed) MHz)" -ForegroundColor White
            }
        }
        catch {
            Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Discos duros
        try {
            $discos = Get-CimInstance -ClassName CIM_DiskDrive -ErrorAction Stop
            Write-Host "`n[Discos duros]" -ForegroundColor Yellow
            $discos | ForEach-Object {
                Write-Host "Disco: $($_.Model) ($([math]::Round($_.Size/1GB, 2)) GB)" -ForegroundColor White
            }
        }
        catch {
            Write-Host "`n[Discos duros]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
function Start-SystemUpdate {
    $progressForm = $null
    try {
        $progressForm = Show-ProgressBar
        $totalSteps = 6
        $currentStep = 0

        Write-Host "`nIniciando proceso de actualización..." -ForegroundColor Cyan
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        # Paso 1: Detener servicio WMI
        Write-Host "`n[Paso 1/$totalSteps] Deteniendo servicio winmgmt..." -ForegroundColor Yellow
        $service = Get-Service -Name "winmgmt" -ErrorAction Stop
        if ($service.Status -eq "Running") {
            Stop-Service -Name "winmgmt" -Force -ErrorAction Stop
            Write-Host "Servicio detenido correctamente." -ForegroundColor Green
        }
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        # Paso 2: Limpieza de repositorio
        Write-Host "`n[Paso 2/$totalSteps] Renombrando carpeta Repository..." -ForegroundColor Yellow
        try {
            $repoPath = Join-Path $env:windir "System32\Wbem\Repository"
            if (Test-Path $repoPath) {
                $newName = "Repository_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Rename-Item -Path $repoPath -NewName $newName -Force -ErrorAction Stop
                Write-Host "Carpeta renombrada: $newName" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Advertencia: No se pudo renombrar la carpeta Repository. Continuando..." -ForegroundColor Yellow
        }
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        
        # Paso 3: Reiniciar servicio
        Write-Host "`n[Paso 3/$totalSteps] Reiniciando servicio winmgmt..." -ForegroundColor Yellow
        net start winmgmt *>&1 | Write-Host
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        # Paso 4: Limpieza de temporales
        Write-Host "`n[Paso 4/$totalSteps] Limpiando archivos temporales (ignorar si hay errores)..." -ForegroundColor Cyan
        $totalDeleted = 0
        $totalDeleted += Clear-TemporaryFiles -folderPath $env:TEMP
        $totalDeleted += Clear-TemporaryFiles -folderPath "$env:SystemDrive\Windows\Temp"
        Write-Host "Total archivos eliminados: $totalDeleted" -ForegroundColor Green
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        # Paso 5: Ejecutar cleanmgr
        Write-Host "`n[Paso 5/$totalSteps] Ejecutando Liberador de espacio..." -ForegroundColor Cyan
        Invoke-DiskCleanup
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        # Paso 6: Mostrar componentes
        Write-Host "`n[Paso 6/$totalSteps] Obteniendo información del sistema..." -ForegroundColor Cyan
        Show-SystemComponents
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps

        Write-Host "`nProceso completado con éxito" -ForegroundColor Green
    }
    catch {
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Detalles: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        [System.Windows.Forms.MessageBox]::Show(
            "Error: $($_.Exception.Message)`nRevise los logs antes de reiniciar.",
            "Error crítico",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        if ($progressForm -ne $null -and -not $progressForm.IsDisposed) {
            Close-ProgressBar $progressForm
        }
    }
}
# Función para mostrar la barra de progreso
function Show-ProgressBar {
                $formProgress = New-Object System.Windows.Forms.Form
                $formProgress.Text = "Progreso"
                $formProgress.Size = New-Object System.Drawing.Size(400, 150)
                $formProgress.StartPosition = "CenterScreen"
                $formProgress.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $formProgress.TopMost = $true
                $formProgress.ControlBox = $false
            
                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Size = New-Object System.Drawing.Size(360, 20)
                $progressBar.Location = New-Object System.Drawing.Point(10, 50)
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $progressBar.Maximum = 100
                $progressBar.SetStyle([System.Windows.Forms.ProgressBarStyle]::Continuous)
                $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $type = $progressBar.GetType()
                $flags = [Reflection.BindingFlags]::NonPublic -bor [Reflection.BindingFlags]::Instance
                $type.GetField("DoubleBuffered", $flags).SetValue($progressBar, $true)

                $lblPercentage = New-Object System.Windows.Forms.Label
                $lblPercentage.Location = New-Object System.Drawing.Point(10, 20)
                $lblPercentage.Size = New-Object System.Drawing.Size(360, 20)
                $lblPercentage.Text = "0% Completado"
                $lblPercentage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            
                # Agregar controles al formulario usando la propiedad Controls nativa
                $formProgress.Controls.Add($progressBar)
                $formProgress.Controls.Add($lblPercentage)
                
                # Exponer los controles como propiedades personalizadas (opcional, si es necesario)
                $formProgress | Add-Member -MemberType NoteProperty -Name ProgressBar -Value $progressBar -Force
                $formProgress | Add-Member -MemberType NoteProperty -Name Label -Value $lblPercentage -Force
            
                $formProgress.Show()
                return $formProgress
}
# Función para actualizar la barra de progreso
function Update-ProgressBar {

    param($ProgressForm, $CurrentStep, $TotalSteps)
    $percent = [math]::Round(($CurrentStep / $TotalSteps) * 100)
    if (-not $ProgressForm.IsDisposed) {
        $ProgressForm.ProgressBar.Value = $percent
        $ProgressForm.Label.Text = "$percent% Completado"
        [System.Windows.Forms.Application]::DoEvents() # Usar con precaución
    }
}

# Función para cerrar la barra de progreso
function Close-ProgressBar {
    param($ProgressForm)
    $ProgressForm.Close()
}
##-------------------------------------------------------------------------------BOTONES ACCIONES
$lblPort.Add_Click({
            if ($lblPort.Text -match "\d+") {  # Asegurarse de que el texto es un número
                $port = $matches[0]  # Extraer el número del texto
                [System.Windows.Forms.Clipboard]::SetText($port)
                Write-Host "Puerto copiado al portapapeles: $port" -ForegroundColor Green
            } else {
                Write-Host "El texto del Label del puerto no contiene un número válido para copiar." -ForegroundColor Red
            }
        })
        $lblPort.Add_MouseEnter($changeColorOnHover)
        $lblPort.Add_MouseLeave($restoreColorOnLeave)
$btnCheckPermissions.Add_Click({
        Write-Host "`nRevisando permisos en C:\NationalSoft" -ForegroundColor Yellow
        Check-Permissions
    })            
$btnForzarActualizacion.Add_Click({
                    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                    Show-SystemComponents
                    
                    # Cargar ensamblado necesario para MessageBox
                    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
                    
                    # Mostrar MessageBox en español
                    $resultado = [System.Windows.Forms.MessageBox]::Show(
                        "¿Desea forzar la actualización de datos?",  # Texto de la pregunta
                        "Confirmación",                              # Título de la ventana
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,  # Botones
                        [System.Windows.Forms.MessageBoxIcon]::Question   # Icono
                    )
                
                    if ($resultado -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Start-SystemUpdate
                        [System.Windows.Forms.MessageBox]::Show(
                            "Actualización completada",
                            "Éxito",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        ) | Out-Null
                    }
                    else {
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
                    }
                })
# SQL MANAGEMENT
$btnSQLManagement.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        
        function Get-SSMSVersions {
            $ssmsPaths = @()
            $possiblePaths = @(
                "${env:ProgramFiles(x86)}\Microsoft SQL Server\*\Tools\Binn\ManagementStudio\Ssms.exe",  
                "${env:ProgramFiles(x86)}\Microsoft SQL Server Management Studio *\Common7\IDE\Ssms.exe"
            )
            foreach ($path in $possiblePaths) {
                $foundPaths = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                if ($foundPaths) {
                    foreach ($foundPath in $foundPaths) {
                        $ssmsPaths += $foundPath.FullName
                    }
                }
            }
            return $ssmsPaths
        }
    
        function Get-SSMSVersionFromPath($path) {
            if ($path -match 'Microsoft SQL Server\\(\d+)') {
                return "SSMS $($matches[1])"
            }
            elseif ($path -match 'SQL Server Management Studio (\d+)') {
                return "SSMS $($matches[1])"
            }
            else {
                return "Versión desconocida"
            }
        }
    
        $ssmsVersions = Get-SSMSVersions
        if ($ssmsVersions.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontró ninguna versión de SQL Server Management Studio instalada.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    
        $formSelectionSSMS = Create-Form -Title "Seleccionar versión de SSMS" -Size (New-Object System.Drawing.Size(350, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
        $labelSSMS = Create-Label -Text "Seleccione la versión de Management:" -Location (New-Object System.Drawing.Point(10, 20)) -Size (New-Object System.Drawing.Size(310, 30)) -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
        $formSelectionSSMS.Controls.Add($labelSSMS)
    
        $labelSelectedVersion = Create-Label -Text "Versión seleccionada: " -Location (New-Object System.Drawing.Point(10, 80))
        $formSelectionSSMS.Controls.Add($labelSelectedVersion)
    
        $comboBoxSSMS = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(310, 20)) -DropDownStyle DropDownList
    
        foreach ($version in $ssmsVersions) {
            $comboBoxSSMS.Items.Add($version)
        }
    
        $comboBoxSSMS.SelectedIndex = 0
        $formSelectionSSMS.Controls.Add($comboBoxSSMS)
    
        # Actualizar la label con la versión real de SSMS extraída de la ruta
        $selectedVersion = $comboBoxSSMS.SelectedItem
        $labelSelectedVersion.Text = "Versión seleccionada: $(Get-SSMSVersionFromPath $selectedVersion)"
    
        $comboBoxSSMS.Add_SelectedIndexChanged({
            $selectedVersion = $comboBoxSSMS.SelectedItem
            $labelSelectedVersion.Text = "Versión seleccionada: $(Get-SSMSVersionFromPath $selectedVersion)"
        })
    
        $buttonOKSSMS = Create-Button -Text "Aceptar" -Location (New-Object System.Drawing.Point(10, 120)) -Size (New-Object System.Drawing.Size(140, 30))
        $buttonOKSSMS.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $buttonCancelSSMS = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(180, 120)) -Size (New-Object System.Drawing.Size(140, 30))
        $buttonCancelSSMS.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $formSelectionSSMS.AcceptButton = $buttonOKSSMS
        $formSelectionSSMS.Controls.Add($buttonOKSSMS)
        $formSelectionSSMS.CancelButton = $buttonCancelSSMS
        $formSelectionSSMS.Controls.Add($buttonCancelSSMS)
    
        $result = $formSelectionSSMS.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedVersion = $comboBoxSSMS.SelectedItem
            try {
                Write-Host "`tEjecutando SQL Server Management Studio desde: $selectedVersion" -ForegroundColor Green
                Start-Process -FilePath $selectedVersion
            } catch {
                Write-Host "`tError al intentar ejecutar SSMS desde la ruta seleccionada." -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("No se pudo iniciar SSMS. Verifique la ruta seleccionada.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
        }
})
#                Profiler:
$btnProfiler.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        $ProfilerUrl = "https://codeplexarchive.org/codeplex/browse/ExpressProfiler/releases/4/ExpressProfiler22wAddinSigned.zip"
        $ProfilerZipPath = "C:\Temp\ExpressProfiler22wAddinSigned.zip"
        $ExtractPath = "C:\Temp\ExpressProfiler2"
        $ExeName = "ExpressProfiler.exe"
        $ValidationPath = "C:\Temp\ExpressProfiler2\ExpressProfiler.exe"

        DownloadAndRun -url $ProfilerUrl -zipPath $ProfilerZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
        if ($disableControls) {        Enable-Controls -parentControl $formPrincipal    }
        }
    )
#                Impresoras printer tool
$btnPrinterTool.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        $PrinterToolUrl = "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip"
        $PrinterToolZipPath = "C:\Temp\RPT-RPI-Printer-Tool-1.zip"
        $ExtractPath = "C:\Temp\RPT-RPI-Printer-Tool-1"
        $ExeName = "POS Printer Test.exe"
        $ValidationPath = "C:\Temp\RPT-RPI-Printer-Tool-1\POS Printer Test.exe"

        DownloadAndRun -url $PrinterToolUrl -zipPath $PrinterToolZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
    })
#                Database 4 net
$btnDatabase.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        $DatabaseUrl = "https://fishcodelib.com/files/DatabaseNet4.zip"
        $DatabaseZipPath = "C:\Temp\DatabaseNet4.zip"
        $ExtractPath = "C:\Temp\Database4"
        $ExeName = "Database4.exe"
        $ValidationPath = "C:\Temp\Database4\Database4.exe"

        DownloadAndRun -url $DatabaseUrl -zipPath $DatabaseZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
    })
#seleccion de managers BOTON
$btnSQLManager.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    
        function Get-SQLServerManagers {
            $managers = @()
            $possiblePaths = @(
                "${env:SystemRoot}\System32\SQLServerManager*.msc",
                "${env:SystemRoot}\SysWOW64\SQLServerManager*.msc"
            )
    
            foreach ($path in $possiblePaths) {
                $foundManagers = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                if ($foundManagers) {
                    $managers += $foundManagers.FullName
                }
            }
            return $managers
        }
    
        $managers = Get-SQLServerManagers
        if ($managers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No se encontró ninguna versión de SQL Server Configuration Manager.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $formSelectionManager = Create-Form -Title "Seleccionar versión de Configuration Manager" -Size (New-Object System.Drawing.Size(350, 250)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
            -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
    
        $labelManager = Create-Label -Text "Seleccione la versión de Configuration Manager:" -Location (New-Object System.Drawing.Point(10, 20)) -Size (New-Object System.Drawing.Size(310, 30)) -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
        $formSelectionManager.Controls.Add($labelManager)
    
        $labelManagerInfo = Create-Label -Text "" -Location (New-Object System.Drawing.Point(10, 80)) -Size (New-Object System.Drawing.Size(310, 30)) -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
        $formSelectionManager.Controls.Add($labelManagerInfo)
    
        function Get-ManagerInfo($path) {
            if ($path -match "SQLServerManager(\d+)") {
                $version = $matches[1]
                if ($path -match "SysWOW64") {
                    return "SQLServerManager${version} 64bits"
                } else {
                    return "SQLServerManager${version} 32bits"
                }
            } else {
                return "Información no disponible"
            }
        }
    
        $UpdateManagerInfo = {
            $selectedManager = $comboBoxManager.SelectedItem
            $managerInfo = Get-ManagerInfo $selectedManager
            $labelManagerInfo.Text = $managerInfo
        }
    
        $comboBoxManager = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(310, 20)) -DropDownStyle DropDownList
    
        foreach ($manager in $managers) {
            $comboBoxManager.Items.Add($manager)
        }
    
        $comboBoxManager.SelectedIndex = 0
        $formSelectionManager.Controls.Add($comboBoxManager)
    
        $UpdateManagerInfo.Invoke() # Actualizar al inicio
    
        $comboBoxManager.Add_SelectedIndexChanged({
            $UpdateManagerInfo.Invoke()
        })
    
        $btnOKManager = Create-Button -Text "Aceptar" -Location (New-Object System.Drawing.Point(10, 120)) -Size (New-Object System.Drawing.Size(140, 30))
        $btnOKManager.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $btnCancelManager = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(180, 120)) -Size (New-Object System.Drawing.Size(140, 30))
        $btnCancelManager.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $formSelectionManager.AcceptButton = $btnOKManager
        $formSelectionManager.Controls.Add($btnOKManager)
        $formSelectionManager.CancelButton = $btnCancelManager
        $formSelectionManager.Controls.Add($btnCancelManager)
    
        $result = $formSelectionManager.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedManager = $comboBoxManager.SelectedItem
            try {
                Write-Host "`tEjecutando SQL Server Configuration Manager desde: $selectedManager" -ForegroundColor Green
                Start-Process -FilePath $selectedManager
            } catch {
                Write-Host "`tError al intentar ejecutar SQL Server Configuration Manager desde la ruta seleccionada." -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("No se pudo iniciar SQL Server Configuration Manager. Verifique la ruta seleccionada.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
        }
})
#Clear Anydesk
$btnClearAnyDesk.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        # Mostrar cuadro de confirmación
        $confirmationResult = [System.Windows.Forms.MessageBox]::Show(
            "¿Estás seguro de renovar AnyDesk?", 
            "Confirmar Renovación", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        # Si el usuario selecciona "Sí"
        if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $filesToDelete = @(
                "C:\ProgramData\AnyDesk\system.conf",
                "C:\ProgramData\AnyDesk\service.conf",
                "$env:APPDATA\AnyDesk\system.conf",
                "$env:APPDATA\AnyDesk\service.conf"
            )

            $deletedFilesCount = 0
            $errors = @()

            # Intentar cerrar el proceso AnyDesk
            try {
                Write-Host "`tCerrando el proceso AnyDesk..." -ForegroundColor Yellow
                Stop-Process -Name "AnyDesk" -Force -ErrorAction Stop
                Write-Host "`tAnyDesk ha sido cerrado correctamente." -ForegroundColor Green
            }
            catch {
                Write-Host "`tError al cerrar el proceso AnyDesk: $_" -ForegroundColor Red
                $errors += "No se pudo cerrar el proceso AnyDesk."
            }

            # Intentar eliminar los archivos
            foreach ($file in $filesToDelete) {
                try {
                    if (Test-Path $file) {
                        Remove-Item -Path $file -Force -ErrorAction Stop
                        Write-Host "`tArchivo eliminado: $file" -ForegroundColor Green
                        $deletedFilesCount++
                    }
                    else {
                        Write-Host "`tArchivo no encontrado: $file" -ForegroundColor Red
                    }
                }
                catch {
                    Write-Host "`nError al eliminar el archivo." -ForegroundColor Red
                }
            }

            # Mostrar el resultado
            if ($errors.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("$deletedFilesCount archivo(s) eliminado(s) correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Se encontraron errores. Revisa la consola para más detalles.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else {
            # Si el usuario selecciona "No", simplemente no hace nada
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
        }
    })
    #                Mostrar impresoras en la consola
$btnShowPrinters.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        try {
            $printers = Get-WmiObject -Query "SELECT * FROM Win32_Printer" | ForEach-Object {
                $printer = $_
                $isShared = $printer.Shared -eq $true
                [PSCustomObject]@{
                    Name = $printer.Name.Substring(0, [Math]::Min(24, $printer.Name.Length))
                    PortName = $printer.PortName.Substring(0, [Math]::Min(19, $printer.PortName.Length))
                    DriverName = $printer.DriverName.Substring(0, [Math]::Min(19, $printer.DriverName.Length))
                    IsShared = if ($isShared) { "Sí" } else { "No" }
                }
            }
            Write-Host "`nImpresoras disponibles en el sistema:"
            # Si hay impresoras, las mostramos en una tabla bien formateada
            if ($printers.Count -gt 0) {
                Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "Nombre", "Puerto", "Driver", "Compartida")
                Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "------", "------", "------", "---------")
            
                $printers | ForEach-Object { 
                    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f $_.Name, $_.PortName, $_.DriverName, $_.IsShared)
                }
            } else {
                Write-Host "`nNo se encontraron impresoras."
            }
        
        } catch {
            Write-Host "`nError al obtener impresoras: $_"
        }
    })
#                Spool cola de impresión
$btnClearPrintJobs.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        try {

            # Ejecutar el script para limpiar los trabajos de impresión y reiniciar la cola de impresión
            Get-Printer | ForEach-Object { 
                Get-PrintJob -PrinterName $_.Name | Remove-PrintJob 
            }
        
            # Reiniciar el servicio de la cola de impresión
                Stop-Service -Name Spooler -Force
                Start-Service -Name Spooler
        
            # Mensaje de confirmación
            [System.Windows.Forms.MessageBox]::Show("Los trabajos de impresión han sido eliminados y el servicio de cola de impresión se ha reiniciado.", "Operación Exitosa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            # Manejar cualquier error que ocurra
            [System.Windows.Forms.MessageBox]::Show("Ocurrió un error al intentar limpiar las impresoras o reiniciar el servicio.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })



#LMZA
$LZMAbtnBuscarCarpeta.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray

    # Ruta del registro
    $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"

    # Verificar existencia de la ruta
    if (-not (Test-Path $LZMAregistryPath)) {
        Write-Host "`nLa ruta del registro no existe: $LZMAregistryPath" -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show(
            "La ruta del registro no existe:`n$LZMAregistryPath",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    try {
        Write-Host "`tLeyendo subcarpetas de LZMA…" -ForegroundColor Gray

        # Obtener carpetas principales
        $LZMcarpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction Stop |
            Where-Object { $_.PSIsContainer }

        if ($LZMcarpetasPrincipales.Count -lt 1) {
            Write-Host "`tNo se encontraron carpetas principales." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show(
                "No se encontraron carpetas principales en la ruta del registro.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # Construir colección de instaladores
        $instaladores = @()
        foreach ($carpeta in $LZMcarpetasPrincipales) {
            $subdirs = Get-ChildItem -Path $carpeta.PSPath | Where-Object { $_.PSIsContainer }
            foreach ($sd in $subdirs) {
                $instaladores += [PSCustomObject]@{
                    Name = $sd.PSChildName
                    Path = $sd.PSPath
                }
            }
        }

        if ($instaladores.Count -lt 1) {
            Write-Host "`tNo se encontraron subcarpetas." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show(
                "No se encontraron subcarpetas en la ruta del registro.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # Ordenar descendente por nombre
        $instaladores = $instaladores | Sort-Object -Property Name -Descending

        # Preparar listas para el ComboBox
        $LZMsubCarpetas  = @("Selecciona instalador a renombrar") + ($instaladores | ForEach-Object { $_.Name })
        $LZMrutasCompletas = $instaladores | ForEach-Object { $_.Path }

        # Crear formulario
        $formLZMA = Create-Form `
            -Title "Carpetas LZMA" `
            -Size (New-Object System.Drawing.Size(400, 200)) `
            -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
            -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) `
            -MaximizeBox $false -MinimizeBox $false

        # ComboBox de subcarpetas
        $LZMcomboBoxCarpetas = Create-ComboBox `
            -Location (New-Object System.Drawing.Point(10, 10)) `
            -Size (New-Object System.Drawing.Size(360, 20)) `
            -DropDownStyle DropDownList `
            -Font $defaultFont

        foreach ($nombre in $LZMsubCarpetas) {
            $LZMcomboBoxCarpetas.Items.Add($nombre)
        }
        $LZMcomboBoxCarpetas.SelectedIndex = 0

        # Label para AI_ExePath
        $lblLZMAExePath = Create-Label `
            -Text "AI_ExePath: -" `
            -Location (New-Object System.Drawing.Point(10, 35)) `
            -Size (New-Object System.Drawing.Size(360, 70)) `
            -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 0, 0)) `
            -Font $defaultFont

        # Botón Renombrar
        $LZMbtnRenombrar = Create-Button `
            -Text "Renombrar" `
            -Location (New-Object System.Drawing.Point(10, 120)) `
            -Size (New-Object System.Drawing.Size(180, 30)) `
            -Enabled $false

        # Botón Salir
        $LMZAbtnSalir = Create-Button `
            -Text "Salir" `
            -Location (New-Object System.Drawing.Point(200, 120)) `
            -Size (New-Object System.Drawing.Size(180, 30))

        # Eventos
        $LZMcomboBoxCarpetas.Add_SelectedIndexChanged({
            $idx = $LZMcomboBoxCarpetas.SelectedIndex
            $LZMbtnRenombrar.Enabled = ($idx -gt 0)
            if ($idx -gt 0) {
                $ruta = $LZMrutasCompletas[$idx - 1]
                $prop = Get-ItemProperty -Path $ruta -Name "AI_ExePath" -ErrorAction SilentlyContinue
                $lblLZMAExePath.Text = if ($prop) { "AI_ExePath: $($prop.AI_ExePath)" } else { "AI_ExePath: No encontrado" }
            } else {
                $lblLZMAExePath.Text = "AI_ExePath: -"
            }
        })

        $LZMbtnRenombrar.Add_Click({
            $idx = $LZMcomboBoxCarpetas.SelectedIndex
            if ($idx -gt 0) {
                $rutaVieja = $LZMrutasCompletas[$idx - 1]
                $nombre = $LZMcomboBoxCarpetas.SelectedItem
                $nuevaNombre = "$nombre.backup"
                $msg = "¿Está seguro de renombrar:`n$rutaVieja`na:`n$nuevaNombre"
                $conf = [System.Windows.Forms.MessageBox]::Show($msg, "Confirmar renombrado", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($conf -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        Rename-Item -Path $rutaVieja -NewName $nuevaNombre
                        [System.Windows.Forms.MessageBox]::Show("Registro renombrado correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        $formLZMA.Close()
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Error al renombrar: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
            }
        })

        $LMZAbtnSalir.Add_Click({
            Write-Host "`tCancelado por el usuario." -ForegroundColor Yellow
            $formLZMA.Close()
        })

        # Agregar controles y mostrar
        $formLZMA.Controls.AddRange(@($LZMcomboBoxCarpetas, $lblLZMAExePath, $LZMbtnRenombrar, $LMZAbtnSalir))
        $formLZMA.ShowDialog()

    } catch {
        Write-Host "`tError accediendo al registro: $_" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show(
            "Error accediendo al registro:`n$_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})


                
# Crear el nuevo formulario para los instaladores de Chocolatey
        $formInstaladoresChoco = Create-Form -Title "Instaladores Choco" -Size (New-Object System.Drawing.Size(500, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(5, 5, 5))   
# Crear los botones dentro del nuevo formulario
    $btnInstallSQL2014 = Create-Button -Text "Install: SQL2014" -Location (New-Object System.Drawing.Point(10, 10)) `
        -ToolTip "Instalación mediante choco de SQL Server 2014 Express." -Enabled $false
    $btnInstallSQL2019 = Create-Button -Text "Install: SQL2019" -Location (New-Object System.Drawing.Point(240, 10)) `
        -ToolTip "Instalación mediante choco de SQL Server 2019 Express."
    $btnInstallSQLManagement = Create-Button -Text "Install: Management14" -Location (New-Object System.Drawing.Point(10, 50)) `
        -ToolTip "Instalación mediante choco de SQL Management 2014."
    $btnExitInstaladores = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(10, 120)) `
        -ToolTip "Salir del formulario de instaladores."
# Agregar los botones al nuevo formulario
    $formInstaladoresChoco.Controls.Add($btnInstallSQL2014)
    $formInstaladoresChoco.Controls.Add($btnInstallSQL2019)
    $formInstaladoresChoco.Controls.Add($btnInstallSQLManagement)
    $formInstaladoresChoco.Controls.Add($btnExitInstaladores)
# Evento para el botón de salir del formulario de instaladores
$btnExitInstaladores.Add_Click({
    $formInstaladoresChoco.Close()
})
# Evento para el botón "Instalar Herramientas"
$btnInstalarHerramientas.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        # Verificar si Chocolatey está instalado
        if (Check-Chocolatey) {
            # Mostrar el formulario de instaladores de Chocolatey
            $formInstaladoresChoco.ShowDialog()
        } else {
            Write-Host "Chocolatey no está instalado. No se puede abrir el menú de instaladores." -ForegroundColor Red
        }
})
#Boton para instalar Management
$btnInstallSQLManagement.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    $response = [System.Windows.Forms.MessageBox]::Show(
        "¿Desea proceder con la instalación de SQL Server Management Studio 2014 Express?",
        "Advertencia de instalación",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($response -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Host "`nEl usuario canceló la instalación." -ForegroundColor Red
        return
    }

    if (!(Check-Chocolatey)) { return } # Sale si Check-Chocolatey retorna falso (cancelado o error)

    Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green

    try {
        Write-Host "`nConfigurando Chocolatey..." -ForegroundColor Yellow
        choco config set cacheLocation C:\Choco\cache

        Write-Host "`nInstalando SQL Server Management Studio 2014 Express usando Chocolatey..." -ForegroundColor Cyan
        Start-Process choco -ArgumentList 'install mssqlservermanagementstudio2014express --confirm --yes' -NoNewWindow -Wait
        Write-Host "`nInstalación completa." -ForegroundColor Green
    } catch {
        Write-Host "`nOcurrió un error durante la instalación: $_" -ForegroundColor Red
    }
})
# Instalador de SQL 2019
$btnInstallSQL2019.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    $response = [System.Windows.Forms.MessageBox]::Show(
        "¿Desea proceder con la instalación de SQL Server 2019 Express?",
        "Advertencia de instalación",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($response -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Host "`nEl usuario canceló la instalación." -ForegroundColor Red
        return
    }

    if (!(Check-Chocolatey)) { return } # Sale si Check-Chocolatey retorna falso

    Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green

    try {
            Write-Host "`nInstalando SQL Server 2019 Express usando Chocolatey..." -ForegroundColor Cyan
            Start-Process choco -ArgumentList 'install sql-server-express -y --version=2019.20190106 --params "/SQLUSER:sa /SQLPASSWORD:National09 /INSTANCENAME:SQL2019 /FEATURES:SQL"' -NoNewWindow -Wait
            Write-Host "`nInstalación completa." -ForegroundColor Green
        Start-Sleep -Seconds 30 # Espera a que la instalación se complete (opcional)
        sqlcmd -S SQL2019 -U sa -P National09 -Q "exec sp_defaultlanguage [sa], 'spanish'"
        [System.Windows.Forms.MessageBox]::Show("SQL Server 2019 Express instalado correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al instalar SQL Server 2019 Express: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
# Instalador de SQL 2014
$btnInstallSQL2014.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $response = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea proceder con la instalación de SQL Server 2014 Express?",
                "Advertencia de instalación",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        
            if ($response -eq [System.Windows.Forms.DialogResult]::No) {
                Write-Host "`nEl usuario canceló la instalación." -ForegroundColor Red
                return
            }
        
            if (!(Check-Chocolatey)) { return } # Sale si Check-Chocolatey retorna falso
        
            Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        
            try {
                # Verificar si la instancia ya existe
                $instanceExists = Get-Service -Name "MSSQL`$NationalSoft" -ErrorAction SilentlyContinue
                if ($instanceExists) {
                    Write-Host "`nLa instancia 'NationalSoft' ya existe. Cancelando la instalación." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("La instancia 'NationalSoft' ya existe. Cancelando la instalación.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return
                }
        
                # Instalar SQL Server 2014 Express
                Write-Host "`nInstalando SQL Server 2014 Express usando Chocolatey..." -ForegroundColor Cyan
                Start-Process choco -ArgumentList 'install sql-server-express -y --version=2014.0.2000.8 --params "/SQLUSER:sa /SQLPASSWORD:National09 /INSTANCENAME:NationalSoft /FEATURES:SQL"' -NoNewWindow -Wait
                Write-Host "`nInstalación completa." -ForegroundColor Green
                Start-Sleep -Seconds 30 # Espera a que la instalación se complete (opcional)
                sqlcmd -S NationalSoft -U sa -P National09 -Q "exec sp_defaultlanguage [sa], 'spanish'"
                [System.Windows.Forms.MessageBox]::Show("SQL Server 2014 Express instalado correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error al instalar SQL Server 2014 Express: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
})
# ------------------------------ Boton para configurar nuevas ips
$btnConfigurarIPs.Add_Click({
            Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
            $formIpAssignAsignacion = Create-Form -Title "Asignación de IPs" -Size (New-Object System.Drawing.Size(400, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false
            #interfaz
            $lblipAssignAdapter = Create-Label -Text "Seleccione el adaptador de red:" -Location (New-Object System.Drawing.Point(10, 20))
            $lblipAssignAdapter.AutoSize = $true
            $formIpAssignAsignacion.Controls.Add($lblipAssignAdapter)
            $ComboBipAssignAdapters = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(360, 20)) -DropDownStyle DropDownList `
                                          -DefaultText "Selecciona 1 adaptador de red"
                        $ComboBipAssignAdapters.Add_SelectedIndexChanged({
                            # Verificar si se ha seleccionado un adaptador distinto de la opción por defecto
                            if ($ComboBipAssignAdapters.SelectedItem -ne "") {
                                # Habilitar los botones si se ha seleccionado un adaptador
                                $btnipAssignAssignIP.Enabled = $true
                                $btnipAssignChangeToDhcp.Enabled = $true
                            } else {
                                # Deshabilitar los botones si no se ha seleccionado un adaptador
                                $btnipAssignAssignIP.Enabled = $false
                                $btnipAssignChangeToDhcp.Enabled = $false
                            }
                        })
                $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
                foreach ($adapter in $adapters) {
                    $ComboBipAssignAdapters.Items.Add($adapter.Name)
                }
            $formIpAssignAsignacion.Controls.Add($ComboBipAssignAdapters)
            $lblipAssignIps = Create-Label -Text "IPs asignadas:" -Location (New-Object System.Drawing.Point(10, 80))
            $lblipAssignIps.AutoSize = $true
            $formIpAssignAsignacion.Controls.Add($lblipAssignIps)
            $btnipAssignAssignIP = Create-Button -Text "Asignar Nueva IP" -Location (New-Object System.Drawing.Point(10, 120)) -Size (New-Object System.Drawing.Size(140, 30)) -Enabled $false
            $btnipAssignAssignIP.Add_Click({
                $selectedAdapterName = $ComboBipAssignAdapters.SelectedItem
                if ($selectedAdapterName -eq "Selecciona 1 adaptador de red") {
                    Write-Host "`nPor favor, selecciona un adaptador de red." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("Por favor, selecciona un adaptador de red.", "Error")
                    return
                }
                $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
                $currentConfig = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
        
                if ($currentConfig) {
                    $isDhcp = ($currentConfig.PrefixOrigin -eq "Dhcp")
                    $currentIPAddress = $currentConfig.IPAddress
                    $currentPrefixLength = $currentConfig.PrefixLength
                    $currentGateway = (Get-NetIPConfiguration -InterfaceAlias $selectedAdapter.Name).IPv4DefaultGateway | Select-Object -ExpandProperty NextHop
        
                    if (-not $isDhcp) {
                        Write-Host "`nEl adaptador ya tiene una IP fija. ¿Desea agregar una nueva IP?" -ForegroundColor Yellow
                        $confirmation = [System.Windows.Forms.MessageBox]::Show("El adaptador ya tiene una IP fija. ¿Desea agregar una nueva IP?", "Confirmación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                            $newIp = Show-NewIpForm
                            if ($newIp) {
                                $existingIp = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.IPAddress -eq $newIp }
                                if ($existingIp) {
                                    Write-Host "`nLa dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name)." -ForegroundColor Red
                                    [System.Windows.Forms.MessageBox]::Show("La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name).", "Error")
                                } else {
                                    try {
                                        New-NetIPAddress -IPAddress $newIp -PrefixLength $currentPrefixLength -InterfaceAlias $selectedAdapter.Name
                                        Write-Host "`nSe agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                        [System.Windows.Forms.MessageBox]::Show("Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name).", "Éxito")
        
                                        # Actualizar la lista de IPs asignadas
                                        $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                                        $ips = $currentIPs.IPAddress -join ", "
                                        $lblipAssignIps.Text = "IPs asignadas: $ips"
                                    } catch {
                                        Write-Host "`nError al agregar la dirección IP adicional: $($_.Exception.Message)" -ForegroundColor Red
                                        [System.Windows.Forms.MessageBox]::Show("Error al agregar la dirección IP adicional: $($_.Exception.Message)", "Error")
                                    }
                                }
                            }
                        }
                    } else {
                        Write-Host "`n¿Desea cambiar a IP fija usando la IP actual ($currentIPAddress)?" -ForegroundColor Yellow
                        $confirmation = [System.Windows.Forms.MessageBox]::Show("¿Desea cambiar a IP fija usando la IP actual ($currentIPAddress)?", "Confirmación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                            try {
                                Set-NetIPInterface -InterfaceAlias $selectedAdapter.Name -Dhcp Disabled
                                New-NetIPAddress -IPAddress $currentIPAddress -PrefixLength $currentPrefixLength -InterfaceAlias $selectedAdapter.Name
        
                                if ($currentGateway) {
                                    Remove-NetRoute -InterfaceAlias $selectedAdapter.Name -NextHop $currentGateway -Confirm:$false -ErrorAction SilentlyContinue
                                    New-NetRoute -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 -NextHop $currentGateway -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue
                                }
        
                                $dnsServers = @("8.8.8.8", "8.8.4.4")
                                Set-DnsClientServerAddress -InterfaceAlias $selectedAdapter.Name -ServerAddresses $dnsServers
        
                                Write-Host "`nSe cambió a IP fija $currentIPAddress en el adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                [System.Windows.Forms.MessageBox]::Show("Se cambió a IP fija $currentIPAddress en el adaptador $($selectedAdapter.Name).", "Éxito")
        
                                # Actualizar la lista de IPs asignadas
                                $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                                $ips = $currentIPs.IPAddress -join ", "
                                $lblipAssignIps.Text = "IPs asignadas: $ips"
        
                                Write-Host "`n¿Desea agregar una dirección IP adicional?" -ForegroundColor Yellow
                                $confirmationAdditionalIP = [System.Windows.Forms.MessageBox]::Show("¿Desea agregar una dirección IP adicional?", "IP Adicional", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                                if ($confirmationAdditionalIP -eq [System.Windows.Forms.DialogResult]::Yes) {
                                    $newIp = Show-NewIpForm
                                    if ($newIp) {
                                        $existingIp = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.IPAddress -eq $newIp }
                                        if ($existingIp) {
                                            Write-Host "`nLa dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name)." -ForegroundColor Red
                                            [System.Windows.Forms.MessageBox]::Show("La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name).", "Error")
                                        } else {
                                            try {
                                                New-NetIPAddress -IPAddress $newIp -PrefixLength $currentPrefixLength -InterfaceAlias $selectedAdapter.Name
                                                Write-Host "`nSe agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                                [System.Windows.Forms.MessageBox]::Show("Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name).", "Éxito")
        
                                                # Actualizar la lista de IPs asignadas
                                                $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                                                $ips = $currentIPs.IPAddress -join ", "
                                                $lblipAssignIps.Text = "IPs asignadas: $ips"
                                            } catch {
                                                Write-Host "`nError al agregar la dirección IP adicional: $($_.Exception.Message)" -ForegroundColor Red
                                                [System.Windows.Forms.MessageBox]::Show("Error al agregar la dirección IP adicional: $($_.Exception.Message)", "Error")
                                            }
                                        }
                                    }
                                }
                            } catch {
                                Write-Host "`nError al cambiar a IP fija: $($_.Exception.Message)" -ForegroundColor Red
                                [System.Windows.Forms.MessageBox]::Show("Error al cambiar a IP fija: $($_.Exception.Message)", "Error")
                            }
                        }
                    }
                } else {
                    Write-Host "`nNo se pudo obtener la configuración actual del adaptador." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("No se pudo obtener la configuración actual del adaptador.", "Error")
                }
            })
            $formIpAssignAsignacion.Controls.Add($btnipAssignAssignIP)
            $btnipAssignChangeToDhcp = Create-Button -Text "Cambiar a DHCP" -Location (New-Object System.Drawing.Point(140, 120)) -Size (New-Object System.Drawing.Size(140, 30)) -Enabled $false
            $btnipAssignChangeToDhcp.Add_Click({
                $selectedAdapterName = $ComboBipAssignAdapters.SelectedItem
                if ($selectedAdapterName -eq "Selecciona 1 adaptador de red") {
                    Write-Host "`nPor favor, selecciona un adaptador de red." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("Por favor, selecciona un adaptador de red.", "Error")
                    return
                }
                $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
                $currentConfig = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
        
                if ($currentConfig) {
                    $isDhcp = ($currentConfig.PrefixOrigin -eq "Dhcp")
                    if ($isDhcp) {
                        Write-Host "`nEl adaptador ya está en DHCP." -ForegroundColor Yellow
                        [System.Windows.Forms.MessageBox]::Show("El adaptador ya está en DHCP.", "Información")
                    } else {
                        Write-Host "`n¿Está seguro de que desea cambiar a DHCP?" -ForegroundColor Yellow
                        $confirmation = [System.Windows.Forms.MessageBox]::Show("¿Está seguro de que desea cambiar a DHCP?", "Confirmación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                        if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                            try {
                                $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -eq "Manual" }
                                foreach ($ip in $currentIPs) {
                                    Remove-NetIPAddress -IPAddress $ip.IPAddress -PrefixLength $ip.PrefixLength -Confirm:$false -ErrorAction SilentlyContinue
                                }
                                Set-NetIPInterface -InterfaceAlias $selectedAdapter.Name -Dhcp Enabled
                                Set-DnsClientServerAddress -InterfaceAlias $selectedAdapter.Name -ResetServerAddresses
                                Write-Host "`nSe cambió a DHCP en el adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                [System.Windows.Forms.MessageBox]::Show("Se cambió a DHCP en el adaptador $($selectedAdapter.Name).", "Éxito")
        
                                # Actualizar la lista de IPs asignadas
                                $lblipAssignIps.Text = "Generando IP por DHCP. Seleccione de nuevo."
                            } catch {
                                Write-Host "`nError al cambiar a DHCP: $($_.Exception.Message)" -ForegroundColor Red
                                [System.Windows.Forms.MessageBox]::Show("Error al cambiar a DHCP: $($_.Exception.Message)", "Error")
                            }
                        }
                    }
                } else {
                    Write-Host "`nNo se pudo obtener la configuración actual del adaptador." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("No se pudo obtener la configuración actual del adaptador.", "Error")
                }
            })
            $formIpAssignAsignacion.Controls.Add($btnipAssignChangeToDhcp)
            # Agregar un botón "Cerrar" al formulario
            $btnCloseFormipAssign = Create-Button -Text "Cerrar" -Location (New-Object System.Drawing.Point(270, 120))  -Size (New-Object System.Drawing.Size(140, 30))
            $btnCloseFormipAssign.Add_Click({
                $formIpAssignAsignacion.Close()
            })
            $formIpAssignAsignacion.Controls.Add($btnCloseFormipAssign)
            $ComboBipAssignAdapters.Add_SelectedIndexChanged({
                $selectedAdapterName = $ComboBipAssignAdapters.SelectedItem
                if ($selectedAdapterName -ne "Selecciona 1 adaptador de red") {
                    $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
                    $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                    $ips = $currentIPs.IPAddress -join ", "
                    $lblipAssignIps.Text = "IPs asignadas: $ips"
                } else {
                    $lblipAssignIps.Text = "IPs asignadas:"
                }
            })
            $formIpAssignAsignacion.ShowDialog()
})
# ICACLS para dar permisos cuando marca error driver de lector
$btnLectorDPicacls.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    try {
        # Ruta de PsExec
        $psexecPath = "C:\Temp\PsExec\PsExec.exe"
        $psexecZip = "C:\Temp\PSTools.zip"
        $psexecUrl = "https://download.sysinternals.com/files/PSTools.zip"
        $psexecExtractPath = "C:\Temp\PsExec"
        # Validar si PsExec.exe existe
        if (-Not (Test-Path $psexecPath)) {
            Write-Host "`tPsExec no encontrado. Descargando desde Sysinternals..." -ForegroundColor Yellow
            # Crear carpeta Temp si no existe
            if (-Not (Test-Path "C:\Temp")) {
                New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
            }
            # Descargar el archivo ZIP
            Invoke-WebRequest -Uri $psexecUrl -OutFile $psexecZip
            # Extraer PsExec.exe
            Write-Host "`tExtrayendo PsExec..." -ForegroundColor Cyan
            Expand-Archive -Path $psexecZip -DestinationPath $psexecExtractPath -Force
            # Verificar si PsExec fue extraído correctamente
            if (-Not (Test-Path $psexecPath)) {
                Write-Host "`tError: No se pudo extraer PsExec.exe." -ForegroundColor Red
                return
            }
            Write-Host "`tPsExec descargado y extraído correctamente." -ForegroundColor Green
        } else {
            Write-Host "`tPsExec ya está instalado en: $psexecPath" -ForegroundColor Green
        }
        # Detectar el nombre correcto del grupo de administradores
        $grupoAdmin = ""
        $gruposLocales = net localgroup | Where-Object { $_ -match "Administrators|Administradores" }
        if ($gruposLocales -match "Administrators") {
            $grupoAdmin = "Administrators"
        } elseif ($gruposLocales -match "Administradores") {
            $grupoAdmin = "Administradores"
        } else {
            Write-Host "`tNo se encontró el grupo de administradores en el sistema." -ForegroundColor Red
            return
        }
        Write-Host "`tGrupo de administradores detectado: " -NoNewline
        Write-Host "$grupoAdmin" -ForegroundColor Green
        # Comandos con el nombre correcto del grupo
        $comando1 = "icacls C:\Windows\System32\en-us /grant `"$grupoAdmin`":F"
        $comando2 = "icacls C:\Windows\System32\en-us /grant `"NT AUTHORITY\SYSTEM`":F"
        # Comandos con formato correcto
        $psexecCmd1 = "`"$psexecPath`" /accepteula /s cmd /c `"$comando1`""
        $psexecCmd2 = "`"$psexecPath`" /accepteula /s cmd /c `"$comando2`""
        Write-Host "`nEjecutando primer comando: $comando1" -ForegroundColor Yellow
        $output1 = & cmd /c $psexecCmd1
        Write-Host $output1
        Write-Host "`nEjecutando segundo comando: $comando2" -ForegroundColor Yellow
        $output2 = & cmd /c $psexecCmd2
        Write-Host $output2
        Write-Host "`nModificación de permisos completada." -ForegroundColor Cyan
            $ResponderDriver = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea descargar e instalar el driver del lector?",
                "Descargar Driver",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($ResponderDriver -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Definir parámetros de la descarga
                $url = "https://softrestaurant.com/drivers?download=120:dp"
                $zipPath = "C:\Temp\Driver_DP.zip"
                $extractPath = "C:\Temp\Driver_DP"
                $exeName = "x64\Setup.msi"
                $validationPath = "C:\Temp\Driver_DP\x64\Setup.msi"
            
                # Llamar a la función de descarga y ejecución
                DownloadAndRun -url $url -zipPath $zipPath -extractPath $extractPath -exeName $exeName -validationPath $validationPath
            }
        
    } catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
})
#AplicacionesNS
$btnAplicacionesNS.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    # Definir una lista para almacenar los resultados
    $resultados = @()

    # Función para extraer valores de un archivo INI
    function Leer-Ini($filePath) {
        if (Test-Path $filePath) {
            $content     = Get-Content $filePath
            $dataSource  = ($content | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
            $catalog     = ($content | Select-String -Pattern "^Catalog=(.*)"    | Select-Object -First 1).Matches.Groups[1].Value
            $authType    = ($content | Select-String -Pattern "^autenticacion=(\d+)").Matches.Groups[1].Value
            $authUser    = if ($authType -eq "2") { "sa" } elseif ($authType -eq "1") { "Windows" } else { "Desconocido" }

            return @{
                DataSource = $dataSource
                Catalog    = $catalog
                Usuario    = $authUser
            }
        }
        return $null
    }

    # Lista de rutas principales con los archivos .ini correspondientes
    $pathsToCheck = @(
        @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
        @{ Path = "C:\NationalSoft\Softrestaurant12.0";    INI = "restaurant.ini"; Nombre = "SR12" },
        @{ Path = "C:\NationalSoft\Softrestaurant11.0";    INI = "restaurant.ini"; Nombre = "SR11" },
        @{ Path = "C:\NationalSoft\Softrestaurant10.0";    INI = "restaurant.ini"; Nombre = "SR10" },
        @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0";INI = "nshoteles.ini";   Nombre = "Hoteles" },
        @{ Path = "C:\NationalSoft\OnTheMinute4.5";        INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
    )

    foreach ($entry in $pathsToCheck) {
        $basePath   = $entry.Path
        $mainIni    = "$basePath\$($entry.INI)"
        $appName    = $entry.Nombre

        # Procesar archivo INI principal
        if (Test-Path $mainIni) {
            $iniData = Leer-Ini $mainIni
            if ($iniData) {
                $resultados += [PSCustomObject]@{
                    Aplicacion = $appName
                    INI        = $entry.INI
                    DataSource = $iniData.DataSource
                    Catalog    = $iniData.Catalog
                    Usuario    = $iniData.Usuario
                }
            }
        } else {
            # Si no se encuentra el INI principal
            $resultados += [PSCustomObject]@{
                Aplicacion = $appName
                INI        = "No encontrado"
                DataSource = "NA"
                Catalog    = "NA"
                Usuario    = "NA"
            }
        }

        # Procesar INIS adicionales sólo si aplica
        $inisFolder = "$basePath\INIS"
        if ($appName -eq "OnTheMinute" -and (Test-Path $inisFolder)) {
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
            if ($iniFiles.Count -gt 1) {
                # Multiempresa: agregar cada INI adicional
                foreach ($iniFile in $iniFiles) {
                    $iniData = Leer-Ini $iniFile.FullName
                    if ($iniData) {
                        $resultados += [PSCustomObject]@{
                            Aplicacion = $appName
                            INI        = $iniFile.Name
                            DataSource = $iniData.DataSource
                            Catalog    = $iniData.Catalog
                            Usuario    = $iniData.Usuario
                        }
                    }
                }
            }
        }
        elseif (Test-Path $inisFolder) {
            # Para todas las demás aplicaciones, conservar el comportamiento anterior
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
            foreach ($iniFile in $iniFiles) {
                $iniData = Leer-Ini $iniFile.FullName
                if ($iniData) {
                    $resultados += [PSCustomObject]@{
                        Aplicacion = $appName
                        INI        = $iniFile.Name
                        DataSource = $iniData.DataSource
                        Catalog    = $iniData.Catalog
                        Usuario    = $iniData.Usuario
                    }
                }
            }
        }
    }

    # Procesar RestCard.ini
    $restCardPath = "C:\NationalSoft\Restcard\RestCard.ini"
    if (Test-Path $restCardPath) {
        $resultados += [PSCustomObject]@{
            Aplicacion = "Restcard"
            INI        = "RestCard.ini"
            DataSource = "existe"
            Catalog    = "existe"
            Usuario    = "existe"
        }
    } else {
        $resultados += [PSCustomObject]@{
            Aplicacion = "Restcard"
            INI        = "No encontrado"
            DataSource = "NA"
            Catalog    = "NA"
            Usuario    = "NA"
        }
    }
    # Cálculo de anchos para presentación
    $columnas = @("Aplicacion","INI","DataSource","Catalog","Usuario")
    $anchos   = @{}
    foreach ($col in $columnas) { $anchos[$col] = $col.Length }
    foreach ($res in $resultados) {
        foreach ($col in $columnas) {
            if ($res.$col.Length -gt $anchos[$col]) {
                $anchos[$col] = $res.$col.Length
            }
        }
    }

    # Mostrar encabezados
    $titulos = $columnas | ForEach-Object { $_.PadRight($anchos[$_] + 2) }
    Write-Host ($titulos -join "") -ForegroundColor Cyan

    # Línea separadora
    $separador = $columnas | ForEach-Object { ("-" * $anchos[$_]).PadRight($anchos[$_] + 2) }
    Write-Host ($separador -join "") -ForegroundColor Cyan

    # Mostrar resultados
    foreach ($res in $resultados) {
        $fila = $columnas | ForEach-Object { $res.$_.PadRight($anchos[$_] + 2) }
        if ($res.INI -eq "No encontrado") {
            Write-Host ($fila -join "") -ForegroundColor Red
        } else {
            Write-Host ($fila -join "")
        }
    }
})
# Cambiar OTM de DBL a SQL y viceverza
$btnCambiarOTM.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    # Ruta de configuración
    $syscfgPath = "C:\Windows\SysWOW64\Syscfg45_2.0.dll"
    $iniPath = "C:\NationalSoft\OnTheMinute4.5"

    # Verificar si existe el archivo de configuración
    if (-not (Test-Path $syscfgPath)) {
        [System.Windows.Forms.MessageBox]::Show("El archivo de configuración no existe.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        Write-Host "`tEl archivo de configuración no existe." -ForegroundColor Red
        return
    }

    # Leer la configuración actual
    $fileContent = Get-Content $syscfgPath
    $isSQL = $fileContent -match "494E5354414C4C=1" -and $fileContent -match "56455253495354454D41=3"
    $isDBF = $fileContent -match "494E5354414C4C=2" -and $fileContent -match "56455253495354454D41=2"

    # Verificar si la configuración es válida
    if (!$isSQL -and !$isDBF) {
        [System.Windows.Forms.MessageBox]::Show("No se detectó una configuración válida de SQL o DBF.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        Write-Host "`tNo se detectó una configuración válida de SQL o DBF." -ForegroundColor Red
        return
    }

    # Obtener todos los archivos .ini en la carpeta
    $iniFiles = Get-ChildItem -Path $iniPath -Filter "*.ini"

    # Si no hay archivos INI, mostrar error
    if ($iniFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos INI en $iniPath.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        Write-Host "`tNo se encontraron archivos INI en $iniPath." -ForegroundColor Red
        return
    }

    # Variables para almacenar los archivos INI detectados
    $iniSQLFile = $null
    $iniDBFFile = $null

    # Iterar sobre los archivos INI y verificar el contenido de cada uno
    foreach ($iniFile in $iniFiles) {
        $content = Get-Content $iniFile.FullName

        # Verificar si contiene la sección Provider para DBF (VFPOLEDB.1) o SQL (SQLOLEDB.1)
        if ($content -match "Provider=VFPOLEDB.1" -and -not $iniDBFFile) {
            $iniDBFFile = $iniFile
        }
        if ($content -match "Provider=SQLOLEDB.1" -and -not $iniSQLFile) {
            $iniSQLFile = $iniFile
        }

        # Si ambos archivos ya han sido encontrados, salir del bucle
        if ($iniSQLFile -and $iniDBFFile) {
            break
        }
    }

    # Si no se encontraron ambos archivos, mostrar mensaje
    if (-not $iniSQLFile -or -not $iniDBFFile) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron los archivos INI esperados.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
        Write-Host "`tNo se encontraron los archivos INI esperados." -ForegroundColor Red
        Write-Host "`tArchivos encontrados:" -ForegroundColor Yellow
        $iniFiles | ForEach-Object { Write-Host "`t- $_.Name" }
        return
    }

    # Confirmar cambio de configuración
    $currentConfig = if ($isSQL) { "SQL" } else { "DBF" }
    $newConfig = if ($isSQL) { "DBF" } else { "SQL" }
    $message = "Actualmente tienes configurado: $currentConfig.`n¿Quieres cambiar a $newConfig?"
    $result = [System.Windows.Forms.MessageBox]::Show($message, "Cambiar Configuración", [System.Windows.Forms.MessageBoxButtons]::YesNo)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Modificar Syscfg45_2.0.dll para cambiar configuración
        if ($newConfig -eq "SQL") {
            Write-Host "`tCambiando a SQL: C:\Windows\SysWOW64\Syscfg45_2.0.dll" -ForegroundColor Yellow
            Write-Host "`t494E5354414C4C=1"
            Write-Host "`t56455253495354454D41=3"
            (Get-Content $syscfgPath) -replace "494E5354414C4C=2", "494E5354414C4C=1" | Set-Content $syscfgPath
            (Get-Content $syscfgPath) -replace "56455253495354454D41=2", "56455253495354454D41=3" | Set-Content $syscfgPath
        } else {
            Write-Host "`tCambiando a DBF: C:\Windows\SysWOW64\Syscfg45_2.0.dll" -ForegroundColor Yellow
            Write-Host "`t494E5354414C4C=2"
            Write-Host "`t56455253495354454D41=1"
            (Get-Content $syscfgPath) -replace "494E5354414C4C=1", "494E5354414C4C=2" | Set-Content $syscfgPath
            (Get-Content $syscfgPath) -replace "56455253495354454D41=3", "56455253495354454D41=2" | Set-Content $syscfgPath
        }

        # Renombrar archivos INI
        if ($newConfig -eq "SQL") {
            Rename-Item -Path $iniDBFFile.FullName -NewName "checadorsql_DBF_old.ini" -ErrorAction Stop
            Rename-Item -Path $iniSQLFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
        } else {
            Rename-Item -Path $iniSQLFile.FullName -NewName "checadorsql_SQL_old.ini" -ErrorAction Stop
            Rename-Item -Path $iniDBFFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
        }

        [System.Windows.Forms.MessageBox]::Show("Configuración cambiada exitosamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK)
        Write-Host "Configuración cambiada exitosamente." -ForegroundColor Green
    }
})


<# ---------------------------------------------------------------------------
  Formulario "Crear Usuario de Windows" con correcciones:
  - Obtención confiable de grupos Administradores y Usuarios estándar
  - Validación de complejidad de contraseña antes de New-LocalUser
  - Manejo adecuado de errores al verificar pertenencia a grupos
--------------------------------------------------------------------------- #>
$btnAddUser.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray

    # Crear formulario
    $formAddUser = Create-Form -Title "Crear Usuario de Windows" -Size (New-Object System.Drawing.Size(450, 250))

    # Controles
    $txtUsername = Create-TextBox -Location (New-Object System.Drawing.Point(120, 20)) -Size (New-Object System.Drawing.Size(290, 30))
    $lblUsername = Create-Label -Text "Nombre:" -Location (New-Object System.Drawing.Point(10, 20))
    $txtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(120, 60)) -Size (New-Object System.Drawing.Size(290, 30)) -UseSystemPasswordChar $true
    $lblPassword = Create-Label -Text "Contraseña:" -Location (New-Object System.Drawing.Point(10, 60))
    $cmbType     = Create-ComboBox -Location (New-Object System.Drawing.Point(120, 100)) -Size (New-Object System.Drawing.Size(290, 30)) -Items @("Usuario estándar","Administrador")
    $lblType     = Create-Label -Text "Tipo:" -Location (New-Object System.Drawing.Point(10, 100))

    # Determinar nombres de grupos estándar y administradores por SID
    $adminGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-544').Name
    $userGroup  = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-545').Name

    # Botones
    $btnCreate = Create-Button -Text "Crear"    -Location (New-Object System.Drawing.Point(10, 150))  -Size (New-Object System.Drawing.Size(130, 30))
    $btnCancel = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(150, 150)) -Size (New-Object System.Drawing.Size(130, 30))
    $btnShow   = Create-Button -Text "Mostrar usuarios" -Location (New-Object System.Drawing.Point(290, 150)) -Size (New-Object System.Drawing.Size(130, 30))
    # Mostrar usuarios existentes
    $btnShow.Add_Click({
        Write-Host "`nUsuarios actuales en el sistema:`n" -ForegroundColor Cyan
            # Obtener todos los usuarios locales
            $users = Get-LocalUser
            # Crear objetos con formato similar al de impresoras
            $usersTable = $users | ForEach-Object {
                $user = $_
                # Determinar estado
                $estado = if ($user.Enabled) { "Habilitado" } else { "Deshabilitado" }
                
                # Determinar tipo de usuario
                $tipoUsuario = "Usuario estándar"
                
                # Verificar si es administrador
                try {
                    $adminMembers = Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop
                    if ($adminMembers | Where-Object { $_.SID -eq $user.SID }) {
                        $tipoUsuario = "Administrador"
                    }
                    else {
                        # Verificar grupo de usuarios estándar
                        $userMembers = Get-LocalGroupMember -Group $userGroup -ErrorAction Stop
                        if (-not ($userMembers | Where-Object { $_.SID -eq $user.SID })) {
                            # Buscar en otros grupos
                            $grupos = Get-LocalGroup | ForEach-Object {
                                if (Get-LocalGroupMember -Group $_ | Where-Object { $_.SID -eq $user.SID }) {
                                    $_.Name
                                }
                            }
                            $tipoUsuario = "Miembro de: " + ($grupos -join ", ")
                        }
                    }
                }
                catch {
                    $tipoUsuario = "Error verificando grupos"
                }
                
                # Acortar texto si es muy largo (como en impresoras)
                $nombre = $user.Name.Substring(0, [Math]::Min(25, $user.Name.Length))
                $tipo = $tipoUsuario.Substring(0, [Math]::Min(40, $tipoUsuario.Length))
                
                [PSCustomObject]@{
                    Nombre = $nombre
                    Tipo   = $tipo
                    Estado = $estado
                }
            }
        
            # Mostrar en tabla formateada
            if ($usersTable.Count -gt 0) {
                Write-Host ("{0,-25} {1,-40} {2,-15}" -f "Nombre", "Tipo", "Estado")
                Write-Host ("{0,-25} {1,-40} {2,-15}" -f "------", "------", "------")
                
                $usersTable | ForEach-Object { 
                    Write-Host ("{0,-25} {1,-40} {2,-15}" -f $_.Nombre, $_.Tipo, $_.Estado) 
                }
            } else {
                Write-Host "No se encontraron usuarios."
            }
        })
    # Crear usuario y asignar grupo
    $btnCreate.Add_Click({
        $username = $txtUsername.Text.Trim()
        $password = $txtPassword.Text
        $type     = $cmbType.SelectedItem

        if (-not $username -or -not $password) {
            Write-Host "`nError: Nombre y contraseña son requeridos" -ForegroundColor Red; return
        }

        # Validar complejidad mínima: 8+ caracteres, mayúscula, minúscula, número y símbolo
        if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '\d' -or $password -notmatch '[^\w]') {
            Write-Host "`nError: La contraseña debe tener al menos 8 caracteres, incluir mayúscula, minúscula, número y símbolo" -ForegroundColor Red; return
        }

        try {
            if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
                Write-Host "`nError: El usuario '$username' ya existe" -ForegroundColor Red; return
            }
            # Crear usuario
            $securePwd = ConvertTo-SecureString $password -AsPlainText -Force
            New-LocalUser -Name $username -Password $securePwd -AccountNeverExpires -PasswordNeverExpires
            Write-Host "`nUsuario '$username' creado exitosamente" -ForegroundColor Green

            # Asignar grupo
            $group = if ($type -eq 'Administrador') { $adminGroup } else { $userGroup }
            Add-LocalGroupMember -Group $group -Member $username
            Write-Host "`tUsuario agregado al grupo $group" -ForegroundColor Cyan

            $formAddUser.Close()
        } catch {
            Write-Host "`nError durante la creación del usuario: $_" -ForegroundColor Red
        }
    })
    # Cancelar
    $btnCancel.Add_Click({ Write-Host "`tOperación cancelada." -ForegroundColor Yellow; $formAddUser.Close() })
    # Agregar controles
    $formAddUser.Controls.AddRange(@($txtUsername,$txtPassword,$cmbType,$btnCreate,$btnCancel,$btnShow,$lblUsername,$lblPassword,$lblType))
    $formAddUser.ShowDialog()
})
# SQL SENTENCIAS QUERIES Y TODO
function Remove-SqlComments {
    param(
        [string]$Query
    )
    # 1) Eliminar comentarios de bloque /* … */ (incluso si abarcan varias líneas)
    $cleanedQuery = $Query -replace '(?s)/\*.*?\*/', ''

    # 2) Eliminar líneas que comienzan con -- (incluyendo espacios iniciales)
    $cleanedQuery = $cleanedQuery -replace '(?m)^\s*--.*\n?', ''

    # 3) Eliminar comentarios en línea con -- (sin afectar URLs)
    $cleanedQuery = $cleanedQuery -replace '(?<!\w)--.*$', ''

    return $cleanedQuery.Trim()
}
function ConvertTo-DataTable {
    param($InputObject)
    $dt = New-Object System.Data.DataTable
    if (-not $InputObject) { return $dt }

    $columns = $InputObject[0].Keys
    foreach ($col in $columns) {
        $dt.Columns.Add($col) | Out-Null
    }

    foreach ($row in $InputObject) {
        $dr = $dt.NewRow()
        foreach ($col in $columns) {
            $dr[$col] = $row[$col]
        }
        $dt.Rows.Add($dr)
    }
    return $dt
}
function Execute-SqlQuery {
    param (
        [string]$server,
        [string]$database,
        [string]$query
    )
    try {
        $connectionString = "Server=$server;Database=$database;User Id=$global:user;Password=$global:password;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        
        # Variable para capturar mensajes
        $infoMessages = New-Object System.Collections.ArrayList
        
        # Evento para capturar mensajes de SQL
        $connection.add_InfoMessage({
            param($sender, $e)
            $infoMessages.Add($e.Message) | Out-Null
        })
        
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query

        # Detectar si es una consulta que devuelve datos
        if ($query -match "(?si)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|RESTORE)") {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable) | Out-Null

            # Obtener mensajes durante la ejecución
            $command.ExecuteNonQuery() | Out-Null

            return @{
                DataTable = $dataTable
                Messages = $infoMessages
            }
        } 
        else {
            $rowsAffected = $command.ExecuteNonQuery()
            return @{
                RowsAffected = $rowsAffected
                Messages = $infoMessages
            }
        }
    }
    catch {
        Write-Host "Error en consulta: $($_.Exception.Message)" -ForegroundColor Red
        throw $_
    }
    finally {
        $connection.Close()
    }
}
#Boton para desconectar de la base de datos
function ConvertTo-DataTable {
    param($InputObject)
    $dt = New-Object System.Data.DataTable
    $InputObject | ForEach-Object {
        if (!$dt.Columns.Count) {
            $_.PSObject.Properties | ForEach-Object {
                $dt.Columns.Add($_.Name, $_.Value.GetType())
            }
        }
        $row = $dt.NewRow()
        $_.PSObject.Properties | ForEach-Object {
            $row[$_.Name] = $_.Value
        }
        $dt.Rows.Add($row)
    }
    return $dt
}

# Función original para mostrar resultados en consola (SIN CAMBIOS)
function Show-ResultsConsole {
    param (
        [string]$query
    )
    try {
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query
        
        if ($results.GetType().Name -eq 'Hashtable') {
            $consoleData = $results.ConsoleData
            if ($consoleData.Count -gt 0) {
                $columns = $consoleData[0].Keys
                $columnWidths = @{}
                foreach ($col in $columns) {
                    $columnWidths[$col] = $col.Length
                }
                
                Write-Host ""
                $header = ""
                foreach ($col in $columns) {
                    $header += $col.PadRight($columnWidths[$col] + 4)
                }
                Write-Host $header
                Write-Host ("-" * $header.Length)
                
                foreach ($row in $consoleData) {
                    $rowText = ""
                    foreach ($col in $columns) {
                        $rowText += ($row[$col].ToString()).PadRight($columnWidths[$col] + 4)
                    }
                    Write-Host $rowText
                }
            } 
            else {
                Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "`nFilas afectadas: $results" -ForegroundColor Green
        }
    } 
    catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
}
#EJECUTANDO
$btnExecute.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    try {
        # Resetear a estilos normales antes de ejecutar
        $dgvResults.DefaultCellStyle.ForeColor = $script:originalForeColor
        $dgvResults.ColumnHeadersDefaultCellStyle.BackColor = $script:originalHeaderBackColor
        $dgvResults.AutoSizeColumnsMode = $script:originalAutoSizeMode
        $dgvResults.DefaultCellStyle.ForeColor = $originalForeColor
        $dgvResults.ColumnHeadersDefaultCellStyle.BackColor = $originalHeaderBackColor
        $dgvResults.AutoSizeColumnsMode = $originalAutoSizeMode
        $toolTip.SetToolTip($dgvResults, $null)
        $dgvResults.DataSource = $null
        $dgvResults.Rows.Clear()
        Clear-Host
        $selectedDb = $cmbDatabases.SelectedItem
        if (-not $selectedDb) { throw "Selecciona una base de datos" }
        
        $rawQuery = $rtbQuery.Text
        $cleanQuery = Remove-SqlComments -Query $rawQuery
        $result = Execute-SqlQuery -server $global:server -database $selectedDb -query $cleanQuery

        # Mostrar mensajes de SQL (incluyendo progreso de RESTORE)
        if ($result.Messages.Count -gt 0) {
            Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
            $result.Messages | ForEach-Object { Write-Host $_ }
        }

        if ($result.DataTable) {
            # Mostrar resultados en DataGridView
            $dgvResults.DataSource = $result.DataTable.DefaultView
            $dgvResults.Enabled = $true
            
            Write-Host "`nColumnas obtenidas: $($result.DataTable.Columns.ColumnName -join ', ')" -ForegroundColor Cyan
            
            # Estilos para resultados exitosos
            $dgvResults.DefaultCellStyle.ForeColor = 'Blue'
            $dgvResults.AlternatingRowsDefaultCellStyle.BackColor = '#F0F8FF'
            
            # Ajuste automático de columnas
            $dgvResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
            foreach ($col in $dgvResults.Columns) {
                $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::DisplayedCells
                $col.Width = [Math]::Max($col.Width, $col.HeaderText.Length * 8)
            }

            # Mostrar en consola si no hay datos
            if ($result.DataTable.Rows.Count -eq 0) {
                Write-Host "La consulta no devolvió resultados" -ForegroundColor Yellow
            }
            else {
                $result.DataTable | Format-Table -AutoSize | Out-String | Write-Host
            }
        }
        else {
            # Mostrar filas afectadas para consultas no SELECT
            Write-Host "`nFilas afectadas: $($result.RowsAffected)" -ForegroundColor Green
        }
    }
    catch {
        # Manejo de errores
        $errorTable = New-Object System.Data.DataTable
        $errorTable.Columns.Add("Tipo") | Out-Null
        $errorTable.Columns.Add("Mensaje") | Out-Null
        $errorTable.Columns.Add("Detalle") | Out-Null
        
        $cleanQuery = $rtbQuery.Text -replace '(?s)/\*.*?\*/', '' -replace '(?m)^\s*--.*'
        $shortQuery = if ($cleanQuery.Length -gt 50) { $cleanQuery.Substring(0,47) + "..." } else { $cleanQuery }
        
        $errorTable.Rows.Add("ERROR SQL", $_.Exception.Message, $shortQuery) | Out-Null
        
        # Configurar DataGridView para errores
        $dgvResults.DataSource = $errorTable
        # 1) Permitir salto de línea en la columna de mensaje de error (índice 1)
        $dgvResults.Columns[1].DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
        $dgvResults.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
        #$dgvResults.RowTemplate.Height = 40
        $dgvResults.AutoSizeColumnsMode = 'Fill'
        $dgvResults.Columns[0].Width = 100
        $dgvResults.Columns[1].Width = 300
        $dgvResults.Columns[2].Width = 200
        
        # Estilos de error
        $dgvResults.DefaultCellStyle.ForeColor = 'Red'
        $dgvResults.ColumnHeadersDefaultCellStyle.BackColor = '#FFB3B3'
        $toolTip.SetToolTip($dgvResults, "Consulta completa:`n$cleanQuery")
        
        # Mensaje en consola
        Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
        Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Consulta: $shortQuery" -ForegroundColor Cyan
        Write-Host "====================================" -ForegroundColor Red
    }
})
$btnConnectDb.Add_Click({
    try {
        $global:server = $txtServer.Text
        $global:user = $txtUser.Text
        $global:password = $txtPassword.Text
        
        if (-not $global:server -or -not $global:user -or -not $global:password) {
            throw "Complete todos los campos de conexión"
        }
        
        $connStr = "Server=$global:server;User Id=$global:user;Password=$global:password;"
        $global:connection = [System.Data.SqlClient.SqlConnection]::new($connStr)
        $global:connection.Open()

        # Obtener bases de datos
        $query = "SELECT name FROM sys.databases WHERE name NOT IN ('tempdb','model','msdb') AND state_desc = 'ONLINE' ORDER BY CASE WHEN name = 'master' THEN 0 ELSE 1 END, name;"
        $result = Execute-SqlQuery -server $global:server -database "master" -query $query
        $cmbDatabases.Items.Clear()
        foreach ($row in $result.DataTable.Rows) {
            $cmbDatabases.Items.Add($row["name"])
        }
        
        $cmbDatabases.Enabled = $true
        $cmbDatabases.SelectedIndex = 0
        $lblConnectionStatus.Text = @"
Conectado a:
Servidor: $($global:server)
Base de datos: $($global:database)
"@.Trim()
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green
        # After successful connection
        $txtServer.Enabled = $false
        $txtUser.Enabled = $false
        $txtPassword.Enabled = $false
        $btnExecute.Enabled = $true
        $cmbQueries.Enabled = $true
        $btnConnectDb.Enabled = $false
        $btnBackup.Enabled    = $True
        $btnDisconnectDb.Enabled = $true
        $btnExecute.Enabled = $true
        $rtbQuery.Enabled = $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error de conexión: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
    }
})
# Actualizar el evento SelectedIndexChanged
$cmbDatabases.Add_SelectedIndexChanged({
    if ($cmbDatabases.SelectedItem) {
        $global:database = $cmbDatabases.SelectedItem
        
        # Actualizar etiqueta
        $lblConnectionStatus.Text = @"
Conectado a:
Servidor: $($global:server)
Base de datos: $($global:database)
"@.Trim()
        #$lblConnectionStatus.Text = "Conectado a: $global:server | BDD: $global:database"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green
        
        Write-Host "Base de datos seleccionada: $($cmbDatabases.SelectedItem)" -ForegroundColor Cyan
    }
})
#boton para desconectar a la base de datos.
$btnDisconnectDb.Add_Click({
    try {
        Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow
        $global:connection.Close()
        $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnectDb.Enabled    = $True
            $btnBackup.Enabled        = $false
            $btnDisconnectDb.Enabled = $false
            $btnExecute.Enabled      = $false
            $rtbQuery.Enabled        = $false
            $txtServer.Enabled = $true
            $txtUser.Enabled = $true
            $txtPassword.Enabled = $true
            $btnExecute.Enabled = $false
            $cmbQueries.Enabled = $false
            $cmbDatabases.Items.Clear()
            $cmbDatabases.Enabled = $false
    }
    catch {
            Write-Host "`nError al desconectar: $_" -ForegroundColor Red
        }
})
#Etiquetas clics------------------------------------------------------------------------------------------------
$lblHostname.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($lblHostname.Text)
        Write-Host "`nNombre del equipo copiado al portapapeles: $($lblHostname.Text)"
    })
    $txt_IpAdress.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($txt_IpAdress.Text)
        Write-Host "`nIP's copiadas al equipo: $($txt_IpAdress.Text)"
    })
            $txt_IpAdress.Add_MouseEnter($changeColorOnHover)
            $txt_IpAdress.Add_MouseLeave($restoreColorOnLeave)
$txt_AdapterStatus.Add_Click({
    # Cambiar sólo las que no sean ya Private
    Get-NetConnectionProfile |
      Where-Object { $_.NetworkCategory -ne 'Private' } |
      ForEach-Object {
          Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private
      }
    Write-Host "Todas las redes se han establecido como Privadas."
    Refresh-AdapterStatus
})
        $txt_AdapterStatus.Add_MouseEnter($changeColorOnHover)
        $txt_AdapterStatus.Add_MouseLeave($restoreColorOnLeave)
#Llamarla pro primera vez
Refresh-AdapterStatus
#Creación del APK
$btnCreateAPK.Add_Click({
    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
    $dllPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\up.dll"
    $infoPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\info.txt"
    
    try {
        Write-Host "`nIniciando proceso de creación de APK..." -ForegroundColor Cyan
        
        # Validar existencia de archivos
        if (-not (Test-Path $dllPath)) {
            Write-Host "Componente necesario no encontrado. Verifique la instalación del Enlace Android." -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("Componente necesario no encontrado. Verifique la instalación del Enlace Android.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        if (-not (Test-Path $infoPath)) {
            Write-Host "Archivo de configuración no encontrado. Verifique la instalación del Enlace Android." -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("Archivo de configuración no encontrado. Verifique la instalación del Enlace Android.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        # Leer y parsear info.txt
        $jsonContent = Get-Content $infoPath -Raw | ConvertFrom-Json
        $versionApp = $jsonContent.versionApp
        Write-Host "Versión detectada: $versionApp" -ForegroundColor Green

        # Confirmación con usuario
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Se creará el APK versión: $versionApp`n¿Desea continuar?", 
            "Confirmación", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "Proceso cancelado por el usuario" -ForegroundColor Yellow
            return
        }

        # Selección de ubicación
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Archivo APK (*.apk)|*.apk"
        $saveDialog.FileName = "SRM_$versionApp.apk"
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        
        if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "Guardado cancelado por el usuario" -ForegroundColor Yellow
            return
        }

        # Copiar archivo
        Copy-Item -Path $dllPath -Destination $saveDialog.FileName -Force
        Write-Host "APK generado exitosamente en: $($saveDialog.FileName)" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("APK creado correctamente!", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    }
    catch {
        Write-Host "Error durante el proceso: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error durante la creación del APK. Consulte la consola para más detalles.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
<# ESTP ESTA EN EL RENGLON 3016---------------------------------------------------------------------------
  Sección "BACKUP" adaptada para uso ad-hoc con Admin Share C$
  - No valida permisos ni crea carpeta remota (SQL se encarga de la escritura)
  - Extrae correctamente nombre de máquina
--------------------------------------------------------------------------- #>
# ############# BACKUP #############
$btnBackup.Add_Click({
# 0. Extraer nombre de servidor para UNC
$serverRaw   = $global:server
$machinePart = $serverRaw.Split('\\')[0]
$machineName = $machinePart.Split(',')[0]
if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }

# Definir carpeta de respaldo remota (Admin Share C$)
$global:tempBackupFolder = "\\$machineName\C$\Temp\SQLBackups"

# ------------------------------- INICIO DE BLOQUE DE CREACIÓN Y PERMISOS -------------------------------
# 1.a) Crear carpeta si no existe
if (-not (Test-Path -Path $global:tempBackupFolder)) {
    try {
        Write-Host "Creando carpeta remota: $global:tempBackupFolder" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $global:tempBackupFolder -Force | Out-Null
        Write-Host "Carpeta creada correctamente." -ForegroundColor Green
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error al crear la carpeta $global:tempBackupFolder:`n$($_.Exception.Message)",
            "Error al crear carpeta de backup",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "Error | No se pudo crear la carpeta remota: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}
else {
    Write-Host "La carpeta remota ya existe: $global:tempBackupFolder" -ForegroundColor Gray
}

# 1.b) Asignar permisos de escritura a "Everyone"
try {
    $acl = Get-Acl -Path $global:tempBackupFolder
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "Everyone",
        "FullControl",
        "ContainerInherit, ObjectInherit",
        "None",
        "Allow"
    )
    if (-not $acl.Access | Where-Object { $_.IdentityReference -eq "Everyone" -and $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl }) {
        $acl.AddAccessRule($rule)
        Set-Acl -Path $global:tempBackupFolder -AclObject $acl
        Write-Host "Permisos asignados a Everyone." -ForegroundColor Green
    }
    else {
        Write-Host "Los permisos para Everyone ya estaban definidos." -ForegroundColor Gray
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Error al asignar permisos en $global:tempBackupFolder:`n$($_.Exception.Message)",
        "Error permisos",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Host "Error | No se pudieron asignar permisos: $($_.Exception.Message)" -ForegroundColor Red
    return
}
# ------------------------------- FIN DEL BLOQUE DE CREACIÓN Y PERMISOS -------------------------------   
    Write-Host "`t- - - Comenzando el proceso - - -" -ForegroundColor Gray

    # 1. Selección de Base de Datos
    $selectedDb = $cmbDatabases.SelectedItem
    if (-not $selectedDb) {
        [System.Windows.Forms.MessageBox]::Show(
            "Seleccione una base de datos primero",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
            Write-Host "Error | Seleccione una base de datos primero" -ForegroundColor Red
        return
    }

    # 2. Ruta del backup dentro de la carpeta remota
    $timestamp         = Get-Date -Format 'yyyyMMdd-HHmmss'
    $bakFileName       = "$selectedDb-$timestamp.bak"
    $global:backupPath = Join-Path $global:tempBackupFolder $bakFileName

    # 3. Crear y ejecutar Job de backup (SQL maneja permisos)
    $script = {
        param($srv,$usr,$pwd,$db,$path)
        $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$srv; Database=master; User Id=$usr; Password=$pwd")
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "BACKUP DATABASE [$db] TO DISK='$path' WITH CHECKSUM"
        $cmd.CommandTimeout = 0
        $cmd.ExecuteNonQuery()
        $conn.Close()
    }
    $global:backupJob = Start-Job -ScriptBlock $script -ArgumentList `
        $global:server, $global:user, $global:password, $selectedDb, $global:backupPath

    # 4. Crear formulario global con ProgressBar “ping‑pong”
    $global:backupProgressForm = New-Object System.Windows.Forms.Form
    $global:backupProgressForm.Text            = "Procesando Backup"
    $global:backupProgressForm.Size            = New-Object System.Drawing.Size(500,160)
    $global:backupProgressForm.StartPosition   = "CenterScreen"
    $global:backupProgressForm.FormBorderStyle = 'FixedDialog'
    $global:backupProgressForm.TopMost         = $true
    # 4.1 Instanciar ProgressBar como variable global
    $global:pb = New-Object System.Windows.Forms.ProgressBar
    $global:pb.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $global:pb.Minimum  = 0
    $global:pb.Maximum  = 100
    $global:pb.Value    = 0
    $global:pb.Size     = New-Object System.Drawing.Size(460,25)
    $global:pb.Location = New-Object System.Drawing.Point(10,20)
    # Label y botón Cancelar
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = "Trabajando, espere..."
    $lbl.AutoSize = $false
    $lbl.Size     = New-Object System.Drawing.Size(460,20)
    $lbl.Location = New-Object System.Drawing.Point(10,55)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = "Cancelar"
    $btnCancel.Size     = New-Object System.Drawing.Size(120,40)
    $btnCancel.Location = New-Object System.Drawing.Point(
        [int](($global:backupProgressForm.ClientSize.Width - 120) / 2),
        90
    )
    $global:backupProgressForm.Controls.AddRange(@($global:pb, $lbl, $btnCancel))
    $global:backupProgressForm.Show()
    # 5. Temporizador de animación “ping‑pong”
    $global:animTimer = New-Object System.Windows.Forms.Timer
    $animTimer.Interval = 400
    # Variable de dirección: +1 sube, -1 baja
    $script:direction = 1  
    $animTimer.Add_Tick({
        # Cuando llegue al máximo, invertimos a bajar
        if ($pb.Value -ge $pb.Maximum) {
            $script:direction = -1
        }
        # Cuando llegue al mínimo, invertimos a subir
        elseif ($pb.Value -le $pb.Minimum) {
            $script:direction = 1
        }
        # Calculamos nuevo valor y lo aplicamos directamente (la ProgressBar limita a min/max)
        $pb.Value += 10 * $script:direction
    })
    $animTimer.Start()
    # 6. Arrancar Job de backup en segundo plano
    $script = {
        param($srv,$usr,$pwd,$db,$path)
        $conn = New-Object System.Data.SqlClient.SqlConnection("
            Server=$srv; Database=master; User Id=$usr; Password=$pwd
        ")
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "
            BACKUP DATABASE [$db] TO DISK='$path' WITH CHECKSUM
        "
        $cmd.CommandTimeout = 0
        $cmd.ExecuteNonQuery()
        $conn.Close()
    }
    $global:backupJob = Start-Job -ScriptBlock $script -ArgumentList `
        $global:server, $global:user, $global:password, $selectedDb, $global:backupPath
# 7. Temporizador para vigilar el Job
    $global:backupTimer = New-Object System.Windows.Forms.Timer
    $global:backupTimer.Interval = 500
        $global:backupTimer.Add_Tick({
            if ($global:backupJob.State -in 'Completed','Failed','Stopped') {
                $animTimer.Stop(); $backupTimer.Stop()
                Receive-Job $global:backupJob | Out-Null
    
                if ($global:backupJob.State -eq 'Completed') {
                    # Mensaje de éxito
                    [System.Windows.Forms.MessageBox]::Show(
                        "Backup completado en servidor:`n$global:backupPath",
                        "Éxito",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-Host "Backup finalizado correctamente. Ruta: $global:backupPath" -ForegroundColor Green
                    # Intentar abrir carpeta remota solo si existe
                    if (Test-Path $global:tempBackupFolder) {
                                $uploadResponse = [System.Windows.Forms.MessageBox]::Show(
                                    "¿Deseas comprimir y subir el respaldo a TransferNow?",
                                    "Subir Respaldo",
                                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                                    [System.Windows.Forms.MessageBoxIcon]::Question)                            
                                if ($uploadResponse -eq [System.Windows.Forms.DialogResult]::Yes) 
                                {
                                    Start-BackupUploadProcess -BackupPath $global:backupPath
                                }
                                Start-Process explorer.exe $global:tempBackupFolder
                        
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show(
                            "No se pudo abrir la carpeta de backup:`n$global:tempBackupFolder",
                            "Atención",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning)
                                                Write-Host "Error | No se pudo abrir la carpeta de backup:`n`t$global:tempBackupFolder`n`tRevise directo en el equipo servidor" -ForegroundColor Red
                    }
                }
                elseif ($global:backupJob.State -eq 'Stopped') {
                    [System.Windows.Forms.MessageBox]::Show("Backup cancelado por el usuario","Cancelado",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
                                        Write-Host "Backup cancelado por el usuario" -ForegroundColor Yellow
                }
                else {
                    $err = Receive-Job $global:backupJob -ErrorAction SilentlyContinue
                    [System.Windows.Forms.MessageBox]::Show("Error en backup:`n$err","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
                    Write-Host "Error | Error en backup:`n$err" -ForegroundColor Red
                }

            # Limpieza
            Remove-Job $global:backupJob -Force
            $global:backupProgressForm.Close()
        }
    })
    $global:backupTimer.Start()
# 8. Cancelar el Job desde el botón
    $btnCancel.Add_Click({
        if ($global:backupJob -and $global:backupJob.State -in @('Running','NotStarted')) {
            Stop-Job $global:backupJob
        }
    })
})
# Cierre seguro al cerrar la aplicación
$formPrincipal.Add_FormClosing({
    if ($global:backupJob -and $global:backupJob.State -eq 'Running') {
        Stop-Job $global:backupJob
    }
    if ($global:animTimer)   { $global:animTimer.Stop()   }
    if ($global:backupTimer) { $global:backupTimer.Stop() }
})
# Botón para salir
$btnExit.Add_Click({
    $formPrincipal.Dispose()
    $formPrincipal.Close()
})


# Instalar megatools con Chocolatey
Write-Host "`nInstalando megatools con Chocolatey..." -ForegroundColor Cyan

# Función para verificar e instalar Chocolatey (si no existe)
function Check-Chocolatey {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "Chocolatey no está instalado. Instalando..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    }
    return $true
}

# Verificar/instalar Chocolatey
if (Check-Chocolatey) {
    try {
        # Instalar megatools
        choco install megatools -y --force
        Write-Host "megatools instalado correctamente." -ForegroundColor Green
        
        # Agregar a PATH
        $env:Path += ";C:\ProgramData\chocolatey\bin"
        [Environment]::SetEnvironmentVariable("Path", $env:Path, "Machine")
    }
    catch {
        Write-Host "Error al instalar megatools: $_" -ForegroundColor Red
    }
}

# Configuración de megatools
$megaConfig = @"
[Login]
Username = gerardo.zermeno@nationalsoft.mx
Password = National09$#
"@

$configPath = "$env:APPDATA\megatools.ini"

try {
    # Crear archivo de configuración
    $megaConfig | Out-File -FilePath $configPath -Encoding utf8 -ErrorAction Stop
    
    # Establecer permisos
    $acl = Get-Acl $configPath
    $acl.SetAccessRuleProtection($true, $false)  # Deshabilitar herencia
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
        "FullControl",
        "Allow"
    )
    $acl.AddAccessRule($rule)
    Set-Acl -Path $configPath -AclObject $acl
    
    Write-Host "Configuración de megatools creada correctamente." -ForegroundColor Green
}
catch {
    Write-Host "Error al crear configuración: $($_.Exception.Message)" -ForegroundColor Red
}

# Reemplazar la función Start-BackupUploadProcess con esta nueva versión
function Start-BackupUploadProcess {
    param([string]$BackupPath)
    
    # Configuración
    $zipPath = "$BackupPath.zip"
    $7zipPath = "C:\Program Files\7-Zip\7z.exe"
    
    try {
        # Comprimir con contraseña (manteniendo tu método actual)
        Write-Host "Comprimiendo respaldo..." -ForegroundColor Cyan
        & $7zipPath a -tzip -p"National09" -mem=AES256 $zipPath $BackupPath

        # Subir a Mega.nz
        Write-Host "Subiendo a Mega.nz..." -ForegroundColor Yellow
        
        # Obtener nombre único para el archivo remoto
        $remoteName = "SRBackup_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".zip"
        
        # Comando megatools para subir
        $uploadCmd = "megatools put `"$zipPath`" --path /Backups/SQL/$remoteName"
        $result = cmd /c $uploadCmd 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            $downloadLink = $result | Select-String -Pattern 'https://mega\.nz/file/[^\s]+' | ForEach-Object { $_.Matches.Value }
            
            [System.Windows.Forms.MessageBox]::Show(
                "Backup subido exitosamente a Mega.nz`nEnlace de descarga:`n$downloadLink`n(Se ha copiado al portapapeles)",
                "Éxito",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Copiar enlace al portapapeles
            $downloadLink | Set-Clipboard
            Write-Host "Enlace copiado al portapapeles: $downloadLink" -ForegroundColor Green
        }
        else {
            Write-Host "Error en subida: $result" -ForegroundColor Red
            throw "Error al subir a Mega.nz"
        }
    }
    catch {
        Write-Host "Error detallado: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show(
            "Error durante la subida a Mega.nz:`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        # Limpieza opcional del ZIP local
        if (Test-Path $zipPath) { 
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
            Write-Host "Archivo ZIP local eliminado" -ForegroundColor Yellow
        }
    }
}



$formPrincipal.Refresh()
$formPrincipal.ShowDialog()
