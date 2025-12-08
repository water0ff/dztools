#requires -Version 5.0
chcp 65001 > $null  # Cambiar codepage de la consola a UTF-8
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding
$global:version = "beta.25.12.03.1046"
try {
    Write-Host "Cargando ensamblados de Windows Forms..." -ForegroundColor Yellow
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "✓ System.Windows.Forms cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando System.Windows.Forms: $_" -ForegroundColor Red
    pause
    exit 1
}
try {
    Add-Type -AssemblyName System.Drawing
    Write-Host "✓ System.Drawing cargado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error cargando System.Drawing: $_" -ForegroundColor Red
    pause
    exit 1
}
try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "✓ VisualStyles habilitado" -ForegroundColor Green
} catch {
    Write-Host "✗ Error habilitando VisualStyles: $_" -ForegroundColor Red
}
if (Get-Command Set-ExecutionPolicy -ErrorAction SilentlyContinue) {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
}
Write-Host "`nImportando módulos..." -ForegroundColor Yellow
$modulesPath = Join-Path $PSScriptRoot "modules"
$modules = @(
    "GUI.psm1",
    "Database.psm1",
    "Utilities.psm1",
    "Installers.psm1"
)
foreach ($module in $modules) {
    $modulePath = Join-Path $modulesPath $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -ErrorAction Stop -DisableNameChecking
            Write-Host "  ✓ $module" -ForegroundColor Green
        } catch {
            Write-Host "  ✗ Error importando módulo: $module" -ForegroundColor Red
            Write-Host "    Ruta    : $modulePath" -ForegroundColor DarkYellow
            Write-Host "    Mensaje : $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "    Detalle : $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkYellow
            Write-Host "    Stack   : $($_.ScriptStackTrace)" -ForegroundColor DarkGray
            throw
        }
    } else {
        Write-Host "  ✗ $module no encontrado" -ForegroundColor Red
    }
}
$global:defaultInstructions = @"
----- CAMBIOS -----
- Carga de INIS en la conexión a BDD.
- Se cambió la instalación de SSMS14 a SSMS21.
- Se deshabilitó la subida a mega.
- Restructura del proceso de Backups (choco).
- Se agregó subida a megaupload.
- Se agregó compresión con contraseña de respaldos
- Se agregó compresión con contraseña de respaldos
- Se agregó consola de cambios y tool tip para botones
- Reorganización de botones
- Query Browser para SQL en pestaña: Base de datos
- - Ahora se pueden agregar comentarios con "-" y entre "/* */"
- - Tabla en consola
- - Obtener columnas en consola
"@
function Register-GlobalErrorHandlers {
    try {
        Write-DzDebug "`t[DEBUG] Registrando manejadores globales de excepciones" -Color DarkGray

        # Modo de captura de excepciones
        [System.Windows.Forms.Application]::SetUnhandledExceptionMode(
            [System.Windows.Forms.UnhandledExceptionMode]::CatchException
        )

        # Manejador de excepciones del thread UI
        [System.Windows.Forms.Application]::add_ThreadException({
                param($sender, $eventArgs)

                $ex = $eventArgs.Exception

                Write-Host "`n===============================================" -ForegroundColor Red
                Write-Host "EXCEPCIÓN DE THREAD UI CAPTURADA" -ForegroundColor Red
                Write-Host "===============================================" -ForegroundColor Red
                Write-Host "Tipo     : $($ex.GetType().FullName)" -ForegroundColor Yellow
                Write-Host "Mensaje  : $($ex.Message)" -ForegroundColor Yellow

                if ($ex.InnerException) {
                    Write-Host "`nExcepción interna:" -ForegroundColor Cyan
                    Write-Host "  Tipo    : $($ex.InnerException.GetType().FullName)" -ForegroundColor Yellow
                    Write-Host "  Mensaje : $($ex.InnerException.Message)" -ForegroundColor Yellow
                }

                if ($ex.StackTrace) {
                    Write-Host "`nStack Trace:" -ForegroundColor Cyan
                    Write-Host $ex.StackTrace -ForegroundColor Gray
                }

                Write-Host "===============================================`n" -ForegroundColor Red

                # Mostrar al usuario
                [System.Windows.Forms.MessageBox]::Show(
                    "Error en la interfaz:`n`n$($ex.Message)`n`nTipo: $($ex.GetType().Name)`n`nRevise la consola para más detalles.",
                    "Error no controlado",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            })

        # Manejador de excepciones del dominio de aplicación
        [System.AppDomain]::CurrentDomain.add_UnhandledException({
                param($sender, $eventArgs)

                $ex = $eventArgs.ExceptionObject

                Write-Host "`n===============================================" -ForegroundColor Red
                Write-Host "EXCEPCIÓN NO CONTROLADA DEL DOMINIO" -ForegroundColor Red
                Write-Host "===============================================" -ForegroundColor Red
                Write-Host "Es terminante: $($eventArgs.IsTerminating)" -ForegroundColor Yellow

                if ($ex -is [System.Exception]) {
                    Write-Host "Tipo     : $($ex.GetType().FullName)" -ForegroundColor Yellow
                    Write-Host "Mensaje  : $($ex.Message)" -ForegroundColor Yellow

                    if ($ex.InnerException) {
                        Write-Host "`nExcepción interna:" -ForegroundColor Cyan
                        Write-Host "  Tipo    : $($ex.InnerException.GetType().FullName)" -ForegroundColor Yellow
                        Write-Host "  Mensaje : $($ex.InnerException.Message)" -ForegroundColor Yellow
                    }

                    if ($ex.StackTrace) {
                        Write-Host "`nStack Trace:" -ForegroundColor Cyan
                        Write-Host $ex.StackTrace -ForegroundColor Gray
                    }
                } else {
                    Write-Host "Objeto de excepción: $ex" -ForegroundColor Yellow
                }

                Write-Host "===============================================`n" -ForegroundColor Red
            })

        Write-DzDebug "`t[DEBUG] Manejadores registrados correctamente" -Color Green

    } catch {
        Write-DzDebug "`t[DEBUG] No se pudieron registrar los manejadores globales: $_" -Color DarkYellow
    }
}
function Initialize-Environment {
    if (!(Test-Path -Path "C:\Temp")) {
        try {
            New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null
            Write-Host "Carpeta 'C:\Temp' creada." -ForegroundColor Green
        } catch {
            Write-Host "Error creando C:\Temp: $_" -ForegroundColor Yellow
        }
    }
    try {
        $debugEnabled = Initialize-DzToolsConfig
        Write-DzDebug "`t[DEBUG]Configuración de debug cargada (debug=$debugEnabled)" -Color DarkGray
    } catch {
        Write-Host "Advertencia: No se pudo inicializar la configuración de debug. $_" -ForegroundColor Yellow
    }
    return $true
}
function New-MainForm {
    Write-Host "Creando formulario principal..." -ForegroundColor Yellow
    try {
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $formPrincipal = New-Object System.Windows.Forms.Form
        $formPrincipal.Size = New-Object System.Drawing.Size(1000, 600)  # Aumentado de 720x400
        $formPrincipal.StartPosition = "CenterScreen"
        $formPrincipal.BackColor = [System.Drawing.Color]::White
        $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $formPrincipal.MaximizeBox = $false
        $formPrincipal.MinimizeBox = $false
        $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
        $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $formPrincipal.Text = "Daniel Tools $version"
        Write-Host "`n=============================================" -ForegroundColor DarkCyan
        Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
        Write-Host "              Versión: $($version)               " -ForegroundColor Green
        Write-Host "=============================================" -ForegroundColor DarkCyan
        Write-Host "`nTodos los derechos reservados para Daniel Tools." -ForegroundColor Cyan
        Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan
        Write-Host "O crea un issue en GitHub. https://github.com/water0ff/dztools/issues/new" -ForegroundColor Cyan
        $toolTip = New-Object System.Windows.Forms.ToolTip
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
        $lblServer = Create-Label -Text "Instancia SQL:" `
            -Location (New-Object System.Drawing.Point(10, 10)) `
            -Size (New-Object System.Drawing.Size(100, 10))
        $txtServer = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) `
            -Size (New-Object System.Drawing.Size(180, 20)) -DropDownStyle "DropDown"
        $txtServer.Text = ".\NationalSoft"
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
        $rtbQuery.Size = New-Object System.Drawing.Size(740, 140)     # Mayor ancho
        $rtbQuery.Multiline = $true
        $rtbQuery.ScrollBars = 'Vertical'
        $rtbQuery.WordWrap = $true
        $global:txtServer = $txtServer
        $global:txtUser = $txtUser
        $global:txtPassword = $txtPassword
        $global:cmbDatabases = $cmbDatabases
        $global:btnConnectDb = $btnConnectDb
        $global:btnDisconnectDb = $btnDisconnectDb
        $global:btnExecute = $btnExecute
        $global:btnBackup = $btnBackup
        $global:cmbQueries = $cmbQueries
        $global:rtbQuery = $rtbQuery
        $global:lblConnectionStatus = $lblConnectionStatus
        $script:SqlKeywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE'
        $script:SqlKeywordRegex = [System.Text.RegularExpressions.Regex]::new(
            "\b($script:SqlKeywords)\b",
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        $script:predefinedQueries = @{
            "Monitor de Servicios | Ventas a subir"           = @"
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
            "BackOffice Estaciones"                           = @"
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
            "SR | Actualizar contraseña de administrador"     = @"
    -- Actualiza la contraseña del primer usuario con rol administrador y retorna el usuario actualizado
    UPDATE usuarios
    SET contraseña = 'A9AE4E13D2A47998AC34'
    OUTPUT inserted.usuario
    WHERE usuario = (SELECT TOP 1 usuario FROM usuarios WHERE administrador = 1);
"@
            "SR | Revisar Pivot Table"                        = @"
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
            "SR | Fecha Revisiones"                           = @"
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
            "OTM | Eliminar Server en OTM"                    = @"
    SELECT serie, ipserver, nombreservidor
    FROM configuracion;
    -- UPDATE configuracion
    --   SET serie='', ipserver='', nombreservidor=''
"@
            "NSH | Eliminar Server en Hoteles"                = @"
    SELECT serievalida, numserie, ipserver, nombreservidor, llave
    FROM configuracion;
    -- UPDATE configuracion
    --   SET serievalida='', numserie='', ipserver='', nombreservidor='', llave=''
"@
            "Restcard | Eliminar Server en Rest Card"         = @"
    -- update tabvariables
    --   SET estacion='', ipservidor='';
"@
            "sql | Listar usuarios e idiomas"                 = @"
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
        $sortedKeys = $script:predefinedQueries.Keys | Sort-Object
        $cmbQueries.Items.Clear()
        foreach ($key in $sortedKeys) {
            $cmbQueries.Items.Add($key) | Out-Null
        }
        $cmbQueries.Add_SelectedIndexChanged({
                $rtbQuery.Text = $script:predefinedQueries[$cmbQueries.SelectedItem]
            })
        $rtbQuery.Add_TextChanged({
                $pos = $rtbQuery.SelectionStart
                $rtbQuery.SuspendLayout()
                $rtbQuery.SelectAll()
                $rtbQuery.SelectionColor = [System.Drawing.Color]::Black
                $commentRanges = @()
                foreach ($c in [regex]::Matches($rtbQuery.Text, '--.*', 'Multiline')) {
                    $rtbQuery.Select($c.Index, $c.Length)
                    $rtbQuery.SelectionColor = [System.Drawing.Color]::Green
                    $commentRanges += [PSCustomObject]@{ Start = $c.Index; End = $c.Index + $c.Length }
                }
                foreach ($b in [regex]::Matches($rtbQuery.Text, '/\*[\s\S]*?\*/', 'Multiline')) {
                    $rtbQuery.Select($b.Index, $b.Length)
                    $rtbQuery.SelectionColor = [System.Drawing.Color]::Green
                    $commentRanges += [PSCustomObject]@{ Start = $b.Index; End = $b.Index + $b.Length }
                }
                $matches = $script:SqlKeywordRegex.Matches($rtbQuery.Text)
                foreach ($m in $matches) {
                    $inComment = $commentRanges | Where-Object { $m.Index -ge $_.Start -and $m.Index -lt $_.End }
                    if (-not $inComment) {
                        $rtbQuery.Select($m.Index, $m.Length)
                        $rtbQuery.SelectionColor = [System.Drawing.Color]::Blue
                    }
                }
                $rtbQuery.Select($pos, 0)
                $rtbQuery.ResumeLayout()
            })
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
        $script:dgvResults = $dgvResults
        $script:originalForeColor = $dgvResults.DefaultCellStyle.ForeColor
        $script:originalHeaderBackColor = $dgvResults.ColumnHeadersDefaultCellStyle.BackColor
        $script:originalAutoSizeMode = $dgvResults.AutoSizeColumnsMode
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
        $btnExecute.Enabled = $false
        $cmbQueries.Enabled = $false
        #Todavía no lo migramos                $btnReloadConnections,  # <-- Agregar este
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
        $btnLectorDPicacls = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(10, 290)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
        $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Instalador LZMA" -Location (New-Object System.Drawing.Point(10, 330)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
        $btnConfigurarIPs = Create-Button -Text "Agregar IPs" -Location (New-Object System.Drawing.Point(10, 370)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
        $btnAddUser = Create-Button -Text "Agregar usuario de Windows" -Location (New-Object System.Drawing.Point(10, 410)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Crear nuevo usuario local en Windows"
        $btnForzarActualizacion = Create-Button -Text "Actualizar datos del sistema" -Location (New-Object System.Drawing.Point(10, 450)) `
            -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Actualiza información de hardware del sistema"
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
        $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(490, 50)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
        $btnCambiarOTM = Create-Button -Text "Cambiar OTM a SQL/DBF" -Location (New-Object System.Drawing.Point(490, 90)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Cambiar la configuración entre SQL y DBF para On The Minute."
        $btnCheckPermissions = Create-Button -Text "Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(490, 130)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
        $btnCreateAPK = Create-Button -Text "Creación de SRM APK" -Location (New-Object System.Drawing.Point(490, 170)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Generar archivo APK para Comandero Móvil"
        # Columna 4: | FIXES Y NOVEDADES |  X:730
        $txt_AdapterStatus = Create-TextBox -Location (New-Object System.Drawing.Point(730, 1)) -Size(New-Object System.Drawing.Size(220, 40)) `
            -BackColor([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) `
            -ScrollBars 'Vertical' -Multiline $true -ReadOnly  $true
        $global:txt_AdapterStatus = $txt_AdapterStatus
        $toolTip.SetToolTip($txt_AdapterStatus, "Lista de adaptadores y su estado. Haga clic en 'Actualizar adaptadores' para refrescar.")
        # Crear el nuevo formulario para los instaladores de Chocolatey
        $script:formInstaladoresChoco = Create-Form -Title "Instaladores Choco" -Size (New-Object System.Drawing.Size(520, 420)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
            -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 80, 80, 85))
        Write-DzDebug "`t[DEBUG] formInstaladoresChoco creado en sección principal."
        Write-DzDebug ("`t[DEBUG]   Es nulo?          : {0}" -f ($null -eq $script:formInstaladoresChoco))
        if ($null -ne $script:formInstaladoresChoco) {
            Write-DzDebug ("`t[DEBUG]   Tipo              : {0}" -f $script:formInstaladoresChoco.GetType().FullName)
        }
        $lblChocoSearch = Create-Label -Text "Buscar en Chocolatey:" -Location (New-Object System.Drawing.Point(10, 10)) `
            -ForeColor ([System.Drawing.Color]::White) -BackColor ([System.Drawing.Color]::Transparent) -Size (New-Object System.Drawing.Size(300, 20))
        $script:txtChocoSearch = Create-TextBox -Location (New-Object System.Drawing.Point(10, 30)) -Size (New-Object System.Drawing.Size(360, 30))
        $btnBuscarChoco = Create-Button -Text "Buscar" -Location (New-Object System.Drawing.Point(380, 28)) -Size (New-Object System.Drawing.Size(120, 32)) `
            -BackColor ([System.Drawing.Color]::FromArgb(76, 175, 80)) -ForeColor ([System.Drawing.Color]::White) -ToolTip "Buscar paquetes disponibles en Chocolatey."
        $script:btnBuscarChoco = $btnBuscarChoco
        $script:lvChocoResults = New-Object System.Windows.Forms.ListView
        $script:lvChocoResults.Location = New-Object System.Drawing.Point(10, 150)
        $script:lvChocoResults.Size = New-Object System.Drawing.Size(490, 200)
        $script:lvChocoResults.View = [System.Windows.Forms.View]::Details
        $script:lvChocoResults.FullRowSelect = $true
        $script:lvChocoResults.GridLines = $true
        $script:lvChocoResults.HideSelection = $false
        $script:lvChocoResults.BackColor = [System.Drawing.Color]::White
        $script:lvChocoResults.ForeColor = [System.Drawing.Color]::Black
        $null = $script:lvChocoResults.Columns.Add("Paquete", 170)
        $null = $script:lvChocoResults.Columns.Add("Versión", 100)
        $null = $script:lvChocoResults.Columns.Add("Descripción", 200)
        $script:lblPresetSSMS = Create-Label -Text "SSMS" -Location (New-Object System.Drawing.Point(10, 70)) `
            -ForeColor ([System.Drawing.Color]::Black) -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 255)) `
            -Size (New-Object System.Drawing.Size(70, 25)) -TextAlign MiddleCenter -BorderStyle FixedSingle
        $script:lblPresetSSMS.Cursor = [System.Windows.Forms.Cursors]::Hand
        $script:lblPresetHeidi = Create-Label -Text "Heidi" -Location (New-Object System.Drawing.Point(90, 70)) `
            -ForeColor ([System.Drawing.Color]::Black) -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 255)) `
            -Size (New-Object System.Drawing.Size(70, 25)) -TextAlign MiddleCenter -BorderStyle FixedSingle
        $script:lblPresetHeidi.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnInstallSelectedChoco = Create-Button -Text "Instalar seleccionado" -Location (New-Object System.Drawing.Point(170, 100)) `
            -Size (New-Object System.Drawing.Size(170, 32)) -ToolTip "Instala el paquete seleccionado de la lista."
        $btnShowInstalledChoco = Create-Button -Text "Mostrar instalados" -Location (New-Object System.Drawing.Point(10, 100)) `
            -Size (New-Object System.Drawing.Size(150, 32)) -ToolTip "Muestra los paquetes instalados con Chocolatey."
        $btnUninstallSelectedChoco = Create-Button -Text "Desinstalar seleccionado" -Location (New-Object System.Drawing.Point(350, 100)) `
            -Size (New-Object System.Drawing.Size(150, 32)) -ToolTip "Desinstala el paquete seleccionado de la lista."
        $btnExitInstaladores = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(10, 365)) `
            -ToolTip "Salir del formulario de instaladores."
        $script:btnInstallSelectedChoco = $btnInstallSelectedChoco
        $script:btnShowInstalledChoco = $btnShowInstalledChoco
        $script:btnUninstallSelectedChoco = $btnUninstallSelectedChoco
        # Agregar los botones al nuevo formulario
        $script:formInstaladoresChoco.Controls.Add($lblChocoSearch)
        $script:formInstaladoresChoco.Controls.Add($txtChocoSearch)
        $script:formInstaladoresChoco.Controls.Add($btnBuscarChoco)
        $script:formInstaladoresChoco.Controls.Add($lblPresetSSMS)
        $script:formInstaladoresChoco.Controls.Add($lblPresetHeidi)
        $script:formInstaladoresChoco.Controls.Add($btnInstallSelectedChoco)
        $script:formInstaladoresChoco.Controls.Add($btnShowInstalledChoco)
        $script:formInstaladoresChoco.Controls.Add($btnUninstallSelectedChoco)
        $script:formInstaladoresChoco.Controls.Add($btnExitInstaladores)
        $script:formInstaladoresChoco.Controls.Add($lvChocoResults)
        $script:chocoPackagePattern = '^(?<name>[A-Za-z0-9\.\+\-_]+)\s+(?<version>[0-9][A-Za-z0-9\.\-]*)\s+(?<description>.+)$'
        $script:addChocoResult = {
            param($line)
            if ([string]::IsNullOrWhiteSpace($line)) { return }
            if ($line -match '^Chocolatey') { return }
            if ($line -match 'packages?\s+found' -or $line -match 'page size') { return }
            if ($line -match $script:chocoPackagePattern) {
                $name = $Matches['name']
                $version = $Matches['version']
                $description = $Matches['description'].Trim()
                $item = New-Object System.Windows.Forms.ListViewItem($name)
                $null = $item.SubItems.Add($version)
                $null = $item.SubItems.Add($description)
                $null = $script:lvChocoResults.Items.Add($item)
            } elseif ($line -match '^(?<name>[A-Za-z0-9\.\+\-_]+)\|(?<version>[0-9][A-Za-z0-9\.\-]*)$') {
                $name = $Matches['name']
                $version = $Matches['version']
                $item = New-Object System.Windows.Forms.ListViewItem($name)
                $null = $item.SubItems.Add($version)
                $null = $item.SubItems.Add("Paquete instalado")
                $null = $script:lvChocoResults.Items.Add($item)
            }
        }
        $script:updateChocoActionButtons = {
            $hasValidSelection = $false
            if ($script:lvChocoResults.SelectedItems.Count -gt 0) {
                $selectedItem = $script:lvChocoResults.SelectedItems[0]
                $packageName = $selectedItem.Text
                $packageVersion = if ($selectedItem.SubItems.Count -gt 1) { $selectedItem.SubItems[1].Text } else { "" }
                if (-not [string]::IsNullOrWhiteSpace($packageName) -and $packageVersion -match '^[0-9]') {
                    $hasValidSelection = $true
                }
            }
            $script:btnInstallSelectedChoco.Enabled = $hasValidSelection
            $script:btnUninstallSelectedChoco.Enabled = $hasValidSelection
        }
        $script:btnInstallSelectedChoco.Enabled = $false
        $script:btnUninstallSelectedChoco.Enabled = $false
        $txt_InfoInstrucciones = Create-TextBox `
            -Location (New-Object System.Drawing.Point(730, 50)) `
            -Size     (New-Object System.Drawing.Size(220, 500)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 1, 36, 86)) `
            -ForeColor ([System.Drawing.Color]::White) `
            -Font      (New-Object System.Drawing.Font("Courier New", 10)) `
            -Multiline $true `
            -ReadOnly  $true
        $txt_InfoInstrucciones.WordWrap = $true
        $txt_InfoInstrucciones.Text = $global:defaultInstructions
        $script:setInstructionText = {
            param(
                [string]$Message
            )

            if ($null -ne $txt_InfoInstrucciones -and
                $txt_InfoInstrucciones.PSObject.Properties.Match('Text').Count -gt 0) {
                $txt_InfoInstrucciones.Text = $Message
            }
        }.GetNewClosure()
        $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(350, 525)) `
            -Size (New-Object System.Drawing.Size(500, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
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
                $txt_InfoInstrucciones,
                $btnCreateAPK
            ))
        Refresh-AdapterStatus
        $formPrincipal.Controls.Add($tabControl)
        $formPrincipal.Controls.Add($btnExit)
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
                    IPAddress   = $_.Address.ToString()
                }
            }
        }
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
        # Obtener el puerto de SQL Server desde el registro
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\NATIONALSOFT\MSSQLServer\SuperSocketNetLib\Tcp"
        $tcpPort = Get-ItemProperty -Path $regKeyPath -Name "TcpPort" -ErrorAction SilentlyContinue
        if ($tcpPort -and $tcpPort.TcpPort) {
            $lblPort.Text = "Puerto SQL \NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $lblPort.Text = "No se encontró puerto o instancia."
        }

        #Acciones en los botones############################################################:
        $changeColorOnHover = {
            param($sender, $eventArgs)
            $sender.BackColor = [System.Drawing.Color]::Orange
        }
        $restoreColorOnLeave = {
            param($sender, $eventArgs)
            $sender.BackColor = [System.Drawing.Color]::Black
        }
        $lblHostname.Add_MouseEnter($changeColorOnHover)
        $lblHostname.Add_MouseLeave($restoreColorOnLeave)
        $buttonsToUpdate = @(
            $LZMAbtnBuscarCarpeta, $btnInstalarHerramientas, $btnProfiler,
            $btnDatabase, $btnSQLManager, $btnSQLManagement, $btnPrinterTool,
            $btnLectorDPicacls, $btnConfigurarIPs, $btnAddUser, $btnForzarActualizacion,
            $btnClearAnyDesk, $btnShowPrinters, $btnClearPrintJobs, $btnAplicacionesNS,
            $btnCheckPermissions, $btnCambiarOTM, $btnCreateAPK
        )
        foreach ($button in $buttonsToUpdate) {
            $button.Add_MouseLeave({
                    if ($script:setInstructionText) {
                        $script:setInstructionText.Invoke($global:defaultInstructions)
                    }
                })
        }
        $LZMAbtnBuscarCarpeta.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
Busca en los registros de Windows el histórico de instalaciones que han fallado,
permitiendo renombrar la carpeta correspondiente para que el instalador genere
un nuevo registro y así evite el mensaje de error conocido:

    Error al crear el archivo en temporales
"@)
                }
            })
        $lblPresetSSMS.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Acceso rápido para buscar SSMS en Chocolatey.
Al hacer clic llenará la búsqueda con 'SSMS'
e iniciará la consulta automáticamente.
"@)
                }
            })
        $lblPresetHeidi.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Acceso rápido para buscar HeidiSQL en Chocolatey.
Al hacer clic llenará la búsqueda con 'Heidi'
e iniciará la consulta automáticamente.
"@)
                }
            })
        $btnInstallSelectedChoco.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Instala el paquete seleccionado de la lista de resultados.
Solicitará confirmación mostrando el paquete,
su versión y descripción.
"@)
                }
            })
        $btnInstalarHerramientas.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Abre el menú de instaladores de Chocolatey para instalar o actualizar
herramientas de línea de comandos y utilerías en el sistema.
"@)
                }
            })
        $btnProfiler.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Ejecuta o descarga ExpressProfiler desde el servidor oficial,
herramienta para monitorear consultas de SQL Server.
"@)
                }
            })
        $btnDatabase.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Ejecuta Database4: si no está instalado, lo descarga automáticamente
y luego lo lanza para la gestión de sus bases de datos.
"@)
                }
            })
        $btnSQLManager.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Ejecuta SQL Server Management Studio (para SQL 2014). Si no lo encuentra,
avisará al usuario dónde descargarlo desde el repositorio oficial.
"@)
                }
            })
        $btnSQLManagement.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Busca SQL Management en el equipo, recupera la versión instalada
y la muestra antes de ejecutarlo.
"@)
                }
            })
        $btnPrinterTool.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Herramienta de Star Micronics para configurar y diagnosticar impresoras POS:
permite probar estado, formatear y configurar parámetros fundamentales.
"@)
                }
            })
        $btnLectorDPicacls.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Repara el error al instalar el Driver del lector DP.
Modifica los permisos de la carpeta C:\Windows\System32\en-us
mediante el comando ICALCS para el driver tenga los permisos necesarios.
"@)
                }
            })
        $btnConfigurarIPs.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Agrega direcciones IP adicionales para configurar impresoras en red
que estén en un segmento diferente al predeterminado.
Convierte de DHCP a ip fija y tambien permite cambiar la configuración de ip fija a DHCP.
"@)
                }
            })
        $btnAddUser.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Crea un nuevo usuario local en Windows con permisos básicos:
útil para sesión independiente en estaciones o terminales.
"@)
                }
            })
        $btnForzarActualizacion.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Para el error de descarga de licencia por no tener datos de equipo como el procesador.
Actualiza la información de hardware del sistema:
reescanea unidades, adaptadores y muestra un resumen de dispositivos.
"@)
                }
            })
        $btnClearAnyDesk.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Detiene el servicio de AnyDesk, elimina los archivos temporales
y forja un nuevo ID para evitar conflictos de acceso remoto.
"@)
                }
            })
        $btnShowPrinters.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Muestra en consola las impresoras instaladas en Windows,
junto con su puerto y driver correspondiente.
"@)
                }
            })
        $btnClearPrintJobs.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Limpia la cola de impresión y reinicia el servicio de spooler
para liberar trabajos atascados.
"@)
                }
            })
        $btnAplicacionesNS.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Busca los archivos INI de National Soft en el equipo
y extrae la información de conexión a bases de datos.
"@)
                }
            })
        $btnCheckPermissions.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Revisa los permisos de la carpeta C:\NationalSoft
y muestra qué usuarios tienen acceso de lectura/escritura.
* Permite asignar permisos heredados a Everyone a dicha carpeta.
"@)
                }
            })
        $btnCambiarOTM.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Cambia la configuración de On The Minute (OTM)
entre SQL Server y DBF según corresponda.
"@)
                }
            })
        $btnCreateAPK.Add_MouseEnter({
                if ($script:setInstructionText) {
                    $script:setInstructionText.Invoke(@"
        Genera el archivo APK para Comandero Móvil:
compila el proyecto y lo coloca en la carpeta de salida.
"@)
                }
            })
        $lblPort.Add_Click({
                param($sender, $e)
                $text = $sender.Text
                Write-DzDebug "`t[DEBUG] lblPort.Text al hacer clic: '$text'"
                $port = [regex]::Match($text, '\d+').Value
                if ([string]::IsNullOrWhiteSpace($port)) {
                    Write-Host "El texto del Label del puerto no contiene un número válido para copiar." -ForegroundColor Red
                    return
                }
                [System.Windows.Forms.Clipboard]::SetText($port)
                Write-Host "Puerto copiado al portapapeles: $port" -ForegroundColor Green
            })
        $lblPort.Add_MouseEnter($changeColorOnHover)
        $lblPort.Add_MouseLeave($restoreColorOnLeave)
        $lblHostname.Add_Click({
                param($sender, $e)
                $hostnameText = $sender.Text  # <- usamos el control que disparó el evento
                if ([string]::IsNullOrWhiteSpace($hostnameText)) {
                    Write-Host "`n[AVISO] El hostname está vacío o nulo, no se copió nada." -ForegroundColor Yellow
                    [System.Windows.Forms.MessageBox]::Show(
                        "El nombre de equipo está vacío, no hay nada que copiar.",
                        "Daniel Tools",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    return
                }
                [System.Windows.Forms.Clipboard]::SetText($hostnameText)
                Write-Host "`nNombre del equipo copiado al portapapeles: $hostnameText" -ForegroundColor Green
            })
        $txt_IpAdress.Add_Click({
                param($sender, $e)
                $ipsText = $sender.Text  # <- usamos el Text del textbox que hizo click
                if ([string]::IsNullOrWhiteSpace($ipsText)) {
                    Write-Host "`n[AVISO] No hay IPs para copiar." -ForegroundColor Yellow
                    return
                }
                [System.Windows.Forms.Clipboard]::SetText($ipsText)
                Write-Host "`nIP's copiadas al portapapeles: $ipsText" -ForegroundColor Green
            })
        $txt_IpAdress.Add_MouseEnter($changeColorOnHover)
        $txt_IpAdress.Add_MouseLeave($restoreColorOnLeave)
        $txt_AdapterStatus.Add_Click({
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
        $btnCheckPermissions.Add_Click({
                Write-Host "`nRevisando permisos en C:\NationalSoft" -ForegroundColor Yellow
                if (-not (Test-Administrator)) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Esta acción requiere permisos de administrador.`r`n" +
                        "Por favor, ejecuta Daniel Tools como administrador.",
                        "Permisos insuficientes",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                    return
                }

                Check-Permissions
            })
        $copyMenu.Add_Click({
                if ($dgvResults.GetCellCount("Selected") -gt 0) {
                    $dataObj = $dgvResults.GetClipboardContent()
                    if ($dataObj) {
                        [Windows.Forms.Clipboard]::SetText($dataObj.GetText())
                    }
                }
            })
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
        $dgvResults.Add_MouseDown({
                param($sender, $args)
                if ($args.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                    $hit = $dgvResults.HitTest($args.X, $args.Y)
                    if ($hit.RowIndex -ge 0 -and $hit.ColumnIndex -ge 0) {
                        $dgvResults.CurrentCell = $dgvResults.Rows[$hit.RowIndex].Cells[$hit.ColumnIndex]
                        if (-not $args.Modifiers.HasFlag([System.Windows.Forms.Keys]::Control)) {
                            $dgvResults.ClearSelection()
                        }
                        $dgvResults.Rows[$hit.RowIndex].Cells[$hit.ColumnIndex].Selected = $true
                    }
                }
            })
        $btnExitInstaladores.Add_Click({
                $script:formInstaladoresChoco.Close()
            })
        $btnInstalarHerramientas.Add_Click({
                Write-Host ""
                Write-DzDebug ("`t[DEBUG] Click en 'Instalar Herramientas' - {0}" -f (Get-Date -Format "HH:mm:ss"))
                Write-Host "`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                if (-not (Check-Chocolatey)) {
                    Write-Host "Chocolatey no está instalado. No se puede abrir el menú de instaladores." -ForegroundColor Red
                    return
                }
                Write-DzDebug ("`t[DEBUG] `\$script:formInstaladoresChoco es nulo? : {0}" -f ($null -eq $script:formInstaladoresChoco))
                if ($null -eq $script:formInstaladoresChoco) {
                    Write-DzDebug "`t[DEBUG] ERROR: `\$script:formInstaladoresChoco es `$null dentro del manejador de clic."
                    return
                }
                $script:formInstaladoresChoco.ShowDialog()
            })
        $script:lvChocoResults.Add_SelectedIndexChanged({
                $script:updateChocoActionButtons.Invoke()
            })
        $btnBuscarChoco.Add_Click({
                $script:lvChocoResults.Items.Clear()
                $script:updateChocoActionButtons.Invoke()
                $query = $script:txtChocoSearch.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($query)) {
                    [System.Windows.Forms.MessageBox]::Show("Ingresa un término para buscar paquetes en Chocolatey.", "Búsqueda de paquetes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    return
                }
                if (-not (Check-Chocolatey)) {
                    return
                }
                $script:btnBuscarChoco.Enabled = $false
                $progressForm = Show-ProgressBar
                $totalSteps = 3
                $currentStep = 1
                Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps -Status "Ejecutando búsqueda..."
                Write-Host ("`tBuscando paquetes para '{0}'..." -f $query) -ForegroundColor Cyan
                Write-DzDebug ("`t[DEBUG] Búsqueda de Chocolatey: término='{0}'" -f $query)
                $searchExitCode = $null
                try {
                    Write-DzDebug "`t[DEBUG] Ejecutando: choco search <query> --page-size=20"
                    $searchOutput = & choco search $query --page-size=20 2>&1
                    $searchExitCode = $LASTEXITCODE
                    Write-DzDebug ("`t[DEBUG] choco search exit code: {0}" -f $searchExitCode)
                    if ($searchOutput) {
                        Write-DzDebug "`t[DEBUG] Salida cruda de choco search:"
                        foreach ($line in $searchOutput) {
                            Write-DzDebug ("`t       > {0}" -f $line)
                        }
                    } else {
                        Write-DzDebug "`t[DEBUG] choco search no devolvió salida."
                    }
                    $currentStep++
                    Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps -Status "Procesando resultados..."
                    foreach ($line in $searchOutput) {
                        $script:addChocoResult.Invoke($line)
                    }
                    if ($script:lvChocoResults.Items.Count -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show("No se encontraron paquetes para la búsqueda realizada.", "Sin resultados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-DzDebug "`t[DEBUG] Búsqueda sin resultados."
                    } else {
                        Write-DzDebug ("`t[DEBUG] Resultados agregados: {0}" -f $script:lvChocoResults.Items.Count)
                    }
                    $currentStep = $totalSteps
                    Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps -Status "Búsqueda finalizada"
                } catch {
                    Write-Error "Error al consultar paquetes de Chocolatey: $_"
                    Write-DzDebug ("`t[DEBUG] Excepción en búsqueda de Chocolatey: {0}" -f $_)
                    Write-DzDebug ("`t[DEBUG] Stack: {0}" -f $_.ScriptStackTrace)
                    if ($searchExitCode -ne $null) {
                        Write-DzDebug ("`t[DEBUG] choco search finalizó con código {0}" -f $searchExitCode)
                    }
                    [System.Windows.Forms.MessageBox]::Show("Ocurrió un error al buscar en Chocolatey. Intenta nuevamente.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                } finally {
                    if ($null -ne $progressForm) {
                        Close-ProgressBar $progressForm
                    }
                    $script:btnBuscarChoco.Enabled = $true
                    $script:updateChocoActionButtons.Invoke()
                }
            })
        $lblPresetSSMS.Add_Click({
                $preset = "ssms"
                Write-DzDebug ("`t[DEBUG] Preset seleccionado: {0}" -f $preset)
                $script:txtChocoSearch.Text = $preset
                $script:btnBuscarChoco.PerformClick()
            })
        $lblPresetHeidi.Add_Click({
                $preset = "heidi"
                Write-DzDebug ("`t[DEBUG] Preset seleccionado: {0}" -f $preset)
                $script:txtChocoSearch.Text = $preset
                $script:btnBuscarChoco.PerformClick()
            })
        $btnShowInstalledChoco.Add_Click({
                $script:lvChocoResults.Items.Clear()
                $script:updateChocoActionButtons.Invoke()
                if (-not (Check-Chocolatey)) {
                    return
                }
                $script:btnShowInstalledChoco.Enabled = $false
                $progressForm = Show-ProgressBar
                $totalSteps = 2
                $currentStep = 1
                Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps -Status "Recuperando instalados..."
                try {
                    Write-DzDebug "`t[DEBUG] Ejecutando: choco list --local-only --limit-output"
                    $installedOutput = & choco list --local-only --limit-output 2>&1
                    $listExitCode = $LASTEXITCODE
                    Write-DzDebug ("`t[DEBUG] choco list exit code: {0}" -f $listExitCode)
                    foreach ($line in $installedOutput) {
                        Write-DzDebug ("`t       > {0}" -f $line)
                    }
                    $currentStep++
                    Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps -Status "Procesando resultados..."
                    foreach ($line in $installedOutput) {
                        $script:addChocoResult.Invoke($line)
                    }
                    if ($script:lvChocoResults.Items.Count -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show("No se encontraron paquetes instalados con Chocolatey.", "Sin resultados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                } catch {
                    Write-Error "Error al consultar paquetes instalados de Chocolatey: $_"
                    [System.Windows.Forms.MessageBox]::Show("Ocurrió un error al consultar paquetes instalados.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                } finally {
                    if ($null -ne $progressForm) {
                        Close-ProgressBar $progressForm
                    }
                    $script:btnShowInstalledChoco.Enabled = $true
                    $script:updateChocoActionButtons.Invoke()
                }
            })
        $btnInstallSelectedChoco.Add_Click({
                try {
                    Write-DzDebug "`n========== INICIO INSTALACIÓN =========="
                    Write-DzDebug "`t[DEBUG] Click en 'Instalar seleccionado'"
                    Write-DzDebug ("`t[DEBUG] Thread ID: {0}" -f [System.Threading.Thread]::CurrentThread.ManagedThreadId)
                    Write-DzDebug ("`t[DEBUG] Items seleccionados: {0}" -f $script:lvChocoResults.SelectedItems.Count)

                    if ($script:lvChocoResults.SelectedItems.Count -eq 0) {
                        Write-DzDebug "`t[DEBUG] Ningún paquete seleccionado"
                        [System.Windows.Forms.MessageBox]::Show(
                            "Seleccione un paquete de la lista antes de instalar.",
                            "Instalación de paquete",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        ) | Out-Null
                        return
                    }

                    $selectedItem = $script:lvChocoResults.SelectedItems[0]
                    $packageName = $selectedItem.Text
                    $packageVersion = if ($selectedItem.SubItems.Count -gt 1) { $selectedItem.SubItems[1].Text } else { "" }

                    Write-DzDebug ("`t[DEBUG] Paquete: {0}, Versión: {1}" -f $packageName, $packageVersion)

                    $confirmationText = "Vas a instalar el paquete: $packageName"
                    if (-not [string]::IsNullOrWhiteSpace($packageVersion)) {
                        $confirmationText += " (versión $packageVersion)"
                    }

                    $response = [System.Windows.Forms.MessageBox]::Show(
                        $confirmationText,
                        "Confirmar instalación",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )

                    if ($response -ne [System.Windows.Forms.DialogResult]::Yes) {
                        Write-DzDebug "`t[DEBUG] Instalación cancelada por usuario"
                        return
                    }

                    Write-DzDebug "`t[DEBUG] Verificando Chocolatey..."
                    if (-not (Check-Chocolatey)) {
                        Write-DzDebug "`t[DEBUG] Check-Chocolatey retornó FALSE"
                        return
                    }

                    Write-DzDebug "`t[DEBUG] Chocolatey OK, preparando instalación..."

                    $arguments = "install $packageName -y"
                    if (-not [string]::IsNullOrWhiteSpace($packageVersion)) {
                        $arguments += " --version=$packageVersion"
                    }

                    Write-DzDebug ("`t[DEBUG] Argumentos: {0}" -f $arguments)
                    Write-DzDebug "`t[DEBUG] Llamando a Invoke-ChocoCommandWithProgress..."

                    $exitCode = Invoke-ChocoCommandWithProgress -Arguments $arguments -OperationTitle "Instalando $packageName"

                    Write-DzDebug ("`t[DEBUG] Retornó código: {0}" -f $exitCode)
                    Write-DzDebug "========== FIN INSTALACIÓN ==========`n"

                    if ($exitCode -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Instalación completada para $packageName.",
                            "Éxito",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        ) | Out-Null
                    } else {
                        [System.Windows.Forms.MessageBox]::Show(
                            "La instalación terminó con código $exitCode para $packageName.",
                            "Aviso",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        ) | Out-Null
                    }

                } catch {
                    Write-DzDebug "`n========== ERROR EN INSTALACIÓN ==========" -Color Red
                    Write-DzDebug ("`t[ERROR] Mensaje: {0}" -f $_.Exception.Message) -Color Red
                    Write-DzDebug ("`t[ERROR] Tipo: {0}" -f $_.Exception.GetType().FullName) -Color Red

                    if ($_.InvocationInfo) {
                        Write-DzDebug ("`t[ERROR] En línea: {0}" -f $_.InvocationInfo.ScriptLineNumber) -Color DarkYellow
                        Write-DzDebug ("`t[ERROR] Comando: {0}" -f $_.InvocationInfo.Line) -Color DarkYellow
                    }

                    if ($_.ScriptStackTrace) {
                        Write-DzDebug "`t[ERROR] Stack trace:" -Color DarkGray
                        Write-DzDebug $_.ScriptStackTrace -Color DarkGray
                    }

                    Write-DzDebug "========== FIN ERROR ==========`n" -Color Red

                    [System.Windows.Forms.MessageBox]::Show(
                        "Error al instalar:`n`n$($_.Exception.Message)`n`nRevise la consola para más detalles.",
                        "Error en instalación",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
        $btnUninstallSelectedChoco.Add_Click({
                try {
                    Write-DzDebug "`t[DEBUG] Click en 'Desinstalar seleccionado' (handler completo)"
                    if ($script:lvChocoResults.SelectedItems.Count -eq 0) {
                        Write-DzDebug "`t[DEBUG] Ningún paquete seleccionado para desinstalar."
                        [System.Windows.Forms.MessageBox]::Show(
                            "Seleccione un paquete de la lista antes de desinstalar.",
                            "Desinstalación de paquete",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        ) | Out-Null
                        return
                    }
                    $selectedItem = $script:lvChocoResults.SelectedItems[0]
                    $packageName = $selectedItem.Text
                    $packageVersion = if ($selectedItem.SubItems.Count -gt 1) { $selectedItem.SubItems[1].Text } else { "" }
                    $packageDescription = if ($selectedItem.SubItems.Count -gt 2) { $selectedItem.SubItems[2].Text } else { "" }
                    $confirmationText = "¿Deseas desinstalar el paquete: $packageName $packageVersion $packageDescription?"
                    $response = [System.Windows.Forms.MessageBox]::Show(
                        $confirmationText,
                        "Confirmar desinstalación",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )
                    if ($response -ne [System.Windows.Forms.DialogResult]::Yes) {
                        Write-DzDebug "`t[DEBUG] Desinstalación cancelada por el usuario."
                        return
                    }
                    if (-not (Check-Chocolatey)) {
                        Write-DzDebug "`t[DEBUG] Chocolatey no está instalado; se cancela la desinstalación."
                        return
                    }
                    $arguments = "uninstall $packageName -y"
                    if (-not [string]::IsNullOrWhiteSpace($packageVersion)) {
                        $arguments = "uninstall $packageName --version=$packageVersion -y"
                    }
                    Write-DzDebug ("`t[DEBUG] Ejecutando desinstalación con argumentos: {0}" -f $arguments)
                    $exitCode = Invoke-ChocoCommandWithProgress -Arguments $arguments -OperationTitle "Desinstalando $packageName"
                    Write-DzDebug ("`t[DEBUG] Invoke-ChocoCommandWithProgress devolvió código {0}" -f $exitCode)
                    if ($exitCode -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Desinstalación completada para $packageName.",
                            "Éxito",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Information
                        ) | Out-Null
                    } else {
                        [System.Windows.Forms.MessageBox]::Show(
                            "La desinstalación terminó con código $exitCode para $packageName.",
                            "Aviso",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        ) | Out-Null
                    }
                } catch {
                    Write-DzDebug ("`t[DEBUG] ERROR en handler 'Desinstalar seleccionado': {0}" -f $_) -Color DarkRed
                    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
                        Write-DzDebug ("`t[DEBUG] Línea: {0}" -f $_.InvocationInfo.PositionMessage) -Color DarkYellow
                    }
                    if ($_.ScriptStackTrace) {
                        Write-DzDebug ("`t[DEBUG] Stack: {0}" -f $_.ScriptStackTrace) -Color DarkGray
                    }
                    [System.Windows.Forms.MessageBox]::Show(
                        "Ocurrió un error al iniciar la desinstalación: $($_.Exception.Message)",
                        "Error en botón Desinstalar",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                }
            })
        $btnForzarActualizacion.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                Show-SystemComponents
                [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
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
                } else {
                    Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
                }
            })
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
                    } elseif ($path -match 'SQL Server Management Studio (\d+)') {
                        return "SSMS $($matches[1])"
                    } else {
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
        $btnProfiler.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $ProfilerUrl = "https://codeplexarchive.org/codeplex/browse/ExpressProfiler/releases/4/ExpressProfiler22wAddinSigned.zip"
                $ProfilerZipPath = "C:\Temp\ExpressProfiler22wAddinSigned.zip"
                $ExtractPath = "C:\Temp\ExpressProfiler2"
                $ExeName = "ExpressProfiler.exe"
                $ValidationPath = "C:\Temp\ExpressProfiler2\ExpressProfiler.exe"

                DownloadAndRun -url $ProfilerUrl -zipPath $ProfilerZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
                if ($disableControls) { Enable-Controls -parentControl $formPrincipal }
            }
        )
        $btnPrinterTool.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $PrinterToolUrl = "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip"
                $PrinterToolZipPath = "C:\Temp\RPT-RPI-Printer-Tool-1.zip"
                $ExtractPath = "C:\Temp\RPT-RPI-Printer-Tool-1"
                $ExeName = "POS Printer Test.exe"
                $ValidationPath = "C:\Temp\RPT-RPI-Printer-Tool-1\POS Printer Test.exe"

                DownloadAndRun -url $PrinterToolUrl -zipPath $PrinterToolZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
            })
        $btnDatabase.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $DatabaseUrl = "https://fishcodelib.com/files/DatabaseNet4.zip"
                $DatabaseZipPath = "C:\Temp\DatabaseNet4.zip"
                $ExtractPath = "C:\Temp\Database4"
                $ExeName = "Database4.exe"
                $ValidationPath = "C:\Temp\Database4\Database4.exe"

                DownloadAndRun -url $DatabaseUrl -zipPath $DatabaseZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
            })
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
        $btnExecute.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                try {
                    if ($null -eq $script:dgvResults) {
                        Write-DzDebug "`t[DEBUG] dgvResults es NULL dentro del Click"
                        throw "DataGridView no inicializado."
                    } else {
                        Write-DzDebug ("`t[DEBUG] dgvResults tipo: {0}" -f ($script:dgvResults.GetType().FullName))
                    }
                    if ($script:originalForeColor) {
                        $script:dgvResults.DefaultCellStyle.ForeColor = $script:originalForeColor
                    }
                    if ($script:originalHeaderBackColor) {
                        $script:dgvResults.ColumnHeadersDefaultCellStyle.BackColor = $script:originalHeaderBackColor
                    }
                    if ($script:originalAutoSizeMode) {
                        $script:dgvResults.AutoSizeColumnsMode = $script:originalAutoSizeMode
                    }
                    if ($script:dgvResults.DataSource) {
                        $script:dgvResults.DataSource = $null
                    }
                    if ($script:dgvResults.Rows -ne $null) {
                        $script:dgvResults.Rows.Clear()
                    }
                    if ($null -ne $toolTip -and ($toolTip | Get-Member -Name 'SetToolTip' -MemberType Method -ErrorAction SilentlyContinue)) {
                        $toolTip.SetToolTip($script:dgvResults, $null)
                    }
                    $selectedDb = $cmbDatabases.SelectedItem
                    if (-not $selectedDb) { throw "Selecciona una base de datos" }
                    $rawQuery = $rtbQuery.Text
                    $cleanQuery = Remove-SqlComments -Query $rawQuery
                    $result = Execute-SqlQuery -server $global:server -database $selectedDb -query $cleanQuery
                    if ($result -and $result.ContainsKey('Messages') -and $result.Messages) {
                        if ($result.Messages.Count -gt 0) {
                            Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
                            $result.Messages | ForEach-Object { Write-Host $_ }
                        }
                    }
                    if ($result -and $result.ContainsKey('DataTable') -and $result.DataTable) {
                        $script:dgvResults.DataSource = $result.DataTable.DefaultView
                        $script:dgvResults.Enabled = $true
                        Write-Host "`nColumnas obtenidas: $($result.DataTable.Columns.ColumnName -join ', ')" -ForegroundColor Cyan
                        $script:dgvResults.DefaultCellStyle.ForeColor = 'Blue'
                        $script:dgvResults.AlternatingRowsDefaultCellStyle.BackColor = '#F0F8FF'
                        $script:dgvResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
                        foreach ($col in $script:dgvResults.Columns) {
                            $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::DisplayedCells
                            $col.Width = [Math]::Max($col.Width, $col.HeaderText.Length * 8)
                        }
                        if ($result.DataTable.Rows.Count -eq 0) {
                            Write-Host "La consulta no devolvió resultados" -ForegroundColor Yellow
                        } else {
                            $result.DataTable | Format-Table -AutoSize | Out-String | Write-Host
                        }
                    } elseif ($result -and $result.ContainsKey('RowsAffected')) {
                        Write-Host "`nFilas afectadas: $($result.RowsAffected)" -ForegroundColor Green
                        $rowsAffectedTable = New-Object System.Data.DataTable
                        $rowsAffectedTable.Columns.Add("Resultado") | Out-Null
                        $rowsAffectedTable.Rows.Add("Filas afectadas: $($result.RowsAffected)") | Out-Null
                        $script:dgvResults.DataSource = $rowsAffectedTable
                        $script:dgvResults.Enabled = $true
                        $script:dgvResults.DefaultCellStyle.ForeColor = 'Green'
                        $script:dgvResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
                    } else {
                        Write-Host "`nNo se recibió DataTable ni RowsAffected en el resultado." -ForegroundColor Yellow
                    }
                } catch {
                    $errorTable = New-Object System.Data.DataTable
                    $errorTable.Columns.Add("Tipo")    | Out-Null
                    $errorTable.Columns.Add("Mensaje") | Out-Null
                    $errorTable.Columns.Add("Detalle") | Out-Null
                    $cleanQuery = $rtbQuery.Text -replace '(?s)/\*.*?\*/', '' -replace '(?m)^\s*--.*'
                    $shortQuery = if ($cleanQuery.Length -gt 50) { $cleanQuery.Substring(0, 47) + "..." } else { $cleanQuery }
                    $errorTable.Rows.Add("ERROR SQL", $_.Exception.Message, $shortQuery) | Out-Null
                    if ($null -ne $script:dgvResults) {
                        $script:dgvResults.DataSource = $errorTable
                        if ($script:dgvResults.Columns.Count -ge 3) {
                            $script:dgvResults.Columns[1].DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
                            $script:dgvResults.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
                            $script:dgvResults.AutoSizeColumnsMode = 'Fill'
                            $script:dgvResults.Columns[0].Width = 100
                            $script:dgvResults.Columns[1].Width = 300
                            $script:dgvResults.Columns[2].Width = 200
                        }
                        $script:dgvResults.DefaultCellStyle.ForeColor = 'Red'
                        $script:dgvResults.ColumnHeadersDefaultCellStyle.BackColor = '#FFB3B3'
                        if ($null -ne $toolTip -and ($toolTip | Get-Member -Name 'SetToolTip' -MemberType Method -ErrorAction SilentlyContinue)) {
                            $toolTip.SetToolTip($script:dgvResults, "Consulta completa:`n$cleanQuery")
                        }
                    }
                    Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
                    Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Host "Consulta: $shortQuery" -ForegroundColor Cyan
                    Write-Host "====================================" -ForegroundColor Red
                }
            })
        $btnClearAnyDesk.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $confirmationResult = [System.Windows.Forms.MessageBox]::Show(
                    "¿Estás seguro de renovar AnyDesk?",
                    "Confirmar Renovación",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($confirmationResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $filesToDelete = @(
                        "C:\ProgramData\AnyDesk\system.conf",
                        "C:\ProgramData\AnyDesk\service.conf",
                        "$env:APPDATA\AnyDesk\system.conf",
                        "$env:APPDATA\AnyDesk\service.conf"
                    )
                    $deletedFilesCount = 0
                    $errors = @()
                    try {
                        Write-Host "`tCerrando el proceso AnyDesk..." -ForegroundColor Yellow
                        Stop-Process -Name "AnyDesk" -Force -ErrorAction Stop
                        Write-Host "`tAnyDesk ha sido cerrado correctamente." -ForegroundColor Green
                    } catch {
                        Write-Host "`tError al cerrar el proceso AnyDesk: $_" -ForegroundColor Red
                        $errors += "No se pudo cerrar el proceso AnyDesk."
                    }
                    foreach ($file in $filesToDelete) {
                        try {
                            if (Test-Path $file) {
                                Remove-Item -Path $file -Force -ErrorAction Stop
                                Write-Host "`tArchivo eliminado: $file" -ForegroundColor Green
                                $deletedFilesCount++
                            } else {
                                Write-Host "`tArchivo no encontrado: $file" -ForegroundColor Red
                            }
                        } catch {
                            Write-Host "`nError al eliminar el archivo." -ForegroundColor Red
                        }
                    }
                    if ($errors.Count -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show("$deletedFilesCount archivo(s) eliminado(s) correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Se encontraron errores. Revisa la consola para más detalles.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } else {
                    Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
                }
            })
        $btnShowPrinters.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                try {
                    $printers = Get-WmiObject -Query "SELECT * FROM Win32_Printer" | ForEach-Object {
                        $printer = $_
                        $isShared = $printer.Shared -eq $true
                        [PSCustomObject]@{
                            Name       = $printer.Name.Substring(0, [Math]::Min(24, $printer.Name.Length))
                            PortName   = $printer.PortName.Substring(0, [Math]::Min(19, $printer.PortName.Length))
                            DriverName = $printer.DriverName.Substring(0, [Math]::Min(19, $printer.DriverName.Length))
                            IsShared   = if ($isShared) { "Sí" } else { "No" }
                        }
                    }
                    Write-Host "`nImpresoras disponibles en el sistema:"
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
        $btnClearPrintJobs.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                try {
                    # 1) Validar que se esté ejecutando como administrador
                    if (-not (Test-Administrator)) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Esta acción requiere permisos de administrador.`r`n" +
                            "Por favor, ejecuta Daniel Tools como administrador.",
                            "Permisos insuficientes",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        )
                        return
                    }
                    # 2) Intentar obtener el servicio Spooler
                    $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
                    if (-not $spooler) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "No se encontró el servicio 'Cola de impresión (Spooler)' en este equipo.",
                            "Servicio no encontrado",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Error
                        )
                        return
                    }
                    # 3) Limpiar trabajos de impresión (sin reventar en cada impresora)
                    try {
                        Get-Printer -ErrorAction Stop | ForEach-Object {
                            try {
                                Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
                            } catch {
                                Write-Host "`tNo se pudieron limpiar trabajos de la impresora '$($_.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
                            }
                        }
                    } catch {
                        Write-Host "`tNo se pudieron enumerar impresoras (Get-Printer). ¿Está instalado el módulo PrintManagement?" -ForegroundColor Yellow
                    }
                    # 4) Detener Spooler (si está corriendo)
                    if ($spooler.Status -eq 'Running') {
                        Write-Host "`tDeteniendo servicio Spooler..." -ForegroundColor DarkYellow
                        Stop-Service -Name Spooler -Force -ErrorAction Stop
                    } else {
                        Write-Host "`tSpooler no está en estado 'Running' (estado actual: $($spooler.Status))." -ForegroundColor DarkYellow
                    }
                    # Refrescar datos por si cambiaron
                    $spooler.Refresh()
                    # 5) Validar si el servicio está deshabilitado
                    if ($spooler.StartType -eq 'Disabled') {
                        [System.Windows.Forms.MessageBox]::Show(
                            "El servicio 'Cola de impresión (Spooler)' está DESHABILITADO." +
                            "`r`nHabilítalo manualmente desde services.msc para poder iniciarlo.",
                            "Spooler deshabilitado",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        )
                        return
                    }
                    # 6) Iniciar Spooler
                    Write-Host "`tIniciando servicio Spooler..." -ForegroundColor DarkYellow
                    Start-Service -Name Spooler -ErrorAction Stop

                    [System.Windows.Forms.MessageBox]::Show(
                        "Los trabajos de impresión han sido eliminados y el servicio de cola de impresión se ha reiniciado correctamente.",
                        "Operación exitosa",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                } catch {
                    Write-Host "`n[ERROR ClearPrintJobs] $($_.Exception.Message)" -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show(
                        "Ocurrió un error al intentar limpiar las impresoras o reiniciar el servicio:" +
                        "`r`n$($_.Exception.Message)",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                }
            })
        $LZMAbtnBuscarCarpeta.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
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
                    $instaladores = $instaladores | Sort-Object -Property Name -Descending
                    $LZMsubCarpetas = @("Selecciona instalador a renombrar") + ($instaladores | ForEach-Object { $_.Name })
                    $LZMrutasCompletas = $instaladores | ForEach-Object { $_.Path }
                    $formLZMA = Create-Form `
                        -Title "Carpetas LZMA" `
                        -Size (New-Object System.Drawing.Size(400, 200)) `
                        -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                        -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) `
                        -MaximizeBox $false -MinimizeBox $false
                    $LZMcomboBoxCarpetas = Create-ComboBox `
                        -Location (New-Object System.Drawing.Point(10, 10)) `
                        -Size (New-Object System.Drawing.Size(360, 20)) `
                        -DropDownStyle DropDownList `
                        -Font $defaultFont
                    foreach ($nombre in $LZMsubCarpetas) {
                        $LZMcomboBoxCarpetas.Items.Add($nombre)
                    }
                    $LZMcomboBoxCarpetas.SelectedIndex = 0
                    $lblLZMAExePath = Create-Label `
                        -Text "AI_ExePath: -" `
                        -Location (New-Object System.Drawing.Point(10, 35)) `
                        -Size (New-Object System.Drawing.Size(360, 70)) `
                        -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 0, 0)) `
                        -Font $defaultFont
                    $LZMbtnRenombrar = Create-Button `
                        -Text "Renombrar" `
                        -Location (New-Object System.Drawing.Point(10, 120)) `
                        -Size (New-Object System.Drawing.Size(180, 30)) `
                        -Enabled $false
                    $LMZAbtnSalir = Create-Button `
                        -Text "Salir" `
                        -Location (New-Object System.Drawing.Point(200, 120)) `
                        -Size (New-Object System.Drawing.Size(180, 30))
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
        $btnConfigurarIPs.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $formIpAssignAsignacion = Create-Form -Title "Asignación de IPs" -Size (New-Object System.Drawing.Size(400, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                    -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false
                $lblipAssignAdapter = Create-Label -Text "Seleccione el adaptador de red:" -Location (New-Object System.Drawing.Point(10, 20))
                $lblipAssignAdapter.AutoSize = $true
                $formIpAssignAsignacion.Controls.Add($lblipAssignAdapter)
                $ComboBipAssignAdapters = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(360, 20)) -DropDownStyle DropDownList `
                    -DefaultText "Selecciona 1 adaptador de red"
                $ComboBipAssignAdapters.Add_SelectedIndexChanged({
                        if ($ComboBipAssignAdapters.SelectedItem -ne "") {
                            $btnipAssignAssignIP.Enabled = $true
                            $btnipAssignChangeToDhcp.Enabled = $true
                        } else {
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
        $btnLectorDPicacls.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                try {
                    $psexecPath = "C:\Temp\PsExec\PsExec.exe"
                    $psexecZip = "C:\Temp\PSTools.zip"
                    $psexecUrl = "https://download.sysinternals.com/files/PSTools.zip"
                    $psexecExtractPath = "C:\Temp\PsExec"
                    if (-Not (Test-Path $psexecPath)) {
                        Write-Host "`tPsExec no encontrado. Descargando desde Sysinternals..." -ForegroundColor Yellow
                        if (-Not (Test-Path "C:\Temp")) {
                            New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
                        }
                        Invoke-WebRequest -Uri $psexecUrl -OutFile $psexecZip
                        Write-Host "`tExtrayendo PsExec..." -ForegroundColor Cyan
                        Expand-Archive -Path $psexecZip -DestinationPath $psexecExtractPath -Force
                        if (-Not (Test-Path $psexecPath)) {
                            Write-Host "`tError: No se pudo extraer PsExec.exe." -ForegroundColor Red
                            return
                        }
                        Write-Host "`tPsExec descargado y extraído correctamente." -ForegroundColor Green
                    } else {
                        Write-Host "`tPsExec ya está instalado en: $psexecPath" -ForegroundColor Green
                    }
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
                    $comando1 = "icacls C:\Windows\System32\en-us /grant `"$grupoAdmin`":F"
                    $comando2 = "icacls C:\Windows\System32\en-us /grant `"NT AUTHORITY\SYSTEM`":F"
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
                        $url = "https://softrestaurant.com/drivers?download=120:dp"
                        $zipPath = "C:\Temp\Driver_DP.zip"
                        $extractPath = "C:\Temp\Driver_DP"
                        $exeName = "x64\Setup.msi"
                        $validationPath = "C:\Temp\Driver_DP\x64\Setup.msi"
                        DownloadAndRun -url $url -zipPath $zipPath -extractPath $extractPath -exeName $exeName -validationPath $validationPath
                    }

                } catch {
                    Write-Host "Error: $_" -ForegroundColor Red
                }
            })
        $btnAplicacionesNS.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $resultados = @()
                function Leer-Ini($filePath) {
                    if (Test-Path $filePath) {
                        $content = Get-Content $filePath
                        $dataSource = ($content | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
                        $catalog = ($content | Select-String -Pattern "^Catalog=(.*)"    | Select-Object -First 1).Matches.Groups[1].Value
                        $authType = ($content | Select-String -Pattern "^autenticacion=(\d+)").Matches.Groups[1].Value
                        $authUser = if ($authType -eq "2") { "sa" } elseif ($authType -eq "1") { "Windows" } else { "Desconocido" }

                        return @{
                            DataSource = $dataSource
                            Catalog    = $catalog
                            Usuario    = $authUser
                        }
                    }
                    return $null
                }
                $pathsToCheck = @(
                    @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
                    @{ Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
                    @{ Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
                    @{ Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
                    @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
                    @{ Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
                )
                foreach ($entry in $pathsToCheck) {
                    $basePath = $entry.Path
                    $mainIni = "$basePath\$($entry.INI)"
                    $appName = $entry.Nombre
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
                        $resultados += [PSCustomObject]@{
                            Aplicacion = $appName
                            INI        = "No encontrado"
                            DataSource = "NA"
                            Catalog    = "NA"
                            Usuario    = "NA"
                        }
                    }
                    $inisFolder = "$basePath\INIS"
                    if ($appName -eq "OnTheMinute" -and (Test-Path $inisFolder)) {
                        $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
                        if ($iniFiles.Count -gt 1) {
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
                    } elseif (Test-Path $inisFolder) {
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
                $columnas = @("Aplicacion", "INI", "DataSource", "Catalog", "Usuario")
                $anchos = @{}
                foreach ($col in $columnas) { $anchos[$col] = $col.Length }
                foreach ($res in $resultados) {
                    foreach ($col in $columnas) {
                        if ($res.$col.Length -gt $anchos[$col]) {
                            $anchos[$col] = $res.$col.Length
                        }
                    }
                }
                $titulos = $columnas | ForEach-Object { $_.PadRight($anchos[$_] + 2) }
                Write-Host ($titulos -join "") -ForegroundColor Cyan
                $separador = $columnas | ForEach-Object { ("-" * $anchos[$_]).PadRight($anchos[$_] + 2) }
                Write-Host ($separador -join "") -ForegroundColor Cyan
                foreach ($res in $resultados) {
                    $fila = $columnas | ForEach-Object { $res.$_.PadRight($anchos[$_] + 2) }
                    if ($res.INI -eq "No encontrado") {
                        Write-Host ($fila -join "") -ForegroundColor Red
                    } else {
                        Write-Host ($fila -join "")
                    }
                }
            })
        $btnCambiarOTM.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $syscfgPath = "C:\Windows\SysWOW64\Syscfg45_2.0.dll"
                $iniPath = "C:\NationalSoft\OnTheMinute4.5"
                if (-not (Test-Path $syscfgPath)) {
                    [System.Windows.Forms.MessageBox]::Show("El archivo de configuración no existe.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
                    Write-Host "`tEl archivo de configuración no existe." -ForegroundColor Red
                    return
                }
                $fileContent = Get-Content $syscfgPath
                $isSQL = $fileContent -match "494E5354414C4C=1" -and $fileContent -match "56455253495354454D41=3"
                $isDBF = $fileContent -match "494E5354414C4C=2" -and $fileContent -match "56455253495354454D41=2"
                if (!$isSQL -and !$isDBF) {
                    [System.Windows.Forms.MessageBox]::Show("No se detectó una configuración válida de SQL o DBF.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
                    Write-Host "`tNo se detectó una configuración válida de SQL o DBF." -ForegroundColor Red
                    return
                }
                $iniFiles = Get-ChildItem -Path $iniPath -Filter "*.ini"
                if ($iniFiles.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos INI en $iniPath.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
                    Write-Host "`tNo se encontraron archivos INI en $iniPath." -ForegroundColor Red
                    return
                }
                $iniSQLFile = $null
                $iniDBFFile = $null
                foreach ($iniFile in $iniFiles) {
                    $content = Get-Content $iniFile.FullName
                    if ($content -match "Provider=VFPOLEDB.1" -and -not $iniDBFFile) {
                        $iniDBFFile = $iniFile
                    }
                    if ($content -match "Provider=SQLOLEDB.1" -and -not $iniSQLFile) {
                        $iniSQLFile = $iniFile
                    }
                    if ($iniSQLFile -and $iniDBFFile) {
                        break
                    }
                }
                if (-not $iniSQLFile -or -not $iniDBFFile) {
                    [System.Windows.Forms.MessageBox]::Show("No se encontraron los archivos INI esperados.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK)
                    Write-Host "`tNo se encontraron los archivos INI esperados." -ForegroundColor Red
                    Write-Host "`tArchivos encontrados:" -ForegroundColor Yellow
                    $iniFiles | ForEach-Object { Write-Host "`t- $_.Name" }
                    return
                }
                $currentConfig = if ($isSQL) { "SQL" } else { "DBF" }
                $newConfig = if ($isSQL) { "DBF" } else { "SQL" }
                $message = "Actualmente tienes configurado: $currentConfig.`n¿Quieres cambiar a $newConfig?"
                $result = [System.Windows.Forms.MessageBox]::Show($message, "Cambiar Configuración", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
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
        $btnAddUser.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $formAddUser = Create-Form -Title "Crear Usuario de Windows" -Size (New-Object System.Drawing.Size(450, 250))
                $txtUsername = Create-TextBox -Location (New-Object System.Drawing.Point(120, 20)) -Size (New-Object System.Drawing.Size(290, 30))
                $lblUsername = Create-Label -Text "Nombre:" -Location (New-Object System.Drawing.Point(10, 20))
                $txtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(120, 60)) -Size (New-Object System.Drawing.Size(290, 30)) -UseSystemPasswordChar $true
                $lblPassword = Create-Label -Text "Contraseña:" -Location (New-Object System.Drawing.Point(10, 60))
                $cmbType = Create-ComboBox -Location (New-Object System.Drawing.Point(120, 100)) -Size (New-Object System.Drawing.Size(290, 30)) -Items @("Usuario estándar", "Administrador")
                $lblType = Create-Label -Text "Tipo:" -Location (New-Object System.Drawing.Point(10, 100))
                $adminGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-544').Name
                $userGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-545').Name
                $btnCreate = Create-Button -Text "Crear"    -Location (New-Object System.Drawing.Point(10, 150))  -Size (New-Object System.Drawing.Size(130, 30))
                $btnCancel = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(150, 150)) -Size (New-Object System.Drawing.Size(130, 30))
                $btnShow = Create-Button -Text "Mostrar usuarios" -Location (New-Object System.Drawing.Point(290, 150)) -Size (New-Object System.Drawing.Size(130, 30))
                $btnShow.Add_Click({
                        Write-Host "`nUsuarios actuales en el sistema:`n" -ForegroundColor Cyan
                        $users = Get-LocalUser
                        $usersTable = $users | ForEach-Object {
                            $user = $_
                            $estado = if ($user.Enabled) { "Habilitado" } else { "Deshabilitado" }
                            $tipoUsuario = "Usuario estándar"
                            try {
                                $adminMembers = Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop
                                if ($adminMembers | Where-Object { $_.SID -eq $user.SID }) {
                                    $tipoUsuario = "Administrador"
                                } else {
                                    $userMembers = Get-LocalGroupMember -Group $userGroup -ErrorAction Stop
                                    if (-not ($userMembers | Where-Object { $_.SID -eq $user.SID })) {
                                        $grupos = Get-LocalGroup | ForEach-Object {
                                            if (Get-LocalGroupMember -Group $_ | Where-Object { $_.SID -eq $user.SID }) {
                                                $_.Name
                                            }
                                        }
                                        $tipoUsuario = "Miembro de: " + ($grupos -join ", ")
                                    }
                                }
                            } catch {
                                $tipoUsuario = "Error verificando grupos"
                            }
                            $nombre = $user.Name.Substring(0, [Math]::Min(25, $user.Name.Length))
                            $tipo = $tipoUsuario.Substring(0, [Math]::Min(40, $tipoUsuario.Length))
                            [PSCustomObject]@{
                                Nombre = $nombre
                                Tipo   = $tipo
                                Estado = $estado
                            }
                        }
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
                $btnCreate.Add_Click({
                        $username = $txtUsername.Text.Trim()
                        $password = $txtPassword.Text
                        $type = $cmbType.SelectedItem

                        if (-not $username -or -not $password) {
                            Write-Host "`nError: Nombre y contraseña son requeridos" -ForegroundColor Red; return
                        }
                        if ($password.Length -lt 8 -or $password -notmatch '[A-Z]' -or $password -notmatch '[a-z]' -or $password -notmatch '\d' -or $password -notmatch '[^\w]') {
                            Write-Host "`nError: La contraseña debe tener al menos 8 caracteres, incluir mayúscula, minúscula, número y símbolo" -ForegroundColor Red; return
                        }
                        try {
                            if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
                                Write-Host "`nError: El usuario '$username' ya existe" -ForegroundColor Red; return
                            }
                            $securePassword = (New-Object System.Net.NetworkCredential('', $passwordText)).SecurePassword
                            New-LocalUser -Name $username -Password $securePassword -AccountNeverExpires -PasswordNeverExpires
                            Write-Host "`nUsuario '$username' creado exitosamente" -ForegroundColor Green
                            $group = if ($type -eq 'Administrador') { $adminGroup } else { $userGroup }
                            Add-LocalGroupMember -Group $group -Member $username
                            Write-Host "`tUsuario agregado al grupo $group" -ForegroundColor Cyan
                            $formAddUser.Close()
                        } catch {
                            Write-Host "`nError durante la creación del usuario: $_" -ForegroundColor Red
                        }
                    })
                $btnCancel.Add_Click({ Write-Host "`tOperación cancelada." -ForegroundColor Yellow; $formAddUser.Close() })
                $formAddUser.Controls.AddRange(@($txtUsername, $txtPassword, $cmbType, $btnCreate, $btnCancel, $btnShow, $lblUsername, $lblPassword, $lblType))
                $formAddUser.ShowDialog()
            })
        $btnConnectDb.Add_Click({
                Write-Host "`nConectando a la instancia..." -ForegroundColor Gray
                try {
                    if ($null -eq $global:txtServer -or
                        $null -eq $global:txtUser -or
                        $null -eq $global:txtPassword) {
                        throw "Error interno: controles de conexión no inicializados."
                    }
                    $serverText = $global:txtServer.Text.Trim()
                    $userText = $global:txtUser.Text.Trim()
                    $passwordText = $global:txtPassword.Text
                    Write-DzDebug "`t[DEBUG] | Server='$serverText' User='$userText' PasswordLen=$($passwordText.Length)"
                    if ([string]::IsNullOrWhiteSpace($serverText) -or
                        [string]::IsNullOrWhiteSpace($userText) -or
                        [string]::IsNullOrWhiteSpace($passwordText)) {
                        throw "Complete todos los campos de conexión"
                    }
                    $securePassword = (New-Object System.Net.NetworkCredential('', $passwordText)).SecurePassword
                    # Usa el valor recién leído del textbox
                    $credential = New-Object System.Management.Automation.PSCredential($userText, $securePassword)
                    $global:server = $serverText
                    $global:user = $userText
                    $global:password = $passwordText
                    $global:dbCredential = $credential
                    $databases = Get-SqlDatabases -Server $serverText -Credential $credential
                    if (-not $databases -or $databases.Count -eq 0) {
                        throw "Conexión correcta, pero no se encontraron bases de datos disponibles."
                    }
                    $global:cmbDatabases.Items.Clear()
                    foreach ($db in $databases) {
                        [void]$global:cmbDatabases.Items.Add($db)
                    }
                    $global:cmbDatabases.Enabled = $true
                    $global:cmbDatabases.SelectedIndex = 0
                    $global:lblConnectionStatus.Text = @"
Conectado a:
Servidor: $serverText
Base de datos: $($global:database)
"@.Trim()
                    $global:lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green
                    Set-ControlEnabled -Control $global:txtServer       -Enabled $false -Name 'txtServer'
                    Set-ControlEnabled -Control $global:txtUser         -Enabled $false -Name 'txtUser'
                    Set-ControlEnabled -Control $global:txtPassword     -Enabled $false -Name 'txtPassword'
                    Set-ControlEnabled -Control $global:btnExecute      -Enabled $true  -Name 'btnExecute'
                    Set-ControlEnabled -Control $global:cmbQueries      -Enabled $true  -Name 'cmbQueries'
                    Set-ControlEnabled -Control $global:btnConnectDb    -Enabled $false -Name 'btnConnectDb'
                    Set-ControlEnabled -Control $global:btnBackup       -Enabled $true  -Name 'btnBackup'
                    Set-ControlEnabled -Control $global:btnDisconnectDb -Enabled $true  -Name 'btnDisconnectDb'
                    Set-ControlEnabled -Control $global:rtbQuery        -Enabled $true  -Name 'rtbQuery'
                } catch {
                    Write-DzDebug "`t[DEBUG][btnConnectDb] CATCH: $($_.Exception.Message)"
                    Write-DzDebug "`t[DEBUG][btnConnectDb] Tipo: $($_.Exception.GetType().FullName)"
                    Write-DzDebug "`t[DEBUG][btnConnectDb] Stack: $($_.ScriptStackTrace)"
                    [System.Windows.Forms.MessageBox]::Show(
                        "Error de conexión: $($_.Exception.Message)",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                    Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
                }
            })
        $btnDisconnectDb.Add_Click({
                try {
                    if ($global:connection -and
                        $global:connection.State -ne [System.Data.ConnectionState]::Closed) {

                        $global:connection.Close()
                        $global:connection.Dispose()
                    }
                    $global:connection = $null
                    $global:dbCredential = $null  # NUEVO: limpiamos la credencial
                    $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
                    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
                    $btnConnectDb.Enabled = $true
                    $btnBackup.Enabled = $false
                    $btnDisconnectDb.Enabled = $false
                    $btnExecute.Enabled = $false
                    $rtbQuery.Enabled = $false
                    $txtServer.Enabled = $true
                    $txtUser.Enabled = $true
                    $txtPassword.Enabled = $true
                    $cmbQueries.Enabled = $false
                    $cmbDatabases.Items.Clear()
                    $cmbDatabases.Enabled = $false
                    Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow
                } catch {
                    Write-Host "`nError al desconectar: $($_.Exception.Message)" -ForegroundColor Red
                }
            })
        $btnCreateAPK.Add_Click({
                Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
                $dllPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\up.dll"
                $infoPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\info.txt"
                try {
                    Write-Host "`nIniciando proceso de creación de APK..." -ForegroundColor Cyan
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
                    $jsonContent = Get-Content $infoPath -Raw | ConvertFrom-Json
                    $versionApp = $jsonContent.versionApp
                    Write-Host "Versión detectada: $versionApp" -ForegroundColor Green
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
                    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                    $saveDialog.Filter = "Archivo APK (*.apk)|*.apk"
                    $saveDialog.FileName = "SRM_$versionApp.apk"
                    $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

                    if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                        Write-Host "Guardado cancelado por el usuario" -ForegroundColor Yellow
                        return
                    }
                    Copy-Item -Path $dllPath -Destination $saveDialog.FileName -Force
                    Write-Host "APK generado exitosamente en: $($saveDialog.FileName)" -ForegroundColor Green
                    [System.Windows.Forms.MessageBox]::Show("APK creado correctamente!", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                } catch {
                    Write-Host "Error durante el proceso: $($_.Exception.Message)" -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("Error durante la creación del APK. Consulte la consola para más detalles.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            })

        $btnBackup.Add_Click({
                $chocoInstalled = Test-ChocolateyInstalled
                if (-not $chocoInstalled) {
                    Write-Host "Chocolatey no está instalado." -ForegroundColor Yellow
                    $messageInstalacionChoco = @"
Chocolatey es necesario SOLAMENTE si deseas:
✓ Comprimir el respaldo (crear ZIP con contraseña)
✓ Subir el respaldo a Mega.nz

Si solo necesitas crear el respaldo básico (archivo .BAK), NO es necesario instalarlo.

¿Deseas instalar Chocolatey ahora para habilitar estas funciones adicionales?
"@
                    $response = [System.Windows.Forms.MessageBox]::Show(
                        $messageInstalacionChoco,
                        "Instalación Requerida",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )

                    if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Write-Host "Instalando Chocolatey..." -ForegroundColor Cyan
                        try {
                            Set-ExecutionPolicy Bypass -Scope Process -Force
                            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                            iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

                            [System.Windows.Forms.MessageBox]::Show(
                                "Chocolatey instalado. Por favor reinicie PowerShell y vuelva a ejecutar la herramienta.",
                                "Reinicio Requerido",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Information
                            )
                            Stop-Process -Id $PID -Force
                        } catch {
                            Write-Host "Error instalando Chocolatey: $_" -ForegroundColor Red
                            [System.Windows.Forms.MessageBox]::Show(
                                "Error instalando Chocolatey: $_",
                                "Error",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Error
                            )
                        }
                        return
                    } else {
                        Write-Host "El usuario omitió la instalación de Chocolatey." -ForegroundColor Yellow
                        [System.Windows.Forms.MessageBox]::Show(
                            "Opciones de compresión/subida deshabilitadas.",
                            "Advertencia",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning
                        )
                    }
                }
                $script:animTimer = $null
                $script:backupTimer = $null
                $serverRaw = $global:server
                $sameHost = Test-SameHost -serverName $serverRaw
                $machinePart = $serverRaw.Split('\')[0]
                $machineName = $machinePart.Split(',')[0]
                if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }
                $global:tempBackupFolder = "\\$machineName\C$\Temp\SQLBackups"
                $formSize = New-Object System.Drawing.Size(480, 400)
                $formBackupOptions = Create-Form -Title "Opciones de Respaldo" `
                    -Size $formSize `
                    -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                    -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog)
                $chkRespaldo = New-Object System.Windows.Forms.CheckBox
                $chkRespaldo.Text = "Respaldar"
                $chkRespaldo.Checked = $true
                $chkRespaldo.Enabled = $false
                $chkRespaldo.AutoSize = $true
                $chkRespaldo.Location = New-Object System.Drawing.Point(20, 20)
                $formBackupOptions.Controls.Add($chkRespaldo)
                $lblNombre = New-Object System.Windows.Forms.Label
                $lblNombre.Text = "Nombre del respaldo:"
                $lblNombre.AutoSize = $true
                $lblNombre.Location = New-Object System.Drawing.Point(20, 50)
                $formBackupOptions.Controls.Add($lblNombre)
                $txtNombre = New-Object System.Windows.Forms.TextBox
                $txtNombre.Size = New-Object System.Drawing.Size(350, 20)
                $txtNombre.Location = New-Object System.Drawing.Point(20, 70)
                $timestampsDefault = Get-Date -Format 'yyyyMMdd-HHmmss'
                $selectedDb = $cmbDatabases.SelectedItem
                if ($selectedDb) {
                    $txtNombre.Text = "$selectedDb-$timestampsDefault.bak"
                } else {
                    $txtNombre.Text = "Backup-$timestampsDefault.bak"
                }
                $formBackupOptions.Controls.Add($txtNombre)
                $tooltipCHK = New-Object System.Windows.Forms.ToolTip
                $chkComprimir = New-Object System.Windows.Forms.CheckBox
                $chkComprimir.Text = "Comprimir"
                $chkComprimir.AutoSize = $true
                $chkComprimir.Location = New-Object System.Drawing.Point(20, 110)
                if (-not $sameHost) {
                    $chkComprimir.Enabled = $false
                    $chkComprimir.Checked = $false
                    $tooltipCHK.SetToolTip($chkComprimir, "Solo disponible si se ejecuta en el mismo host que el servidor.")
                } else {
                    $chkComprimir.Enabled = $true
                }
                $formBackupOptions.Controls.Add($chkComprimir)
                $chkComprimir.Enabled = $chocoInstalled  # <-- Nueva línea
                if (-not $chocoInstalled) {
                    $tooltipCHK.SetToolTip($chkComprimir, "Requiere Chocolatey instalado")
                }
                $lblPassword = New-Object System.Windows.Forms.Label
                $lblPassword.Text = "Contraseña (opcional) para ZIP:"
                $lblPassword.AutoSize = $true
                $lblPassword.Location = New-Object System.Drawing.Point(40, 135)
                $formBackupOptions.Controls.Add($lblPassword)
                $txtPassword = New-Object System.Windows.Forms.TextBox
                $txtPassword.Size = New-Object System.Drawing.Size(250, 20)
                $txtPassword.Location = New-Object System.Drawing.Point(40, 155)
                $txtPassword.UseSystemPasswordChar = $true
                $txtPassword.Enabled = $false
                $formBackupOptions.Controls.Add($txtPassword)
                $chkComprimir.Add_CheckedChanged({
                        if ($chkComprimir.Checked) {
                            $txtPassword.Enabled = $true
                        } else {
                            $txtPassword.Enabled = $false
                            $txtPassword.Text = ""
                            $chkSubir.Checked = $false
                            $chkSubir.Enabled = $false
                        }
                    })
                $chkSubir = New-Object System.Windows.Forms.CheckBox
                $chkSubir.Text = "Subir a Mega.nz"
                $chkSubir.AutoSize = $true
                $chkSubir.Location = New-Object System.Drawing.Point(20, 195)
                $chkSubir.Checked = $false
                $chkSubir.Enabled = $false  # inicialmente deshabilitado; se activará al chequeo de "Comprimir"
                $formBackupOptions.Controls.Add($chkSubir)
                $chkSubir.Enabled = $chocoInstalled  # <-- Nueva línea
                if (-not $chocoInstalled) {
                    $tooltipCHK.SetToolTip($chkSubir, "Requiere Chocolatey instalado")
                }
                $chkComprimir.Add_CheckedChanged({
                        if ($chkComprimir.Checked) {
                            if ($sameHost) {
                                #$chkSubir.Enabled = $true
                                $chkSubir.Enabled = $false
                                $tooltipCHK.SetToolTip($chkSubir, "Activar para subir respaldo comprimido a Mega.nz.")
                            } else {
                                $chkSubir.Enabled = $false
                                $chkSubir.Checked = $false
                                $tooltipCHK.SetToolTip($chkSubir, "No disponible: debe ejecutar desde el mismo host que el servidor.")
                            }
                        } else {
                            $chkSubir.Enabled = $false
                            $chkSubir.Checked = $false
                        }
                    })
                $pbBackup = New-Object System.Windows.Forms.ProgressBar
                $pbBackup.Location = New-Object System.Drawing.Point(20, 240)
                $pbBackup.Size = New-Object System.Drawing.Size(420, 20)
                $pbBackup.Minimum = 0
                $pbBackup.Maximum = 100
                $pbBackup.Value = 0
                $pbBackup.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $pbBackup.Visible = $false
                $formBackupOptions.Controls.Add($pbBackup)
                $btnAceptar = Create-Button -Text "Aceptar" `
                    -Size (New-Object System.Drawing.Size(120, 30)) `
                    -Location (New-Object System.Drawing.Point(20, 270))
                $formBackupOptions.Controls.Add($btnAceptar)
                $btnAbrirCarpeta = Create-Button -Text "Abrir Carpeta" `
                    -Size (New-Object System.Drawing.Size(120, 30)) `
                    -Location (New-Object System.Drawing.Point(160, 270))
                $formBackupOptions.Controls.Add($btnAbrirCarpeta)
                $btnCerrar = Create-Button -Text "Cerrar" `
                    -Size (New-Object System.Drawing.Size(120, 30)) `
                    -Location (New-Object System.Drawing.Point(340, 270))
                $formBackupOptions.Controls.Add($btnCerrar)
                $btnAbrirCarpeta.Add_Click({
                        if (Test-Path $global:tempBackupFolder) {
                            Start-Process explorer.exe $global:tempBackupFolder
                        } else {
                            [System.Windows.Forms.MessageBox]::Show(
                                "La carpeta de respaldos no existe todavía.",
                                "Atención",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                        }
                    })
                $btnCerrar.Add_Click({
                        $formBackupOptions.Close()
                    })
                $btnAceptar.Add_Click({
                        $chkComprimir.Enabled = $false
                        $chkSubir.Enabled = $false
                        $txtNombre.Enabled = $false
                        $txtPassword.Enabled = $false
                        $btnAceptar.Enabled = $false
                        $btnAbrirCarpeta.Enabled = $false
                        $btnCerrar.Enabled = $false
                        $script:lblTrabajando = New-Object System.Windows.Forms.Label
                        $script:lblTrabajando.Text = "Iniciando respaldo..."
                        $script:lblTrabajando.AutoSize = $false
                        $script:lblTrabajando.Size = New-Object System.Drawing.Size(420, 20)
                        $script:lblTrabajando.Location = New-Object System.Drawing.Point(20, 215)
                        $formBackupOptions.Controls.Add($script:lblTrabajando)
                        $pbBackup.Visible = $true
                        $selectedDb = $cmbDatabases.SelectedItem
                        if (-not $selectedDb) {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Seleccione una base de datos primero.",
                                "Error",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Warning
                            )
                            $formBackupOptions.Close()
                            return
                        }
                        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                        $inputName = $txtNombre.Text.Trim()
                        if (-not $inputName.ToLower().EndsWith(".bak")) {
                            $bakFileName = "$inputName.bak"
                        } else {
                            $bakFileName = $inputName
                        }
                        $global:backupPath = Join-Path $global:tempBackupFolder $bakFileName
                        if (-not (Test-Path -Path $global:tempBackupFolder)) {
                            try {
                                New-Item -ItemType Directory -Path $global:tempBackupFolder -Force | Out-Null
                            } catch {
                                [System.Windows.Forms.MessageBox]::Show(
                                    "No se pudo crear la carpeta remota: $global:tempBackupFolder.`n$_",
                                    "Error",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Error
                                )
                                $formBackupOptions.Close()
                                return
                            }
                        }
                        $scriptBackup = {
                            param($srv, $usr, $pwd, $db, $pathBak)
                            $conn = New-Object System.Data.SqlClient.SqlConnection("Server=$srv; Database=master; User Id=$usr; Password=$pwd")
                            $conn.Open()
                            $cmd = $conn.CreateCommand()
                            $cmd.CommandText = "BACKUP DATABASE [$db] TO DISK='$pathBak' WITH CHECKSUM"
                            $cmd.CommandTimeout = 0
                            $cmd.ExecuteNonQuery()
                            $conn.Close()
                        }
                        $global:backupJob = Start-Job -ScriptBlock $scriptBackup -ArgumentList `
                            $global:server, $global:user, $global:password, $selectedDb, $global:backupPath
                        $script:animTimer = New-Object System.Windows.Forms.Timer
                        $script:animTimer.Interval = 400
                        $direction = 1
                        $script:animTimer.Add_Tick({
                                if ($pbBackup.Value -ge $pbBackup.Maximum) {
                                    $direction = -1
                                } elseif ($pbBackup.Value -le $pbBackup.Minimum) {
                                    $direction = 1
                                }
                                $pbBackup.Value += 10 * $direction
                            })
                        $script:animTimer.Start()
                        $script:backupTimer = New-Object System.Windows.Forms.Timer
                        $script:backupTimer.Interval = 500
                        $script:backupTimer.Add_Tick({
                                if ($global:backupJob.State -in 'Completed', 'Failed', 'Stopped') {
                                    if ($script:animTimer) { $script:animTimer.Stop() }
                                    if ($script:backupTimer) { $script:backupTimer.Stop() }
                                    Receive-Job $global:backupJob | Out-Null
                                    Remove-Job $global:backupJob -Force
                                    if ($formBackupOptions.InvokeRequired) {
                                        $formBackupOptions.Invoke([action] { $formBackupOptions.Enabled = $false })
                                    } else {
                                        $formBackupOptions.Enabled = $false
                                    }
                                    if ($global:backupJob.State -eq 'Completed') {
                                        Write-Host "Backup finalizado correctamente." -ForegroundColor Green
                                        if ($chkComprimir.Checked) {
                                            if (-not (Test-7ZipInstalled)) {
                                                Write-Host "7-Zip no encontrado. Intentando instalar con Chocolatey..."
                                                try {
                                                    if (Get-Command choco -ErrorAction SilentlyContinue) {
                                                        choco install 7zip -y | Out-Null
                                                        Start-Sleep -Seconds 2  # Dar un momento para que termine la instalación
                                                        if (-not (Test-7ZipInstalled)) {
                                                            throw "La instalación de 7-Zip no completó correctamente."
                                                        } else {
                                                            Write-Host "7-Zip instalado correctamente en 'C:\Program Files\7-Zip\7z.exe'."
                                                        }
                                                    } else {
                                                        throw "Chocolatey no está instalado. Imposible instalar 7-Zip automáticamente."
                                                    }
                                                } catch {
                                                    [System.Windows.Forms.MessageBox]::Show(
                                                        "Error instalando 7-Zip:`n$($_.Exception.Message)",
                                                        "Error",
                                                        [System.Windows.Forms.MessageBoxButtons]::OK,
                                                        [System.Windows.Forms.MessageBoxIcon]::Error
                                                    )
                                                    return
                                                }
                                            }
                                            $zipPath = "$global:backupPath.zip"
                                            $script:lblTrabajando.Text = "Comprimiendo respaldo..."
                                            if ($txtPassword.Text.Trim().Length -gt 0) {
                                                & "C:\Program Files\7-Zip\7z.exe" a -tzip -p"$($txtPassword.Text.Trim())" -mem=AES256 $zipPath $global:backupPath
                                            } else {
                                                & "C:\Program Files\7-Zip\7z.exe" a -tzip $zipPath $global:backupPath
                                            }
                                            Write-Host "Respaldo comprimido en: $zipPath" -ForegroundColor Green
                                            if ($chkSubir.Checked) {
                                                if (-not (Test-MegaToolsInstalled)) {
                                                    Write-Host "MegaTools no encontrado. Intentando instalar con Chocolatey..."
                                                    try {
                                                        if (Get-Command choco -ErrorAction SilentlyContinue) {
                                                            choco install megatools -y | Out-Null
                                                            Start-Sleep -Seconds 2
                                                            if (-not (Test-MegaToolsInstalled)) {
                                                                throw "La instalación de megatools falló"
                                                            }
                                                        } else {
                                                            throw "Chocolatey no está instalado"
                                                        }
                                                    } catch {
                                                        [System.Windows.Forms.MessageBox]::Show(
                                                            "Error instalando megatools: $($_.Exception.Message)",
                                                            "Error",
                                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                                            [System.Windows.Forms.MessageBoxIcon]::Error
                                                        )
                                                        $chkSubir.Checked = $false
                                                        return
                                                    }
                                                }
                                            }
                                            if ($chkSubir.Checked) {
                                                if (-not (Test-MegaToolsInstalled)) {
                                                    [System.Windows.Forms.MessageBox]::Show(
                                                        "megatools no está instalado. No puede subir a Mega.nz.",
                                                        "Error",
                                                        [System.Windows.Forms.MessageBoxButtons]::OK,
                                                        [System.Windows.Forms.MessageBoxIcon]::Error
                                                    )
                                                } else {
                                                    $script:lblTrabajando.Text = "Iniciando subida a Mega.nz..."
                                                    $pbBackup.Value = 0
                                                    Start-Sleep -Milliseconds 300
                                                    for ($i = 0; $i -le 30; $i += 10) {
                                                        $pbBackup.Value = $i
                                                        Start-Sleep -Milliseconds 200
                                                    }
                                                    $MegaUser = "gerardo.zermeno@nationalsoft.mx"
                                                    $MegaPass = "National.09$#"
                                                    $configPath = "$env:APPDATA\megatools.ini"
                                                    if (-not (Test-Path $configPath)) {
                                                        $configDir = Split-Path -Path $configPath -Parent
                                                        if (-not (Test-Path $configDir)) {
                                                            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
                                                        }
                                                        $megaConfig = @"
[Login]
Username = $MegaUser
Password = $MegaPass
"@
                                                        $megaConfig | Out-File -FilePath $configPath -Encoding utf8 -Force
                                                    }
                                                    for ($i = 30; $i -le 60; $i += 10) {
                                                        $pbBackup.Value = $i
                                                        Start-Sleep -Milliseconds 200
                                                    }
                                                    $script:lblTrabajando.Text = "Subiendo archivo comprimido..."
                                                    $zipToUpload = "$global:backupPath.zip"
                                                    $uploadCmd = "megatools put --username `"$MegaUser`" --password `"$MegaPass`" `"$zipToUpload`""
                                                    $uploadResult = cmd /c $uploadCmd 2>&1
                                                    for ($i = 60; $i -le 100; $i += 10) {
                                                        $pbBackup.Value = $i
                                                        Start-Sleep -Milliseconds 200
                                                    }
                                                    $downloadLink = $null
                                                    $uploadResult | ForEach-Object {
                                                        if ($_ -match 'https://mega\.nz/\S+') {
                                                            $downloadLink = $matches[0]
                                                        }
                                                    }
                                                    if (-not $downloadLink) {
                                                        $fileName = [System.IO.Path]::GetFileName($zipToUpload)
                                                        $exportCmd = "megatools export --username `"$MegaUser`" --password `"$MegaPass`" /Root/$fileName"
                                                        $exportResult = cmd /c $exportCmd 2>&1
                                                        if ($exportResult -match 'https://mega\.nz/\S+') {
                                                            $downloadLink = $matches[0]
                                                        }
                                                    }
                                                    if ($downloadLink) {
                                                        $cleanLink = $downloadLink -replace '[^\x20-\x7E]', ''
                                                        [System.Windows.Forms.MessageBox]::Show(
                                                            "Subida exitosa.`nEnlace: $cleanLink`n(Copiado al portapapeles)",
                                                            "Éxito",
                                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                                            [System.Windows.Forms.MessageBoxIcon]::Information
                                                        )
                                                        $cleanLink | Set-Clipboard
                                                    } else {
                                                        [System.Windows.Forms.MessageBox]::Show(
                                                            "Se completó la subida, pero no se pudo extraer el enlace.",
                                                            "Atención",
                                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                                            [System.Windows.Forms.MessageBoxIcon]::Warning
                                                        )
                                                    }
                                                    if (Test-Path $zipToUpload) {
                                                        Remove-Item $zipToUpload -Force
                                                    }
                                                }
                                            }
                                        }
                                        [System.Windows.Forms.Application]::DoEvents()
                                        $formBackupOptions.Close()
                                    } elseif ($global:backupJob.State -eq 'Stopped') {
                                        [System.Windows.Forms.MessageBox]::Show(
                                            "Backup cancelado por el usuario.",
                                            "Cancelado",
                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                            [System.Windows.Forms.MessageBoxIcon]::Information
                                        )
                                        $formBackupOptions.Close()
                                    } else {
                                        $errorMessage = Receive-Job $global:backupJob -ErrorAction SilentlyContinue
                                        [System.Windows.Forms.MessageBox]::Show(
                                            "Error en backup:`n$errorMessage",
                                            "Error",
                                            [System.Windows.Forms.MessageBoxButtons]::OK,
                                            [System.Windows.Forms.MessageBoxIcon]::Error
                                        )
                                    }
                                }
                            })
                        $script:backupTimer.Start()
                    })
                $formBackupOptions.ShowDialog()
            })
        $btnReloadConnections = Create-Button -Text "Recargar Conexiones" -Location (New-Object System.Drawing.Point(10, 180)) `
            -Size (New-Object System.Drawing.Size(180, 30)) `
            -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 255)) `
            -ToolTip "Recargar la lista de conexiones desde archivos INI"
        $btnReloadConnections.Add_Click({
                Write-Host "Recargando conexiones desde archivos INI..." -ForegroundColor Cyan
                Load-IniConnectionsToComboBox
            })
        $btnExit.Add_Click({
                $form = $this.FindForm()
                if ($form -ne $null) {
                    $form.Close()
                }
            })
        #Load-IniConnectionsToComboBox
        return $formPrincipal

    } catch {
        Write-Host "`n===============================================" -ForegroundColor Red
        Write-Host "ERROR FATAL NO CAPTURADO" -ForegroundColor Red
        Write-Host "===============================================" -ForegroundColor Red
        Write-Host "Tipo     : $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
        Write-Host "Mensaje  : $($_.Exception.Message)" -ForegroundColor Yellow

        if ($_.Exception.InnerException) {
            Write-Host "`nExcepción interna:" -ForegroundColor Cyan
            Write-Host "  $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
        }

        if ($_.InvocationInfo) {
            Write-Host "`nUbicación:" -ForegroundColor Cyan
            Write-Host "  $($_.InvocationInfo.PositionMessage)" -ForegroundColor Yellow
        }

        if ($_.Exception.StackTrace) {
            Write-Host "`nStack Trace:" -ForegroundColor Cyan
            Write-Host $_.Exception.StackTrace -ForegroundColor Gray
        }

        Write-Host "===============================================`n" -ForegroundColor Red

        Write-Host "`nPresione una tecla para salir..." -ForegroundColor Gray
        pause
        exit 1

    } finally {
        Write-Host "`nScript finalizado." -ForegroundColor Gray
    }
}
function Start-Application {
    Write-Host "Iniciando aplicación..." -ForegroundColor Cyan

    if (-not (Initialize-Environment)) {
        Write-Host "Error inicializando entorno. Saliendo..." -ForegroundColor Red
        return
    }

    Register-GlobalErrorHandlers

    $mainForm = $null

    try {
        $mainForm = New-MainForm

        if ($mainForm -eq $null) {
            Write-Host "Error: No se pudo crear el formulario principal" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "No se pudo crear la interfaz gráfica. Verifique los logs.",
                "Error crítico",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        Write-Host "Mostrando formulario..." -ForegroundColor Yellow
        $result = $mainForm.ShowDialog()

        Write-Host "`nAplicación cerrada normalmente." -ForegroundColor Green
        Write-Host "Resultado del diálogo: $result" -ForegroundColor Gray

    } catch {
        Write-Host "`n===============================================" -ForegroundColor Red
        Write-Host "ERROR EN Start-Application" -ForegroundColor Red
        Write-Host "===============================================" -ForegroundColor Red
        Write-Host "Tipo     : $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
        Write-Host "Mensaje  : $($_.Exception.Message)" -ForegroundColor Yellow

        if ($_.Exception.InnerException) {
            Write-Host "`nExcepción interna:" -ForegroundColor Cyan
            Write-Host "  Tipo    : $($_.Exception.InnerException.GetType().FullName)" -ForegroundColor Yellow
            Write-Host "  Mensaje : $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
        }

        if ($_.InvocationInfo) {
            Write-Host "`nUbicación del error:" -ForegroundColor Cyan
            Write-Host "  Archivo : $($_.InvocationInfo.ScriptName)" -ForegroundColor Yellow
            Write-Host "  Línea   : $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
            Write-Host "  Comando : $($_.InvocationInfo.Line.Trim())" -ForegroundColor Yellow
        }

        if ($_.ScriptStackTrace) {
            Write-Host "`nStack Trace completo:" -ForegroundColor Cyan
            Write-Host $_.ScriptStackTrace -ForegroundColor Gray
        }

        Write-Host "===============================================`n" -ForegroundColor Red

        [System.Windows.Forms.MessageBox]::Show(
            "Error mostrando formulario:`n`n$($_.Exception.Message)`n`nRevise la consola para detalles completos.",
            "Error en la aplicación",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )

    } finally {
        Write-Host "`nLimpiando recursos..." -ForegroundColor Cyan

        if ($mainForm -ne $null) {
            try {
                if (-not $mainForm.IsDisposed) {
                    $mainForm.Dispose()
                    Write-Host "  ✓ Formulario principal liberado" -ForegroundColor Green
                }
            } catch {
                Write-Host "  ✗ Error liberando formulario: $_" -ForegroundColor Yellow
            }
        }
    }
}
try {
    Start-Application
} catch {
    Write-Host "Error inesperado al iniciar la aplicación: $($_.Exception.Message)" -ForegroundColor Red

    if ($_.InvocationInfo) {
        Write-Host "Ubicación: $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    }

    throw
}