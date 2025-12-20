#requires -Version 5.0

function Get-PredefinedQueries {
    return @{
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
ORDER BY t.IsEnabled DESC, t.Name;
"@
        "SR | Actualizar contraseña de administrador"     = @"
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
        "SR SYNC | nsplatformcontrol"                     = @"
BEGIN TRY
    BEGIN TRANSACTION;
    SELECT WorkspaceId, EntityType, OperationType, 0 AS IsSync, 0 AS Attempts, CreateDate
    INTO #tempcontroltaxes
    FROM nsplatformcontrol
    WHERE EntityType = 1;
    TRUNCATE TABLE nsplatformcontrol;
    INSERT INTO nsplatformcontrol (WorkspaceId, EntityType, OperationType, IsSync, Attempts, CreateDate)
    SELECT WorkspaceId, EntityType, OperationType, IsSync, Attempts, CreateDate
    FROM #tempcontroltaxes;
    DROP TABLE #tempcontroltaxes;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
"@
        "SR | Memoria Insuficiente"                       = @"
UPDATE empresas
        SET nombre='', razonsocial='',direccion='',sucursal='',
        rfc='',curp='', telefono='',giro='',contacto='',fax='',
        email='',idhardware='',web='',ciudad='',estado='',pais='',
        ciudadsucursal='',estadosucursal='',codigopostal='',a86ed5f9d02ec5b3='',
        codigopostalsucursal='',uid=newid();
GO
DELETE FROM registro_licencias;
"@
        "OTM | Eliminar Server en OTM"                    = @"
SELECT serie, ipserver, nombreservidor
    FROM configuracion;
"@
        "NSH | Eliminar Server en Hoteles"                = @"
SELECT serievalida, numserie, ipserver, nombreservidor, llave
    FROM configuracion;
"@
        "Restcard | Eliminar Server en Rest Card"         = @"
SELECT estacion, ipservidor FROM tabvariables;
"@
        "sql | Listar usuarios e idiomas"                 = @"
SELECT
    p.name AS Usuario,
    l.default_language_name AS Idioma
FROM
    sys.server_principals p
LEFT JOIN
    sys.sql_logins l ON p.principal_id = l.principal_id
WHERE
    p.type IN ('S', 'U')
"@
    }
}

# En Queries.psm1, reemplaza la función Initialize-PredefinedQueries:

function Initialize-PredefinedQueries {
    param(
        [Parameter(Mandatory = $true)]
        $ComboQueries,
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [hashtable]$Queries
    )

    # Detectar si es WPF o WinForms
    $isWPF = $RichTextBox.GetType().FullName -like "*System.Windows.Controls*"

    if ($isWPF) {
        Write-DzDebug "`t[DEBUG] RichTextBox detectado: WPF"
    } else {
        Write-DzDebug "`t[DEBUG] RichTextBox detectado: Windows Forms"
    }

    # Agregar queries al ComboBox
    $ComboQueries.Items.Add("Selecciona una consulta predefinida") | Out-Null
    foreach ($key in ($Queries.Keys | Sort-Object)) {
        $ComboQueries.Items.Add($key) | Out-Null
    }
    $ComboQueries.SelectedIndex = 0

    # Evento de selección
    if ($isWPF) {
        # WPF: SelectionChanged
        $ComboQueries.Add_SelectionChanged({
                $selectedQuery = $ComboQueries.SelectedItem
                if ($selectedQuery -and $selectedQuery -ne "Selecciona una consulta predefinida") {
                    if ($Queries.ContainsKey($selectedQuery)) {
                        $queryText = $Queries[$selectedQuery]

                        # Limpiar y establecer texto en WPF RichTextBox
                        $RichTextBox.Document.Blocks.Clear()
                        $paragraph = New-Object System.Windows.Documents.Paragraph
                        $run = New-Object System.Windows.Documents.Run($queryText)
                        $paragraph.Inlines.Add($run)
                        $RichTextBox.Document.Blocks.Add($paragraph)
                    }
                }
            })
    } else {
        # Windows Forms: SelectedIndexChanged
        $ComboQueries.Add_SelectedIndexChanged({
                $selectedQuery = $ComboQueries.SelectedItem
                if ($selectedQuery -and $selectedQuery -ne "Selecciona una consulta predefinida") {
                    if ($Queries.ContainsKey($selectedQuery)) {
                        $RichTextBox.Text = $Queries[$selectedQuery]
                    }
                }
            })
    }
}

Export-ModuleMember -Function @(
    'Get-PredefinedQueries',
    'Initialize-PredefinedQueries',
    'Remove-SqlComments'
)