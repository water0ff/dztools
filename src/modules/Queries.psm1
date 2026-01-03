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
function Get-TextPointerFromOffset {
    param(
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [int]$Offset
    )

    $pointer = $RichTextBox.Document.ContentStart
    $count = 0
    while ($null -ne $pointer) {
        if ($pointer.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward) -eq [System.Windows.Documents.TextPointerContext]::Text) {
            $textRun = $pointer.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
            if ($count + $textRun.Length -ge $Offset) {
                return $pointer.GetPositionAtOffset($Offset - $count)
            }
            $count += $textRun.Length
        }
        $pointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
    }

    return $RichTextBox.Document.ContentEnd
}

function Get-DzThemeBrush {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Hex,
        [Parameter(Mandatory = $true)]
        [System.Windows.Media.Brush]$Fallback
    )

    if ([string]::IsNullOrWhiteSpace($Hex)) {
        return $Fallback
    }

    try {
        $brush = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Hex)
        if ($brush -is [System.Windows.Freezable] -and $brush.CanFreeze) { $brush.Freeze() }
        return $brush
    } catch {
        return $Fallback
    }
}

function Set-WpfSqlHighlighting {
    param(
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [string]$Keywords
    )

    if ($null -eq $RichTextBox -or $null -eq $RichTextBox.Document) {
        return
    }

    $script:isHighlightingQuery = $true
    $theme = Get-DzUiTheme
    $defaultBrush = Get-DzThemeBrush -Hex $theme.ControlForeground -Fallback ([System.Windows.Media.Brushes]::Black)
    $commentBrush = Get-DzThemeBrush -Hex $theme.AccentMuted -Fallback ([System.Windows.Media.Brushes]::DarkGreen)
    $keywordBrush = Get-DzThemeBrush -Hex $theme.AccentPrimary -Fallback ([System.Windows.Media.Brushes]::Blue)

    $textRange = New-Object System.Windows.Documents.TextRange($RichTextBox.Document.ContentStart, $RichTextBox.Document.ContentEnd)
    $textRange.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, $defaultBrush)
    $text = $textRange.Text

    $commentRanges = @()
    foreach ($c in [regex]::Matches($text, '--.*', 'Multiline')) {
        $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $c.Index
        $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($c.Index + $c.Length)
        (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
            [System.Windows.Documents.TextElement]::ForegroundProperty,
            $commentBrush
        )
        $commentRanges += [PSCustomObject]@{ Start = $c.Index; End = $c.Index + $c.Length }
    }

    foreach ($b in [regex]::Matches($text, '/\*[\s\S]*?\*/', 'Multiline')) {
        $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $b.Index
        $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($b.Index + $b.Length)
        (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
            [System.Windows.Documents.TextElement]::ForegroundProperty,
            $commentBrush
        )
        $commentRanges += [PSCustomObject]@{ Start = $b.Index; End = $b.Index + $b.Length }
    }

    foreach ($m in [regex]::Matches($text, "\b($Keywords)\b", 'IgnoreCase')) {
        $inComment = $commentRanges | Where-Object { $m.Index -ge $_.Start -and $m.Index -lt $_.End }
        if (-not $inComment) {
            $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $m.Index
            $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($m.Index + $m.Length)
            (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
                [System.Windows.Documents.TextElement]::ForegroundProperty,
                $keywordBrush
            )
        }
    }

    $script:isHighlightingQuery = $false
}
function Initialize-PredefinedQueries {
    param(
        [Parameter(Mandatory = $true)]
        $ComboQueries,
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [hashtable]$Queries,
        [Parameter(Mandatory = $false)]
        $Window
    )

    # Detectar si es WPF o WinForms
    $isWPF = $RichTextBox.GetType().FullName -like "*System.Windows.Controls*"

    if ($isWPF) {
        Write-DzDebug "`t[DEBUG] RichTextBox detectado: WPF"
        Write-DzDebug "`t[DEBUG] RichTextBox es null: $($null -eq $RichTextBox)"
    } else {
        Write-DzDebug "`t[DEBUG] RichTextBox detectado: Windows Forms"
    }

    # Agregar queries al ComboBox
    $ComboQueries.Items.Clear()
    $ComboQueries.Items.Add("Selecciona una consulta predefinida") | Out-Null
    foreach ($key in ($Queries.Keys | Sort-Object)) {
        $ComboQueries.Items.Add($key) | Out-Null
    }
    $ComboQueries.SelectedIndex = 0

    # Definir las palabras clave SQL
    $sqlKeywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE'

    # Variable para controlar el resaltado
    $isHighlightingQuery = $false

    # Crear script blocks usando GetNewClosure para capturar las variables
    $selectionChangedScript = {
        param($sender, $e)

        try {
            $selectedQuery = $sender.SelectedItem
            if ($selectedQuery -and $selectedQuery -ne "Selecciona una consulta predefinida") {
                # Usar la variable capturada
                if ($this.Queries.ContainsKey($selectedQuery)) {
                    $queryText = $this.Queries[$selectedQuery]

                    # Limpiar y establecer texto en WPF RichTextBox
                    $this.RichTextBox.Document.Blocks.Clear()
                    $paragraph = New-Object System.Windows.Documents.Paragraph
                    $run = New-Object System.Windows.Documents.Run($queryText)
                    $paragraph.Inlines.Add($run)
                    $this.RichTextBox.Document.Blocks.Add($paragraph)

                    # Llamar a la función de resaltado
                    Set-WpfSqlHighlighting -RichTextBox $this.RichTextBox -Keywords $this.SqlKeywords
                }
            }
        } catch {
            Write-Host "`t[DEBUG] Error en SelectionChanged: $_" -ForegroundColor Red
        }
    }.GetNewClosure()

    # Preparar el objeto para almacenar referencias
    $comboData = New-Object PSObject -Property @{
        RichTextBox = $RichTextBox
        Queries     = $Queries
        SqlKeywords = $sqlKeywords
    }

    # Asignar el objeto como Tag del ComboBox
    $ComboQueries.Tag = $comboData

    # Agregar propiedades al ComboBox
    $ComboQueries | Add-Member -NotePropertyName Queries -NotePropertyValue $Queries -Force
    $ComboQueries | Add-Member -NotePropertyName SqlKeywords -NotePropertyValue $sqlKeywords -Force
    $ComboQueries | Add-Member -NotePropertyName RichTextBox -NotePropertyValue $RichTextBox -Force

    # Asignar el evento con el contexto correcto
    $ComboQueries.Add_SelectionChanged($selectionChangedScript)

    # Evento para resaltado en tiempo real - versión simplificada que evita el error
    $textChangedScript = {
        param($sender, $e)

        # Solo procesar si no estamos en medio de otro resaltado
        if (-not $global:isHighlightingQuery) {
            try {
                $global:isHighlightingQuery = $true
                $keywords = 'ADD|ALL|ALTER|AND|ANY|AS|ASC|AUTHORIZATION|BACKUP|BETWEEN|BIGINT|BINARY|BIT|BY|CASE|CHECK|COLUMN|CONSTRAINT|CREATE|CROSS|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|DATABASE|DEFAULT|DELETE|DESC|DISTINCT|DROP|EXEC|EXECUTE|EXISTS|FOREIGN|FROM|FULL|FUNCTION|GROUP|HAVING|IN|INDEX|INNER|INSERT|INT|INTO|IS|JOIN|KEY|LEFT|LIKE|LIMIT|NOT|NULL|ON|OR|ORDER|OUTER|PRIMARY|PROCEDURE|REFERENCES|RETURN|RIGHT|ROWNUM|SELECT|SET|SMALLINT|TABLE|TOP|TRUNCATE|UNION|UNIQUE|UPDATE|VALUES|VIEW|WHERE|WITH|RESTORE'

                if ([string]::IsNullOrEmpty($keywords)) {
                    Write-Host "`t[DEBUG] Advertencia: Keywords está vacío" -ForegroundColor Yellow
                    return
                }

                Set-WpfSqlHighlighting -RichTextBox $sender -Keywords $keywords
            } catch {
                Write-Host "`t[DEBUG] Error en TextChanged: $_" -ForegroundColor Red
            } finally {
                $global:isHighlightingQuery = $false
            }
        }
    }

    # Asignar el evento
    $RichTextBox.Add_TextChanged($textChangedScript)
}

function Set-WpfSqlHighlighting {
    param(
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [string]$Keywords
    )

    if ($null -eq $RichTextBox -or $null -eq $RichTextBox.Document) {
        Write-DzDebug "`t[DEBUG] RichTextBox o Document es null" -Color Red
        return
    }

    $script:isHighlightingQuery = $true
    try {
        $theme = Get-DzUiTheme
        $defaultBrush = Get-DzThemeBrush -Hex $theme.ControlForeground -Fallback ([System.Windows.Media.Brushes]::Black)
        $commentBrush = Get-DzThemeBrush -Hex $theme.AccentMuted -Fallback ([System.Windows.Media.Brushes]::DarkGreen)
        $keywordBrush = Get-DzThemeBrush -Hex $theme.AccentPrimary -Fallback ([System.Windows.Media.Brushes]::Blue)

        $textRange = New-Object System.Windows.Documents.TextRange(
            $RichTextBox.Document.ContentStart,
            $RichTextBox.Document.ContentEnd
        )
        $textRange.ApplyPropertyValue(
            [System.Windows.Documents.TextElement]::ForegroundProperty,
            $defaultBrush
        )

        $text = $textRange.Text
        $commentRanges = @()

        # Procesar comentarios de línea
        foreach ($c in [regex]::Matches($text, '--.*', 'Multiline')) {
            $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $c.Index
            $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($c.Index + $c.Length)
            if ($start -ne $null -and $end -ne $null) {
                (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
                    [System.Windows.Documents.TextElement]::ForegroundProperty,
                    $commentBrush
                )
                $commentRanges += [PSCustomObject]@{ Start = $c.Index; End = $c.Index + $c.Length }
            }
        }

        # Procesar comentarios de bloque
        foreach ($b in [regex]::Matches($text, '/\*[\s\S]*?\*/', 'Multiline')) {
            $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $b.Index
            $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($b.Index + $b.Length)
            if ($start -ne $null -and $end -ne $null) {
                (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
                    [System.Windows.Documents.TextElement]::ForegroundProperty,
                    $commentBrush
                )
                $commentRanges += [PSCustomObject]@{ Start = $b.Index; End = $b.Index + $b.Length }
            }
        }

        # Procesar palabras clave SQL
        foreach ($m in [regex]::Matches($text, "\b($Keywords)\b", 'IgnoreCase')) {
            $inComment = $commentRanges | Where-Object { $m.Index -ge $_.Start -and $m.Index -lt $_.End }
            if (-not $inComment) {
                $start = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset $m.Index
                $end = Get-TextPointerFromOffset -RichTextBox $RichTextBox -Offset ($m.Index + $m.Length)
                if ($start -ne $null -and $end -ne $null) {
                    (New-Object System.Windows.Documents.TextRange($start, $end)).ApplyPropertyValue(
                        [System.Windows.Documents.TextElement]::ForegroundProperty,
                        $keywordBrush
                    )
                }
            }
        }
    } catch {
        Write-DzDebug "`t[DEBUG] Error en Set-WpfSqlHighlighting: $_" -Color Red
    } finally {
        $script:isHighlightingQuery = $false
    }
}
function Get-TextPointerFromOffset {
    param(
        [Parameter(Mandatory = $true)]
        $RichTextBox,
        [Parameter(Mandatory = $true)]
        [int]$Offset
    )

    if ($null -eq $RichTextBox -or $null -eq $RichTextBox.Document) {
        return $null
    }

    $pointer = $RichTextBox.Document.ContentStart
    $count = 0

    while ($null -ne $pointer) {
        if ($pointer.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward) -eq
            [System.Windows.Documents.TextPointerContext]::Text) {

            $textRun = $pointer.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
            if ($count + $textRun.Length -ge $Offset) {
                return $pointer.GetPositionAtOffset($Offset - $count)
            }
            $count += $textRun.Length
        }

        $nextPointer = $pointer.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
        if ($nextPointer -eq $pointer) {
            break
        }
        $pointer = $nextPointer
    }

    return $RichTextBox.Document.ContentEnd
}
Export-ModuleMember -Function @(
    'Get-PredefinedQueries',
    'Initialize-PredefinedQueries',
    'Remove-SqlComments',
    'Set-WpfSqlHighlighting',
    'Get-TextPointerFromOffset'
)
