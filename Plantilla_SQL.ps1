# Crear la carpeta 'C:\Temp' si no existe
if (!(Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    Write-Host "Carpeta 'C:\Temp' creada correctamente."
}
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
# Crear el formulario
    $formPrincipal = New-Object System.Windows.Forms.Form
    $formPrincipal.Size = New-Object System.Drawing.Size(500, 460)
    $formPrincipal.StartPosition = "CenterScreen"
    $formPrincipal.BackColor = [System.Drawing.Color]::White
    $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formPrincipal.MaximizeBox = $false
    $formPrincipal.MinimizeBox = $false
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                                                                                                        $version = "Alfa 250201.2359"  # Valor predeterminado para la versión
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
            [string]$ToolTipText = $null  # Nuevo parámetro para el ToolTip
        )
        # Pasar esto a los parametros de arriba.
        $buttonStyle = @{
            Size      = New-Object System.Drawing.Size(220, 35)
            FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
            Font      = $defaultFont
        }
        # Eventos de mouse (definidos dentro de la función)
        $button_MouseEnter = {
            $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)  # Cambia el color al pasar el mouse
            $this.Font = $boldFont
        }
        $button_MouseLeave = {
            $this.BackColor = $this.Tag  # Restaura el color original almacenado en Tag
            $this.Font = $defaultFont
        }
        # Crear el botón
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = $buttonStyle.Size
        $button.Location = $Location
        $button.BackColor = $BackColor
        $button.ForeColor = $ForeColor
        $button.Font = $buttonStyle.Font
        $button.FlatStyle = $buttonStyle.FlatStyle
        $button.Tag = $BackColor  # Almacena el color original en Tag
        $button.Add_MouseEnter($button_MouseEnter)
        $button.Add_MouseLeave($button_MouseLeave)
        # Agregar ToolTip si se proporciona
        if ($ToolTipText) {
            $toolTip.SetToolTip($button, $ToolTipText)
        }
        return $button
    }
# Crear las pestañas (TabControl)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(480, 300) #X,Y
    $tabControl.Location = New-Object System.Drawing.Point(0,0)
    $tabControl.BackColor = [System.Drawing.Color]::LightGray
# Crear las tres pestañas (Aplicaciones, Consultas y Pro)
    $tabAplicaciones = New-Object System.Windows.Forms.TabPage
    $tabAplicaciones.Text = "Aplicaciones"
    $tabProSql = New-Object System.Windows.Forms.TabPage
    $tabProSql.Text = "Pro"
# Añadir las pestañas al TabControl
    $tabControl.TabPages.Add($tabAplicaciones)
    $tabControl.TabPages.Add($tabProSql)
# Crear los botones utilizando la función
    $btnInstallSQLManagement = Create-Button -Text "Instalar Management2014" -Location (New-Object System.Drawing.Point(10, 10)) -ToolTip "Instalación mediante choco de SQL Management 2014."
    $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 50)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 90)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 130)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "De momento solo si es SQL 2014."
    $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 170)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Busca SQL Management en tu equipo y te confirma la versión previo a ejecutarlo."
    $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 210)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Herramienta de Star con funciones multiples para impresoras POS."
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(240, 10)) -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(240, 50)) -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(240, 90)) -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(240, 130)) -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
    $btnConfigurarIPs = Create-Button -Text "Configurar IPs" -Location (New-Object System.Drawing.Point(240, 170)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
    $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 40)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))
    $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(240, 40)) -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))
    $btnDisconnectDb.Enabled = $false  # Deshabilitado inicialmente
    $btnReviewPivot = Create-Button -Text "Revisar Pivot Table" -Location (New-Object System.Drawing.Point(10, 110))
    $btnReviewPivot.Enabled = $false  # Deshabilitado inicialmente
    $btnFechaRevEstaciones = Create-Button -Text "Fecha de revisiones" -Location (New-Object System.Drawing.Point(10, 150))
    $btnFechaRevEstaciones.Enabled = $false  # Deshabilitado inicialmente
    $btnRespaldarRestcard = Create-Button -Text "Respaldar restcard" -Location (New-Object System.Drawing.Point(10, 190))
    $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(120, 310)) -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
# Crear el CheckBox chkSqlServer
    $chkSqlServer = New-Object System.Windows.Forms.CheckBox
    $chkSqlServer.Text = "Instalar SQL Tools (opcional)"
    $chkSqlServer.Size = New-Object System.Drawing.Size(290, 30)
    $chkSqlServer.Location = New-Object System.Drawing.Point(10, 10)
# Label para mostrar conexión a la base de datos
    $lblConnectionStatus = New-Object System.Windows.Forms.Label
    $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
    $lblConnectionStatus.Font = $defaultFont
    $lblConnectionStatus.Size = New-Object System.Drawing.Size(290, 30)
    $lblConnectionStatus.Location = New-Object System.Drawing.Point(10, 250)
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::RED
# Crear el Label para mostrar el nombre del equipo fuera de las pestañas
    $lblHostname = New-Object System.Windows.Forms.Label
    $lblHostname.Text = [System.Net.Dns]::GetHostName()
    $lblHostname.Size = New-Object System.Drawing.Size(240, 35)
    $lblHostname.Font = $defaultFont
    $lblHostname.Location = New-Object System.Drawing.Point(2, 350)
    $lblHostname.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblHostname.BackColor = [System.Drawing.Color]::White
    $lblHostname.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $lblHostname.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
    $toolTipHostname = New-Object System.Windows.Forms.ToolTip
    $toolTipHostname.SetToolTip($lblHostname, "Haz clic para copiar el Hostname al portapapeles.")
# Crear el Label para mostrar el puerto
    $lblPort = New-Object System.Windows.Forms.Label
    $lblPort.Size = New-Object System.Drawing.Size(236, 35)
    $lblPort.Font = $defaultFont
    $lblPort.Location = New-Object System.Drawing.Point(245, 350)  # Alineado a la derecha del hostname
    $lblPort.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblPort.BackColor = [System.Drawing.Color]::White
    $lblPort.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $lblPort.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
    $toolTip2 = New-Object System.Windows.Forms.ToolTip
    $toolTip2.SetToolTip($lblPort, "Haz clic para copiar el Puerto al portapapeles.")
# Crear el Label para mostrar las IPs y adaptadores
    $lbIpAdress = New-Object System.Windows.Forms.Label
    $lbIpAdress.Size = New-Object System.Drawing.Size(240, 100)  # Tamaño inicial
    $lbIpAdress.Font = $defaultFont
    $lbIpAdress.Location = New-Object System.Drawing.Point(2, 390)
    $lbIpAdress.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $lbIpAdress.BackColor = [System.Drawing.Color]::White
    $lbIpAdress.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $lbIpAdress.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
# Crear el ToolTip
    $toolTip2.SetToolTip($lbIpAdress, "Haz clic para copiar las IPs al portapapeles.")
# Agregar botones a la pestaña de aplicaciones
    $tabAplicaciones.Controls.Add($btnInstallSQLManagement)
    $tabAplicaciones.Controls.Add($btnProfiler)
    $tabAplicaciones.Controls.Add($btnDatabase)
    $tabAplicaciones.Controls.Add($btnSQLManager)
    $tabAplicaciones.Controls.Add($btnSQLManagement)
    $tabAplicaciones.Controls.Add($btnClearPrintJobs)
    $tabAplicaciones.Controls.Add($btnClearAnyDesk)
    $tabAplicaciones.Controls.Add($btnShowPrinters)
    $tabAplicaciones.Controls.Add($btnPrinterTool)
    $tabAplicaciones.Controls.Add($btnAplicacionesNS)
    $tabAplicaciones.Controls.Add($btnConfigurarIPs)
# Agregar controles a la pestaña Pro
    $tabProSql.Controls.Add($chkSqlServer)
    $tabProSql.Controls.Add($btnReviewPivot)
    $tabProSql.Controls.Add($btnRespaldarRestcard)
    $tabProSql.Controls.Add($btnFechaRevEstaciones)
    $tabProSql.Controls.Add($lblConnectionStatus)
    $tabProSql.Controls.Add($btnConnectDb)
    $tabProSql.Controls.Add($btnDisconnectDb)
#Funcion para copiar el puerto al portapapeles
$lblPort.Add_Click({
    if ($lblPort.Text -match "\d+") {  # Asegurarse de que el texto es un número
        $port = $matches[0]  # Extraer el número del texto
        [System.Windows.Forms.Clipboard]::SetText($port)
        Write-Host "Puerto copiado al portapapeles: $port" -ForegroundColor Green
    } else {
        Write-Host "El texto del Label del puerto no contiene un número válido para copiar." -ForegroundColor Red
    }
})
$lblHostname.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($lblHostname.Text)
        Write-Host "`nNombre del equipo copiado al portapapeles: $($lblHostname.Text)"
    })
$lbIpAdress.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($lbIpAdress.Text)
        Write-Host "`nIP's copiadas al equipo: $($lbIpAdress.Text)"
    })
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
    # Formatear las IPs en el formato requerido para el portapapeles (ip1, ip2, ip3, etc.)
    $ipsTextForClipboard = ($ipsWithAdapters | ForEach-Object {
        $_.IPAddress
    }) -join ", "
    # Copiar las IPs al portapapeles en el formato adecuado
    [System.Windows.Forms.Clipboard]::SetText($ipsTextForClipboard)
    # Construir el texto para mostrar en el Label
    $ipsTextForLabel = ($ipsWithAdapters | ForEach-Object {
        "Adaptador: $($_.AdapterName)`nIP: $($_.IPAddress)`n"
    }) -join "`n"
    # Asignar el texto al label
    $lbIpAdress.Text = "$ipsTextForLabel"
} else {
    $lbIpAdress.Text = "No se encontraron direcciones IP"
}
# Configuración dinámica del tamaño del Label según la cantidad de líneas
    $lineHeight = 15
    $maxLines = $lbIpAdress.Text.Split("`n").Count
    $labelHeight = [Math]::Min(400, $lineHeight * $maxLines)
    $lbIpAdress.Size = New-Object System.Drawing.Size(240, $labelHeight)
# Ajustar la altura del formulario según el Label de IPs
    $formHeight = $formPrincipal.Size.Height + $labelHeight - 20
    $formPrincipal.Size = New-Object System.Drawing.Size($formPrincipal.Size.Width, $formHeight)
# Función para obtener adaptadores y sus estados (modificada)
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
# Función para cambiar el estado de la red
function Set-NetworkCategory {
    param (
        [string]$category,
        [int]$interfaceIndex,
        [System.Windows.Forms.Label]$label
    )
    # Obtener el estado anterior
    $profile = Get-NetConnectionProfile | Where-Object { $_.InterfaceIndex -eq $interfaceIndex }
    $previousCategory = if ($profile) { $profile.NetworkCategory } else { "Desconocido" }
    # Cambiar el tipo de red
    if ($category -eq "Privado") {
        Set-NetConnectionProfile -InterfaceIndex $interfaceIndex -NetworkCategory Private
        Write-Host "Estado cambiado a Privado."
        $label.ForeColor = [System.Drawing.Color]::Green
        $label.Text = "$($label.Text.Split(' - ')[0]) - Privado"  # Actualizar el texto de la etiqueta
    } elseif ($category -eq "Público") {
        Set-NetConnectionProfile -InterfaceIndex $interfaceIndex -NetworkCategory Public
        Write-Host "Estado cambiado a Público."
        $label.ForeColor = [System.Drawing.Color]::Red
        $label.Text = "$($label.Text.Split(' - ')[0]) - Público"  # Actualizar el texto de la etiqueta
    }
    # Mostrar el cambio en consola con categorías en español
    Write-Host "Categoría anterior: $previousCategory"
    Write-Host "Categoría nueva: $category"
}
# Crear la etiqueta para mostrar los adaptadores y su estado
$lblPerfilDeRed = New-Object System.Windows.Forms.Label
$lblPerfilDeRed.Text = "Estado de los Adaptadores:"
$lblPerfilDeRed.Size = New-Object System.Drawing.Size(236, 35)
$lblPerfilDeRed.Location = New-Object System.Drawing.Point(245, 390)
# Llenar el contenido de la etiqueta con el nombre del adaptador y su estado
$networkAdapters = Get-NetworkAdapterStatus
$adapterInfo = ""
# Usamos un contador para ubicar los labels
$index = 0
foreach ($adapter in $networkAdapters) {
    $text = ""
    $color = [System.Drawing.Color]::Green

    if ($adapter.NetworkCategory -eq "Private") {
        $text = "$($adapter.AdapterName) - Privado"
        $color = [System.Drawing.Color]::Green
    } elseif ($adapter.NetworkCategory -eq "Public") {
        $text = "$($adapter.AdapterName) - Público"
        $color = [System.Drawing.Color]::Red
    }
    # Crear un Label con la palabra "Público" o "Privado" clickeable
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.ForeColor = $color
    $label.Cursor = [System.Windows.Forms.Cursors]::Hand
    $label.Size = New-Object System.Drawing.Size(236, 20)
    # Ajustar la ubicación para las etiquetas
    $label.Location = New-Object System.Drawing.Point(245, (390 + (30 * $index)))  # Ajustar el desplazamiento de acuerdo con el índice
    # Evento para manejar el clic
    $label.Add_Click({
        # Asegurarse de que se maneje el idioma correctamente
        $category = if ($adapter.NetworkCategory -eq "Private") { "Público" } else { "Privado" }
        
        # Confirmar el cambio y llamar a la función de cambio
        Write-Host "Deseas cambiar el estado a $category?"  # Mostrar en consola la solicitud
        $result = [System.Windows.Forms.MessageBox]::Show("¿Deseas cambiar el estado a $category?", "Confirmar cambio", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Set-NetworkCategory -category $category -interfaceIndex $adapter.InterfaceIndex -label $label
        }
    })
    $adapterInfo += $label.Text + "`n"
    $formPrincipal.Controls.Add($label)
    # Incrementar el índice para la siguiente posición del label
    $index++
}
# Agregar los controles al formulario
    $formPrincipal.Controls.Add($tabControl)
    $formPrincipal.Controls.Add($lblHostname)
    $formPrincipal.Controls.Add($lblPort)
    $formPrincipal.Controls.Add($lbIpAdress)
    $formPrincipal.Controls.Add($lblPerfilDeRed)
    $formPrincipal.Controls.Add($btnExit)

# Acción para el CheckBox, si el usuario lo marca manualmente
$chkSqlServer.Add_CheckedChanged({
    if ($chkSqlServer.Checked) {
        # Confirmación con MsgBox
        $result = [System.Windows.Forms.MessageBox]::Show("¿Estás seguro que deseas instalar las herramientas de SQL Server?", "Confirmar instalación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "`nInstalando herramientas de SQL Server..."
            $chkSqlServer.Enabled = $false
            $installedSql = Check-SqlServerInstallation
            if ($installedSql) {
                $chkSqlServer.Checked = $true
                $chkSqlServer.Enabled = $false  # Deshabilitar la edición si SQL Server está instalado
            } else {
                $chkSqlServer.Checked = $false
            }
            Install-Module -Name SqlServer -Force -Scope CurrentUser -AllowClobber
            Write-Host "`nHerramientas de SQL Server instaladas."
        } else {
            # Desmarcar el checkbox si el usuario cancela
            $chkSqlServer.Checked = $false
            Write-Host "`nInstalación cancelada."
        }
    } else {
        # Si el usuario desmarca el checkbox, se habilita para futuras instalaciones
        $chkSqlServer.Enabled = $true
    }
    # Habilitar el botón de Conectar a BDD solo si las herramientas SQL Server están habilitadas
    $btnConnectDb.Enabled = $chkSqlServer.Checked
})
# Obtener el puerto de SQL Server desde el registro
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\NATIONALSOFT\MSSQLServer\SuperSocketNetLib\Tcp"
        $tcpPort = Get-ItemProperty -Path $regKeyPath -Name "TcpPort" -ErrorAction SilentlyContinue

        if ($tcpPort -and $tcpPort.TcpPort) {
            $lblPort.Text = "Puerto SQL \NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $lblPort.Text = "No se encontró puerto o instancia."
        }
##-------------------- FUNCIONES                                                          -------#
function Execute-SqlQuery {
    param (
        [string]$server,
        [string]$database,
        [string]$query
    )
    try {
        # Cadena de conexión
        $connectionString = "Server=$server;Database=$database;User Id=sa;Password=$($global:password);"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()
        # Ejecutar consulta
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $reader = $command.ExecuteReader()
        # Obtener los nombres de las columnas
        $columns = @()
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $columns += $reader.GetName($i)
        }
        # Leer los resultados
        $results = @()
        while ($reader.Read()) {
            $row = @{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $row[$columns[$i]] = $reader[$i]
            }
            $results += $row
        }
        # Cerrar la conexión y liberar recursos
        $connection.Close()
        $connection.Dispose()
        return $results
    } catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
}
function Show-ResultsConsole {
    param (
        [string]$query
    )
    try {
        # Ejecutar la consulta y obtener los resultados
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query

        if ($results.Count -gt 0) {
            # Mostrar los resultados de la consulta
            $columns = $results[0].Keys
            $columnWidths = @{}
            foreach ($col in $columns) {
                $columnWidths[$col] = $col.Length
            }
            # Mostrar los encabezados
            $header = ""
            foreach ($col in $columns) {
                $header += $col.PadRight($columnWidths[$col] + 4)
            }
            Write-Host $header
            Write-Host ("-" * $header.Length)
            # Mostrar las filas de resultados
            foreach ($row in $results) {
                $rowText = ""
                foreach ($col in $columns) {
                    $rowText += ($row[$col].ToString()).PadRight($columnWidths[$col] + 4)
                }
                Write-Host $rowText
            }
        } else {
            Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
}
##---------------OTROS BOTONES Y FUNCIONES OMITIDAS AQUI----------------------------------------------------------------BOTONES#
$btnConnectDb.Add_Click({
        # Crear el formulario para pedir los datos de conexión
        $connectionForm = New-Object System.Windows.Forms.Form
        $connectionForm.Text = "Conexión a SQL Server"
        $connectionForm.Size = New-Object System.Drawing.Size(400, 200)
        $connectionForm.StartPosition = "CenterScreen"
        $connectionForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $connectionForm.MaximizeBox = $false
        $connectionForm.MinimizeBox = $false

        # Crear las etiquetas y cajas de texto
        $labelProfile = New-Object System.Windows.Forms.Label
        $labelProfile.Text = "Perfil de conexión"
        $labelProfile.Location = New-Object System.Drawing.Point(10, 20)
        $labelProfile.Size = New-Object System.Drawing.Size(100, 20)

        $cmbProfiles = New-Object System.Windows.Forms.ComboBox
        $cmbProfiles.Location = New-Object System.Drawing.Point(120, 20)
        $cmbProfiles.Size = New-Object System.Drawing.Size(250, 20)
        $cmbProfiles.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

        # Cargar archivos INI desde las rutas especificadas
        $profiles = @{ }
        $iniFiles = @(
            "C:\NationalSoft\OnTheMinute4.5\checadorsql.ini",
            "C:\NationalSoft\Softrestaurant11.0\INIS\*.ini",
            "C:\NationalSoft\Softrestaurant10.0\INIS\*.ini"
        )

            foreach ($path in $iniFiles) {
                $files = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                foreach ($file in $files) {
                    # Obtener la ruta de las dos carpetas anteriores
                    $relativePath = $file.DirectoryName -replace "^.*\\([^\\]+\\[^\\]+)\\.*$", '$1'
        
                    # Concatenar la ruta relativa con el nombre del archivo
                    $profileName = "$relativePath\$($file.Name)"
        
                    # Agregar el perfil a la lista
                    $profiles[$profileName] = $file.FullName
                    $cmbProfiles.Items.Add($profileName)
                }
            }

        # Agregar opción "Personalizado" si no existe ya en la lista
        if (-not $cmbProfiles.Items.Contains("Personalizado")) {
            $cmbProfiles.Items.Add("Personalizado")
        }
        # Crear las demás etiquetas y campos de texto
        $labelServer = New-Object System.Windows.Forms.Label
        $labelServer.Text = "Servidor SQL"
        $labelServer.Location = New-Object System.Drawing.Point(10, 50)
        $labelServer.Size = New-Object System.Drawing.Size(100, 20)
    
        $txtServer = New-Object System.Windows.Forms.TextBox
        $txtServer.Location = New-Object System.Drawing.Point(120, 50)
        $txtServer.Size = New-Object System.Drawing.Size(250, 20)

        $labelDatabase = New-Object System.Windows.Forms.Label
        $labelDatabase.Text = "Base de Datos"
        $labelDatabase.Location = New-Object System.Drawing.Point(10, 80)
        $labelDatabase.Size = New-Object System.Drawing.Size(100, 20)

        $txtDatabase = New-Object System.Windows.Forms.TextBox
        $txtDatabase.Location = New-Object System.Drawing.Point(120, 80)
        $txtDatabase.Size = New-Object System.Drawing.Size(250, 20)

        $labelPassword = New-Object System.Windows.Forms.Label
        $labelPassword.Text = "Contraseña"
        $labelPassword.Location = New-Object System.Drawing.Point(10, 110)
        $labelPassword.Size = New-Object System.Drawing.Size(100, 20)

        $txtPassword = New-Object System.Windows.Forms.TextBox
        $txtPassword.Location = New-Object System.Drawing.Point(120, 110)
        $txtPassword.Size = New-Object System.Drawing.Size(250, 20)
        $txtPassword.UseSystemPasswordChar = $true

    # Habilitar el botón "Conectar" si la contraseña tiene al menos un carácter
    $txtPassword.Add_TextChanged({
        if ($txtPassword.Text.Length -ge 1) {
            $btnOK.Enabled = $true
        } else {
            $btnOK.Enabled = $false
        }
    })
        # Manejar selección del ComboBox
        $cmbProfiles.Add_SelectedIndexChanged({
            if ($cmbProfiles.SelectedItem -eq "Personalizado") {
                $txtServer.Clear()
                $txtDatabase.Clear()
            } else {
                $selectedFile = $profiles[$cmbProfiles.SelectedItem]
                if (Test-Path $selectedFile) {
                    $content = Get-Content -Path $selectedFile -ErrorAction SilentlyContinue
                    if ($content.Length -gt 0) {
                        # Inicializar valores por defecto
                        $server = "No especificado"
                        $database = "No especificado"

                        # Obtener la primera coincidencia de DataSource y Catalog usando Select-String
                        $dataSourceMatch = $content | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1
                        if ($dataSourceMatch) {
                            $server = $dataSourceMatch.Matches.Groups[1].Value
                        }

                        $catalogMatch = $content | Select-String -Pattern "^Catalog=(.*)" | Select-Object -First 1
                        if ($catalogMatch) {
                            $database = $catalogMatch.Matches.Groups[1].Value
                        }

                        # Verificar que los valores sean correctos antes de asignarlos
                        if ($server -ne "No especificado") {
                            $txtServer.Text = $server
                        }

                        if ($database -ne "No especificado") {
                            $txtDatabase.Text = $database
                        }
                    }
                }
            }
        })

        # Crear el botón para conectar
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "Conectar"
        $btnOK.Size = New-Object System.Drawing.Size(100, 30)
        $btnOK.Location = New-Object System.Drawing.Point(150, 140)  # Ajusta la posición según necesites
        $btnOK.Enabled = $false  # Deshabilitar el botón inicialmente
    # Variables globales para guardar la información de conexión
    $global:server
    $global:database
    $global:password
    $global:connection  # Variable global para la conexión
$btnOK.Add_Click({
        try {
            # Cadena de conexión
            $connectionString = "Server=$($txtServer.Text);Database=$($txtDatabase.Text);User Id=sa;Password=$($txtPassword.Text);"
            $global:connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)  # Asignar la conexión a la variable global
            $global:connection.Open()

            Write-Host "`nConexión exitosa" -ForegroundColor Green

            # Guardar la información de conexión en variables globales
            $global:server = $txtServer.Text
            $global:database = $txtDatabase.Text
            $global:password = $txtPassword.Text

            # Cerrar la ventana de conexión
            $connectionForm.Close()

            # Actualizar el texto del label de conexión
            $lblConnectionStatus.Text = "Conectado a BDD: $($txtDatabase.Text)"
            $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green

            # Habilitar o deshabilitar botones cuando hay conexiones existosas
            $btnReviewPivot.Enabled = $true
            $btnFechaRevEstaciones.Enabled = $true
            $btnConnectDb.Enabled = $false
            $btnDisconnectDb.Enabled = $true

        } catch {
            Write-Host "`nError de conexión: $_" -ForegroundColor Red
            $lblConnectionStatus.Text = "Conexión fallida"
        }
    })
        # Agregar los controles al formulario
        $connectionForm.Controls.Add($labelProfile)
        $connectionForm.Controls.Add($cmbProfiles)
        $connectionForm.Controls.Add($labelServer)
        $connectionForm.Controls.Add($txtServer)
        $connectionForm.Controls.Add($labelDatabase)
        $connectionForm.Controls.Add($txtDatabase)
        $connectionForm.Controls.Add($labelPassword)
        $connectionForm.Controls.Add($txtPassword)
        $connectionForm.Controls.Add($btnOK)
        $connectionForm.ShowDialog()
    })
$btnDisconnectDb.Add_Click({
    try {
        # Cerrar la conexión
        $global:connection.Close()
        Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow

        # Restaurar el label al estado original
        $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red

        # Habilitar el botón de conectar y deshabilitar el de desconectar
        $btnConnectDb.Enabled = $true
        $btnDisconnectDb.Enabled = $false
        $btnFechaRevEstaciones.Enabled = $false
        $btnReviewPivot.Enabled = $false
    } catch {
        Write-Host "`nError al desconectar: $_" -ForegroundColor Red
    }
})
$btnReviewPivot.Add_Click({
    try {
        if (-not $global:server -or -not $global:database -or -not $global:password) {
            Write-Host "`nNo hay una conexión válida." -ForegroundColor Red
            return
        }

        # Consultas SQL
        $query1 = "SELECT field, COUNT(*) FROM app_settings GROUP BY field HAVING COUNT(*) > 1;"
        $query2 = "SELECT * FROM app_settings where field = 'activar_restcard_srm';"

        # Ejecutar y analizar la primera consulta
        $resultsQuery1 = Execute-SqlQuery -server $global:server -database $global:database -query $query1

        if ($resultsQuery1.Count -gt 0) {
            # Si hay duplicados, mostrar mensaje en rojo y ejecutar query2
            Write-Host "`nSELECT field, COUNT(*) FROM app_settings GROUP BY field HAVING COUNT(*) > 1;" -ForegroundColor white
            Write-Host "`nSe recomienda correr proceso de limpieza pivot table" -ForegroundColor Red
            # Ejecutar la segunda consulta
            Show-ResultsConsole -query $query2
        } else {
            # Si no hay duplicados, mostrar mensaje en verde
            Write-Host "`nSELECT field, COUNT(*) FROM app_settings GROUP BY field HAVING COUNT(*) > 1;" -ForegroundColor white
            Write-Host "`nNo hay duplicados (pivot table)" -ForegroundColor Green
        }

    } catch {
        Write-Host "`nError al ejecutar consulta: $($_.Exception.Message)" -ForegroundColor Red
    }
})
$btnFechaRevEstaciones.Add_Click({
    try {
        if (-not $global:server -or -not $global:database -or -not $global:password) {
            Write-Host "`nNo hay una conexión válida." -ForegroundColor Red
            return
        }
        # Consultas SQL
$query1 = "SELECT e.FECHAREV, 
                   b.estacion as Estacion, 
                   CONVERT(varchar, b.fecha, 23) AS UltimaUso
            FROM bitacorasistema b
            INNER JOIN (
                SELECT estacion, MAX(fecha) AS max_fecha
                FROM bitacorasistema
                GROUP BY estacion
            ) latest_bitacora 
                ON b.estacion = latest_bitacora.estacion 
                AND b.fecha = latest_bitacora.max_fecha
            INNER JOIN estaciones e 
                ON b.estacion = e.idestacion
            ORDER BY b.fecha DESC;"
        # Ejecutar y analizar la primera consulta
        $resultsQuery1 = Execute-SqlQuery -server $global:server -database $global:database -query $query1
Write-Host "`nSELECT e.FECHAREV, ` 
            b.estacion as Estacion, `
            CONVERT(varchar, b.fecha, 23) AS UltimaUso, `
            FROM bitacorasistema b `
            INNER JOIN (SELECT estacion, MAX(fecha) AS max_fecha FROM bitacorasistema GROUP BY estacion) latest_bitacora `
            ON b.estacion = latest_bitacora.estacion AND b.fecha = latest_bitacora.max_fecha `
            INNER JOIN estaciones e ON b.estacion = e.idestacion ORDER BY b.fecha DESC;`n" -ForegroundColor Yellow
                    Show-ResultsConsole -query $query1
    } catch {
        Write-Host "`nError al ejecutar consulta: $($_.Exception.Message)" -ForegroundColor Red
    }
})
$btnExit.Add_Click({
    $formPrincipal.Dispose()
    $formPrincipal.Close()
})
$formPrincipal.Refresh()
# Mostrar el formulario principal
$formPrincipal.ShowDialog()