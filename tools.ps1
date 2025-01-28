# Crear la carpeta 'C:\Temp' si no existe
if (!(Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    Write-Host "Carpeta 'C:\Temp' creada correctamente."
}

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

# Crear el formulario
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(500, 460)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
# Crear un TextBox para ingresar la versión manualmente
                                                                $version = "Alfa 250128.1312"  # Valor predeterminado para la versión
$form.Text = "Daniel Tools v$version"

Write-Host "`n=============================================" -ForegroundColor DarkCyan
Write-Host "       Daniel Tools - Suite de Utilidades       " -ForegroundColor Green
Write-Host "              Versión: v$($version)               " -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor DarkCyan

Write-Host "`nTodos los derechos reservados para Daniel Tools." -ForegroundColor Cyan
Write-Host "Para reportar errores o sugerencias, contacte vía Teams." -ForegroundColor Cyan

# Crear un estilo base para los botones
    $buttonStyle = New-Object System.Windows.Forms.Button
    $buttonStyle.Size = New-Object System.Drawing.Size(220, 35)
    $buttonStyle.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonStyle.BackColor = [System.Drawing.Color]::LightGray
    $buttonStyle.ForeColor = [System.Drawing.Color]::Black
    $buttonStyle.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)

# Crear las pestañas (TabControl)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(480, 300) #X,Y
    $tabControl.Location = New-Object System.Drawing.Point(0,0)
    $tabControl.BackColor = [System.Drawing.Color]::LightGray


# Crear las tres pestañas (Aplicaciones, Consultas y Pro)
$tabAplicaciones = New-Object System.Windows.Forms.TabPage
$tabAplicaciones.Text = "Aplicaciones"
    #$tabConsultas = New-Object System.Windows.Forms.TabPage
    #$tabConsultas.Text = "Consultas"
$tabProSql = New-Object System.Windows.Forms.TabPage
$tabProSql.Text = "Pro"

# Añadir las pestañas al TabControl
$tabControl.TabPages.Add($tabAplicaciones)
    #$tabControl.TabPages.Add($tabConsultas)
$tabControl.TabPages.Add($tabProSql)

# Crear los botones dentro de la pestaña de "Aplicaciones"
    $btnInstallSQLManagement = New-Object System.Windows.Forms.Button
    $btnInstallSQLManagement.Text = "Instalar Management2014"
    $btnInstallSQLManagement.Size = $buttonStyle.Size
    $btnInstallSQLManagement.Location = New-Object System.Drawing.Point(10, 10)
#Profiler
    $btnProfiler = New-Object System.Windows.Forms.Button
    $btnProfiler.Text = "Ejecutar ExpressProfiler2.0"
    $btnProfiler.Size = $buttonStyle.Size
    $btnProfiler.Location = New-Object System.Drawing.Point(10, 50)
#Database
    $btnDatabase = New-Object System.Windows.Forms.Button
    $btnDatabase.Text = "Ejecutar Database4"
    $btnDatabase.Size = $buttonStyle.Size
    $btnDatabase.Location = New-Object System.Drawing.Point(10, 90)
#SqlManager
    $btnSQLManager = New-Object System.Windows.Forms.Button
    $btnSQLManager.Text = "Ejecutar Manager"
    $btnSQLManager.Size = $buttonStyle.Size
    $btnSQLManager.Location = New-Object System.Drawing.Point(10, 130)
#SqlManagement
    $btnSQLManagement = New-Object System.Windows.Forms.Button
    $btnSQLManagement.Text = "Ejecutar Management"
    $btnSQLManagement.Size = $buttonStyle.Size
    $btnSQLManagement.Location = New-Object System.Drawing.Point(10, 170)
# Anydesk
    $btnClearAnyDesk = New-Object System.Windows.Forms.Button
    $btnClearAnyDesk.Text = "Clear AnyDesk"
    $btnClearAnyDesk.Size = $buttonStyle.Size
    $btnClearAnyDesk.Location = New-Object System.Drawing.Point(240, 10)
    $btnClearAnyDesk.BackColor = [System.Drawing.Color]::FromArgb(255, 76, 76)  # Rojo claro (FF4C4C)
# Impresoras show
    $buttonShowPrinters = New-Object System.Windows.Forms.Button
    $buttonShowPrinters.Text = "Mostrar Impresoras"
    $buttonShowPrinters.Size = $buttonStyle.Size
    $buttonShowPrinters.Location = New-Object System.Drawing.Point(240, 50)
# cola de impresion
    $btnClearPrintJobs = New-Object System.Windows.Forms.Button
    $btnClearPrintJobs.Text = "Limpiar Impresiones y Reiniciar Cola"
    $btnClearPrintJobs.Size = $buttonStyle.Size
    $btnClearPrintJobs.Location = New-Object System.Drawing.Point(240, 90)
# PrinterTool
    $btnPrinterTool = New-Object System.Windows.Forms.Button
    $btnPrinterTool.Text = "Printer Tools"
    $btnPrinterTool.Size = $buttonStyle.Size
    $btnPrinterTool.Location = New-Object System.Drawing.Point(240, 130)
# aplicaciones ns
    $btnAplicacionesNS = New-Object System.Windows.Forms.Button
    $btnAplicacionesNS.Text = "Aplicaciones National Soft"
    $btnAplicacionesNS.Size = $buttonStyle.Size
    $btnAplicacionesNS.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 150)  # Naranja tenue (RGB: 200, 150, 100)
    $btnAplicacionesNS.Location = New-Object System.Drawing.Point(240, 170)
#Agregar botones a la de aplicaciones
    $tabAplicaciones.Controls.Add($btnInstallSQLManagement)
    $tabAplicaciones.Controls.Add($btnProfiler)
    $tabAplicaciones.Controls.Add($btnDatabase)
    $tabAplicaciones.Controls.Add($btnSQLManager)
    $tabAplicaciones.Controls.Add($btnSQLManagement)
    $tabAplicaciones.Controls.Add($btnClearPrintJobs)
    $tabAplicaciones.Controls.Add($btnClearAnyDesk)
    $tabAplicaciones.Controls.Add($buttonShowPrinters)
    $tabAplicaciones.Controls.Add($btnPrinterTool)
    $tabAplicaciones.Controls.Add($btnAplicacionesNS)
# Crear el CheckBox
    $chkSqlServer = New-Object System.Windows.Forms.CheckBox
    $chkSqlServer.Text = "Instalar SQL Tools (opcional)"
    $chkSqlServer.Size = New-Object System.Drawing.Size(290, 30)
    $chkSqlServer.Location = New-Object System.Drawing.Point(10, 10)
# Crear el Botón para conectar a la base de datos
    $btnConnectDb = New-Object System.Windows.Forms.Button
    $btnConnectDb.Text = "Conectar a BDD"
    $btnConnectDb.Size = $buttonStyle.Size
    $btnConnectDb.Location = New-Object System.Drawing.Point(10, 40)
    $btnConnectDb.Enabled = $true
#2
    $btnDisconnectDb = New-Object System.Windows.Forms.Button
    $btnDisconnectDb.Text = "Desconectar de BDD"
    $btnDisconnectDb.Size = $buttonStyle.Size
    $btnDisconnectDb.Location = New-Object System.Drawing.Point(240, 40)
    $btnDisconnectDb.Enabled = $false  # Deshabilitado inicialmente
# Crear el Botón para revisar Pivot Table
    $btnReviewPivot = New-Object System.Windows.Forms.Button
    $btnReviewPivot.Text = "Revisar Pivot Table"
    $btnReviewPivot.Size = $buttonStyle.Size
    $btnReviewPivot.Location = New-Object System.Drawing.Point(10, 110)
    $btnReviewPivot.Enabled = $false  # Deshabilitado inicialmente
# Crear el Botón para revisar Pivot Table
    $btnFechaRevEstaciones = New-Object System.Windows.Forms.Button
    $btnFechaRevEstaciones.Text = "Fecha de revisiones"
    $btnFechaRevEstaciones.Size = $buttonStyle.Size
    $btnFechaRevEstaciones.Location = New-Object System.Drawing.Point(10, 150)
    $btnFechaRevEstaciones.Enabled = $false  # Deshabilitado inicialmente
# Label para mostrar conexión a la base de datos
    $lblConnectionStatus = New-Object System.Windows.Forms.Label
    $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
    $lblConnectionStatus.Size = New-Object System.Drawing.Size(290, 30)
    $lblConnectionStatus.Location = New-Object System.Drawing.Point(10, 250)
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::RED
# Agregar controles a la pestaña Pro
    $tabProSql.Controls.Add($chkSqlServer)  # Agregar el CheckBox
    $tabProSql.Controls.Add($btnReviewPivot)  # Agregar el Botón para revisar Pivot Table
    $tabProSql.Controls.Add($btnFechaRevEstaciones)  
    $tabProSql.Controls.Add($lblConnectionStatus)  # Agregar el Label de estado de conexión
    $tabProSql.Controls.Add($btnConnectDb)  # Agregar el Botón para conectar
    $tabProSql.Controls.Add($btnDisconnectDb)
# Crear el botón "Salir" fuera de las pestañas
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Salir"
    $btnExit.Size = $buttonStyle.Size
    $btnExit.Location = New-Object System.Drawing.Point(120, 310)
    $btnExit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExit.BackColor = [System.Drawing.Color]::DarkGray
    $btnExit.ForeColor = [System.Drawing.Color]::White
# Crear el Label para mostrar el nombre del equipo fuera de las pestañas
    $labelHostname = New-Object System.Windows.Forms.Label
    $labelHostname.Text = [System.Net.Dns]::GetHostName()
    $labelHostname.Size = New-Object System.Drawing.Size(240, 35)
    $labelHostname.Location = New-Object System.Drawing.Point(2, 350)
    $labelHostname.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelHostname.BackColor = [System.Drawing.Color]::White
    $labelHostname.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $labelHostname.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
    $toolTipHostname = New-Object System.Windows.Forms.ToolTip
    $toolTipHostname.SetToolTip($labelHostname, "Haz clic para copiar el Hostname al portapapeles.")
# Crear el Label para mostrar el puerto
    $labelPort = New-Object System.Windows.Forms.Label
    $labelPort.Size = New-Object System.Drawing.Size(236, 35)
    $labelPort.Location = New-Object System.Drawing.Point(245, 350)  # Alineado a la derecha del hostname
    $labelPort.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $labelPort.BackColor = [System.Drawing.Color]::White
    $labelPort.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $labelPort.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($labelPort, "Haz clic para copiar el Puerto al portapapeles.")
# Crear el Label para mostrar las IPs y adaptadores
    $labelipADress = New-Object System.Windows.Forms.Label
    $labelipADress.Size = New-Object System.Drawing.Size(240, 100)  # Tamaño inicial
    $labelipADress.Location = New-Object System.Drawing.Point(2, 390)
    $labelipADress.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $labelipADress.BackColor = [System.Drawing.Color]::White
    $labelipADress.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $labelipADress.Cursor = [System.Windows.Forms.Cursors]::Hand  # Cambiar el cursor para que se vea como clickeable
# Crear el ToolTip
    $toolTip.SetToolTip($labelipADress, "Haz clic para copiar las IPs al portapapeles.")
#Funcion para copiar el puerto al portapapeles
$labelPort.Add_Click({
    if ($labelPort.Text -match "\d+") {  # Asegurarse de que el texto es un número
        $port = $matches[0]  # Extraer el número del texto
        [System.Windows.Forms.Clipboard]::SetText($port)
        Write-Host "Puerto copiado al portapapeles: $port" -ForegroundColor Green
    } else {
        Write-Host "El texto del Label del puerto no contiene un número válido para copiar." -ForegroundColor Red
    }
})
$labelHostname.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($labelHostname.Text)
        Write-Host "`nNombre del equipo copiado al portapapeles: $($labelHostname.Text)"
    })
$labelipADress.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($labelipADress.Text)
        Write-Host "`nIP's copiadas al equipo: $($labelipADress.Text)"
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
    $labelipADress.Text = "$ipsTextForLabel"
} else {
    $labelipADress.Text = "No se encontraron direcciones IP"
}

# Configuración dinámica del tamaño del Label según la cantidad de líneas
    $lineHeight = 11
    $maxLines = $labelipADress.Text.Split("`n").Count
    $labelHeight = [Math]::Min(400, $lineHeight * $maxLines)
    $labelipADress.Size = New-Object System.Drawing.Size(240, $labelHeight)
# Ajustar la altura del formulario según el Label de IPs
    $formHeight = $form.Size.Height + $labelHeight - 26
    $form.Size = New-Object System.Drawing.Size($form.Size.Width, $formHeight)





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

    # Mostrar el mensaje con el estado anterior y nuevo
    [System.Windows.Forms.MessageBox]::Show("Categoría anterior: $previousCategory`nCategoría nueva: $category", "Cambio de categoría", [System.Windows.Forms.MessageBoxButtons]::OK)
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

    if ($adapter.NetworkCategory -ne "Private") {
        $text = "$($adapter.AdapterName) - Público"
        $color = [System.Drawing.Color]::Red
    } else {
        $text = "$($adapter.AdapterName) - Privado"
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
        
        # Mostrar el mensaje con el idioma correcto
        $result = [System.Windows.Forms.MessageBox]::Show("¿Deseas cambiar el estado a $category?", "Confirmar cambio", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Set-NetworkCategory -category $category -interfaceIndex $adapter.InterfaceIndex -label $label
        }
    })

    $adapterInfo += $label.Text + "`n"
    $form.Controls.Add($label)

    # Incrementar el índice para la siguiente posición del label
    $index++
}

# Agregar los controles al formulario
    $form.Controls.Add($tabControl)
    $form.Controls.Add($labelHostname)
    $form.Controls.Add($labelPort)
    $form.Controls.Add($labelipADress)
    $form.Controls.Add($lblPerfilDeRed)
    $form.Controls.Add($btnExit)



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
            $labelPort.Text = "Puerto SQL en instancia NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $labelPort.Text = "No se encontró puerto o instancia."
        }


##-------------------- FUNCIONES                                                          -------#
function Check-SqlServerInstallation {
    $sqlInstalled = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft SQL Server%'" | Select-Object -First 1
    return $sqlInstalled
}
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
            Write-Host "`nEl usuario canceló la operación."  -ForegroundColor Red
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
            Write-Host "`nEliminando archivos anteriores..."
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::No) {
            # Si selecciona "No", abrir el programa sin eliminar archivos
            $exePath = Join-Path -Path $extractPath -ChildPath $exeName
            if (Test-Path -Path $exePath) {
                Write-Host "`nEjecutando el archivo ya descargado..."
                Start-Process -FilePath $exePath #-Wait   # Se quitó para ver si se usaban múltiples apps.
                Write-Host "`n$exeName se está ejecutando."
                return
            } else {
                Write-Host "`nNo se pudo encontrar el archivo ejecutable."
                return
            }
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # Si selecciona "Cancelar", no hacer nada y decir que el usuario canceló
            Write-Host "`nEl usuario canceló la operación."
            return  # Aquí se termina la ejecución si el usuario cancela
        }
    }
    # Proceder con la descarga si no fue cancelada
    Write-Host "`nDescargando desde: $url"
    # Obtener el tamaño total del archivo antes de la descarga
    $response = Invoke-WebRequest -Uri $url -Method Head
    $totalSize = $response.Headers["Content-Length"]
    $totalSizeKB = [math]::round($totalSize / 1KB, 2)
    Write-Host "`nTamaño total: $totalSizeKB KB"
    # Descargar el archivo con barra de progreso
    $downloaded = 0
    $request = Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
    foreach ($chunk in $request.Content) {
        $downloaded += $chunk.Length
        $downloadedKB = [math]::round($downloaded / 1KB, 2)
        $progress = [math]::round(($downloaded / $totalSize) * 100, 2)
        Write-Progress -PercentComplete $progress -Status "Descargando..." -Activity "Progreso de la descarga" -CurrentOperation "$downloadedKB KB de $totalSizeKB KB descargados"
    }
    Write-Host "`nDescarga completada."
    # Crear directorio de extracción si no existe
    if (!(Test-Path -Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }
    Write-Host "`nExtrayendo archivos..."
    try {
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Host "`nArchivos extraídos correctamente."
    } catch {
        Write-Host "`nError al descomprimir el archivo: $_"
    }
    $exePath = Join-Path -Path $extractPath -ChildPath $exeName
    if (Test-Path -Path $exePath) {
        Write-Host "`nEjecutando $exeName..."
        Start-Process -FilePath $exePath #-Wait
        Write-Host "`n$exeName se está ejecutando."
    } else {
        Write-Host "`nNo se pudo encontrar el archivo ejecutable."
    }
}

# Función para manejar MouseEnter y cambiar el color
$changeColorOnHover = {
    param($sender, $eventArgs)
    $sender.BackColor = [System.Drawing.Color]::Orange
}

# Función para manejar MouseLeave y restaurar el color
$restoreColorOnLeave = {
    param($sender, $eventArgs)
    $sender.BackColor = [System.Drawing.Color]::White
}
    $labelHostname.Add_MouseEnter($changeColorOnHover)
    $labelHostname.Add_MouseLeave($restoreColorOnLeave)
    $labelPort.Add_MouseEnter($changeColorOnHover)
    $labelPort.Add_MouseLeave($restoreColorOnLeave)
    $labelipADress.Add_MouseEnter($changeColorOnHover)
    $labelipADress.Add_MouseLeave($restoreColorOnLeave)
##-------------------------------------------------------------------------------BOTONES#
$btnProfiler.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        $ProfilerUrl = "https://codeplexarchive.org/codeplex/browse/ExpressProfiler/releases/4/ExpressProfiler22wAddinSigned.zip"
        $ProfilerZipPath = "C:\Temp\ExpressProfiler22wAddinSigned.zip"
        $ExtractPath = "C:\Temp\ExpressProfiler2"
        $ExeName = "ExpressProfiler.exe"
        $ValidationPath = "C:\Temp\ExpressProfiler2\ExpressProfiler.exe"

        DownloadAndRun -url $ProfilerUrl -zipPath $ProfilerZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
        if ($disableControls) {        Enable-Controls -parentControl $form    }
        }
    )
$btnPrinterTool.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        $PrinterToolUrl = "https://3nstar.com/wp-content/uploads/2023/07/RPT-RPI-Printer-Tool-1.zip"
        $PrinterToolZipPath = "C:\Temp\RPT-RPI-Printer-Tool-1.zip"
        $ExtractPath = "C:\Temp\RPT-RPI-Printer-Tool-1"
        $ExeName = "POS Printer Test.exe"
        $ValidationPath = "C:\Temp\RPT-RPI-Printer-Tool-1\POS Printer Test.exe"

        DownloadAndRun -url $PrinterToolUrl -zipPath $PrinterToolZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
    })
$btnDatabase.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        $DatabaseUrl = "https://fishcodelib.com/files/DatabaseNet4.zip"
        $DatabaseZipPath = "C:\Temp\DatabaseNet4.zip"
        $ExtractPath = "C:\Temp\Database4"
        $ExeName = "Database4.exe"
        $ValidationPath = "C:\Temp\Database4\Database4.exe"

        DownloadAndRun -url $DatabaseUrl -zipPath $DatabaseZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
    })
$btnSQLManager.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        try {
            Write-Host "`nEjecutando SQL Server Configuration Manager..."
            Start-Process "SQLServerManager12.msc"
        }
        catch {
            Write-Host "`nError al ejecutar SQL Server Configuration Manager: $_"
            [System.Windows.Forms.MessageBox]::Show("No se pudo abrir SQL Server Configuration Manager.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
$btnSQLManagement.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        try {
            Write-Host "`nEjecutando SQL Server Management Studio..."
            Start-Process "ssms.exe"
        }
        catch {
            Write-Host "`nError al ejecutar SQL Server Management Studio: $_"
            [System.Windows.Forms.MessageBox]::Show("No se pudo abrir SQL Server Management Studio.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
$btnClearAnyDesk.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
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
                Write-Host "`nCerrando el proceso AnyDesk..."
                Stop-Process -Name "AnyDesk" -Force -ErrorAction Stop
                Write-Host "`nAnyDesk ha sido cerrado correctamente."
            }
            catch {
                Write-Host "`nError al cerrar el proceso AnyDesk: $_"
                $errors += "No se pudo cerrar el proceso AnyDesk."
            }

            # Intentar eliminar los archivos
            foreach ($file in $filesToDelete) {
                try {
                    if (Test-Path $file) {
                        Remove-Item -Path $file -Force -ErrorAction Stop
                        Write-Host "`nArchivo eliminado: $file"
                        $deletedFilesCount++
                    }
                    else {
                        Write-Host "`nArchivo no encontrado: $file"
                    }
                }
                catch {
                    Write-Host "`nError al eliminar el archivo."
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
            Write-Host "`nRenovación de AnyDesk cancelada por el usuario."
        }
    })
$buttonShowPrinters.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
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
$btnClearPrintJobs.Add_Click({
            Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
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
$btnAplicacionesNS.Add_Click({
            Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green

    # Proceso 1: Buscar y analizar archivos .ini 
    # Proceso 1: Buscar y analizar archivos .ini 
    $pathsToCheck = @(
        "C:\NationalSoft\Softrestaurant11.0\restaurant.ini",
        "C:\NationalSoft\Softrestaurant10.0\restaurant.ini",
        "C:\NationalSoft\NationalSoftHoteles3.0\nshoteles.ini"
    )

    foreach ($path in $pathsToCheck) {
        if (Test-Path $path) {
            Write-Host "`nArchivo encontrado: $path" -ForegroundColor Green
            $content = Get-Content $path

            # Obtener solo la primera coincidencia de DataSource y Catalog
            $dataSource = ($content | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
            $catalog = ($content | Select-String -Pattern "^Catalog=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
            $authType = $content | Select-String -Pattern "^autenticacion=(\d+)" | ForEach-Object { $_.Matches.Groups[1].Value }

            if ($authType -eq "2") {
                $authUser = "Usuario: sa"
            } elseif ($authType -eq "1") {
                $authUser = "Usuario: Windows"
            } else {
                $authUser = "Autenticación desconocida"
            }

            Write-Host "  DataSource:" -ForegroundColor Cyan
            Write-Host "    $dataSource" -ForegroundColor White
            Write-Host "  Catalog:" -ForegroundColor Cyan
            Write-Host "    $catalog" -ForegroundColor White
            Write-Host "  $authUser" -ForegroundColor Cyan

            # Revisar subcarpeta INIS
            $inisPath = [System.IO.Path]::Combine((Get-Item $path).DirectoryName, "INIS")
            if (Test-Path $inisPath) {
                $iniFiles = Get-ChildItem -Path $inisPath -Filter "*.ini"
                if ($iniFiles.Count -gt 1) {
                    Write-Host "`nEstá utilizando multiempresas, revisa los INIS en la carpeta: $inisPath" -ForegroundColor Red
                } elseif ($iniFiles.Count -eq 1) {
                    $firstIniContent = Get-Content $iniFiles[0].FullName
                    $firstDataSource = ($firstIniContent | Select-String -Pattern "^DataSource=(.*)" | Select-Object -First 1).Matches.Groups[1].Value
                    $firstCatalog = ($firstIniContent | Select-String -Pattern "^Catalog=(.*)" | Select-Object -First 1).Matches.Groups[1].Value

                    # Comparar con los datos de los archivos encontrados previamente
                    if ($dataSource -eq $firstDataSource -and $catalog -eq $firstCatalog) {
                        Write-Host "`nLos INIs apuntan al mismo servidor." -ForegroundColor Green
                    } else {
                        Write-Host "`nLos INIs tienen incongruencias, revisa los INIS en la carpeta: $inisPath" -ForegroundColor Red
                        Write-Host "  DataSource del INI original:" -ForegroundColor Cyan
                        Write-Host "    $dataSource" -ForegroundColor White
                        Write-Host "  Catalog del INI original:" -ForegroundColor Cyan
                        Write-Host "    $catalog" -ForegroundColor White
                        Write-Host "  DataSource del INI encontrado:" -ForegroundColor Cyan
                        Write-Host "    $firstDataSource" -ForegroundColor White
                        Write-Host "  Catalog del INI encontrado:" -ForegroundColor Cyan
                        Write-Host "    $firstCatalog" -ForegroundColor White
                    }
                }
            }
        } else {
            Write-Host "`nArchivo no encontrado: $path" -ForegroundColor Red
        }
    }
            # Proceso 2: Detectar y validar archivo checadorsql.ini
            $checadorsqlPath = "C:\NationalSoft\OnTheMinute4.5\checadorsql.ini"
            if (Test-Path $checadorsqlPath) {
                Write-Host "`nArchivo encontrado: $checadorsqlPath" -ForegroundColor Green
                $checadorsqlContent = Get-Content $checadorsqlPath

                $provider = ($checadorsqlContent | Select-String -Pattern "^Provider=(.*)" | ForEach-Object { $_.Matches.Groups[1].Value }).Trim()

                if ($provider -eq "SQLOLEDB.1") {
                    Write-Host "`nEl proveedor es SQL. Ejecutando Proceso 1." -ForegroundColor White

                    # Ejecutar Proceso 1 en caso de proveedor SQL
                    $dataSource = $checadorsqlContent | Select-String -Pattern "^DataSource=(.*)" | ForEach-Object { $_.Matches.Groups[1].Value }
                    $catalog = $checadorsqlContent | Select-String -Pattern "^Catalog=(.*)" | ForEach-Object { $_.Matches.Groups[1].Value }
                    $authType = $checadorsqlContent | Select-String -Pattern "^autenticacion=(\d+)" | ForEach-Object { $_.Matches.Groups[1].Value }

                    if ($authType -eq "2") {
                        $authUser = "Usuario: sa"
                    } elseif ($authType -eq "1") {
                        $authUser = "Usuario: Windows"
                    } else {
                        $authUser = "Autenticación desconocida"
                    }

                    Write-Host "  DataSource:" -ForegroundColor Cyan
                    Write-Host "    $dataSource" -ForegroundColor White
                    Write-Host "  Catalog:" -ForegroundColor Cyan
                    Write-Host "    $catalog" -ForegroundColor White
                    Write-Host "  $authUser" -ForegroundColor Cyan

                } elseif ($provider -eq "VFPOLEDB.1") {
                    $dataSource = $checadorsqlContent | Select-String -Pattern "^DataSource=(.*)" | ForEach-Object { $_.Matches.Groups[1].Value }
                    Write-Host "`nEl proveedor es DBF. Ruta de datos: $dataSource" -ForegroundColor White
                } else {
                    Write-Host "`nProveedor desconocido en checadorsql.ini" -ForegroundColor Red
                }
            } else {
                Write-Host "`nArchivo no encontrado: $checadorsqlPath" -ForegroundColor Red
            }

            # Proceso 3: Detectar si existe RestCard.ini
            $restCardPath = "C:\NationalSoft\Restcard\RestCard.ini"
            if (Test-Path $restCardPath) {
                Write-Host "`nArchivo encontrado: $restCardPath" -ForegroundColor Green
            } else {
                Write-Host "`nArchivo no encontrado: $restCardPath" -ForegroundColor Red
            }
        })
$btnInstallSQLManagement.Add_Click({
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

    Write-Host "`nVerificando si Chocolatey está instalado..." -ForegroundColor Yellow

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "`nChocolatey no está instalado. Instalándolo ahora..." -ForegroundColor Cyan
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

            Write-Host "`nChocolatey se instaló correctamente." -ForegroundColor Green
            [System.Windows.Forms.MessageBox]::Show(
                "Chocolatey se instaló correctamente. Por favor, reinicie PowerShell antes de continuar.",
                "Reinicio requerido",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        } catch {
            Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "Error al instalar Chocolatey. Por favor, inténtelo manualmente.",
                "Error de instalación",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }
    } else {
        Write-Host "`nChocolatey ya está instalado." -ForegroundColor Green
    }

    Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green

    try {
        Write-Host "`nInstalando SQL Server Management Studio 2014 Express usando Chocolatey..." -ForegroundColor Cyan
        Start-Process choco -ArgumentList 'install mssqlservermanagementstudio2014express --confirm --yes' -NoNewWindow -Wait
        Write-Host "`nInstalación completa." -ForegroundColor Green
    } catch {
        Write-Host "`nOcurrió un error durante la instalación: $_" -ForegroundColor Red
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
            $lblConnectionStatus.Font = New-Object System.Drawing.Font($lblConnectionStatus.Font, [System.Drawing.FontStyle]::Bold)


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

$btnExit.Add_Click({
        $form.Close()
        $form.Dispose()
})


$form.Refresh()

# Mostrar el formulario principal
$form.ShowDialog()
