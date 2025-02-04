##############################﻿# Crear la carpeta 'C:\Temp' si no existe
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
                                                                    $version = "Alfa SQL.2034"  # Valor predeterminado para la versión
    $formPrincipal.Text = "Daniel Tools v$version"
    Write-Host "              Versión: v$($version)               " -ForegroundColor Green
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
    $btnInstallSQLManagement = Create-Button -Text "Instalar Management2014" -Location (New-Object System.Drawing.Point(10, 10))
    $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 50))
    $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 90))
    $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 130))
    $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 170))
    $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 210))
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(240, 10))
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(240, 50))
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(240, 90))
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(240, 130))
    $btnConfigurarIPs = Create-Button -Text "Configurar IPs" -Location (New-Object System.Drawing.Point(240, 170))
    $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 40))
    $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(240, 40))
    $btnDisconnectDb.Enabled = $false  # Deshabilitado inicialmente
    $btnReviewPivot = Create-Button -Text "Revisar Pivot Table" -Location (New-Object System.Drawing.Point(10, 110))
    $btnReviewPivot.Enabled = $false  # Deshabilitado inicialmente
    $btnFechaRevEstaciones = Create-Button -Text "Fecha de revisiones" -Location (New-Object System.Drawing.Point(10, 150))
    $btnFechaRevEstaciones.Enabled = $false  # Deshabilitado inicialmente
# Crear el botón btnEliminarServidorBDD
$btnEliminarServidorBDD = Create-Button -Text "Eliminar Server de BDD" -Location (New-Object System.Drawing.Point(240, 80)) -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 200, 200))
$btnEliminarServidorBDD.Enabled = $false  # Deshabilitado inicialmente
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
# Agregar el botón a la pestaña Pro
$tabProSql.Controls.Add($btnEliminarServidorBDD)
# Agregar los controles al formulario
    $formPrincipal.Controls.Add($tabControl)
    $formPrincipal.Controls.Add($lblHostname)
    $formPrincipal.Controls.Add($lblPort)
    $formPrincipal.Controls.Add($lbIpAdress)
    $formPrincipal.Controls.Add($lblPerfilDeRed)
    $formPrincipal.Controls.Add($btnExit)
##-------------------- FUNCIONES                                                          -------#
# Función para ejecutar consultas SQL
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
        $reader.Close()
        $connection.Close()
        $connection.Dispose()

        return $results
    } catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
        return $null
    }
}
# Función para mostrar los resultados en la consola en columnas
function Show-ResultsConsole {
    param (
        [string]$query
    )
    try {
        # Ejecutar la consulta y obtener los resultados
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query

        if ($results -and $results.Count -gt 0) {
            # Obtener los nombres de las columnas
            $columns = $results[0].Keys

            # Calcular el ancho máximo de cada columna
            $columnWidths = @{}
            foreach ($col in $columns) {
                $maxLength = $col.Length
                foreach ($row in $results) {
                    if ($row[$col] -and $row[$col].ToString().Length -gt $maxLength) {
                        $maxLength = $row[$col].ToString().Length
                    }
                }
                $columnWidths[$col] = $maxLength
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
#Boton para revisar estaciones
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
        Write-Host "`n$query1`n" -ForegroundColor Yellow
                    Show-ResultsConsole -query $query1
    } catch {
        Write-Host "`nError al ejecutar consulta: $($_.Exception.Message)" -ForegroundColor Red
    }
})








# Función para manejar el evento Click del botón btnEliminarServidorBDD
$btnEliminarServidorBDD.Add_Click({
    # Crear el formulario para eliminar el servidor de la base de datos
    $formEliminarServidor = New-Object System.Windows.Forms.Form
    $formEliminarServidor.Text = "Eliminar Servidor de BDD"
    $formEliminarServidor.Size = New-Object System.Drawing.Size(400, 200)
    $formEliminarServidor.StartPosition = "CenterScreen"
    $formEliminarServidor.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formEliminarServidor.MaximizeBox = $false
    $formEliminarServidor.MinimizeBox = $false

    # Crear el ComboBox con las opciones
    $cmbOpciones = New-Object System.Windows.Forms.ComboBox
    $cmbOpciones.Location = New-Object System.Drawing.Point(10, 20)
    $cmbOpciones.Size = New-Object System.Drawing.Size(360, 20)
    $cmbOpciones.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbOpciones.Items.AddRange(@("Seleccione 1", "On The minute", "NS Hoteles", "Rest Card"))

    # Crear los botones Eliminar y Cancelar
    $btnEliminar = New-Object System.Windows.Forms.Button
    $btnEliminar.Text = "Eliminar"
    $btnEliminar.Size = New-Object System.Drawing.Size(100, 30)
    $btnEliminar.Location = New-Object System.Drawing.Point(150, 60)
    $btnEliminar.Enabled = $false  # Deshabilitado inicialmente

    $btnCancelar = New-Object System.Windows.Forms.Button
    $btnCancelar.Text = "Cancelar"
    $btnCancelar.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancelar.Location = New-Object System.Drawing.Point(260, 60)

    # Habilitar el botón Eliminar si se selecciona una opción válida
    $cmbOpciones.Add_SelectedIndexChanged({
        if ($cmbOpciones.SelectedIndex -gt 0) {
            $btnEliminar.Enabled = $true
        } else {
            $btnEliminar.Enabled = $false
        }
    })

    # Manejar el evento Click del botón Eliminar
    $btnEliminar.Add_Click({
        $opcionSeleccionada = $cmbOpciones.SelectedItem
        $confirmacion = [System.Windows.Forms.MessageBox]::Show("¿Está seguro de que desea eliminar el servidor de la base de datos para $opcionSeleccionada?", "Confirmar Eliminación", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($confirmacion -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                switch ($opcionSeleccionada) {
                    "On The minute" {
                        $query = "UPDATE configuracion SET serie='', ipserver='', nombreservidor=''"
                        Execute-SqlQuery -server $global:server -database $global:database -query $query
                        Write-Host "Servidor de BDD eliminado para On The minute." -ForegroundColor Green
                    }
                    "NS Hoteles" {
                        $query = "UPDATE configuracion SET serievalida='', numserie='', ipserver='', nombreservidor='', llave=''"
                        Execute-SqlQuery -server $global:server -database $global:database -query $query
                        Write-Host "Servidor de BDD eliminado para NS Hoteles." -ForegroundColor Green
                    }
                    "Rest Card" {
                        # Pedir al usuario los datos de conexión a MySQL
                        $usuarioRestcard = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese el usuario de MySQL:", "Usuario MySQL")
                        $passwordRestcard = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese la contraseña de MySQL:", "Contraseña MySQL")
                        $hostnameRestcard = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese el hostname de MySQL:", "Hostname MySQL")
                        $baseDeDatosRestcard = [Microsoft.VisualBasic.Interaction]::InputBox("Ingrese el nombre de la base de datos:", "Base de Datos MySQL")

                        # Ejecutar el comando mysqldump
                        $argumentos = "-u $usuarioRestcard -p$passwordRestcard -h $hostnameRestcard $baseDeDatosRestcard --result-file=`"$rutaRespaldo`""
                        $process = Start-Process -FilePath "mysqldump" -ArgumentList $argumentos -NoNewWindow -Wait -PassThru

                        if ($process.ExitCode -eq 0) {
                            $query = "UPDATE tabvariables SET estacion='', ipservidor=''"
                            Execute-SqlQuery -server $global:server -database $global:database -query $query
                            Write-Host "Servidor de BDD eliminado para Rest Card." -ForegroundColor Green
                        } else {
                            Write-Host "Error al ejecutar mysqldump." -ForegroundColor Red
                        }
                    }
                }
                $formEliminarServidor.Close()
            } catch {
                Write-Host "Error al eliminar el servidor de BDD: $_" -ForegroundColor Red
            }
        }
    })

    # Manejar el evento Click del botón Cancelar
    $btnCancelar.Add_Click({
        $formEliminarServidor.Close()
    })

    # Agregar los controles al formulario
    $formEliminarServidor.Controls.Add($cmbOpciones)
    $formEliminarServidor.Controls.Add($btnEliminar)
    $formEliminarServidor.Controls.Add($btnCancelar)

    # Mostrar el formulario
    $formEliminarServidor.ShowDialog()
})

# Habilitar el botón btnEliminarServidorBDD cuando se conecte a la base de datos
$btnConnectDb.Add_Click({
    # ... (código existente para conectar a la base de datos)
    $btnEliminarServidorBDD.Enabled = $true
})

# Deshabilitar el botón btnEliminarServidorBDD cuando se desconecte de la base de datos
$btnDisconnectDb.Add_Click({
    # ... (código existente para desconectar de la base de datos)
    $btnEliminarServidorBDD.Enabled = $false
})
#SALIR DEL SISTEMA------------------------------------------------
$btnExit.Add_Click({
                        $formPrincipal.Dispose()
                        $formPrincipal.Close()
                    })
$formPrincipal.Refresh()
# Mostrar el formulario principal
$formPrincipal.ShowDialog()
