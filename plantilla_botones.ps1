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
                                                                    $version = "Alfa SQL.1454"  # Valor predeterminado para la versión
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
                    [string]$ToolTipText = $null,  # Nuevo parámetro para el ToolTip
                    [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35))  # Tamaño personalizable
                )
                # Estilo del botón
                $buttonStyle = @{
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
                $button.Size = $Size  # Usar el tamaño personalizado
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
#BotonCreado
    $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Carpeta LZMA" -Location (New-Object System.Drawing.Point(240, 210))
    $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 40))
    $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(240, 40))
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
    $tabAplicaciones.Controls.Add($LZMAbtnBuscarCarpeta)
# Agregar controles a la pestaña Pro
    $tabProSql.Controls.Add($chkSqlServer)
    $tabProSql.Controls.Add($btnReviewPivot)
    $tabProSql.Controls.Add($btnRespaldarRestcard)
    $tabProSql.Controls.Add($btnFechaRevEstaciones)
    $tabProSql.Controls.Add($lblConnectionStatus)
    $tabProSql.Controls.Add($btnConnectDb)
    $tabProSql.Controls.Add($btnDisconnectDb)
# Agregar los controles al formulario
    $formPrincipal.Controls.Add($tabControl)
    $formPrincipal.Controls.Add($lblHostname)
    $formPrincipal.Controls.Add($lblPort)
    $formPrincipal.Controls.Add($lbIpAdress)
    $formPrincipal.Controls.Add($lblPerfilDeRed)
    $formPrincipal.Controls.Add($btnExit)
##-------------------- FUNCIONES                                                          -------#
##---------------OTROS BOTONES Y FUNCIONES OMITIDAS AQUI----------------------------------------------------------------BOTONES#
                $LZMAbtnBuscarCarpeta.Add_Click({
         Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
                   $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
                    try {
                        # Intentar obtener las carpetas principales
                        $LZMcarpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }
                
                        # Verificar si hay al menos una carpeta principal
                        if ($LZMcarpetasPrincipales.Count -ge 1) {
                            # Crear una lista para almacenar las subcarpetas y sus rutas completas
                            $LZMsubCarpetas = @("Selecciona instalador a renombrar")  # Opción por defecto
                            $LZMrutasCompletas = @()
                            # Recorrer cada carpeta principal y obtener sus subcarpetas
                            foreach ($LZMcarpetaPrincipal in $LZMcarpetasPrincipales) {
                                $LZMsubCarpetasPrincipal = Get-ChildItem -Path $LZMcarpetaPrincipal.PSPath | Where-Object { $_.PSIsContainer }
                                foreach ($LZMsubCarpeta in $LZMsubCarpetasPrincipal) {
                                    $LZMsubCarpetas += $LZMsubCarpeta.PSChildName
                                    $LZMrutasCompletas += $LZMsubCarpeta.PSPath
                                }
                            }
                            # Verificar si hay al menos una subcarpeta
                            if ($LZMsubCarpetas.Count -gt 1) {
                                # Crear un nuevo formulario para mostrar las subcarpetas
                                $formLZMA = New-Object System.Windows.Forms.Form
                                $formLZMA.Text = "Carpetas LZMA"
                                $formLZMA.Size = New-Object System.Drawing.Size(400, 200)
                                $formLZMA.StartPosition = "CenterScreen"
                                $formLZMA.MaximizeBox = $false  # Deshabilitar botón de maximizar
                                $formLZMA.MinimizeBox = $false  # Deshabilitar botón de minimizar
                                $formLZMA.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog  # Evitar redimensionar
                                # Crear un ComboBox para mostrar las subcarpetas
                                $LZMcomboBoxCarpetas = New-Object System.Windows.Forms.ComboBox
                                $LZMcomboBoxCarpetas.Location = New-Object System.Drawing.Point(10, 10)
                                $LZMcomboBoxCarpetas.Size = New-Object System.Drawing.Size(360, 20)
                                $LZMcomboBoxCarpetas.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                                $LZMcomboBoxCarpetas.Font = $defaultFont  # Usar la fuente predeterminada
                                # Agregar las subcarpetas al ComboBox
                                foreach ($LZMsubCarpeta in $LZMsubCarpetas) {
                                    $LZMcomboBoxCarpetas.Items.Add($LZMsubCarpeta)
                                }
                                # Seleccionar la primera opción por defecto
                                $LZMcomboBoxCarpetas.SelectedIndex = 0
                                # Crear un Label para mostrar el valor de AI_ExePath
                                $LZMlblExePath = New-Object System.Windows.Forms.Label
                                $LZMlblExePath.Location = New-Object System.Drawing.Point(10, 40)
                                $LZMlblExePath.Size = New-Object System.Drawing.Size(360, 60)  # Aumentar la altura para 3 líneas
                                $LZMlblExePath.Font = $defaultFont  # Usar la fuente predeterminada
                                $LZMlblExePath.Text = "AI_ExePath: -"
                                # Evento cuando se selecciona una subcarpeta en el ComboBox
                                $LZMcomboBoxCarpetas.Add_SelectedIndexChanged({
                                    $indiceSeleccionado = $LZMcomboBoxCarpetas.SelectedIndex
                                    if ($indiceSeleccionado -gt 0) {  # Ignorar la opción por defecto
                                        $LZMrutaCompleta = $LZMrutasCompletas[$indiceSeleccionado - 1]  # Ajustar índice
                                        $valorExePath = Get-ItemProperty -Path $LZMrutaCompleta -Name "AI_ExePath" -ErrorAction SilentlyContinue
                                        if ($valorExePath) {
                                            $LZMlblExePath.Text = "AI_ExePath: $($valorExePath.AI_ExePath)"
                                        } else {
                                            $LZMlblExePath.Text = "AI_ExePath: No encontrado"
                                        }
                                    } else {
                                        $LZMlblExePath.Text = "AI_ExePath: -"
                                    }
                                })
                                # Crear botón para renombrar usando la función Create-Button
                                $LZMbtnRenombrar = Create-Button -Text "Renombrar" -Location (New-Object System.Drawing.Point(10, 100)) -Size (New-Object System.Drawing.Size(170, 40))
                                $LZMbtnRenombrar.Enabled = $false  # Deshabilitar inicialmente
                                # Evento Click del botón Renombrar
                                $LZMbtnRenombrar.Add_Click({
                                    $indiceSeleccionado = $LZMcomboBoxCarpetas.SelectedIndex
                                    if ($indiceSeleccionado -gt 0) {  # Ignorar la opción por defecto
                                        $LZMrutaCompleta = $LZMrutasCompletas[$indiceSeleccionado - 1]  # Ajustar índice
                                        $nuevaRuta = "$LZMrutaCompleta.backup"  # Nueva ruta con .backup
                                        Write-Host "`t¿Estás seguro de que deseas renombrar la ruta del registro?`n$LZMrutaCompleta`nA:`n$nuevaRuta" -ForegroundColor Yellow
                                        $confirmacion = [System.Windows.Forms.MessageBox]::Show(
                                            "¿Estás seguro de que deseas renombrar la ruta del registro?`n$LZMrutaCompleta`nA:`n$nuevaRuta",
                                            "Confirmar renombrado",
                                            [System.Windows.Forms.MessageBoxButtons]::YesNo,
                                            [System.Windows.Forms.MessageBoxIcon]::Warning
                                        )
                                        if ($confirmacion -eq [System.Windows.Forms.DialogResult]::Yes) {
                                            try {
                                                Rename-Item -Path $LZMrutaCompleta -NewName "$($LZMcomboBoxCarpetas.SelectedItem).backup"
                                                [System.Windows.Forms.MessageBox]::Show("Registro renombrado correctamente.", "Éxito", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                                                Write-Host "`tRegistro renombrado correctamente." -ForegroundColor Yellow
                                                $formLZMA.Close()  # Cerrar el formulario después de renombrar
                                            } catch {
                                                Write-Host "`tError al renombrar el registro." -ForegroundColor Red
                                                [System.Windows.Forms.MessageBox]::Show("Error al renombrar el registro: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                                            }
                                        }
                                    }
                                })
                                # Crear botón para salir usando la función Create-Button
                                $LMZAbtnSalir = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(190, 100)) -Size (New-Object System.Drawing.Size(170, 40))
                                # Evento Click del botón Salir
                                $LMZAbtnSalir.Add_Click({
                                                Write-Host "`tCancelado por el usuario." -ForegroundColor Yellow
                                    $formLZMA.Close()
                                })
                                # Habilitar el botón Renombrar solo si se selecciona una opción válida
                                $LZMcomboBoxCarpetas.Add_SelectedIndexChanged({
                                    $LZMbtnRenombrar.Enabled = ($LZMcomboBoxCarpetas.SelectedIndex -gt 0)
                                })
                                # Agregar controles al formulario
                                $formLZMA.Controls.Add($LZMcomboBoxCarpetas)
                                $formLZMA.Controls.Add($LZMbtnRenombrar)
                                $formLZMA.Controls.Add($LMZAbtnSalir)
                                $formLZMA.Controls.Add($LZMlblExePath)
                                # Mostrar el formulario
                                $formLZMA.ShowDialog()
                            } else {
                                Write-Host "`tNo se encontraron subcarpetas en la ruta del registro." -ForegroundColor Yellow
                                [System.Windows.Forms.MessageBox]::Show("No se encontraron subcarpetas en la ruta del registro.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            }
                        } else {
                            Write-Host "`tNo se encontraron carpetas principales en la ruta del registro." -ForegroundColor Yellow
                            [System.Windows.Forms.MessageBox]::Show("No se encontraron carpetas principales en la ruta del registro.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                    } catch {
                        # Capturar la excepción si la ruta no existe
                        Write-Host "`tLa ruta del registro no existe: $LZMAregistryPath" -ForegroundColor Yellow
                        [System.Windows.Forms.MessageBox]::Show("La ruta del registro no existe: $LZMAregistryPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                })
#SALIR DEL SISTEMA------------------------------------------------
        $btnExit.Add_Click({
                        $formPrincipal.Dispose()
                        $formPrincipal.Close()
                    })
$formPrincipal.Refresh()
# Mostrar el formulario principal
$formPrincipal.ShowDialog()
