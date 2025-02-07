# Crear la carpeta 'C:\Temp' si no existe
if (!(Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    Write-Host "Carpeta 'C:\Temp' creada correctamente."
}
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
# Crear el formulario
    $formPrincipal = New-Object System.Windows.Forms.Form
    $formPrincipal.Size = New-Object System.Drawing.Size(500, 470)
    $formPrincipal.StartPosition = "CenterScreen"
    $formPrincipal.BackColor = [System.Drawing.Color]::White
    $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formPrincipal.MaximizeBox = $false
    $formPrincipal.MinimizeBox = $false
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                                                                                                        $version = "permi_250207.1057"  # Valor predeterminado para la versión
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
                    [string]$ToolTipText = $null,
                    [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35)),
                    [bool]$Enabled = $true
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
                $button.Enabled = $Enabled
                # Agregar ToolTip si se proporciona
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
                
                    # Crear la etiqueta
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
                
                # Crear el formulario
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
        
            # Crear el ComboBox
            $comboBox = New-Object System.Windows.Forms.ComboBox
            $comboBox.Location = $Location
            $comboBox.Size = $Size
            $comboBox.DropDownStyle = $DropDownStyle
            $comboBox.Font = $Font
        
            # Agregar elementos si hay disponibles
            if ($Items.Count -gt 0) {
                $comboBox.Items.AddRange($Items)
                $comboBox.SelectedIndex = $SelectedIndex
            }
        
            # Definir texto por defecto si se especifica
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

    # Crear el TextBox
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

    # Si se requiere usar el sistema de contraseña, configurar el PasswordChar
    if ($UseSystemPasswordChar) {
        $textBox.UseSystemPasswordChar = $true
    }

    return $textBox
}
# Crear las pestañas (TabControl)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size(480, 315) #X,Y
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
    $btnInstallSQLManagement = Create-Button -Text "Instalar Management2014" -Location (New-Object System.Drawing.Point(10, 10)) `
                                -ToolTip "Instalación mediante choco de SQL Management 2014."
    $btnProfiler = Create-Button -Text "Ejecutar ExpressProfiler" -Location (New-Object System.Drawing.Point(10, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnDatabase = Create-Button -Text "Ejecutar Database4" -Location (New-Object System.Drawing.Point(10, 90)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Ejecuta o Descarga la herramienta desde el servidor oficial."
    $btnSQLManager = Create-Button -Text "Ejecutar Manager" -Location (New-Object System.Drawing.Point(10, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "De momento solo si es SQL 2014."
    $btnSQLManagement = Create-Button -Text "Ejecutar Management" -Location (New-Object System.Drawing.Point(10, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Busca SQL Management en tu equipo y te confirma la versión previo a ejecutarlo."
    $btnPrinterTool = Create-Button -Text "Printer Tools" -Location (New-Object System.Drawing.Point(10, 210)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Herramienta de Star con funciones multiples para impresoras POS."
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(240, 10)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(240, 50)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(240, 90)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(240, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
    $btnConfigurarIPs = Create-Button -Text "Configurar IPs" -Location (New-Object System.Drawing.Point(240, 170)) `
                                -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
    $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Carpeta LZMA" -Location (New-Object System.Drawing.Point(240, 210)) `
                                -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
    $btnModificarPermisos = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(240, 250)) `
                                -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
    $btnConnectDb = Create-Button -Text "Conectar a BDD" -Location (New-Object System.Drawing.Point(10, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255))
    $btnDisconnectDb = Create-Button -Text "Desconectar de BDD" -Location (New-Object System.Drawing.Point(240, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -Enabled $false
    $btnReviewPivot = Create-Button -Text "Revisar Pivot Table" -Location (New-Object System.Drawing.Point(10, 90)) `
                                -ToolTip "Para SR, busca y elimina duplicados en app_settings" -Enabled $false
    $btnEliminarServidorBDD = Create-Button -Text "Eliminar Server de BDD" -Location (New-Object System.Drawing.Point(240, 90)) `
                                -ToolTip "Quitar servidor asignado a la base de datos." -Enabled $false
    $btnFechaRevEstaciones = Create-Button -Text "Fecha de revisiones" -Location (New-Object System.Drawing.Point(10, 130)) `
                                -ToolTip "Para SR, revision, ultimo uso y estación." -Enabled $false
    $btnRespaldarRestcard = Create-Button -Text "Respaldar restcard" -Location (New-Object System.Drawing.Point(10, 210)) `
                                -ToolTip "Respaldo de Restcard, puede requerir MySQL instalado."
    $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(120, 320)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
# Crear el botón para revisar permisos
$btnCheckPermissions = Create-Button -Text "Revisar Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(10, 250)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
# Crear el CheckBox chkSqlServer
    $chkSqlServer = New-Object System.Windows.Forms.CheckBox
    $chkSqlServer.Text = "Instalar SQL Tools (opcional)"
    $chkSqlServer.Size = New-Object System.Drawing.Size(290, 30)
    $chkSqlServer.Location = New-Object System.Drawing.Point(10, 10)
# Usar la función Create-Label para crear la label de conexión
    $lblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" -Location (New-Object System.Drawing.Point(10, 260)) -Size (New-Object System.Drawing.Size(290, 30)) `
                                     -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 0, 0)) -Font $defaultFont
# Crear el Label para mostrar el nombre del equipo fuera de las pestañas
    $lblHostname = Create-Label -Text ([System.Net.Dns]::GetHostName()) -Location (New-Object System.Drawing.Point(2, 360)) -Size (New-Object System.Drawing.Size(240, 35)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Hostname al portapapeles."
# Crear el Label para mostrar el puerto
    $lblPort = Create-Label -Text "Puerto: No disponible" -Location (New-Object System.Drawing.Point(245, 360)) -Size (New-Object System.Drawing.Size(236, 35)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Puerto al portapapeles."
# Crear el Label para mostrar las IPs y adaptadores
    $lbIpAdress = Create-Label -Text "Obteniendo IPs..." -Location (New-Object System.Drawing.Point(2, 400)) -Size (New-Object System.Drawing.Size(240, 100)) -BorderStyle FixedSingle -TextAlign TopLeft -ToolTipText "Haz clic para copiar las IPs al portapapeles."

                    # Función para revisar permisos y agregar Full Control a "Everyone" si es necesario
                    function Check-Permissions {
                        $folderPath = "C:\NationalSoft"
                        $acl = Get-Acl -Path $folderPath
                        $permissions = @()
                    
                        $everyoneHasFullControl = $false
                    
                        foreach ($access in $acl.Access) {
                            $permissions += [PSCustomObject]@{
                                Usuario = $access.IdentityReference
                                Permiso = $access.FileSystemRights
                                Tipo    = $access.AccessControlType
                            }
                    
                            # Verificar si "Everyone" tiene FullControl o Control total
                            if ($access.IdentityReference -eq "Everyone" -and 
                                ($access.FileSystemRights -match "FullControl" -or $access.FileSystemRights -match "Control total")) {
                                $everyoneHasFullControl = $true
                            }
                        }
                    
                        # Mostrar los permisos en la consola
                        $permissions | ForEach-Object { 
                            Write-Host "$($_.Usuario) - $($_.Permiso) - $($_.Tipo)"
                        }
                    
                        # Si "Everyone" no tiene Full Control o Control total, preguntar si se desea concederlo
                        if (-not $everyoneHasFullControl) {
                            $message = "El usuario 'Everyone' no tiene permisos de 'Full Control' (o 'Control total'). ¿Deseas concederlo?"
                            $title = "Permisos 'Everyone'"
                            $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
                            $icon = [System.Windows.Forms.MessageBoxIcon]::Question
                    
                            $result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
                    
                            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                                # Aquí puedes agregar el código para conceder el permiso
                                # Por ejemplo, agregar Full Control a "Everyone" en la carpeta
                                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "Allow")
                                $acl.AddAccessRule($accessRule)
                                Set-Acl -Path $folderPath -AclObject $acl
                                Write-Host "Se ha concedido 'Full Control' a 'Everyone'."
                            }
                        }
                    }
                    
                    # Asignar la función al botón (si sigue siendo necesario)
                    $btnCheckPermissions.Add_Click({
                        Check-Permissions
                    })
                    
                    # Agregar el botón a la pestaña de Aplicaciones (si sigue siendo necesario)
                    $tabAplicaciones.Controls.Add($btnCheckPermissions)




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
    $tabAplicaciones.Controls.Add($btnModificarPermisos)

#Funcion para copiar el puerto al portapapeles
    $lblPort.Add_Click({
    })
$lblHostname.Add_Click({
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
                [System.Windows.Forms.Label]$lblAdaptadorStatus
            )
            
            # Obtener el estado anterior
            $profile = Get-NetConnectionProfile | Where-Object { $_.InterfaceIndex -eq $interfaceIndex }
            $previousCategory = if ($profile) { $profile.NetworkCategory } else { "Desconocido" }
            
            # Solo cambiar si la red es pública
            if ($previousCategory -eq "Public") {
                if ($category -eq "Privado") {
                    Set-NetConnectionProfile -InterfaceIndex $interfaceIndex -NetworkCategory Private
                    Write-Host "Estado cambiado a Privado."
                    $lblAdaptadorStatus.ForeColor = [System.Drawing.Color]::Green
                    $lblAdaptadorStatus.Text = "$($lblAdaptadorStatus.Text.Split(' - ')[0]) - Privado"  # Actualizar el texto de la etiqueta
                }
            } else {
                Write-Host "La red ya es privada o no es pública, no se realizará ningún cambio."
            }
        }
        
        # Crear la etiqueta para mostrar los adaptadores y su estado
        $lblPerfilDeRed = Create-Label -Text "Estado de los Adaptadores:" -Location (New-Object System.Drawing.Point(245, 400)) -Size (New-Object System.Drawing.Size(236, 25)) -TextAlign MiddleCenter -ToolTipText "Haz clic para cambiar la red a privada."
        
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
            
            $lblAdaptadorStatus = Create-Label -Text $text -Location (New-Object System.Drawing.Point(245, (425 + (30 * $index)))) -Size (New-Object System.Drawing.Size(236, 20)) -ForeColor $color -BorderStyle FixedSingle
            $lblAdaptadorStatus.Add_MouseEnter($changeColorOnHover)
            $lblAdaptadorStatus.Add_MouseLeave($restoreColorOnLeave)

            # Función de cierre para capturar el adaptador actual
            $adapterIndex = $adapter.InterfaceIndex
            $lblAdaptadorStatus.Add_Click({
                # Obtener el adaptador asociado a este label
                $currentCategory = $adapter.NetworkCategory
                
                # Solo cambiar si la red es pública
                if ($currentCategory -eq "Public") {
                    # Confirmar el cambio y llamar a la función de cambio
                    $result = [System.Windows.Forms.MessageBox]::Show("¿Deseas cambiar el estado a Privado?", "Confirmar cambio", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                    
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Set-NetworkCategory -category "Privado" -interfaceIndex $adapterIndex -lblAdaptadorStatus $lblAdaptadorStatus
                    }
                } else {
                    Write-Host "La red ya es privada o no es pública, no se realizará ningún cambio."
                }
            })
        
            $adapterInfo += $lblAdaptadorStatus.Text + "`n"
            $formPrincipal.Controls.Add($lblAdaptadorStatus)
            
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
# Obtener el puerto de SQL Server desde el registro
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\NATIONALSOFT\MSSQLServer\SuperSocketNetLib\Tcp"
        $tcpPort = Get-ItemProperty -Path $regKeyPath -Name "TcpPort" -ErrorAction SilentlyContinue

        if ($tcpPort -and $tcpPort.TcpPort) {
            $lblPort.Text = "Puerto SQL \NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $lblPort.Text = "No se encontró puerto o instancia."
        }
    $lblHostname.Add_MouseEnter($changeColorOnHover)
    $lblHostname.Add_MouseLeave($restoreColorOnLeave)
    $lblPort.Add_MouseEnter($changeColorOnHover)
    $lblPort.Add_MouseLeave($restoreColorOnLeave)
    $lbIpAdress.Add_MouseEnter($changeColorOnHover)
    $lbIpAdress.Add_MouseLeave($restoreColorOnLeave)

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
    
#Boton para salir
    $btnExit.Add_Click({
        $formPrincipal.Dispose()
        $formPrincipal.Close()
    })
$formPrincipal.Refresh()
$formPrincipal.ShowDialog()
