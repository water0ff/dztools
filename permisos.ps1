# Crear la carpeta 'C:\Temp' si no existe
if (!(Test-Path -Path "C:\Temp")) {
    New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    Write-Host "Carpeta 'C:\Temp' creada correctamente."
}
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
# Crear el formulario
    $formPrincipal = New-Object System.Windows.Forms.Form
    $formPrincipal.Size = New-Object System.Drawing.Size(720, 400)
    $formPrincipal.StartPosition = "CenterScreen"
    $formPrincipal.BackColor = [System.Drawing.Color]::White
    $formPrincipal.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formPrincipal.MaximizeBox = $false
    $formPrincipal.MinimizeBox = $false
    $defaultFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    $boldFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                                                                                                        $version = "Alfa 250207.1346"  # Valor predeterminado para la versión
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
    $tabControl.Size = New-Object System.Drawing.Size(710, 315) #X,Y
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
    $btnInstallSQLManagement = Create-Button -Text "Instalar Management2014" -Location (New-Object System.Drawing.Point(10, 50)) `
                                -ToolTip "Instalación mediante choco de SQL Management 2014."
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
    $btnClearAnyDesk = Create-Button -Text "Clear AnyDesk" -Location (New-Object System.Drawing.Point(240, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 76, 76)) -ToolTip "Detiene el programa y elimina los archivos para crear nuevos IDS."
    $btnShowPrinters = Create-Button -Text "Mostrar Impresoras" -Location (New-Object System.Drawing.Point(240, 90)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Muestra en consola: Impresora, Puerto y Driver instaladas en Windows."
    $btnClearPrintJobs = Create-Button -Text "Limpia y Reinicia Cola de Impresión" -Location (New-Object System.Drawing.Point(240, 130)) `
                                -BackColor ([System.Drawing.Color]::White) -ToolTip "Limpia las impresiones pendientes y reinicia la cola de impresión."
    $btnAplicacionesNS = Create-Button -Text "Aplicaciones National Soft" -Location (New-Object System.Drawing.Point(240, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 200, 150)) -ToolTip "Busca los INIS en el equipo y brinda información de conexión a sus BDDs."
    $btnCheckPermissions = Create-Button -Text "Permisos C:\NationalSoft" -Location (New-Object System.Drawing.Point(470, 50)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Revisa los permisos de los usuarios en la carpeta C:\NationalSoft."
    $btnModificarPermisos = Create-Button -Text "Lector DP - Permisos" -Location (New-Object System.Drawing.Point(470, 90)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Modifica los permisos de la carpeta C:\Windows\System32\en-us."
    $LZMAbtnBuscarCarpeta = Create-Button -Text "Buscar Instalador LZMA" -Location (New-Object System.Drawing.Point(470, 130)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Para el error de instalación, renombra en REGEDIT la carpeta del instalador."
    $btnConfigurarIPs = Create-Button -Text "Agregar IPs" -Location (New-Object System.Drawing.Point(470, 170)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(150, 200, 255)) -ToolTip "Agregar IPS para configurar impresoras en red en segmento diferente."
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
    $btnExit = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(240, 320)) `
                                -BackColor ([System.Drawing.Color]::FromArgb(255, 169, 169, 169))
# Crear el CheckBox chkSqlServer
    $chkSqlServer = New-Object System.Windows.Forms.CheckBox
    $chkSqlServer.Text = "Instalar SQL Tools (opcional)"
    $chkSqlServer.Size = New-Object System.Drawing.Size(290, 30)
    $chkSqlServer.Location = New-Object System.Drawing.Point(10, 10)
    $lblConnectionStatus = Create-Label -Text "Conectado a BDD: Ninguna" -Location (New-Object System.Drawing.Point(10, 260)) -Size (New-Object System.Drawing.Size(290, 30)) `
                                     -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 0, 0)) -Font $defaultFont
# Usar la función Create-Label para crear la label de conexión
    $lblHostname = Create-Label -Text ([System.Net.Dns]::GetHostName()) -Location (New-Object System.Drawing.Point(10, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Hostname al portapapeles."
    $lblPort = Create-Label -Text "Puerto: No disponible" -Location (New-Object System.Drawing.Point(240, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign MiddleCenter -ToolTipText "Haz clic para copiar el Puerto al portapapeles."
    $lbIpAdress = Create-Label -Text "Obteniendo IPs..." -Location (New-Object System.Drawing.Point(470, 1)) -Size (New-Object System.Drawing.Size(220, 40)) `
        -BackColor ([System.Drawing.Color]::FromArgb(255, 0, 0, 0)) -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 255, 255)) -BorderStyle FixedSingle -TextAlign TopLeft -ToolTipText "Haz clic para copiar las IPs al portapapeles."
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
    $tabAplicaciones.Controls.Add($lblHostname)
    $tabAplicaciones.Controls.Add($lblPort)
    $tabAplicaciones.Controls.Add($lbIpAdress)
# Agregar controles a la pestaña Pro
    $tabProSql.Controls.Add($chkSqlServer)
    $tabProSql.Controls.Add($btnReviewPivot)
    $tabProSql.Controls.Add($btnRespaldarRestcard)
    $tabProSql.Controls.Add($btnFechaRevEstaciones)
    $tabProSql.Controls.Add($btnEliminarServidorBDD)
    $tabProSql.Controls.Add($lblConnectionStatus)
    $tabProSql.Controls.Add($btnConnectDb)
    $tabProSql.Controls.Add($btnDisconnectDb)
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

# Función para revisar permisos y agregar Full Control a "Everyone" si es necesario
            function Check-Permissions {
                $folderPath = "C:\NationalSoft"
                $acl = Get-Acl -Path $folderPath
                $permissions = @()
                # Obtener el SID universal de "Everyone"
                $everyoneSid = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::WorldSid, $null)

                $everyonePermissions = @()
                $everyoneHasFullControl = $false
            
                foreach ($access in $acl.Access) {
                    # Almacenar los permisos de todos los usuarios
                    $permissions += [PSCustomObject]@{
                        Usuario = $access.IdentityReference
                        Permiso = $access.FileSystemRights
                        Tipo    = $access.AccessControlType
                    }
                    foreach ($access in $acl.Access) 
                    {
                        # Obtener el SID del usuario en la ACL
                        $userSid = (New-Object System.Security.Principal.NTAccount($access.IdentityReference)).Translate([System.Security.Principal.SecurityIdentifier])
                        # Comparar usando el SID universal de "Everyone"
                        if ($userSid -eq $everyoneSid) 
                        {
                            $everyonePermissions += $access.FileSystemRights
                            if ($access.FileSystemRights -match "FullControl") 
                            {
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
                    Write-Host "`n\tEveryone tiene los siguientes permisos: $($everyonePermissions -join ', ')" -ForegroundColor Green
                } else {
                    Write-Host "`n\tNo hay permisos para 'Everyone'" -ForegroundColor Red
                }
            
                # Si "Everyone" no tiene Full Control o Control total, preguntar si se desea concederlo
                if (-not $everyoneHasFullControl) {
                    $message = "El usuario 'Everyone' no tiene permisos de 'Full Control' (o 'Control total'). ¿Deseas concederlo?"
                    $title = "Permisos 'Everyone'"
                    $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
                    $icon = [System.Windows.Forms.MessageBoxIcon]::Question
            
                    $result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
            
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Agregar Full Control a "Everyone" en la carpeta
                        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                        $acl.AddAccessRule($accessRule)
            
                        # Forzar la herencia para subcarpetas y archivos
                        $acl.SetAccessRuleProtection($false, $true)
            
                        Set-Acl -Path $folderPath -AclObject $acl
                        Write-Host "Se ha concedido 'Full Control' a 'Everyone'." -ForegroundColor Green
                    }
                }
            }            
            # Asignar la función al botón (si sigue siendo necesario)
            $btnCheckPermissions.Add_Click({
                Write-Host "`nRevisando permisos en C:\NationalSoft" -ForegroundColor Yellow
                Check-Permissions
            })            
            # Agregar el botón a la pestaña de Aplicaciones (si sigue siendo necesario)
            $tabAplicaciones.Controls.Add($btnCheckPermissions)


$lblHostname.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($lblHostname.Text)
        Write-Host "`nNombre del equipo copiado al portapapeles: $($lblHostname.Text)"
    })
$lbIpAdress.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($lbIpAdress.Text)
        Write-Host "`nIP's copiadas al equipo: $($lbIpAdress.Text)"
    })
    $lbIpAdress.Add_MouseEnter($changeColorOnHover)
    $lbIpAdress.Add_MouseLeave($restoreColorOnLeave)
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
                            "Dongle $($_.AdapterName) - IP: $($_.IPAddress)`n"
                        }) -join "`n"
                        # Asignar el texto al label
                        $lbIpAdress.Text = "$ipsTextForLabel"
                    } else {
                        $lbIpAdress.Text = "No se encontraron direcciones IP"
                    }
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
            
            $lblAdaptadorStatus = Create-Label -Text $text -Location (New-Object System.Drawing.Point(10, (320 + (30 * $index)))) -Size (New-Object System.Drawing.Size(236, 20)) -ForeColor $color -BorderStyle FixedSingle
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
            $formPrincipal.Controls.Add($btnExit)
# Obtener el puerto de SQL Server desde el registro
        $regKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\NATIONALSOFT\MSSQLServer\SuperSocketNetLib\Tcp"
        $tcpPort = Get-ItemProperty -Path $regKeyPath -Name "TcpPort" -ErrorAction SilentlyContinue

        if ($tcpPort -and $tcpPort.TcpPort) {
            $lblPort.Text = "Puerto SQL \NationalSoft: $($tcpPort.TcpPort)"
        } else {
            $lblPort.Text = "No se encontró puerto o instancia."
        }
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
        $lblPort.Add_MouseEnter($changeColorOnHover)
        $lblPort.Add_MouseLeave($restoreColorOnLeave)
##-------------------- FUNCIONES                                                          -------#
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
    #$btnConnectDb.Enabled = $chkSqlServer.Checked
})
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
        
                # Determinar si la consulta es de tipo SELECT (devuelve resultados) o no (afecta filas)
                if ($query -match "^\s*SELECT") {
                    # Consulta que devuelve resultados (SELECT)
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
        
                    return $results  # Devolver los resultados como una lista de objetos
                } else {
                    # Consulta que no devuelve resultados (UPDATE, DELETE, etc.)
                    $command = $connection.CreateCommand()
                    $command.CommandText = $query
                    $rowsAffected = $command.ExecuteNonQuery()  # Obtener el número de filas afectadas
        
                    # Cerrar la conexión y liberar recursos
                    $connection.Close()
                    $connection.Dispose()
        
                    return $rowsAffected  # Devolver el número de filas afectadas
                }
            } catch {
                Write-Host "Error al ejecutar la consulta: $_" -ForegroundColor Red
                return $null  # Devolver $null en caso de error
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
##-------------------------------------------------------------------------------BOTONES#
$btnSQLManagement.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        
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
            Write-Host "`tOperación cancelada por el usuario." -ForegroundColor Yellow
        }
})
#Profiler:
$btnProfiler.Add_Click({
        Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
        $ProfilerUrl = "https://codeplexarchive.org/codeplex/browse/ExpressProfiler/releases/4/ExpressProfiler22wAddinSigned.zip"
        $ProfilerZipPath = "C:\Temp\ExpressProfiler22wAddinSigned.zip"
        $ExtractPath = "C:\Temp\ExpressProfiler2"
        $ExeName = "ExpressProfiler.exe"
        $ValidationPath = "C:\Temp\ExpressProfiler2\ExpressProfiler.exe"

        DownloadAndRun -url $ProfilerUrl -zipPath $ProfilerZipPath -extractPath $ExtractPath -exeName $ExeName -validationPath $ValidationPath
        if ($disableControls) {        Enable-Controls -parentControl $formPrincipal    }
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
            Write-Host "`tEjecutando SQL Server Configuration Manager..."
            Start-Process "SQLServerManager12.msc"
        }
        catch {
            Write-Host "`tError al ejecutar SQL Server Configuration Manager: $_"  -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("No se pudo abrir SQL Server Configuration Manager.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
            Write-Host "`tRenovación de AnyDesk cancelada por el usuario."
        }
    })
$btnShowPrinters.Add_Click({
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
#LMZA
$LZMAbtnBuscarCarpeta.Add_Click({
         Write-Host "`nComenzando el proceso, por favor espere..." -ForegroundColor Green
            # Usar la ruta correcta del registro
            $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
            
            # Verificar si la ruta existe
            if (-not (Test-Path $LZMAregistryPath)) {
                Write-Host "`nLa ruta del registro no existe: $LZMAregistryPath" -ForegroundColor Yellow
                [System.Windows.Forms.MessageBox]::Show("La ruta del registro no existe: $LZMAregistryPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            try {
                         Write-Host "`t HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"
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
                                $formLZMA = Create-Form -Title "Carpetas LZMA" -Size (New-Object System.Drawing.Size(400, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                                                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false
                                # Crear un ComboBox para mostrar las subcarpetas
                                $LZMcomboBoxCarpetas = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 10)) -Size (New-Object System.Drawing.Size(360, 20)) `
                                       -DropDownStyle DropDownList -Font $defaultFont
                                foreach ($LZMsubCarpeta in $LZMsubCarpetas) {
                                    $LZMcomboBoxCarpetas.Items.Add($LZMsubCarpeta)
                                }
                                # Seleccionar la primera opción por defecto
                                $LZMcomboBoxCarpetas.SelectedIndex = 0
                                # Crear un Label para mostrar el valor de AI_ExePath
                                $lblLZMAExePath = Create-Label -Text "AI_ExePath: -" -Location (New-Object System.Drawing.Point(10, 40)) -Size (New-Object System.Drawing.Size(360, 60)) `
                                     -ForeColor ([System.Drawing.Color]::FromArgb(255, 255, 0, 0)) -Font $defaultFont
                                # Evento cuando se selecciona una subcarpeta en el ComboBox
                                $LZMcomboBoxCarpetas.Add_SelectedIndexChanged({
                                    $indiceSeleccionado = $LZMcomboBoxCarpetas.SelectedIndex
                                    if ($indiceSeleccionado -gt 0) {  # Ignorar la opción por defecto
                                        $LZMrutaCompleta = $LZMrutasCompletas[$indiceSeleccionado - 1]  # Ajustar índice
                                        $valorExePath = Get-ItemProperty -Path $LZMrutaCompleta -Name "AI_ExePath" -ErrorAction SilentlyContinue
                                        if ($valorExePath) {
                                            $lblLZMAExePath.Text = "AI_ExePath: $($valorExePath.AI_ExePath)"
                                        } else {
                                            $lblLZMAExePath.Text = "AI_ExePath: No encontrado"
                                        }
                                    } else {
                                        $lblLZMAExePath.Text = "AI_ExePath: -"
                                    }
                                })
                                # Crear botón para renombrar usando la función Create-Button
                                $LZMbtnRenombrar = Create-Button -Text "Renombrar" -Location (New-Object System.Drawing.Point(10, 100)) -Size (New-Object System.Drawing.Size(140, 30)) -Enabled $false
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
                                $LMZAbtnSalir = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(210, 100)) -Size (New-Object System.Drawing.Size(140, 30))
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
                                $formLZMA.Controls.Add($lblLZMAExePath)
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
#AplicacionesNS
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
#Pivot new
                                    $btnReviewPivot.Add_Click({
                                        try {
                                            if (-not $global:server -or -not $global:database -or -not $global:password) {
                                                Write-Host "`nNo hay una conexión válida." -ForegroundColor Red
                                                return
                                            }
                                    
                                            Write-Host "`nConectado a la base de datos: $global:database en el servidor: $global:server" -ForegroundColor Green
                                    
                                            # Consulta SQL para verificar duplicados
                                            $queryCheckDuplicates = @"
                                            SELECT app_id, field, COUNT(*) AS DuplicateCount
                                            FROM app_settings
                                            GROUP BY app_id, field
                                            HAVING COUNT(*) > 1
"@
                                    
                                            # Mostrar la consulta en la consola en amarillo
                                            Write-Host "`nEjecutando la consulta para verificar duplicados..." -ForegroundColor Yellow
                                            Write-Host "`t$queryCheckDuplicates`n" -ForegroundColor Yellow
                                    
                                            # Ejecutar la consulta para verificar duplicados
                                            $duplicates = Execute-SqlQuery -server $global:server -database $global:database -query $queryCheckDuplicates
                                    
                                            if ($duplicates.Count -eq 0) {
                                                Write-Host "`nNo se encontraron duplicados en la tabla app_settings." -ForegroundColor Green
                                            } else {
                                                Write-Host "`nSe encontraron los siguientes duplicados:" -ForegroundColor Green
                                                foreach ($dup in $duplicates) {
                                                    Write-Host "app_id: $($dup.app_id), field: $($dup.field), Veces duplicado: $($dup.DuplicateCount)" -ForegroundColor Cyan
                                                }
                                    
                                                # Mostrar un MessageBox preguntando si desea eliminar los duplicados
                                                $result = [System.Windows.Forms.MessageBox]::Show("¿Desea eliminar los duplicados mostrados en la consola?", "Eliminar Duplicados", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                                    
                                                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                                                    # Consulta SQL para eliminar duplicados
                                                    $queryDeleteDuplicates = @"
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
"@
                                    
                                                    # Mostrar la consulta en la consola en amarillo
                                                    Write-Host "`nEliminando duplicados..." -ForegroundColor Yellow
                                                    Write-Host "`t$queryDeleteDuplicates" -ForegroundColor Yellow
                                    
                                                    # Ejecutar la consulta para eliminar duplicados
                                                    $rowsAffected = Execute-SqlQuery -server $global:server -database $global:database -query $queryDeleteDuplicates
                                    
                                                    if ($rowsAffected -gt 0) {
                                                        Write-Host "`nDuplicados eliminados correctamente. Filas afectadas: $rowsAffected" -ForegroundColor Green
                                                    } else {
                                                        Write-Host "`nNo se eliminaron duplicados." -ForegroundColor Yellow
                                                    }
                                                } else {
                                                    Write-Host "`nEl usuario decidió no eliminar los duplicados." -ForegroundColor Red
                                                }
                                            }
                                    
                                        } catch {
                                            Write-Host "`nError al ejecutar consulta: $($_.Exception.Message)" -ForegroundColor Red
                                        }
                                    })
#Estaciones new
                                            $btnFechaRevEstaciones.Add_Click({
                                                try {
                                                    if (-not $global:server -or -not $global:database -or -not $global:password) {
                                                        Write-Host "`nNo hay una conexión válida." -ForegroundColor Red
                                                        return
                                                    }
                                                    # Consultas SQL
                                                    $query1 = @"
                                                    SELECT e.FECHAREV, 
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
                                                    ORDER BY b.fecha DESC;
"@
                                                    # Ejecutar y analizar la primera consulta
                                                    $resultsQuery1 = Execute-SqlQuery -server $global:server -database $global:database -query $query1
                                                    Write-Host "`n$query1`n" -ForegroundColor Yellow
                                                    Show-ResultsConsole -query $query1
                                                } catch {
                                                    Write-Host "`nError al ejecutar consulta: $($_.Exception.Message)" -ForegroundColor Red
                                                }
                                            })
#Boton para actualizar los datos del servidor (borrarlo basicamente)
    $btnEliminarServidorBDD.Add_Click({
        $formEliminarServidor = Create-Form -Title "Eliminar Servidor de BDD" -Size (New-Object System.Drawing.Size(400, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))   
            $cmbOpciones = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 20)) -Size (New-Object System.Drawing.Size(360, 20)) `
                               -DropDownStyle DropDownList -Items @("Seleccione una opción", "On The minute", "NS Hoteles", "Rest Card") -SelectedIndex 0
            $btnEliminar = Create-Button -Text "Eliminar" -Location (New-Object System.Drawing.Point(150, 60)) -Size (New-Object System.Drawing.Size(140, 30)) -Enabled $false
            $btnCancelar = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(260, 60)) -Size (New-Object System.Drawing.Size(140, 30))
            $cmbOpciones.Add_SelectedIndexChanged({
                if ($cmbOpciones.SelectedIndex -gt 0) {
                    $btnEliminar.Enabled = $true
                } else {
                    $btnEliminar.Enabled = $false
                }
            })
            $btnEliminar.Add_Click({
                $opcionSeleccionada = $cmbOpciones.SelectedItem
                $confirmacion = [System.Windows.Forms.MessageBox]::Show("¿Está seguro de que desea eliminar el servidor de la base de datos para $opcionSeleccionada?", "Confirmar Eliminación", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($confirmacion -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        # Definir la consulta SQL según la opción seleccionada
                        $query = $null
                        switch ($opcionSeleccionada) {
                            "On The minute" {
                                Write-Host "`tEjecutando Query" -ForegroundColor Yellow
                                $query = "UPDATE configuracion SET serie='', ipserver='', nombreservidor=''"
                                Write-Host "`t$query"
                            }
                            "NS Hoteles" {
                                Write-Host "`tEjecutando Query" -ForegroundColor Yellow
                                $query = "UPDATE configuracion SET serievalida='', numserie='', ipserver='', nombreservidor='', llave=''"
                                Write-Host "`t$query"
                            }
                            "Rest Card" {
                                Write-Host "`nFunción deshabilitada, ejecuta el Query en la base de datos:" -ForegroundColor Yellow
                                Write-Host "`tupdate tabvariables set estacion='', ipservidor='';" -ForegroundColor Yellow
                            }
                        }
            
                        if ($query) {
                            # Ejecutar la consulta y obtener el número de filas afectadas
                            $rowsAffected = Execute-SqlQuery -server $global:server -database $global:database -query $query
            
                            if ($rowsAffected -gt 0) {
                                Write-Host "Servidor de BDD eliminado para $opcionSeleccionada." -ForegroundColor Green
                            } elseif ($rowsAffected -eq 0) {
                                Write-Host "No se encontraron filas para actualizar en la tabla configuracion." -ForegroundColor Yellow
                            } else {
                                Write-Host "No fue posible eliminar el servidor de BDD para $opcionSeleccionada." -ForegroundColor Red
                            }
                        }
                    } catch {
                        Write-Host "Error al eliminar el servidor de BDD: $_" -ForegroundColor Red
                    }
                    $formEliminarServidor.Close()
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
#Boton para conectar a la base de datos
    $btnConnectDb.Add_Click({
            # Crear el formulario para pedir los datos de conexión
                $formBddConnection = Create-Form -Title "Conexión a SQL Server" -Size (New-Object System.Drawing.Size(400, 200)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
    
            # Crear las etiquetas y cajas de texto
            $lblProfile = Create-Label -Text "Perfil de conexión" -Location (New-Object System.Drawing.Point(10, 20)) -Size (New-Object System.Drawing.Size(100, 20)) `
            $cmbProfiles = Create-ComboBox -Location (New-Object System.Drawing.Point(120, 20)) -Size (New-Object System.Drawing.Size(250, 20)) -DropDownStyle DropDownList
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
            $labelServer = Create-Label -Text "Servidor SQL" -Location (New-Object System.Drawing.Point(10, 50)) -Size (New-Object System.Drawing.Size(100, 20)) `
            $txtServer = Create-TextBox -Location (New-Object System.Drawing.Point(120, 50)) -Size (New-Object System.Drawing.Size(250, 20))
            $lblDatabase = Create-Label -Text "Base de Datos" -Location (New-Object System.Drawing.Point(10, 80)) -Size (New-Object System.Drawing.Size(100, 20)) `
            $txtDatabase = Create-TextBox -Location (New-Object System.Drawing.Point(120, 80)) -Size (New-Object System.Drawing.Size(250, 20))
            $lblPassword = Create-Label -Text "Contraseña" -Location (New-Object System.Drawing.Point(10, 110)) -Size (New-Object System.Drawing.Size(100, 20)) `
            $txtPassword = Create-TextBox -Location (New-Object System.Drawing.Point(120, 110)) -Size (New-Object System.Drawing.Size(250, 20)) -UseSystemPasswordChar $true

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
            $btnOK = Create-Button -Text "Conectar" -Location (New-Object System.Drawing.Point(150, 140)) -Size (New-Object System.Drawing.Size(140, 30)) -Enabled $false
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
                    $formBddConnection.Close()
        
                    # Actualizar el texto del label de conexión
                    $lblConnectionStatus.Text = "Conectado a BDD: $($txtDatabase.Text)"
                    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green
                    Write-Host "`t$global:database"
        
                    # Habilitar o deshabilitar botones cuando hay conexiones existosas
                    $btnReviewPivot.Enabled = $true
                    $btnFechaRevEstaciones.Enabled = $true
                    $btnEliminarServidorBDD.Enabled = $true
                    $btnConnectDb.Enabled = $false
                    $btnDisconnectDb.Enabled = $true
        
                } catch {
                    Write-Host "`nError de conexión: $_" -ForegroundColor Red
                    $lblConnectionStatus.Text = "Conexión fallida"
                }
            })
                # Agregar los controles al formulario
                $formBddConnection.Controls.Add($lblProfile)
                $formBddConnection.Controls.Add($cmbProfiles)
                $formBddConnection.Controls.Add($labelServer)
                $formBddConnection.Controls.Add($txtServer)
                $formBddConnection.Controls.Add($lblDatabase)
                $formBddConnection.Controls.Add($txtDatabase)
                $formBddConnection.Controls.Add($lblPassword)
                $formBddConnection.Controls.Add($txtPassword)
                $formBddConnection.Controls.Add($btnOK)
                $formBddConnection.ShowDialog()
            })
#Boton para desconectar de la base de datos
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
            $btnEliminarServidorBDD.Enabled = $false
            $btnReviewPivot.Enabled = $false
        } catch {
            Write-Host "`nError al desconectar: $_" -ForegroundColor Red
        }
    })
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
# ------------------------------ funcion para impresoras
    $btnConfigurarIPs.Add_Click({
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
$btnModificarPermisos.Add_Click({
    Write-Host "`nIniciando modificación de permisos..." -ForegroundColor Cyan

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
# Evento de clic para el botón de respaldo
    $btnRespaldarRestcard.Add_Click({
        Write-Host "En espera de los datos de conexión" -ForegroundColor Gray
        # Crear la segunda ventana para ingresar los datos de conexión
        $formRespaldarRestcard = Create-Form -Title "Datos de Conexión para Respaldar" -Size (New-Object System.Drawing.Size(350, 210)) -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
                -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))    
        # Etiquetas y controles para ingresar la información de conexión
        $lblUsuarioRestcard = Create-Label -Text "Usuario:" -Location (New-Object System.Drawing.Point(20, 40))    
        $txtUsuarioRestcard = Create-TextBox -Location (New-Object System.Drawing.Point(120, 40)) -Size (New-Object System.Drawing.Size(200, 20))
        $lblBaseDeDatosRestcard = Create-Label -Text "Base de Datos:" -Location (New-Object System.Drawing.Point(20, 65))        
        $txtBaseDeDatosRestcard = Create-TextBox -Location (New-Object System.Drawing.Point(120, 65)) -Size (New-Object System.Drawing.Size(200, 20))
        $lblPasswordRestcard = Create-Label -Text "Contraseña:" -Location (New-Object System.Drawing.Point(20, 90))
        $txtPasswordRestcard = Create-TextBox -Location (New-Object System.Drawing.Point(120, 90)) -Size (New-Object System.Drawing.Size(200, 20)) -UseSystemPasswordChar $true
        $lblHostnameRestcard = Create-Label -Text "Hostname:" -Location (New-Object System.Drawing.Point(20, 115))    
        $txtHostnameRestcard = Create-TextBox -Location (New-Object System.Drawing.Point(120, 115)) -Size (New-Object System.Drawing.Size(200, 20))
    
        # Crear el checkbox para llenar los datos por omisión
        $chkLlenarDatos = New-Object System.Windows.Forms.CheckBox
        $chkLlenarDatos.Text = "Usar los datos por omisión"
        $chkLlenarDatos.Location = New-Object System.Drawing.Point(5, 20)
        $chkLlenarDatos.AutoSize = $true
    
        # Evento de cambio de estado del checkbox
        $chkLlenarDatos.Add_CheckedChanged({
            if ($chkLlenarDatos.Checked) {
                # Llenar los datos por omisión
                $txtUsuarioRestcard.Text = "root"
                $txtBaseDeDatosRestcard.Text = "restcard"
                $txtPasswordRestcard.Text = "national"
                $txtHostnameRestcard.Text = "localhost"
            } else {
                # Limpiar los campos
                $txtUsuarioRestcard.Clear()
                $txtBaseDeDatosRestcard.Clear()
                $txtPasswordRestcard.Clear()
                $txtHostnameRestcard.Clear()
            }
        })
    
        # Crear botón para ejecutar el respaldo
            $btnRespaldar = Create-Button -Text "Respaldar" -Location (New-Object System.Drawing.Point(20, 140))  -Size (New-Object System.Drawing.Size(140, 30))
    
        # Evento de clic para el botón de respaldo
                    $btnRespaldar.Add_Click({
                        # Obtener los valores del formulario
                        $usuarioRestcard = $txtUsuarioRestcard.Text
                        $baseDeDatosRestcard = $txtBaseDeDatosRestcard.Text
                        $passwordRestcard = $txtPasswordRestcard.Text
                        $hostnameRestcard = $txtHostnameRestcard.Text
                    
                        # Validar que la información no esté vacía
                        if ($usuarioRestcard -eq "" -or $baseDeDatosRestcard -eq "" -or $passwordRestcard -eq "" -or $hostnameRestcard -eq "") {
                            [System.Windows.Forms.MessageBox]::Show("Por favor, complete toda la información.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            return
                        }
                    
                        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
                        $folderDialog.Description = "Selecciona la carpeta donde guardar el respaldo"
                        Write-Host "Selecciona la carpeta donde guardar el respaldo" -ForegroundColor Yellow
                        if ($folderDialog.ShowDialog() -eq "OK") {
                            # Obtener la ruta seleccionada
                            Write-Host "Realizando respaldo para la base de datos." -ForegroundColor Green
                            Write-Host "`tBase de datos:`t $baseDeDatosRestcard"
                            Write-Host "`tEn el servidor:`t $hostnameRestcard"
                            Write-Host "`tCon el usuario:`t $usuarioRestcard"
                            $folderPath = $folderDialog.SelectedPath
                            # Crear la ruta con el timestamp
                            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                            $rutaRespaldo = "$folderPath\respaldo_restcard_$timestamp.sql"
                    
                            # Ejecutar el comando mysqldump para hacer el respaldo
                            $argumentos = "-u $usuarioRestcard -p$passwordRestcard -h $hostnameRestcard $baseDeDatosRestcard --result-file=`"$rutaRespaldo`""
                            $process = Start-Process -FilePath "mysqldump" -ArgumentList $argumentos -NoNewWindow -Wait -PassThru
                    
                            # Verificar si el respaldo se realizó correctamente
                            if ($process.ExitCode -eq 0) {
                                # Obtener el tamaño del archivo generado
                                $tamañoArchivo = (Get-Item $rutaRespaldo).Length / 1KB
                                $tamañoArchivo = [math]::Round($tamañoArchivo, 2)
                                Write-Host "Respaldo completado correctamente. Tamaño del archivo: $tamañoArchivo KB" -ForegroundColor Green
                            } else {
                                # Mostrar el error en rojo con tabulación
                                Write-Host "`tError: No se pudo realizar el respaldo. Verifica los datos de conexión y la base de datos." -ForegroundColor Red
                                # Eliminar el archivo de respaldo si el proceso falló
                                if (Test-Path $rutaRespaldo) {
                                    Remove-Item $rutaRespaldo
                                    Write-Host "`tArchivo de respaldo eliminado debido a un error." -ForegroundColor Yellow
                                }
                            }
                    
                            # Cerrar la segunda ventana después de completar el respaldo
                            $formRespaldarRestcard.Close()
                        }
                    })
    
        # Crear botón para salir
        $btnSalirRestcard = Create-Button -Text "Salir" -Location (New-Object System.Drawing.Point(185, 140))  -Size (New-Object System.Drawing.Size(140, 30))
    
        # Evento de clic para el botón de salir
        $btnSalirRestcard.Add_Click({
            Write-Host "`tSalió sin realizar respaldo." -ForegroundColor Red
            $formRespaldarRestcard.Close()
        })
    
        # Agregar controles a la segunda ventana
        $formRespaldarRestcard.Controls.Add($txtUsuarioRestcard)
        $formRespaldarRestcard.Controls.Add($txtBaseDeDatosRestcard)
        $formRespaldarRestcard.Controls.Add($txtPasswordRestcard)
        $formRespaldarRestcard.Controls.Add($txtHostnameRestcard)
        $formRespaldarRestcard.Controls.Add($lblUsuarioRestcard)
        $formRespaldarRestcard.Controls.Add($lblBaseDeDatosRestcard)
        $formRespaldarRestcard.Controls.Add($lblPasswordRestcard)
        $formRespaldarRestcard.Controls.Add($lblHostnameRestcard)
        $formRespaldarRestcard.Controls.Add($chkLlenarDatos)
        $formRespaldarRestcard.Controls.Add($btnRespaldar)
        $formRespaldarRestcard.Controls.Add($btnSalirRestcard)
    
        # Mostrar la segunda ventana
        $formRespaldarRestcard.ShowDialog()
    })
#Boton para salir
    $btnExit.Add_Click({
        $formPrincipal.Dispose()
        $formPrincipal.Close()
    })
$formPrincipal.Refresh()
$formPrincipal.ShowDialog()
