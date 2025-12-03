

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

#btnInstallSQLManagement Click
$btnInstallSQLManagement.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray

        # Verificar Chocolatey
        if (!(Check-Chocolatey)) { return }

        # Mostrar selector
        $choice = Show-SSMSInstallerDialog
        if (-not $choice) {
            Write-Host "`nInstalación cancelada por el usuario." -ForegroundColor Yellow
            return
        }

        # Confirmación
        $texto = if ($choice -eq "latest") {
            "¿Desea instalar el SSMS 'Último disponible'?"
        } else {
            "¿Desea instalar SSMS 14 (2014)?"
        }
        $response = [System.Windows.Forms.MessageBox]::Show(
            $texto, "Advertencia de instalación",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($response -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "`nEl usuario canceló la instalación." -ForegroundColor Red
            return
        }

        try {
            Write-Host "`nConfigurando Chocolatey..." -ForegroundColor Yellow
            choco config set cacheLocation C:\Choco\cache | Out-Null

            if ($choice -eq "latest") {
                Write-Host "`nInstalando 'Último disponible' (sql-server-management-studio)..." -ForegroundColor Cyan
                Start-Process choco -ArgumentList 'install sql-server-management-studio -y' -NoNewWindow -Wait
                [System.Windows.Forms.MessageBox]::Show(
                    "SSMS (último disponible) instalado correctamente.",
                    "Éxito",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            } elseif ($choice -eq "ssms14") {
                Write-Host "`nInstalando SSMS 14 (mssqlservermanagementstudio2014express)..." -ForegroundColor Cyan
                Start-Process choco -ArgumentList 'install mssqlservermanagementstudio2014express -y' -NoNewWindow -Wait
                [System.Windows.Forms.MessageBox]::Show(
                    "SSMS 14 (2014) instalado correctamente.",
                    "Éxito",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                ) | Out-Null
            }
            Write-Host "`nInstalación completada." -ForegroundColor Green
        } catch {
            Write-Host "`nOcurrió un error durante la instalación: $_" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "Error al instalar SSMS: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
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


$btnReloadConnections = Create-Button -Text "Recargar Conexiones" -Location (New-Object System.Drawing.Point(10, 180)) `
    -Size (New-Object System.Drawing.Size(180, 30)) `
    -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 255)) `
    -ToolTip "Recargar la lista de conexiones desde archivos INI"
$btnReloadConnections.Add_Click({
        Write-Host "Recargando conexiones desde archivos INI..." -ForegroundColor Cyan
        Load-IniConnectionsToComboBox
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

        # Lista de rutas principales con los archivos .ini correspondientes
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
            } elseif (Test-Path $inisFolder) {
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
                    $securePwd = ConvertTo-SecureString $password -AsPlainText -Force
                    New-LocalUser -Name $username -Password $securePwd -AccountNeverExpires -PasswordNeverExpires
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
$btnExecute.Add_Click({
        Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray
        try {
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
            if ($result.Messages.Count -gt 0) {
                Write-Host "`nMensajes de SQL:" -ForegroundColor Cyan
                $result.Messages | ForEach-Object { Write-Host $_ }
            }
            if ($result.DataTable) {
                $dgvResults.DataSource = $result.DataTable.DefaultView
                $dgvResults.Enabled = $true
                Write-Host "`nColumnas obtenidas: $($result.DataTable.Columns.ColumnName -join ', ')" -ForegroundColor Cyan
                $dgvResults.DefaultCellStyle.ForeColor = 'Blue'
                $dgvResults.AlternatingRowsDefaultCellStyle.BackColor = '#F0F8FF'
                $dgvResults.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
                foreach ($col in $dgvResults.Columns) {
                    $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::DisplayedCells
                    $col.Width = [Math]::Max($col.Width, $col.HeaderText.Length * 8)
                }
                if ($result.DataTable.Rows.Count -eq 0) {
                    Write-Host "La consulta no devolvió resultados" -ForegroundColor Yellow
                } else {
                    $result.DataTable | Format-Table -AutoSize | Out-String | Write-Host
                }
            } else {
                Write-Host "`nFilas afectadas: $($result.RowsAffected)" -ForegroundColor Green
            }
        } catch {
            $errorTable = New-Object System.Data.DataTable
            $errorTable.Columns.Add("Tipo") | Out-Null
            $errorTable.Columns.Add("Mensaje") | Out-Null
            $errorTable.Columns.Add("Detalle") | Out-Null
            $cleanQuery = $rtbQuery.Text -replace '(?s)/\*.*?\*/', '' -replace '(?m)^\s*--.*'
            $shortQuery = if ($cleanQuery.Length -gt 50) { $cleanQuery.Substring(0, 47) + "..." } else { $cleanQuery }
            $errorTable.Rows.Add("ERROR SQL", $_.Exception.Message, $shortQuery) | Out-Null
            $dgvResults.DataSource = $errorTable
            $dgvResults.Columns[1].DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
            $dgvResults.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
            $dgvResults.AutoSizeColumnsMode = 'Fill'
            $dgvResults.Columns[0].Width = 100
            $dgvResults.Columns[1].Width = 300
            $dgvResults.Columns[2].Width = 200
            $dgvResults.DefaultCellStyle.ForeColor = 'Red'
            $dgvResults.ColumnHeadersDefaultCellStyle.BackColor = '#FFB3B3'
            $toolTip.SetToolTip($dgvResults, "Consulta completa:`n$cleanQuery")
            Write-Host "`n=============== ERROR ==============" -ForegroundColor Red
            Write-Host "Mensaje: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Consulta: $shortQuery" -ForegroundColor Cyan
            Write-Host "====================================" -ForegroundColor Red
        }
    })
$btnConnectDb.Add_Click({
        Write-Host "`nConectando a la instancia..." -ForegroundColor Gray
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
            $txtServer.Enabled = $false
            $txtUser.Enabled = $false
            $txtPassword.Enabled = $false
            $btnExecute.Enabled = $true
            $cmbQueries.Enabled = $true
            $btnConnectDb.Enabled = $false
            $btnBackup.Enabled = $True
            $btnDisconnectDb.Enabled = $true
            $btnExecute.Enabled = $true
            $rtbQuery.Enabled = $true
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error de conexión: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-Host "Error | Error de conexión: $($_.Exception.Message)" -ForegroundColor Red
        }
    })
$cmbDatabases.Add_SelectedIndexChanged({
        if ($cmbDatabases.SelectedItem) {
            $global:database = $cmbDatabases.SelectedItem
            $lblConnectionStatus.Text = @"
Conectado a:
Servidor: $($global:server)
Base de datos: $($global:database)
"@.Trim()
            $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Green

            Write-Host "`nBase de datos seleccionada:`t $($cmbDatabases.SelectedItem)" -ForegroundColor Cyan
        }
    })
$btnDisconnectDb.Add_Click({
        try {
            Write-Host "`nDesconexión exitosa" -ForegroundColor Yellow
            $global:connection.Close()
            $lblConnectionStatus.Text = "Conectado a BDD: Ninguna"
            $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
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
            $cmbDatabases.Items.Clear()
            $cmbDatabases.Enabled = $false
        } catch {
            Write-Host "`nError al desconectar: $_" -ForegroundColor Red
        }
    })
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
Refresh-AdapterStatus
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

$formPrincipal.Refresh()
$formPrincipal.ShowDialog()