# ------------------------------ Boton para configurar nuevas ips
$btnReloadConnections = Create-Button -Text "Recargar Conexiones" -Location (New-Object System.Drawing.Point(10, 180)) `
    -Size (New-Object System.Drawing.Size(180, 30)) `
    -BackColor ([System.Drawing.Color]::FromArgb(200, 230, 255)) `
    -ToolTip "Recargar la lista de conexiones desde archivos INI"
$btnReloadConnections.Add_Click({
        Write-Host "Recargando conexiones desde archivos INI..." -ForegroundColor Cyan
        Load-IniConnectionsToComboBox
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