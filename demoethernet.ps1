Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-NewIpForm {
    $ipAssignForm = New-Object System.Windows.Forms.Form
    $ipAssignForm.Text = "Agregar IP Adicional"
    $ipAssignForm.Size = New-Object System.Drawing.Size(350, 150)
    $ipAssignForm.StartPosition = "CenterScreen"

    $ipAssignLabel = New-Object System.Windows.Forms.Label
    $ipAssignLabel.Text = "Ingrese la nueva dirección IP:"
    $ipAssignLabel.Location = New-Object System.Drawing.Point(10, 20)
    $ipAssignLabel.AutoSize = $true
    $ipAssignForm.Controls.Add($ipAssignLabel)

    $ipAssignTextBox1 = New-Object System.Windows.Forms.TextBox
    $ipAssignTextBox1.Location = New-Object System.Drawing.Point(10, 50)
    $ipAssignTextBox1.Size = New-Object System.Drawing.Size(50, 20)
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
    $ipAssignForm.Controls.Add($ipAssignTextBox1)

    $ipAssignLabelDot1 = New-Object System.Windows.Forms.Label
    $ipAssignLabelDot1.Text = "."
    $ipAssignLabelDot1.Location = New-Object System.Drawing.Point(65, 53)
    $ipAssignLabelDot1.AutoSize = $true
    $ipAssignForm.Controls.Add($ipAssignLabelDot1)

    $ipAssignTextBox2 = New-Object System.Windows.Forms.TextBox
    $ipAssignTextBox2.Location = New-Object System.Drawing.Point(80, 50)
    $ipAssignTextBox2.Size = New-Object System.Drawing.Size(50, 20)
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
    $ipAssignForm.Controls.Add($ipAssignTextBox2)

    $ipAssignLabelDot2 = New-Object System.Windows.Forms.Label
    $ipAssignLabelDot2.Text = "."
    $ipAssignLabelDot2.Location = New-Object System.Drawing.Point(135, 53)
    $ipAssignLabelDot2.AutoSize = $true
    $ipAssignForm.Controls.Add($ipAssignLabelDot2)

    $ipAssignTextBox3 = New-Object System.Windows.Forms.TextBox
    $ipAssignTextBox3.Location = New-Object System.Drawing.Point(150, 50)
    $ipAssignTextBox3.Size = New-Object System.Drawing.Size(50, 20)
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
    $ipAssignForm.Controls.Add($ipAssignTextBox3)

    $ipAssignLabelDot3 = New-Object System.Windows.Forms.Label
    $ipAssignLabelDot3.Text = "."
    $ipAssignLabelDot3.Location = New-Object System.Drawing.Point(205, 53)
    $ipAssignLabelDot3.AutoSize = $true
    $ipAssignForm.Controls.Add($ipAssignLabelDot3)

    $ipAssignTextBox4 = New-Object System.Windows.Forms.TextBox
    $ipAssignTextBox4.Location = New-Object System.Drawing.Point(220, 50)
    $ipAssignTextBox4.Size = New-Object System.Drawing.Size(50, 20)
    $ipAssignTextBox4.MaxLength = 3
    $ipAssignTextBox4.Add_KeyPress({
        if (-not [char]::IsDigit($_.KeyChar) -and $_.KeyChar -ne 8) { $_.Handled = $true }
    })
    $ipAssignForm.Controls.Add($ipAssignTextBox4)

    $ipAssignButton = New-Object System.Windows.Forms.Button
    $ipAssignButton.Text = "Aceptar"
    $ipAssignButton.Location = New-Object System.Drawing.Point(100, 80)
    $ipAssignButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ipAssignForm.AcceptButton = $ipAssignButton
    $ipAssignForm.Controls.Add($ipAssignButton)

    $result = $ipAssignForm.ShowDialog()

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

$ipAssignFormDemo = New-Object System.Windows.Forms.Form
$ipAssignFormDemo.Text = "Configuración de IP"
$ipAssignFormDemo.Size = New-Object System.Drawing.Size(400, 300)
$ipAssignFormDemo.StartPosition = "CenterScreen"

$ipAssignButtonAsignacion = New-Object System.Windows.Forms.Button
$ipAssignButtonAsignacion.Text = "Asignación de IPs"
$ipAssignButtonAsignacion.Location = New-Object System.Drawing.Point(10, 20)
$ipAssignButtonAsignacion.Size = New-Object System.Drawing.Size(120, 30)
$ipAssignButtonAsignacion.Add_Click({
    $ipAssignFormAsignacion = New-Object System.Windows.Forms.Form
    $ipAssignFormAsignacion.Text = "Asignación de IPs"
    $ipAssignFormAsignacion.Size = New-Object System.Drawing.Size(400, 300)
    $ipAssignFormAsignacion.StartPosition = "CenterScreen"

    $ipAssignLabelAdapter = New-Object System.Windows.Forms.Label
    $ipAssignLabelAdapter.Text = "Seleccione el adaptador de red:"
    $ipAssignLabelAdapter.Location = New-Object System.Drawing.Point(10, 20)
    $ipAssignLabelAdapter.AutoSize = $true
    $ipAssignFormAsignacion.Controls.Add($ipAssignLabelAdapter)

    $ipAssignComboBoxAdapters = New-Object System.Windows.Forms.ComboBox
    $ipAssignComboBoxAdapters.Location = New-Object System.Drawing.Point(10, 50)
    $ipAssignComboBoxAdapters.Size = New-Object System.Drawing.Size(360, 20)
    $ipAssignComboBoxAdapters.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $ipAssignComboBoxAdapters.Text = "Selecciona 1 adaptador de red"
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    foreach ($adapter in $adapters) {
        $ipAssignComboBoxAdapters.Items.Add($adapter.Name)
    }
    $ipAssignFormAsignacion.Controls.Add($ipAssignComboBoxAdapters)

    $ipAssignLabelIps = New-Object System.Windows.Forms.Label
    $ipAssignLabelIps.Text = "IPs asignadas:"
    $ipAssignLabelIps.Location = New-Object System.Drawing.Point(10, 80)
    $ipAssignLabelIps.AutoSize = $true
    $ipAssignFormAsignacion.Controls.Add($ipAssignLabelIps)

    $ipAssignButtonAssignIP = New-Object System.Windows.Forms.Button
    $ipAssignButtonAssignIP.Text = "Asignar Nueva IP"
    $ipAssignButtonAssignIP.Location = New-Object System.Drawing.Point(10, 120)
    $ipAssignButtonAssignIP.Size = New-Object System.Drawing.Size(120, 30)
    $ipAssignButtonAssignIP.Add_Click({
        $selectedAdapterName = $ipAssignComboBoxAdapters.SelectedItem
        if ($selectedAdapterName -eq "Selecciona 1 adaptador de red") {
            Write-Host "Por favor, selecciona un adaptador de red." -ForegroundColor Red
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
                Write-Host "El adaptador ya tiene una IP fija. ¿Desea agregar una nueva IP?" -ForegroundColor Yellow
                $confirmation = [System.Windows.Forms.MessageBox]::Show("El adaptador ya tiene una IP fija. ¿Desea agregar una nueva IP?", "Confirmación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                    $newIp = Show-NewIpForm
                    if ($newIp) {
                        $existingIp = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.IPAddress -eq $newIp }
                        if ($existingIp) {
                            Write-Host "La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name)." -ForegroundColor Red
                            [System.Windows.Forms.MessageBox]::Show("La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name).", "Error")
                        } else {
                            try {
                                New-NetIPAddress -IPAddress $newIp -PrefixLength $currentPrefixLength -InterfaceAlias $selectedAdapter.Name
                                Write-Host "Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                [System.Windows.Forms.MessageBox]::Show("Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name).", "Éxito")

                                # Actualizar la lista de IPs asignadas
                                $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                                $ips = $currentIPs.IPAddress -join ", "
                                $ipAssignLabelIps.Text = "IPs asignadas: $ips"
                            } catch {
                                Write-Host "Error al agregar la dirección IP adicional: $($_.Exception.Message)" -ForegroundColor Red
                                [System.Windows.Forms.MessageBox]::Show("Error al agregar la dirección IP adicional: $($_.Exception.Message)", "Error")
                            }
                        }
                    }
                }
            } else {
                Write-Host "¿Desea cambiar a IP fija usando la IP actual ($currentIPAddress)?" -ForegroundColor Yellow
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

                        Write-Host "Se cambió a IP fija $currentIPAddress en el adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                        [System.Windows.Forms.MessageBox]::Show("Se cambió a IP fija $currentIPAddress en el adaptador $($selectedAdapter.Name).", "Éxito")

                        # Actualizar la lista de IPs asignadas
                        $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                        $ips = $currentIPs.IPAddress -join ", "
                        $ipAssignLabelIps.Text = "IPs asignadas: $ips"

                        Write-Host "¿Desea agregar una dirección IP adicional?" -ForegroundColor Yellow
                        $confirmationAdditionalIP = [System.Windows.Forms.MessageBox]::Show("¿Desea agregar una dirección IP adicional?", "IP Adicional", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                        if ($confirmationAdditionalIP -eq [System.Windows.Forms.DialogResult]::Yes) {
                            $newIp = Show-NewIpForm
                            if ($newIp) {
                                $existingIp = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.IPAddress -eq $newIp }
                                if ($existingIp) {
                                    Write-Host "La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name)." -ForegroundColor Red
                                    [System.Windows.Forms.MessageBox]::Show("La dirección IP $newIp ya está asignada al adaptador $($selectedAdapter.Name).", "Error")
                                } else {
                                    try {
                                        New-NetIPAddress -IPAddress $newIp -PrefixLength $currentPrefixLength -InterfaceAlias $selectedAdapter.Name
                                        Write-Host "Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                                        [System.Windows.Forms.MessageBox]::Show("Se agregó la dirección IP adicional $newIp al adaptador $($selectedAdapter.Name).", "Éxito")

                                        # Actualizar la lista de IPs asignadas
                                        $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                                        $ips = $currentIPs.IPAddress -join ", "
                                        $ipAssignLabelIps.Text = "IPs asignadas: $ips"
                                    } catch {
                                        Write-Host "Error al agregar la dirección IP adicional: $($_.Exception.Message)" -ForegroundColor Red
                                        [System.Windows.Forms.MessageBox]::Show("Error al agregar la dirección IP adicional: $($_.Exception.Message)", "Error")
                                    }
                                }
                            }
                        }
                    } catch {
                        Write-Host "Error al cambiar a IP fija: $($_.Exception.Message)" -ForegroundColor Red
                        [System.Windows.Forms.MessageBox]::Show("Error al cambiar a IP fija: $($_.Exception.Message)", "Error")
                    }
                }
            }
        } else {
            Write-Host "No se pudo obtener la configuración actual del adaptador." -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("No se pudo obtener la configuración actual del adaptador.", "Error")
        }
    })
    $ipAssignFormAsignacion.Controls.Add($ipAssignButtonAssignIP)

    $ipAssignButtonChangeToDhcp = New-Object System.Windows.Forms.Button
    $ipAssignButtonChangeToDhcp.Text = "Cambiar a DHCP"
    $ipAssignButtonChangeToDhcp.Location = New-Object System.Drawing.Point(140, 120)
    $ipAssignButtonChangeToDhcp.Size = New-Object System.Drawing.Size(120, 30)
    $ipAssignButtonChangeToDhcp.Add_Click({
        $selectedAdapterName = $ipAssignComboBoxAdapters.SelectedItem
        if ($selectedAdapterName -eq "Selecciona 1 adaptador de red") {
            Write-Host "Por favor, selecciona un adaptador de red." -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("Por favor, selecciona un adaptador de red.", "Error")
            return
        }
        $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
        $currentConfig = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue

        if ($currentConfig) {
            $isDhcp = ($currentConfig.PrefixOrigin -eq "Dhcp")
            if ($isDhcp) {
                Write-Host "El adaptador ya está en DHCP." -ForegroundColor Yellow
                [System.Windows.Forms.MessageBox]::Show("El adaptador ya está en DHCP.", "Información")
            } else {
                Write-Host "¿Está seguro de que desea cambiar a DHCP?" -ForegroundColor Yellow
                $confirmation = [System.Windows.Forms.MessageBox]::Show("¿Está seguro de que desea cambiar a DHCP?", "Confirmación", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -eq "Manual" }
                        foreach ($ip in $currentIPs) {
                            Remove-NetIPAddress -IPAddress $ip.IPAddress -PrefixLength $ip.PrefixLength -Confirm:$false -ErrorAction SilentlyContinue
                        }
                        Set-NetIPInterface -InterfaceAlias $selectedAdapter.Name -Dhcp Enabled
                        Set-DnsClientServerAddress -InterfaceAlias $selectedAdapter.Name -ResetServerAddresses
                        Write-Host "Se cambió a DHCP en el adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                        [System.Windows.Forms.MessageBox]::Show("Se cambió a DHCP en el adaptador $($selectedAdapter.Name).", "Éxito")

                        # Actualizar la lista de IPs asignadas
                        $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
                        $ips = $currentIPs.IPAddress -join ", "
                        $ipAssignLabelIps.Text = "IPs asignadas: $ips"
                    } catch {
                        Write-Host "Error al cambiar a DHCP: $($_.Exception.Message)" -ForegroundColor Red
                        [System.Windows.Forms.MessageBox]::Show("Error al cambiar a DHCP: $($_.Exception.Message)", "Error")
                    }
                }
            }
        } else {
            Write-Host "No se pudo obtener la configuración actual del adaptador." -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("No se pudo obtener la configuración actual del adaptador.", "Error")
        }
    })
    $ipAssignFormAsignacion.Controls.Add($ipAssignButtonChangeToDhcp)

    $ipAssignComboBoxAdapters.Add_SelectedIndexChanged({
        $selectedAdapterName = $ipAssignComboBoxAdapters.SelectedItem
        if ($selectedAdapterName -ne "Selecciona 1 adaptador de red") {
            $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
            $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4
            $ips = $currentIPs.IPAddress -join ", "
            $ipAssignLabelIps.Text = "IPs asignadas: $ips"
        } else {
            $ipAssignLabelIps.Text = "IPs asignadas:"
        }
    })

    $ipAssignFormAsignacion.ShowDialog()
})
$ipAssignFormDemo.Controls.Add($ipAssignButtonAsignacion)

$ipAssignButtonClose = New-Object System.Windows.Forms.Button
$ipAssignButtonClose.Text = "Cerrar"
$ipAssignButtonClose.Location = New-Object System.Drawing.Point(10, 60)
$ipAssignButtonClose.Size = New-Object System.Drawing.Size(120, 30)
$ipAssignButtonClose.Add_Click({
    Write-Host "Cerrando la aplicación..." -ForegroundColor Yellow
    $ipAssignFormDemo.Close()
})
$ipAssignFormDemo.Controls.Add($ipAssignButtonClose)

$ipAssignFormDemo.ShowDialog()