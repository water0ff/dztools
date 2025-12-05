#requires -Version 5.0
$script:toolTip = New-Object System.Windows.Forms.ToolTip
function Create-Form {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(350, 200)),
        [Parameter()]
        [System.Windows.Forms.FormStartPosition]$StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen,
        [Parameter()]
        [System.Windows.Forms.FormBorderStyle]$FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog,
        [Parameter()]
        [bool]$MaximizeBox = $false,
        [Parameter()]
        [bool]$MinimizeBox = $false,
        [Parameter()]
        [bool]$TopMost = $false,
        [Parameter()]
        [bool]$ControlBox = $true,
        [Parameter()]
        [System.Drawing.Icon]$Icon = $null,
        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.SystemColors]::Control
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = $Size
    $form.StartPosition = $StartPosition
    $form.FormBorderStyle = $FormBorderStyle
    $form.MaximizeBox = $MaximizeBox
    $form.MinimizeBox = $MinimizeBox
    $form.TopMost = $TopMost
    $form.ControlBox = $ControlBox
    if ($Icon) {
        $form.Icon = $Icon
    }
    $form.BackColor = $BackColor
    return $form
}
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
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Size = $Size
    $label.Location = $Location
    $label.BackColor = $BackColor
    $label.ForeColor = $ForeColor
    $label.Font = $Font
    $label.BorderStyle = $BorderStyle
    $label.TextAlign = $TextAlign
    if ($ToolTipText) { $script:toolTip.SetToolTip($label, $ToolTipText) }
    return $label
}
function Create-Button {
    param (
        [string]$Text,
        [System.Drawing.Point]$Location,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
        [string]$ToolTipText = $null,
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35)),
        [System.Drawing.Font]$Font = $defaultFont, # Agregar parámetro Font con valor predeterminado
        [bool]$Enabled = $true
    )
    $buttonStyle = @{
        FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
        Font      = $defaultFont
    }
    $button_MouseEnter = {
        $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)  # Cambia el color al pasar el mouse
        $this.Font = $boldFont
    }
    $button_MouseLeave = {
        $this.BackColor = $this.Tag  # Restaura el color original almacenado en Tag
        $this.Font = $defaultFont
    }
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = $Size  # Usar el tamaño personalizado
    $button.Location = $Location
    $button.BackColor = $BackColor
    $button.ForeColor = $ForeColor
    $button.Font = $Font # Usar el parámetro Font
    $button.FlatStyle = $buttonStyle.FlatStyle
    $button.Tag = $BackColor  # Almacena el color original en Tag
    $button.Add_MouseEnter($button_MouseEnter)
    $button.Add_MouseLeave($button_MouseLeave)
    $button.Enabled = $Enabled
    if ($ToolTipText) {
        $script:toolTip.SetToolTip($button, $ToolTipText)
    }
    if ($PSBoundParameters.ContainsKey('DialogResult')) {
        $button.DialogResult = $DialogResult
    }
    return $button
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
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = $Location
    $comboBox.Size = $Size
    $comboBox.DropDownStyle = $DropDownStyle
    $comboBox.Font = $Font
    if ($Items.Count -gt 0) {
        $comboBox.Items.AddRange($Items)
        $comboBox.SelectedIndex = $SelectedIndex
    }
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
    $textBox.WordWrap = $false
    if ($UseSystemPasswordChar) {
        $textBox.UseSystemPasswordChar = $true
    }
    return $textBox
}
function Show-ProgressBar {
    # Crear formulario de progreso
    $sizeProgress = New-Object System.Drawing.Size(400, 150)
    $formProgress = Create-Form `
        -Title "Progreso" `
        -Size $sizeProgress `
        -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
        -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) `
        -TopMost $true `
        -ControlBox $false  # Cambiado a false para evitar que el usuario lo cierre
    # Barra de progreso
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size(360, 20)
    $progressBar.Location = New-Object System.Drawing.Point(10, 50)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $progressBar.Maximum = 100
    # Label de porcentaje
    $lblPercentage = New-Object System.Windows.Forms.Label
    $lblPercentage.Location = New-Object System.Drawing.Point(10, 20)
    $lblPercentage.Size = New-Object System.Drawing.Size(360, 20)
    $lblPercentage.Text = "0% Completado"
    $lblPercentage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    # Label de estado (nuevo)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(10, 80)
    $lblStatus.Size = New-Object System.Drawing.Size(360, 40)
    $lblStatus.Text = "Iniciando proceso..."
    $lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
    # Agregar controles
    $formProgress.Controls.Add($progressBar)
    $formProgress.Controls.Add($lblPercentage)
    $formProgress.Controls.Add($lblStatus)
    # Exponer las propiedades para Update-ProgressBar
    $formProgress | Add-Member -MemberType NoteProperty -Name ProgressBar -Value $progressBar -Force
    $formProgress | Add-Member -MemberType NoteProperty -Name Label       -Value $lblPercentage -Force
    $formProgress | Add-Member -MemberType NoteProperty -Name StatusLabel -Value $lblStatus -Force
    $formProgress.Show()
    $formProgress.Refresh()  # Forzar actualización inicial
    return $formProgress
}
function Update-ProgressBar {
    param(
        $ProgressForm,
        $CurrentStep,
        $TotalSteps,
        [string]$Status = ""
    )
    if ($null -eq $ProgressForm -or $ProgressForm.IsDisposed) {
        return
    }
    try {
        $percent = [math]::Round(($CurrentStep / $TotalSteps) * 100)
        # Actualizar barra de progreso
        $ProgressForm.ProgressBar.Value = $percent
        # Actualizar label de porcentaje
        $ProgressForm.Label.Text = "$percent% Completado"
        # Actualizar status si se proporcionó y el control existe
        if ($Status -ne "") {
            if ($ProgressForm.PSObject.Properties.Name -contains 'StatusLabel') {
                $ProgressForm.StatusLabel.Text = $Status
            }
        }
        # IMPORTANTE: Múltiples métodos para forzar actualización de UI
        $ProgressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        # Pequeño delay para permitir que la UI se actualice
        Start-Sleep -Milliseconds 100
    } catch {
        # Silenciar errores en modo release, solo mostrar en debug
        Write-DzDebug "`t[DEBUG]Update-ProgressBar: Error: $($_.Exception.Message)" Red
    }
}
function Close-ProgressBar {
    param($ProgressForm)

    if ($null -eq $ProgressForm -or $ProgressForm.IsDisposed) {
        return
    }
    try {
        $ProgressForm.Close()
        $ProgressForm.Dispose()
    } catch {
        Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)"
    }
}
function Show-SSMSInstallerDialog {
    $form = Create-Form -Title "Instalar SSMS" `
        -Size (New-Object System.Drawing.Size(360, 180)) `
        -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
        -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) `
        -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
    $lbl = Create-Label -Text "Elige la versión a instalar:" -Location (New-Object System.Drawing.Point(10, 15)) -Size (New-Object System.Drawing.Size(320, 20))
    $cmb = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 40)) -Size (New-Object System.Drawing.Size(320, 22)) -DropDownStyle DropDownList
    $null = $cmb.Items.Add("Último disponible.")
    $null = $cmb.Items.Add("SSMS 14 (2014)")
    $cmb.SelectedIndex = 0
    $btnOK = Create-Button -Text "Instalar" -Location (New-Object System.Drawing.Point(10, 80)) -Size (New-Object System.Drawing.Size(140, 30))
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(190, 80)) -Size (New-Object System.Drawing.Size(140, 30))
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($lbl, $cmb, $btnOK, $btnCancel))

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    switch ($cmb.SelectedIndex) {
        0 { return "latest" }
        1 { return "ssms14" }
    }
}
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
function New-FormBuilder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(1000, 600)),
        [Parameter()]
        [System.Windows.Forms.FormStartPosition]$StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen,
        [Parameter()]
        [System.Windows.Forms.FormBorderStyle]$FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog,
        [Parameter()]
        [bool]$MaximizeBox = $false,
        [Parameter()]
        [bool]$MinimizeBox = $false,
        [Parameter()]
        [bool]$TopMost = $false,
        [Parameter()]
        [bool]$ControlBox = $true,
        [Parameter()]
        [System.Drawing.Icon]$Icon = $null,
        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = $Size
    $form.StartPosition = $StartPosition
    $form.FormBorderStyle = $FormBorderStyle
    $form.MaximizeBox = $MaximizeBox
    $form.MinimizeBox = $MinimizeBox
    $form.TopMost = $TopMost
    $form.ControlBox = $ControlBox
    if ($Icon) {
        $form.Icon = $Icon
    }
    $form.BackColor = $BackColor
    return $form
}
function New-Button {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [System.Drawing.Point]$Location,
        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,
        [Parameter()]
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
        [Parameter()]
        [string]$ToolTipText = $null,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35)),
        [Parameter()]
        [System.Drawing.Font]$Font = $null,
        [Parameter()]
        [bool]$Enabled = $true
    )

    if (-not $Font) {
        $Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    }
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = $Size
    $button.Location = $Location
    $button.BackColor = $BackColor
    $button.ForeColor = $ForeColor
    $button.Font = $Font
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $button.Enabled = $Enabled
    $button.Add_MouseEnter({
            $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
            $this.Font = New-Object System.Drawing.Font($this.Font.Name, $this.Font.Size, [System.Drawing.FontStyle]::Bold)
        })
    $button.Add_MouseLeave({
            $this.BackColor = $BackColor
            $this.Font = $Font
        })
    if ($ToolTipText) {
        $script:toolTip.SetToolTip($button, $ToolTipText)
    }
    return $button
}
function New-Label {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [System.Drawing.Point]$Location,
        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::Transparent,
        [Parameter()]
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
        [Parameter()]
        [string]$ToolTipText = $null,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
        [Parameter()]
        [System.Drawing.Font]$Font = $null,
        [Parameter()]
        [System.Windows.Forms.BorderStyle]$BorderStyle = [System.Windows.Forms.BorderStyle]::None,
        [Parameter()]
        [System.Drawing.ContentAlignment]$TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    )
    if (-not $Font) {
        $Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    }
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Size = $Size
    $label.Location = $Location
    $label.BackColor = $BackColor
    $label.ForeColor = $ForeColor
    $label.Font = $Font
    $label.BorderStyle = $BorderStyle
    $label.TextAlign = $TextAlign
    if ($ToolTipText) {
        $script:toolTip.SetToolTip($label, $ToolTipText)
    }
    return $label
}
function New-TextBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Point]$Location,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,
        [Parameter()]
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Black,
        [Parameter()]
        [System.Drawing.Font]$Font = $null,
        [Parameter()]
        [string]$Text = "",
        [Parameter()]
        [bool]$Multiline = $false,
        [Parameter()]
        [System.Windows.Forms.ScrollBars]$ScrollBars = [System.Windows.Forms.ScrollBars]::None,
        [Parameter()]
        [bool]$ReadOnly = $false,
        [Parameter()]
        [bool]$UseSystemPasswordChar = $false
    )
    if (-not $Font) {
        $Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    }
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

    if ($UseSystemPasswordChar) {
        $textBox.UseSystemPasswordChar = $true
    }
    return $textBox
}
function New-ComboBox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Point]$Location,
        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),
        [Parameter()]
        [System.Windows.Forms.ComboBoxStyle]$DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList,
        [Parameter()]
        [System.Drawing.Font]$Font = $null,
        [Parameter()]
        [string[]]$Items = @(),
        [Parameter()]
        [int]$SelectedIndex = -1,
        [Parameter()]
        [string]$DefaultText = $null
    )

    if (-not $Font) {
        $Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
    }
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = $Location
    $comboBox.Size = $Size
    $comboBox.DropDownStyle = $DropDownStyle
    $comboBox.Font = $Font

    if ($Items.Count -gt 0) {
        $comboBox.Items.AddRange($Items)
        $comboBox.SelectedIndex = $SelectedIndex
    }
    if ($DefaultText) {
        $comboBox.Text = $DefaultText
    }
    return $comboBox
}
function Show-ProgressDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter()]
        [string]$Message = "Procesando...",

        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(400, 150))
    )
    $form = New-FormBuilder -Title $Title -Size $Size -TopMost $true -ControlBox $false
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size(360, 20)
    $progressBar.Location = New-Object System.Drawing.Point(20, 50)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progressBar.MarqueeAnimationSpeed = 30
    $label = New-Label -Text $Message -Location (New-Object System.Drawing.Point(20, 20)) -Size (New-Object System.Drawing.Size(360, 20))
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.AddRange(@($progressBar, $label))
    return $form
}
function Set-ControlEnabled {
    param(
        [object]$Control,
        [bool]$Enabled,
        [string]$Name
    )

    if ($null -eq $Control) {
        Write-DzDebug "`t[DEBUG][Set-ControlEnabled] $Name es NULL"
        return
    }

    if ($Control -is [System.Windows.Forms.Control]) {
        $Control.Enabled = $Enabled
        Write-DzDebug "`t[DEBUG][Set-ControlEnabled] $Name ($($Control.GetType().Name)) Enabled=$Enabled"
    } else {
        Write-DzDebug "`t[DEBUG][Set-ControlEnabled] $Name tipo inesperado: $($Control.GetType().FullName)"
    }
}


Export-ModuleMember -Function New-FormBuilder, New-Button, New-Label, New-TextBox, New-ComboBox,
Show-ProgressDialog, Create-Form, Create-Button, Create-Label, Create-TextBox, Create-ComboBox,
Show-ProgressBar, Update-ProgressBar, Close-ProgressBar, Show-SSMSInstallerDialog, Show-NewIpForm, Set-ControlEnabled