#requires -Version 5.0

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
    $form.Text            = $Title
    $form.Size            = $Size
    $form.StartPosition   = $StartPosition
    $form.FormBorderStyle = $FormBorderStyle
    $form.MaximizeBox     = $MaximizeBox
    $form.MinimizeBox     = $MinimizeBox
    $form.TopMost         = $TopMost
    $form.ControlBox      = $ControlBox
    
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
    
    # Configurar eventos hover
    $button.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
        $this.Font = New-Object System.Drawing.Font($this.Font.Name, $this.Font.Size, [System.Drawing.FontStyle]::Bold)
    })
    
    $button.Add_MouseLeave({
        $this.BackColor = $BackColor
        $this.Font = $Font
    })
    
    if ($ToolTipText) {
        $toolTip.SetToolTip($button, $ToolTipText)
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
        $toolTip.SetToolTip($label, $ToolTipText)
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

Export-ModuleMember -Function New-FormBuilder, New-Button, New-Label, New-TextBox, New-ComboBox, Show-ProgressDialog
