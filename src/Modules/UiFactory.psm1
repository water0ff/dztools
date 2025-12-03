function New-MainForm {
    <#
        .SYNOPSIS
            Crea la ventana principal con configuración visual consistente.
        .PARAMETER Title
            Texto a mostrar en la barra de título.
        .PARAMETER Size
            Tamaño inicial del formulario.
        .PARAMETER BackColor
            Color de fondo aplicado al formulario.
        .PARAMETER DefaultFont
            Fuente usada por defecto en los controles.
        .PARAMETER BoldFont
            Fuente en negritas usada para efectos hover.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter()]
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(1000, 600)),

        [Parameter()]
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::White,

        [Parameter()]
        [System.Drawing.Font]$DefaultFont = $null,

        [Parameter()]
        [System.Drawing.Font]$BoldFont = $null
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = $Size
    $form.StartPosition = "CenterScreen"
    $form.BackColor = $BackColor
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    if ($DefaultFont) {
        $form.Tag = @{ DefaultFont = $DefaultFont; BoldFont = $BoldFont }
    }

    return $form
}

function New-Button {
    <#
        .SYNOPSIS
            Genera un botón estándar con estilos y tooltip opcional.
        .PARAMETER Text
            Etiqueta mostrada en el botón.
        .PARAMETER Location
            Posición en el formulario.
        .PARAMETER BackColor
            Color de fondo.
        .PARAMETER ForeColor
            Color del texto.
        .PARAMETER Size
            Dimensiones del botón.
        .PARAMETER Font
            Fuente a aplicar; si no se especifica usa Segoe UI 10.
        .PARAMETER Enabled
            Define si el botón inicia habilitado.
        .PARAMETER ToolTipText
            Texto del tooltip a asociar.
        .PARAMETER ToolTip
            Instancia de ToolTip para asociar la descripción.
    #>
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
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(220, 35)),

        [Parameter()]
        [System.Drawing.Font]$Font = $null,

        [Parameter()]
        [bool]$Enabled = $true,

        [Parameter()]
        [string]$ToolTipText = $null,

        [Parameter()]
        [System.Windows.Forms.ToolTip]$ToolTip
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
    $button.Tag = @{ BaseColor = $BackColor; BaseFont = $Font }

    $button.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 255)
        if ($this.Tag.BaseFont) {
            $this.Font = New-Object System.Drawing.Font($this.Tag.BaseFont.Name, $this.Tag.BaseFont.Size, [System.Drawing.FontStyle]::Bold)
        }
    })

    $button.Add_MouseLeave({
        $this.BackColor = $this.Tag.BaseColor
        if ($this.Tag.BaseFont) {
            $this.Font = $this.Tag.BaseFont
        }
    })

    if ($ToolTip -and $ToolTipText) {
        $ToolTip.SetToolTip($button, $ToolTipText)
    }

    return $button
}

function New-Label {
    <#
        .SYNOPSIS
            Genera una etiqueta configurable.
        .PARAMETER Text
            Contenido mostrado en el label.
        .PARAMETER Location
            Posición en el formulario.
        .PARAMETER BackColor
            Color de fondo.
        .PARAMETER ForeColor
            Color del texto.
        .PARAMETER Size
            Dimensiones del control.
        .PARAMETER Font
            Fuente a aplicar.
        .PARAMETER BorderStyle
            Tipo de borde a utilizar.
        .PARAMETER TextAlign
            Alineación del texto.
        .PARAMETER ToolTipText
            Texto del tooltip.
        .PARAMETER ToolTip
            Instancia para asociar el tooltip.
    #>
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
        [System.Drawing.Size]$Size = (New-Object System.Drawing.Size(200, 30)),

        [Parameter()]
        [System.Drawing.Font]$Font = $null,

        [Parameter()]
        [System.Windows.Forms.BorderStyle]$BorderStyle = [System.Windows.Forms.BorderStyle]::None,

        [Parameter()]
        [System.Drawing.ContentAlignment]$TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft,

        [Parameter()]
        [string]$ToolTipText = $null,

        [Parameter()]
        [System.Windows.Forms.ToolTip]$ToolTip
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

    if ($ToolTip -and $ToolTipText) {
        $ToolTip.SetToolTip($label, $ToolTipText)
    }

    return $label
}

function New-Form {
    [CmdletBinding()]
    param(
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
    $textBox.WordWrap = $false

    if ($UseSystemPasswordChar) {
        $textBox.UseSystemPasswordChar = $true
    }

    return $textBox
}

Set-Alias -Name Create-Button -Value New-Button
Set-Alias -Name Create-Label -Value New-Label
Set-Alias -Name Create-Form -Value New-Form
Set-Alias -Name Create-ComboBox -Value New-ComboBox
Set-Alias -Name Create-TextBox -Value New-TextBox
Set-Alias -Name Create-MainForm -Value New-MainForm

Export-ModuleMember -Function New-MainForm, New-Button, New-Label, New-Form, New-ComboBox, New-TextBox -Alias *
