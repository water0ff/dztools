#requires -Version 5.0

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

#region Funciones Base WPF

function New-WpfWindow {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xaml,
        [switch]$PassThru
    )
    try {
        $reader = New-Object System.Xml.XmlNodeReader $Xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)

        if ($PassThru) {
            $controls = @{}
            $Xaml.SelectNodes("//*[@Name]") | ForEach-Object {
                $controls[$_.Name] = $window.FindName($_.Name)
            }
            return @{
                Window   = $window
                Controls = $controls
            }
        }
        return $window
    } catch {
        Write-Error "Error cargando XAML: $_"
        throw
    }
}

function Show-WpfMessageBox {
    param(
        [string]$Message,
        [string]$Title = "Mensaje",
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")]
        [string]$Buttons = "OK",
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [string]$Icon = "Information"
    )

    $buttonMap = @{
        "OK"          = [System.Windows.MessageBoxButton]::OK
        "OKCancel"    = [System.Windows.MessageBoxButton]::OKCancel
        "YesNo"       = [System.Windows.MessageBoxButton]::YesNo
        "YesNoCancel" = [System.Windows.MessageBoxButton]::YesNoCancel
    }

    $iconMap = @{
        "Information" = [System.Windows.MessageBoxImage]::Information
        "Warning"     = [System.Windows.MessageBoxImage]::Warning
        "Error"       = [System.Windows.MessageBoxImage]::Error
        "Question"    = [System.Windows.MessageBoxImage]::Question
    }

    return [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        $buttonMap[$Buttons],
        $iconMap[$Icon]
    )
}

#endregion

#region Barra de Progreso WPF Mejorada

function Show-WpfProgressBar {
    <#
    .SYNOPSIS
        Muestra una barra de progreso WPF no bloqueante.
    #>
    param(
        [string]$Title = "Procesando",
        [string]$Message = "Por favor espere..."
    )

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title"
        Height="220" Width="500"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="False">
    <Border Background="White"
            CornerRadius="8"
            BorderBrush="#2196F3"
            BorderThickness="2"
            Padding="20">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="5" BlurRadius="15" Opacity="0.3"/>
        </Border.Effect>
        <StackPanel>
            <TextBlock Text="$Title"
                       FontSize="18"
                       FontWeight="Bold"
                       Foreground="#2196F3"
                       HorizontalAlignment="Center"
                       Margin="0,0,0,20"/>

            <TextBlock Name="lblMessage"
                       Text="$Message"
                       FontSize="12"
                       Foreground="#757575"
                       TextAlignment="Center"
                       TextWrapping="Wrap"
                       Margin="0,0,0,15"
                       MinHeight="30"/>

            <ProgressBar Name="progressBar"
                         Height="25"
                         Minimum="0"
                         Maximum="100"
                         Value="0"
                         Foreground="#4CAF50"
                         Background="#E0E0E0"
                         Margin="0,0,0,10"/>

            <TextBlock Name="lblPercent"
                       Text="0%"
                       FontSize="14"
                       FontWeight="Bold"
                       Foreground="#2196F3"
                       HorizontalAlignment="Center"/>
        </StackPanel>
    </Border>
</Window>
"@

    try {
        $result = New-WpfWindow -Xaml $xaml -PassThru
        $window = $result.Window

        $window | Add-Member -MemberType NoteProperty -Name ProgressBar   -Value $result.Controls['progressBar'] | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name MessageLabel  -Value $result.Controls['lblMessage']  | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name PercentLabel  -Value $result.Controls['lblPercent']  | Out-Null
        $window | Add-Member -MemberType NoteProperty -Name IsClosed      -Value $false                          | Out-Null
        $window.Add_Closed({
                $window.IsClosed = $true
            })
        $window.Show()

        # Forzar render inicial (siempre con el dispatcher de la misma ventana)
        $window.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action] {}
        ) | Out-Null
        return $window
    } catch {
        Write-Error "Error al crear barra de progreso: $_"
        return $null
    }
}

function Update-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)] $Window,
        [Parameter(Mandatory = $true)][ValidateRange(0, 100)][int] $Percent,
        [string] $Message = $null
    )

    if ($null -eq $Window) { return }
    if (-not ($Window -is [System.Windows.Window])) { return }
    if ($Window.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $Window.IsClosed) { return }

    # Capturar valores (sin cambiar variables externas para compatibilidad)
    $pLocal = [Math]::Min($Percent, 100)
    $mLocal = $Message

    # Copia local del window para pasarla como argumento al delegate
    $wLocal = $Window

    try {
        # Tipado del delegate para poder pasar parámetros (evita closures raros en PS5)
        $action = [Action[object, int, string]] {
            param($w, $p, $m)

            if ($null -eq $w) { return }
            if ($w.PSObject.Properties.Match('IsClosed').Count -gt 0 -and $w.IsClosed) { return }

            # ProgressBar
            if ($w.PSObject.Properties.Match('ProgressBar').Count -gt 0 -and $w.ProgressBar) {
                $w.ProgressBar.IsIndeterminate = $false
                $w.ProgressBar.Value = $p
            }

            # Porcentaje
            if ($w.PSObject.Properties.Match('PercentLabel').Count -gt 0 -and $w.PercentLabel) {
                $w.PercentLabel.Text = "$p%"
            }

            # Mensaje
            if (-not [string]::IsNullOrWhiteSpace($m) -and
                $w.PSObject.Properties.Match('MessageLabel').Count -gt 0 -and $w.MessageLabel) {
                $w.MessageLabel.Text = $m
            }

            $w.UpdateLayout()
        }

        # Invoke con prioridad de render, pasando args en lugar de capturar variables
        $wLocal.Dispatcher.Invoke(
            $action,
            [System.Windows.Threading.DispatcherPriority]::Render,
            $wLocal, $pLocal, $mLocal
        ) | Out-Null

    } catch {
        Write-Warning "Error actualizando barra de progreso: $($_.Exception.Message)"
    }
}

function Close-WpfProgressBar {
    param([Parameter(Mandatory = $true)] $Window)

    if ($null -eq $Window) { return }

    if (-not ($Window -is [System.Windows.Window])) {
        Write-Warning "Close-WpfProgressBar: El objeto recibido NO es WPF Window. Tipo: $($Window.GetType().FullName)"
        return
    }

    if ($Window.IsClosed) { return }

    if ($null -eq $Window.Dispatcher -or
        $Window.Dispatcher.HasShutdownStarted -or
        $Window.Dispatcher.HasShutdownFinished) { return }

    try {
        $Window.Dispatcher.Invoke([action] {
                if (-not $Window.IsClosed) { $Window.Close() }
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
    } catch {
        Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)"
    }
}
function Show-ProgressBar {
    return Show-WpfProgressBar -Title "Progreso de Actualización" -Message "Iniciando proceso..."
}

function Update-ProgressBar {
    param(
        $ProgressForm,
        $CurrentStep,
        $TotalSteps,
        [string]$Status = ""
    )

    if ($null -eq $ProgressForm -or $ProgressForm.IsClosed) {
        return
    }

    try {
        $percent = [math]::Round(($CurrentStep / $TotalSteps) * 100)
        Update-WpfProgressBar -Window $ProgressForm -Percent $percent -Message $Status
    } catch {
        Write-Warning "Error actualizando barra de progreso: $($_.Exception.Message)"
    }
}

function Close-ProgressBar {
    param($ProgressForm)

    if ($null -eq $ProgressForm) {
        return
    }

    try {
        Close-WpfProgressBar -Window $ProgressForm
    } catch {
        Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)"
    }
}

#endregion

#region Utilidades WPF

function Set-WpfControlEnabled {
    param(
        [Parameter(Mandatory = $true)]
        $Control,
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    if ($null -eq $Control) {
        Write-Warning "Control es null."
        return
    }

    try {
        if ($Control.Dispatcher.CheckAccess()) {
            $Control.IsEnabled = $Enabled
        } else {
            $Control.Dispatcher.Invoke([action] {
                    $Control.IsEnabled = $Enabled
                })
        }
    } catch {
        Write-Warning "Error cambiando estado del control: $_"
    }
}

function Set-ControlEnabled {
    param(
        [object]$Control,
        [bool]$Enabled,
        [string]$Name
    )

    if ($null -eq $Control) {
        Write-Warning "Control $Name es NULL"
        return
    }

    Set-WpfControlEnabled -Control $Control -Enabled $Enabled
}

#endregion

#region Diálogos de Entrada

function New-WpfInputDialog {
    param(
        [string]$Title = "Entrada",
        [string]$Prompt = "Ingrese un valor:",
        [string]$DefaultValue = ""
    )

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title"
        Height="180" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        ShowInTaskbar="False">
    <StackPanel Margin="20">
        <TextBlock Text="$Prompt" FontSize="12" Margin="0,0,0,10"/>
        <TextBox Name="txtInput" Text="$DefaultValue" FontSize="12" Padding="5" Margin="0,0,0,20"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnOK" Content="Aceptar" Width="80" Margin="0,0,10,0" IsDefault="True"/>
            <Button Name="btnCancel" Content="Cancelar" Width="80" IsCancel="True"/>
        </StackPanel>
    </StackPanel>
</Window>
"@

    $result = New-WpfWindow -Xaml $xaml -PassThru
    $window = $result.Window
    $controls = $result.Controls

    $script:inputValue = $null

    $controls['btnOK'].Add_Click({
            $script:inputValue = $controls['txtInput'].Text
            $window.DialogResult = $true
            $window.Close()
        })

    $controls['btnCancel'].Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })

    $dialogResult = $window.ShowDialog()

    if ($dialogResult) {
        return $script:inputValue
    }

    return $null
}

function Get-WpfPasswordBoxText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.PasswordBox]$PasswordBox
    )
    return $PasswordBox.Password
}

function Add-WpfComboBoxItems {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.ComboBox]$ComboBox,
        [Parameter(Mandatory = $true)]
        [string[]]$Items,
        [int]$SelectedIndex = -1
    )

    $ComboBox.Items.Clear()
    foreach ($item in $Items) {
        $ComboBox.Items.Add($item) | Out-Null
    }

    if ($SelectedIndex -ge 0 -and $SelectedIndex -lt $Items.Count) {
        $ComboBox.SelectedIndex = $SelectedIndex
    }
}

#endregion

#region Diálogos de Archivos

function Show-WpfFileDialog {
    param(
        [ValidateSet("Open", "Save")]
        [string]$Type = "Open",
        [string]$Filter = "Todos los archivos (*.*)|*.*",
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$DefaultFileName = ""
    )

    Add-Type -AssemblyName Microsoft.Win32

    if ($Type -eq "Open") {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
    } else {
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.FileName = $DefaultFileName
    }

    $dialog.Filter = $Filter
    $dialog.InitialDirectory = $InitialDirectory

    $result = $dialog.ShowDialog()

    if ($result) {
        return $dialog.FileName
    }

    return $null
}

function Show-WpfFolderDialog {
    param(
        [string]$Description = "Seleccione una carpeta",
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop')
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.SelectedPath = $InitialDirectory

    $result = $dialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

function Show-NewIpForm {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Agregar IP Adicional"
        Height="180" Width="350"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#F5F5F5">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Ingrese la nueva dirección IP:" FontSize="12" Margin="0,0,0,15"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center">
            <TextBox Name="txt1" Width="50" MaxLength="3" FontSize="12" Padding="5" Margin="0,0,5,0"/>
            <TextBlock Text="." FontSize="16" VerticalAlignment="Center" Margin="0,0,5,0"/>
            <TextBox Name="txt2" Width="50" MaxLength="3" FontSize="12" Padding="5" Margin="0,0,5,0"/>
            <TextBlock Text="." FontSize="16" VerticalAlignment="Center" Margin="0,0,5,0"/>
            <TextBox Name="txt3" Width="50" MaxLength="3" FontSize="12" Padding="5" Margin="0,0,5,0"/>
            <TextBlock Text="." FontSize="16" VerticalAlignment="Center" Margin="0,0,5,0"/>
            <TextBox Name="txt4" Width="50" MaxLength="3" FontSize="12" Padding="5"/>
        </StackPanel>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="btnOK" Content="Aceptar" Width="100" Margin="0,0,10,0" Padding="8"/>
            <Button Name="btnCancel" Content="Cancelar" Width="100" Padding="8"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $result = New-WpfWindow -Xaml $xaml -PassThru
    $window = $result.Window
    $controls = $result.Controls
    $script:newIpValue = $null

    # Validación de entrada numérica
    $numericValidation = {
        param($sender, $e)
        if (-not [char]::IsDigit($e.Text[0])) {
            $e.Handled = $true
        }
    }

    $controls['txt1'].Add_PreviewTextInput($numericValidation)
    $controls['txt2'].Add_PreviewTextInput($numericValidation)
    $controls['txt3'].Add_PreviewTextInput($numericValidation)
    $controls['txt4'].Add_PreviewTextInput($numericValidation)

    # Auto-focus al siguiente campo
    $controls['txt1'].Add_TextChanged({
            if ($controls['txt1'].Text.Length -eq 3) { $controls['txt2'].Focus() }
        })
    $controls['txt2'].Add_TextChanged({
            if ($controls['txt2'].Text.Length -eq 3) { $controls['txt3'].Focus() }
        })
    $controls['txt3'].Add_TextChanged({
            if ($controls['txt3'].Text.Length -eq 3) { $controls['txt4'].Focus() }
        })

    $controls['btnOK'].Add_Click({
            try {
                $octet1 = [int]$controls['txt1'].Text
                $octet2 = [int]$controls['txt2'].Text
                $octet3 = [int]$controls['txt3'].Text
                $octet4 = [int]$controls['txt4'].Text

                if ($octet1 -ge 0 -and $octet1 -le 255 -and
                    $octet2 -ge 0 -and $octet2 -le 255 -and
                    $octet3 -ge 0 -and $octet3 -le 255 -and
                    $octet4 -ge 0 -and $octet4 -le 255) {

                    $newIp = "$octet1.$octet2.$octet3.$octet4"

                    if ($newIp -eq "0.0.0.0") {
                        [System.Windows.MessageBox]::Show("La dirección IP no puede ser 0.0.0.0.", "Error")
                        return
                    }

                    $script:newIpValue = $newIp
                    $window.DialogResult = $true
                    $window.Close()
                } else {
                    [System.Windows.MessageBox]::Show("Octetos fuera del rango válido (0-255).", "Error")
                }
            } catch {
                [System.Windows.MessageBox]::Show("Complete todos los campos con valores numéricos.", "Error")
            }
        })

    $controls['btnCancel'].Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })

    $dialogResult = $window.ShowDialog()

    if ($dialogResult) {
        return $script:newIpValue
    }

    return $null
}

function Show-AddUserDialog {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Crear Usuario de Windows" Height="250" Width="450"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <Label Content="Nombre:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0"/>
        <TextBox Name="txtUsername" HorizontalAlignment="Left" VerticalAlignment="Top"
                 Width="290" Height="25" Margin="110,0,0,0"/>

        <Label Content="Contraseña:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,40,0,0"/>
        <PasswordBox Name="txtPassword" HorizontalAlignment="Left" VerticalAlignment="Top"
                     Width="290" Height="25" Margin="110,40,0,0"/>

        <Label Content="Tipo:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,80,0,0"/>
        <ComboBox Name="cmbType" HorizontalAlignment="Left" VerticalAlignment="Top"
                  Width="290" Height="25" Margin="110,80,0,0">
            <ComboBoxItem Content="Usuario estándar"/>
            <ComboBoxItem Content="Administrador"/>
        </ComboBox>

        <Button Content="Crear" Name="btnCreate" Width="130" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,130,0,0"/>
        <Button Content="Cancelar" Name="btnCancel" Width="130" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="140,130,0,0"/>
        <Button Content="Mostrar usuarios" Name="btnShow" Width="130" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="280,130,0,0"/>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $txtUsername = $window.FindName("txtUsername")
    $txtPassword = $window.FindName("txtPassword")
    $cmbType = $window.FindName("cmbType")
    $btnCreate = $window.FindName("btnCreate")
    $btnCancel = $window.FindName("btnCancel")
    $btnShow = $window.FindName("btnShow")

    $cmbType.SelectedIndex = 0

    $adminGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-544').Name
    $userGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-545').Name

    $btnShow.Add_Click({
            Write-Host "`nUsuarios actuales:" -ForegroundColor Cyan
            $users = Get-LocalUser
            $usersTable = $users | ForEach-Object {
                $user = $_
                $estado = if ($user.Enabled) { "Habilitado" } else { "Deshabilitado" }
                $tipoUsuario = "Usuario estándar"
                try {
                    $adminMembers = Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop
                    if ($adminMembers | Where-Object { $_.SID -eq $user.SID }) {
                        $tipoUsuario = "Administrador"
                    }
                } catch { }
                [PSCustomObject]@{
                    Nombre = $user.Name
                    Tipo   = $tipoUsuario
                    Estado = $estado
                }
            }
            $usersTable | Format-Table -AutoSize | Out-String | Write-Host
        })

    $btnCreate.Add_Click({
            $username = $txtUsername.Text.Trim()
            $password = $txtPassword.Password
            $type = $cmbType.Text

            if (-not $username -or -not $password) {
                Write-Host "Error: Nombre y contraseña requeridos" -ForegroundColor Red
                return
            }

            if ($password.Length -lt 8) {
                Write-Host "Error: Contraseña debe tener al menos 8 caracteres" -ForegroundColor Red
                return
            }

            try {
                if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
                    Write-Host "Error: Usuario '$username' ya existe" -ForegroundColor Red
                    return
                }

                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                New-LocalUser -Name $username -Password $securePassword -AccountNeverExpires -PasswordNeverExpires
                Write-Host "Usuario '$username' creado exitosamente" -ForegroundColor Green

                $group = if ($type -eq 'Administrador') { $adminGroup } else { $userGroup }
                Add-LocalGroupMember -Group $group -Member $username
                Write-Host "Usuario agregado al grupo $group" -ForegroundColor Cyan

                $window.Close()
            } catch {
                Write-Host "Error: $_" -ForegroundColor Red
            }
        })

    $btnCancel.Add_Click({
            Write-Host "Operación cancelada" -ForegroundColor Yellow
            $window.Close()
        })

    $window.ShowDialog() | Out-Null
}

function Show-IPConfigDialog {
    # [Contenido original de Show-IPConfigDialog - mantener igual]
    Write-Host "Show-IPConfigDialog: Implementación pendiente o usar la original"
}

function Show-SQLselector {
    param(
        [array]$Managers,
        [array]$SSMSVersions
    )

    # Helper: intenta asignar owner ($window) para modal real
    function Set-DialogOwner {
        param([System.Windows.Window]$Dialog)

        try {
            if (Get-Variable -Name window -Scope Global -ErrorAction SilentlyContinue) {
                $Dialog.Owner = $Global:window
                return
            }
            if (Get-Variable -Name window -Scope Script -ErrorAction SilentlyContinue) {
                $Dialog.Owner = $script:window
                return
            }
        } catch { }
    }

    # Helper: bits para managers
    function Get-ManagerBits {
        param([string]$Path)

        # System32 = 64-bit, SysWOW64 = 32-bit
        if ($Path -match "\\SysWOW64\\") { return "32 bits" }
        return "64 bits"
    }

    # Helper: versión del manager (SQLServerManager15.msc => 15)
    function Get-ManagerVersion {
        param([string]$Path)
        if ($Path -match "SQLServerManager(\d+)\.msc") { return $matches[1] }
        return "?"
    }

    # Helper: crea items para listbox (Display + Path)
    function New-SelectorItem {
        param(
            [string]$Path,
            [string]$Display
        )
        [PSCustomObject]@{
            Path    = $Path
            Display = $Display
        }
    }

    # Helper: construye un diálogo genérico de selección
    function Show-PathSelectionDialog {
        param(
            [string]$Title,
            [string]$Prompt,
            [array]$Items,               # array de PSCustomObject {Path, Display}
            [scriptblock]$OnExecute,     # recibe $SelectedPath
            [string]$ExecuteButtonText = "Ejecutar"
        )

        $dialog = New-Object System.Windows.Window
        $dialog.Title = $Title
        $dialog.Width = 780
        $dialog.Height = 420
        $dialog.WindowStartupLocation = "CenterOwner"
        $dialog.ResizeMode = "NoResize"
        if ($Global:window -is [System.Windows.Window]) {
            $dialog.Owner = $Global:window
        }

        # 2) Centrar respecto al owner
        $dialog.WindowStartupLocation = "CenterOwner"

        # 3) Fallback: si no hay Owner, centrar en pantalla
        if (-not $dialog.Owner) {
            $dialog.WindowStartupLocation = "CenterScreen"
        }
        Set-DialogOwner -Dialog $dialog

        $root = New-Object System.Windows.Controls.StackPanel
        $root.Margin = New-Object System.Windows.Thickness(10)

        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $Prompt
        $label.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $label.FontSize = 13
        $root.Children.Add($label) | Out-Null

        # ListBox: aquí va lo "legible": version/bits + RUTA COMPLETA
        $listBox = New-Object System.Windows.Controls.ListBox
        $listBox.Height = 250
        $listBox.FontSize = 12
        $listBox.DisplayMemberPath = "Display"
        $listBox.SelectedValuePath = "Path"
        foreach ($it in $Items) { $null = $listBox.Items.Add($it) }
        $listBox.SelectedIndex = 0
        $root.Children.Add($listBox) | Out-Null

        # Ruta seleccionada (para copiar fácil)
        $pathLabelTitle = New-Object System.Windows.Controls.TextBlock
        $pathLabelTitle.Text = "Ruta seleccionada:"
        $pathLabelTitle.Margin = New-Object System.Windows.Thickness(0, 10, 0, 2)
        $pathLabelTitle.FontSize = 11
        $root.Children.Add($pathLabelTitle) | Out-Null

        $pathLabel = New-Object System.Windows.Controls.TextBlock
        $pathLabel.Text = ""
        $pathLabel.FontSize = 11
        $pathLabel.FontFamily = "Consolas"
        $pathLabel.TextWrapping = "Wrap"
        $pathLabel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $root.Children.Add($pathLabel) | Out-Null

        $updatePath = {
            if ($listBox.SelectedValue) { $pathLabel.Text = $listBox.SelectedValue }
            else { $pathLabel.Text = "" }
        }
        & $updatePath
        $listBox.Add_SelectionChanged({ & $updatePath })

        $btnPanel = New-Object System.Windows.Controls.StackPanel
        $btnPanel.Orientation = "Horizontal"
        $btnPanel.HorizontalAlignment = "Right"

        $cancelButton = New-Object System.Windows.Controls.Button
        $cancelButton.Content = "Cancelar"
        $cancelButton.Width = 95
        $cancelButton.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
        $cancelButton.Add_Click({
                $dialog.DialogResult = $false
                $dialog.Close()
            })

        $okButton = New-Object System.Windows.Controls.Button
        $okButton.Content = $ExecuteButtonText
        $okButton.Width = 95
        $okButton.IsDefault = $true
        $okButton.Add_Click({
                if ($listBox.SelectedValue) {
                    try {
                        & $OnExecute $listBox.SelectedValue
                        $dialog.DialogResult = $true
                        $dialog.Close()
                    } catch {
                        [System.Windows.MessageBox]::Show(
                            "Error al ejecutar:`n$($_.Exception.Message)",
                            "Error",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Error
                        ) | Out-Null
                    }
                }
            })

        $btnPanel.Children.Add($cancelButton) | Out-Null
        $btnPanel.Children.Add($okButton) | Out-Null
        $root.Children.Add($btnPanel) | Out-Null

        $dialog.Content = $root
        $null = $dialog.ShowDialog()
    }

    # =========================
    # 1) MANAGERS (SQLServerManager*.msc)
    # =========================
    if ($Managers -and $Managers.Count -gt 0) {

        $items = @()

        $unique = $Managers | Where-Object { $_ } | Select-Object -Unique
        foreach ($m in $unique) {
            $ver = Get-ManagerVersion -Path $m
            $bits = Get-ManagerBits    -Path $m

            # Display legible (incluye RUTA COMPLETA)
            $display = "SQLServerManager$ver  |  $bits  |  $m"
            $items += (New-SelectorItem -Path $m -Display $display)
        }

        Show-PathSelectionDialog `
            -Title  "Seleccionar Configuration Manager" `
            -Prompt "Seleccione la versión de SQL Server Configuration Manager a ejecutar:" `
            -Items  $items `
            -OnExecute {
            param($selectedPath)
            Write-Host "`tEjecutando SQL Server Configuration Manager desde: $selectedPath" -ForegroundColor Green
            Start-Process -FilePath $selectedPath
        } `
            -ExecuteButtonText "Abrir"

        return
    }

    # =========================
    # 2) SSMS (Ssms.exe)
    # =========================
    if ($SSMSVersions -and $SSMSVersions.Count -gt 0) {

        $items = @()
        $unique = $SSMSVersions | Where-Object { $_ } | Select-Object -Unique

        foreach ($p in $unique) {
            # Display legible: producto + versión + ruta completa
            try {
                $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($p)
                $prod = if ($vi.ProductName) { $vi.ProductName } else { "SSMS" }
                $ver = if ($vi.FileVersion) { $vi.FileVersion }  else { "" }

                $display = "$prod  |  $ver  |  $p"
                $items += (New-SelectorItem -Path $p -Display $display)
            } catch {
                $items += (New-SelectorItem -Path $p -Display "SSMS  |  $p")
            }
        }

        Show-PathSelectionDialog `
            -Title  "Seleccionar SSMS" `
            -Prompt "Seleccione la versión de SQL Server Management Studio a ejecutar:" `
            -Items  $items `
            -OnExecute {
            param($selectedPath)
            Write-Host "`tEjecutando: $selectedPath" -ForegroundColor Green
            Start-Process -FilePath $selectedPath
        } `
            -ExecuteButtonText "Ejecutar"

        return
    }

    Write-Host "Show-SQLselector: No se recibieron rutas para Managers ni para SSMS." -ForegroundColor Yellow
}


function Show-LZMADialog {
    param([array]$Instaladores)
    Write-Host "Show-LZMADialog: Implementación pendiente o usar la original"
}



Export-ModuleMember -Function @(
    'New-WpfWindow',
    'Show-WpfMessageBox',
    'New-WpfInputDialog',
    'Show-WpfProgressBar',
    'Update-WpfProgressBar',
    'Close-WpfProgressBar',
    'Set-WpfControlEnabled',
    'Get-WpfPasswordBoxText',
    'Add-WpfComboBoxItems',
    'Show-WpfFileDialog',
    'Show-WpfFolderDialog',
    'Show-ProgressBar',
    'Update-ProgressBar',
    'Close-ProgressBar',
    'Set-ControlEnabled',
    'Show-NewIpForm',
    'Show-AddUserDialog',
    'Show-IPConfigDialog',
    'Show-SQLselector',
    'Show-LZMADialog'
)