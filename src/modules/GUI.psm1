#requires -Version 5.0

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

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
        ResizeMode="NoResize">
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

function Show-WpfProgressBar {
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
        Topmost="True">
    <Border Background="White"
            CornerRadius="8"
            BorderBrush="#2196F3"
            BorderThickness="2"
            Padding="20">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="5" BlurRadius="15" Opacity="0.3"/>
        </Border.Effect>
        <StackPanel>
            <TextBlock Text="$Title" FontSize="18" FontWeight="Bold" Foreground="#2196F3" HorizontalAlignment="Center" Margin="0,0,0,20"/>
            <TextBlock Name="lblMessage" Text="$Message" FontSize="12" Foreground="#757575" TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,15"/>
            <ProgressBar Name="progressBar" Height="25" Minimum="0" Maximum="100" Value="0" Foreground="#4CAF50" Background="#E0E0E0" Margin="0,0,0,10"/>
            <TextBlock Name="lblPercent" Text="0%" FontSize="14" FontWeight="Bold" Foreground="#2196F3" HorizontalAlignment="Center"/>
        </StackPanel>
    </Border>
</Window>
"@
    $result = New-WpfWindow -Xaml $xaml -PassThru
    $window = $result.Window
    $window | Add-Member -MemberType NoteProperty -Name ProgressBar -Value $result.Controls['progressBar']
    $window | Add-Member -MemberType NoteProperty -Name MessageLabel -Value $result.Controls['lblMessage']
    $window | Add-Member -MemberType NoteProperty -Name PercentLabel -Value $result.Controls['lblPercent']
    $window.Show()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action] {}
    )
    return $window
}

function Update-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)]
        $Window,
        [Parameter(Mandatory = $true)]
        [int]$Percent,
        [string]$Message = $null
    )
    if ($null -eq $Window -or -not $Window.IsVisible) {
        return
    }
    $Window.ProgressBar.Value = [Math]::Min($Percent, 100)
    $Window.PercentLabel.Text = "$Percent%"
    if ($Message) {
        $Window.MessageLabel.Text = $Message
    }
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [action] {}
    )
}

function Close-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)]
        $Window
    )
    if ($Window -and $Window.IsVisible) {
        $Window.Close()
    }
}

function Set-WpfControlEnabled {
    param(
        [Parameter(Mandatory = $true)]
        $Control,
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )
    if ($null -ne $Control) {
        $Control.IsEnabled = $Enabled
    }
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
    Add-Type -AssemblyName System.Windows.Forms
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
    $controls['txt1'].Add_PreviewTextInput({
            param($sender, $e)
            if (-not [char]::IsDigit($e.Text[0])) {
                $e.Handled = $true
            }
        })
    $controls['txt2'].Add_PreviewTextInput({
            param($sender, $e)
            if (-not [char]::IsDigit($e.Text[0])) {
                $e.Handled = $true
            }
        })
    $controls['txt3'].Add_PreviewTextInput({
            param($sender, $e)
            if (-not [char]::IsDigit($e.Text[0])) {
                $e.Handled = $true
            }
        })
    $controls['txt4'].Add_PreviewTextInput({
            param($sender, $e)
            if (-not [char]::IsDigit($e.Text[0])) {
                $e.Handled = $true
            }
        })
    $controls['txt1'].Add_TextChanged({
            if ($controls['txt1'].Text.Length -eq 3) {
                $controls['txt2'].Focus()
            }
        })
    $controls['txt2'].Add_TextChanged({
            if ($controls['txt2'].Text.Length -eq 3) {
                $controls['txt3'].Focus()
            }
        })
    $controls['txt3'].Add_TextChanged({
            if ($controls['txt3'].Text.Length -eq 3) {
                $controls['txt4'].Focus()
            }
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
                        [System.Windows.MessageBox]::Show("La dirección IP no puede ser 0.0.0.0.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        return
                    }
                    $script:newIpValue = $newIp
                    $window.DialogResult = $true
                    $window.Close()
                } else {
                    [System.Windows.MessageBox]::Show("Uno o más octetos están fuera del rango válido (0-255).", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            } catch {
                [System.Windows.MessageBox]::Show("Por favor, complete todos los campos con valores numéricos.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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
    if ($null -eq $ProgressForm -or -not $ProgressForm.IsVisible) {
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
    if ($null -eq $ProgressForm -or -not $ProgressForm.IsVisible) {
        return
    }
    try {
        Close-WpfProgressBar -Window $ProgressForm
    } catch {
        Write-Warning "Error cerrando barra de progreso: $($_.Exception.Message)"
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
                    }
                } catch {}
                [PSCustomObject]@{
                    Nombre = $user.Name.Substring(0, [Math]::Min(25, $user.Name.Length))
                    Tipo   = $tipoUsuario
                    Estado = $estado
                }
            }
            if ($usersTable.Count -gt 0) {
                Write-Host ("{0,-25} {1,-40} {2,-15}" -f "Nombre", "Tipo", "Estado")
                Write-Host ("{0,-25} {1,-40} {2,-15}" -f "------", "------", "------")
                $usersTable | ForEach-Object {
                    Write-Host ("{0,-25} {1,-40} {2,-15}" -f $_.Nombre, $_.Tipo, $_.Estado)
                }
            }
        })

    $btnCreate.Add_Click({
            $username = $txtUsername.Text.Trim()
            $password = $txtPassword.Password
            $type = $cmbType.Text

            if (-not $username -or -not $password) {
                Write-Host "`nError: Nombre y contraseña son requeridos" -ForegroundColor Red
                return
            }
            if ($password.Length -lt 8) {
                Write-Host "`nError: La contraseña debe tener al menos 8 caracteres" -ForegroundColor Red
                return
            }
            try {
                if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) {
                    Write-Host "`nError: El usuario '$username' ya existe" -ForegroundColor Red
                    return
                }
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                New-LocalUser -Name $username -Password $securePassword -AccountNeverExpires -PasswordNeverExpires
                Write-Host "`nUsuario '$username' creado exitosamente" -ForegroundColor Green

                $group = if ($type -eq 'Administrador') { $adminGroup } else { $userGroup }
                Add-LocalGroupMember -Group $group -Member $username
                Write-Host "`tUsuario agregado al grupo $group" -ForegroundColor Cyan
                $window.Close()
            } catch {
                Write-Host "`nError durante la creación del usuario: $_" -ForegroundColor Red
            }
        })

    $btnCancel.Add_Click({
            Write-Host "`tOperación cancelada." -ForegroundColor Yellow
            $window.Close()
        })

    $window.ShowDialog() | Out-Null
}

function Show-IPConfigDialog {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Asignación de IPs" Height="250" Width="400"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <Label Content="Seleccione el adaptador de red:" HorizontalAlignment="Left"
               VerticalAlignment="Top" Margin="0,0,0,0"/>
        <ComboBox Name="cmbAdapters" HorizontalAlignment="Left" VerticalAlignment="Top"
                  Width="360" Height="25" Margin="0,30,0,0"/>

        <Label Name="lblIps" Content="IPs asignadas:" HorizontalAlignment="Left"
               VerticalAlignment="Top" Margin="0,70,0,0"/>

        <Button Content="Asignar Nueva IP" Name="btnAssignIP" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,110,0,0" IsEnabled="False"/>
        <Button Content="Cambiar a DHCP" Name="btnChangeToDhcp" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="150,110,0,0" IsEnabled="False"/>
        <Button Content="Cerrar" Name="btnClose" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,150,0,0"/>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $cmbAdapters = $window.FindName("cmbAdapters")
    $lblIps = $window.FindName("lblIps")
    $btnAssignIP = $window.FindName("btnAssignIP")
    $btnChangeToDhcp = $window.FindName("btnChangeToDhcp")
    $btnClose = $window.FindName("btnClose")

    $cmbAdapters.Items.Add("Selecciona 1 adaptador de red") | Out-Null
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    foreach ($adapter in $adapters) {
        $cmbAdapters.Items.Add($adapter.Name) | Out-Null
    }
    $cmbAdapters.SelectedIndex = 0

    $cmbAdapters.Add_SelectionChanged({
            if ($cmbAdapters.SelectedIndex -gt 0) {
                $btnAssignIP.IsEnabled = $true
                $btnChangeToDhcp.IsEnabled = $true

                $selectedAdapter = Get-NetAdapter -Name $cmbAdapters.SelectedItem
                $currentIPs = Get-NetIPAddress -InterfaceAlias $selectedAdapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
                if ($currentIPs) {
                    $ips = ($currentIPs.IPAddress) -join ", "
                    $lblIps.Content = "IPs asignadas: $ips"
                }
            } else {
                $btnAssignIP.IsEnabled = $false
                $btnChangeToDhcp.IsEnabled = $false
                $lblIps.Content = "IPs asignadas:"
            }
        })

    $btnAssignIP.Add_Click({
            [System.Windows.MessageBox]::Show("Funcionalidad de asignar IP en desarrollo", "Info",
                [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        })

    $btnChangeToDhcp.Add_Click({
            $selectedAdapterName = $cmbAdapters.SelectedItem
            if ($selectedAdapterName -eq "Selecciona 1 adaptador de red") {
                return
            }

            $result = [System.Windows.MessageBox]::Show("¿Está seguro de cambiar a DHCP?", "Confirmación",
                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                try {
                    $selectedAdapter = Get-NetAdapter -Name $selectedAdapterName
                    Set-NetIPInterface -InterfaceAlias $selectedAdapter.Name -Dhcp Enabled
                    Set-DnsClientServerAddress -InterfaceAlias $selectedAdapter.Name -ResetServerAddresses
                    Write-Host "`nSe cambió a DHCP en el adaptador $($selectedAdapter.Name)." -ForegroundColor Green
                    [System.Windows.MessageBox]::Show("Se cambió a DHCP correctamente.", "Éxito",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                } catch {
                    Write-Host "`nError al cambiar a DHCP: $_" -ForegroundColor Red
                    [System.Windows.MessageBox]::Show("Error: $_", "Error",
                        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
            }
        })

    $btnClose.Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}

function Show-SSMSSelectionDialog {
    param(
        [array]$Managers,
        [array]$SSMSVersions
    )

    if ($Managers) {
        # Diálogo para SQL Server Configuration Manager
        [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Seleccionar Configuration Manager" Height="250" Width="450"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <Label Content="Seleccione la versión de Configuration Manager:"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0"/>
        <ComboBox Name="cmbManager" HorizontalAlignment="Left" VerticalAlignment="Top"
                  Width="410" Height="25" Margin="0,30,0,0"/>
        <Label Name="lblInfo" Content="" HorizontalAlignment="Left" VerticalAlignment="Top"
               Margin="0,65,0,0" Width="410"/>

        <Button Content="Aceptar" Name="btnOK" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,110,0,0"/>
        <Button Content="Cancelar" Name="btnCancel" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="150,110,0,0"/>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $cmbManager = $window.FindName("cmbManager")
        $lblInfo = $window.FindName("lblInfo")
        $btnOK = $window.FindName("btnOK")
        $btnCancel = $window.FindName("btnCancel")

        foreach ($manager in $Managers) {
            $cmbManager.Items.Add($manager) | Out-Null
        }
        $cmbManager.SelectedIndex = 0

        $cmbManager.Add_SelectionChanged({
                $path = $cmbManager.SelectedItem
                if ($path -match "SQLServerManager(\d+)") {
                    $version = $matches[1]
                    $arch = if ($path -match "SysWOW64") { "32bits" } else { "64bits" }
                    $lblInfo.Content = "SQLServerManager$version $arch"
                }
            })

        $btnOK.Add_Click({
                $selectedManager = $cmbManager.SelectedItem
                try {
                    Write-Host "`tEjecutando SQL Server Configuration Manager desde: $selectedManager" -ForegroundColor Green
                    Start-Process -FilePath $selectedManager
                    $window.Close()
                } catch {
                    Write-Host "`tError al ejecutar Manager." -ForegroundColor Red
                }
            })

        $btnCancel.Add_Click({
                Write-Host "`tOperación cancelada." -ForegroundColor Red
                $window.Close()
            })

        $window.ShowDialog() | Out-Null
    } elseif ($SSMSVersions) {
        # Diálogo para SSMS
        [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Seleccionar SSMS" Height="250" Width="450"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <Label Content="Seleccione la versión de Management Studio:"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0"/>
        <ComboBox Name="cmbSSMS" HorizontalAlignment="Left" VerticalAlignment="Top"
                  Width="410" Height="25" Margin="0,30,0,0"/>
        <Label Name="lblVersion" Content="Versión seleccionada:" HorizontalAlignment="Left"
               VerticalAlignment="Top" Margin="0,65,0,0"/>

        <Button Content="Aceptar" Name="btnOK" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,110,0,0"/>
        <Button Content="Cancelar" Name="btnCancel" Width="140" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="150,110,0,0"/>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $cmbSSMS = $window.FindName("cmbSSMS")
        $lblVersion = $window.FindName("lblVersion")
        $btnOK = $window.FindName("btnOK")
        $btnCancel = $window.FindName("btnCancel")

        foreach ($version in $SSMSVersions) {
            $cmbSSMS.Items.Add($version) | Out-Null
        }
        $cmbSSMS.SelectedIndex = 0

        $cmbSSMS.Add_SelectionChanged({
                $path = $cmbSSMS.SelectedItem
                if ($path -match 'Microsoft SQL Server\\(\d+)') {
                    $lblVersion.Content = "Versión seleccionada: SSMS $($matches[1])"
                } elseif ($path -match 'Management Studio (\d+)') {
                    $lblVersion.Content = "Versión seleccionada: SSMS $($matches[1])"
                }
            })

        $btnOK.Add_Click({
                $selectedVersion = $cmbSSMS.SelectedItem
                try {
                    Write-Host "`tEjecutando SSMS desde: $selectedVersion" -ForegroundColor Green
                    Start-Process -FilePath $selectedVersion
                    $window.Close()
                } catch {
                    Write-Host "`tError al ejecutar SSMS." -ForegroundColor Red
                }
            })

        $btnCancel.Add_Click({
                Write-Host "`tOperación cancelada." -ForegroundColor Red
                $window.Close()
            })

        $window.ShowDialog() | Out-Null
    }
}

function Show-LZMADialog {
    param([array]$Instaladores)

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Carpetas LZMA" Height="250" Width="400"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <ComboBox Name="cmbCarpetas" HorizontalAlignment="Left" VerticalAlignment="Top"
                  Width="360" Height="25" Margin="0,10,0,0"/>
        <Label Name="lblExePath" Content="AI_ExePath: -" HorizontalAlignment="Left"
               VerticalAlignment="Top" Margin="0,45,0,0" Width="360" Height="70"
               Foreground="Red"/>

        <Button Content="Renombrar" Name="btnRenombrar" Width="180" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,130,0,0" IsEnabled="False"/>
        <Button Content="Salir" Name="btnSalir" Width="180" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="190,130,0,0"/>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $cmbCarpetas = $window.FindName("cmbCarpetas")
    $lblExePath = $window.FindName("lblExePath")
    $btnRenombrar = $window.FindName("btnRenombrar")
    $btnSalir = $window.FindName("btnSalir")

    $cmbCarpetas.Items.Add("Selecciona instalador a renombrar") | Out-Null
    foreach ($inst in $Instaladores) {
        $cmbCarpetas.Items.Add($inst.Name) | Out-Null
    }
    $cmbCarpetas.SelectedIndex = 0

    $cmbCarpetas.Add_SelectionChanged({
            $idx = $cmbCarpetas.SelectedIndex
            $btnRenombrar.IsEnabled = ($idx -gt 0)

            if ($idx -gt 0) {
                $ruta = $Instaladores[$idx - 1].Path
                $prop = Get-ItemProperty -Path $ruta -Name "AI_ExePath" -ErrorAction SilentlyContinue
                $lblExePath.Content = if ($prop) { "AI_ExePath: $($prop.AI_ExePath)" } else { "AI_ExePath: No encontrado" }
            } else {
                $lblExePath.Content = "AI_ExePath: -"
            }
        })

    $btnRenombrar.Add_Click({
            $idx = $cmbCarpetas.SelectedIndex
            if ($idx -gt 0) {
                $rutaVieja = $Instaladores[$idx - 1].Path
                $nombre = $cmbCarpetas.SelectedItem
                $nuevaNombre = "$nombre.backup"

                $result = [System.Windows.MessageBox]::Show("¿Renombrar a $nuevaNombre?", "Confirmar",
                    [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)

                if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                    try {
                        Rename-Item -Path $rutaVieja -NewName $nuevaNombre
                        [System.Windows.MessageBox]::Show("Registro renombrado correctamente.", "Éxito",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                        $window.Close()
                    } catch {
                        [System.Windows.MessageBox]::Show("Error: $_", "Error",
                            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    }
                }
            }
        })

    $btnSalir.Add_Click({
            Write-Host "`tCancelado por el usuario." -ForegroundColor Yellow
            $window.Close()
        })

    $window.ShowDialog() | Out-Null
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
    'Show-NewIpForm',
    'Show-ProgressBar',
    'Update-ProgressBar',
    'Close-ProgressBar',
    'Set-ControlEnabled',
    'Create-WpfWindow',
    'Show-WpfMessageBox',
    'Show-WpfProgressBar',
    'Update-WpfProgressBar',
    'Close-WpfProgressBar',
    'Show-AddUserDialog',
    'Show-IPConfigDialog',
    'Show-SSMSSelectionDialog',
    'Show-LZMADialog'
)