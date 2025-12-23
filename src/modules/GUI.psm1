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


        # Manejador para cuando se cierra la ventana
        $window.Add_Closed({
                param($sender, $e)
                $sender.IsClosed = $true
            })

        # Mostrar de forma no modal
        $window.Show()

        # Forzar render inicial (siempre con el dispatcher de la misma ventana)
        $window.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action] {}
        ) | Out-Null

        # IMPORTANTE: retornar exactamente 1 objeto (Window), no array
        return [System.Windows.Window]$window

    } catch {
        Write-Error "Error al crear barra de progreso: $_"
        return $null
    }
}
function Update-WpfProgressBar {
    param(
        [Parameter(Mandatory = $true)] $Window,
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 100)] [int] $Percent,
        [string] $Message = $null
    )

    if ($null -eq $Window) { return }

    # Asegurar que sea ventana WPF
    if (-not ($Window -is [System.Windows.Window])) {
        Write-Warning "Update-WpfProgressBar: El objeto recibido NO es WPF Window. Tipo: $($Window.GetType().FullName)"
        return
    }

    if ($Window.IsClosed) { return }

    # Asegurar dispatcher vivo
    if ($null -eq $Window.Dispatcher -or
        $Window.Dispatcher.HasShutdownStarted -or
        $Window.Dispatcher.HasShutdownFinished) {
        Write-Warning "Update-WpfProgressBar: Dispatcher no disponible."
        return
    }

    try {
        # Pasa valores por parámetro para evitar problemas de closure
        $Window.Dispatcher.Invoke(
            [action[object, int, string]] {
                param($w, $p, $m)
                $w.ProgressBar.Value = [Math]::Min($p, 100)
                $w.PercentLabel.Text = "$p%"

                if (-not [string]::IsNullOrWhiteSpace($m)) {
                    $w.MessageLabel.Text = $m
                }
                $w.UpdateLayout()
            },
            [System.Windows.Threading.DispatcherPriority]::Render,
            @($Window, $Percent, $Message)
        )
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

function Show-SSMSSelectionDialog {
    param(
        [array]$Managers,
        [array]$SSMSVersions
    )

    # [Contenido original de Show-SSMSSelectionDialog - mantener igual]
    Write-Host "Show-SSMSSelectionDialog: Implementación pendiente o usar la original"
}

function Show-LZMADialog {
    param([array]$Instaladores)

    # [Contenido original de Show-LZMADialog - mantener igual]
    Write-Host "Show-LZMADialog: Implementación pendiente o usar la original"
}

function Show-ChocolateyInstallerMenu {
    <#
    .SYNOPSIS
        Menú de instalación de paquetes Chocolatey con búsqueda mejorada.
    #>

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Instaladores Choco" Height="420" Width="520"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10" Background="#505055">
        <Label Content="Buscar en Chocolatey:" HorizontalAlignment="Left" VerticalAlignment="Top"
               Margin="0,0,0,0" Foreground="White"/>
        <TextBox Name="txtChocoSearch" HorizontalAlignment="Left" VerticalAlignment="Top"
                 Width="360" Height="25" Margin="0,25,0,0"/>
        <Button Content="Buscar" Name="btnBuscarChoco" Width="120" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="370,23,0,0"
                Background="#4CAF50" Foreground="White"/>

        <Label Content="SSMS" Name="lblPresetSSMS" Width="70" Height="25"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,65,0,0"
               Background="#C8E6FF" HorizontalContentAlignment="Center"
               VerticalContentAlignment="Center" BorderBrush="Black" BorderThickness="1" Cursor="Hand"/>
        <Label Content="Heidi" Name="lblPresetHeidi" Width="70" Height="25"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="80,65,0,0"
               Background="#C8E6FF" HorizontalContentAlignment="Center"
               VerticalContentAlignment="Center" BorderBrush="Black" BorderThickness="1" Cursor="Hand"/>

        <Button Content="Mostrar instalados" Name="btnShowInstalledChoco" Width="150" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,100,0,0"/>
        <Button Content="Instalar seleccionado" Name="btnInstallSelectedChoco" Width="170" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="160,100,0,0" IsEnabled="False"/>
        <Button Content="Desinstalar seleccionado" Name="btnUninstallSelectedChoco" Width="150" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="340,100,0,0" IsEnabled="False"/>
        <DataGrid Name="dgvChocoResults" HorizontalAlignment="Left" VerticalAlignment="Top"
                Width="490" Height="200" Margin="0,145,0,0" IsReadOnly="True"
                AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Paquete" Binding="{Binding Name}" Width="170"/>
                <DataGridTextColumn Header="Versión" Binding="{Binding Version}" Width="100"/>
                <DataGridTextColumn Header="Descripción" Binding="{Binding Description}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <Button Content="Salir" Name="btnExitInstaladores" Width="490" Height="30"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,355,0,0"/>
    </Grid>
</Window>
"@

    try {
        $result = New-WpfWindow -Xaml $xaml -PassThru
        $window = $result.Window
    } catch {
        Write-Host "Error creando ventana: $_" -ForegroundColor Red
        return
    }

    # Obtener controles
    $txtChocoSearch = $window.FindName("txtChocoSearch")
    $btnBuscarChoco = $window.FindName("btnBuscarChoco")
    $lblPresetSSMS = $window.FindName("lblPresetSSMS")
    $lblPresetHeidi = $window.FindName("lblPresetHeidi")
    $btnShowInstalledChoco = $window.FindName("btnShowInstalledChoco")
    $btnInstallSelectedChoco = $window.FindName("btnInstallSelectedChoco")
    $btnUninstallSelectedChoco = $window.FindName("btnUninstallSelectedChoco")
    $dgvChocoResults = $window.FindName("dgvChocoResults")
    $btnExitInstaladores = $window.FindName("btnExitInstaladores")

    # Colección observable
    $chocoResultsCollection = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
    $dgvChocoResults.ItemsSource = $chocoResultsCollection

    # Función auxiliar para agregar resultados (SEARCH)
    $addChocoResult = {
        param($line)
        if ([string]::IsNullOrWhiteSpace($line)) { return }
        if ($line -match '^Chocolatey') { return }
        if ($line -match 'packages?\s+found' -or $line -match 'page size') { return }

        if ($line -match '^(?<name>[A-Za-z0-9\.\+\-_]+)\s+(?<version>[0-9][A-Za-z0-9\.\-]*)\s+(?<description>.+)$') {
            $window.Dispatcher.Invoke([action] {
                    $chocoResultsCollection.Add([PSCustomObject]@{
                            Name        = $Matches['name']
                            Version     = $Matches['version']
                            Description = $Matches['description'].Trim()
                        })
                }) | Out-Null
        } elseif ($line -match '^(?<name>[A-Za-z0-9\.\+\-_]+)\s+\|\s+(?<version>[0-9][A-Za-z0-9\.\-]*)$') {
            $window.Dispatcher.Invoke([action] {
                    $chocoResultsCollection.Add([PSCustomObject]@{
                            Name        = $Matches['name']
                            Version     = $Matches['version']
                            Description = "Paquete instalado"
                        })
                }) | Out-Null
        }
    }

    # Función auxiliar para agregar resultados (INSTALADOS)  ✅ NUEVA
    # Formato esperado con --limit-output: paquete|version
    $addChocoInstalled = {
        param($line)

        if ([string]::IsNullOrWhiteSpace($line)) { return }
        if ($line -match '^Chocolatey') { return }

        if ($line -match '^(?<name>[^|]+)\|(?<version>.+)$') {
            $name = $Matches['name'].Trim()
            $ver = $Matches['version'].Trim()

            $window.Dispatcher.Invoke([action] {
                    $chocoResultsCollection.Add([PSCustomObject]@{
                            Name        = $name
                            Version     = $ver
                            Description = "Paquete instalado"
                        })
                }) | Out-Null
        }
    }

    # Actualizar botones de acción
    $updateActionButtons = {
        $hasValidSelection = $false
        if ($dgvChocoResults.SelectedItem) {
            $selectedItem = $dgvChocoResults.SelectedItem
            if ($selectedItem.Name -and $selectedItem.Version -match '^[0-9]') {
                $hasValidSelection = $true
            }
        }
        $btnInstallSelectedChoco.IsEnabled = $hasValidSelection
        $btnUninstallSelectedChoco.IsEnabled = $hasValidSelection
    }

    $dgvChocoResults.Add_SelectionChanged({ & $updateActionButtons })

    # Botón Buscar - MEJORADO CON BARRA DE PROGRESO
    $btnBuscarChoco.Add_Click({
            $chocoResultsCollection.Clear()
            & $updateActionButtons

            $query = $txtChocoSearch.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($query)) {
                [System.Windows.MessageBox]::Show("Ingresa un término para buscar", "Búsqueda")
                return
            }

            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                [System.Windows.MessageBox]::Show("Chocolatey no está instalado", "Error")
                return
            }

            $btnBuscarChoco.IsEnabled = $false

            # USAR BARRA DE PROGRESO MEJORADA
            $progress = Show-WpfProgressBar -Title "Buscando paquetes" -Message "Iniciando búsqueda..."

            try {
                Update-WpfProgressBar -Window $progress -Percent 20 -Message "Verificando Chocolatey..."
                Start-Sleep -Milliseconds 300

                Update-WpfProgressBar -Window $progress -Percent 40 -Message "Buscando '$query'..."
                $searchOutput = & choco search $query --page-size=20 2>&1

                Update-WpfProgressBar -Window $progress -Percent 70 -Message "Procesando resultados..."

                foreach ($line in $searchOutput) {
                    & $addChocoResult $line
                }

                Update-WpfProgressBar -Window $progress -Percent 100 -Message "Búsqueda completada"
                Start-Sleep -Milliseconds 300

                if ($chocoResultsCollection.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No se encontraron paquetes", "Sin resultados")
                }
            } catch {
                Write-Error "Error: $_"
                [System.Windows.MessageBox]::Show("Error durante la búsqueda: $_", "Error")
            } finally {
                Close-WpfProgressBar -Window $progress
                $btnBuscarChoco.IsEnabled = $true
            }
        })

    # Presets
    $lblPresetSSMS.Add_MouseLeftButtonDown({
            $txtChocoSearch.Text = "ssms"
            $btnBuscarChoco.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        })

    $lblPresetHeidi.Add_MouseLeftButtonDown({
            $txtChocoSearch.Text = "heidi"
            $btnBuscarChoco.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        })

    # Mostrar instalados - FIX (parse correcto de choco list)
    $btnShowInstalledChoco.Add_Click({
            $chocoResultsCollection.Clear()
            & $updateActionButtons

            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                [System.Windows.MessageBox]::Show("Chocolatey no está instalado", "Error")
                return
            }

            $btnShowInstalledChoco.IsEnabled = $false
            $progress = Show-WpfProgressBar -Title "Listando instalados" -Message "Recuperando paquetes..."

            try {
                Update-WpfProgressBar -Window $progress -Percent 30 -Message "Consultando paquetes instalados..."

                # ✅ Este formato es el más fácil de parsear:
                # choco list --local-only --limit-output  => paquete|version
                $installedOutput = & choco list --local-only --limit-output 2>&1

                Update-WpfProgressBar -Window $progress -Percent 70 -Message "Procesando resultados..."

                foreach ($line in $installedOutput) {
                    & $addChocoInstalled $line
                }

                Update-WpfProgressBar -Window $progress -Percent 100 -Message "Completado"
                Start-Sleep -Milliseconds 200

                if ($chocoResultsCollection.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No hay paquetes instalados (o no se pudo leer la salida).", "Sin resultados")
                }
            } catch {
                Write-Error "Error: $_"
                [System.Windows.MessageBox]::Show("Error consultando paquetes: $_", "Error")
            } finally {
                Close-WpfProgressBar -Window $progress
                $btnShowInstalledChoco.IsEnabled = $true
            }
        })

    $btnInstallSelectedChoco.Add_Click({
            if (-not $dgvChocoResults.SelectedItem) {
                [System.Windows.MessageBox]::Show("Seleccione un paquete", "Instalación")
                return
            }

            $packageName = $dgvChocoResults.SelectedItem.Name

            $result = [System.Windows.MessageBox]::Show(
                "¿Instalar $packageName?",
                "Confirmar",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )

            if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                return
            }

            $progress = Show-WpfProgressBar -Title "Instalando" -Message "Preparando instalación..."

            try {
                Update-WpfProgressBar -Window $progress -Percent 20 -Message "Verificando paquete..."
                Start-Sleep -Milliseconds 300

                Update-WpfProgressBar -Window $progress -Percent 40 -Message "Instalando $packageName..."

                $installProcess = Start-Process -FilePath "choco" `
                    -ArgumentList "install", $packageName, "-y" `
                    -NoNewWindow -PassThru -Wait

                Update-WpfProgressBar -Window $progress -Percent 90 -Message "Verificando instalación..."
                Start-Sleep -Milliseconds 500

                if ($installProcess.ExitCode -eq 0) {
                    Update-WpfProgressBar -Window $progress -Percent 100 -Message "Instalación completada"
                    Start-Sleep -Milliseconds 500
                    [System.Windows.MessageBox]::Show("Paquete instalado exitosamente", "Éxito")
                } else {
                    throw "Error de instalación: código $($installProcess.ExitCode)"
                }
            } catch {
                Write-Error $_
                [System.Windows.MessageBox]::Show("Error: $_", "Error")
            } finally {
                Close-WpfProgressBar -Window $progress
            }
        })

    # Desinstalar - MEJORADO
    $btnUninstallSelectedChoco.Add_Click({
            if (-not $dgvChocoResults.SelectedItem) {
                [System.Windows.MessageBox]::Show("Seleccione un paquete", "Desinstalación")
                return
            }

            $packageName = $dgvChocoResults.SelectedItem.Name

            $result = [System.Windows.MessageBox]::Show(
                "¿Desinstalar $packageName?",
                "Confirmar",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )

            if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                return
            }

            $progress = Show-WpfProgressBar -Title "Desinstalando" -Message "Preparando desinstalación..."

            try {
                Update-WpfProgressBar -Window $progress -Percent 30 -Message "Desinstalando $packageName..."

                $uninstallProcess = Start-Process -FilePath "choco" `
                    -ArgumentList "uninstall", $packageName, "-y" `
                    -NoNewWindow -PassThru -Wait

                Update-WpfProgressBar -Window $progress -Percent 90 -Message "Verificando desinstalación..."
                Start-Sleep -Milliseconds 500

                if ($uninstallProcess.ExitCode -eq 0) {
                    Update-WpfProgressBar -Window $progress -Percent 100 -Message "Desinstalación completada"
                    Start-Sleep -Milliseconds 500
                    [System.Windows.MessageBox]::Show("Paquete desinstalado exitosamente", "Éxito")
                } else {
                    throw "Error: código $($uninstallProcess.ExitCode)"
                }
            } catch {
                Write-Error $_
                [System.Windows.MessageBox]::Show("Error: $_", "Error")
            } finally {
                Close-WpfProgressBar -Window $progress
            }
        })

    # Salir
    $btnExitInstaladores.Add_Click({ $window.Close() })

    # Mostrar ventana
    $window.ShowDialog() | Out-Null
}

#endregion

# Exportar todas las funciones
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
    'Show-SSMSSelectionDialog',
    'Show-LZMADialog',
    'Show-ChocolateyInstallerMenu'
)