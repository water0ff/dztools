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
Export-ModuleMember -Function @(
    'New-WpfWindow',
    'Show-WpfMessageBox',
    'New-WpfInputDialog',
    'Show-WpfProgressBar',
    'Update-WpfProgressBar',
    'Close-WpfProgressBar',
    'Set-WpfControlEnabled',
    'Get-WpfPasswordBoxText',
    'Show-WpfFolderDialog',
    'Show-ProgressBar',
    'Show-AddUserDialog'
)