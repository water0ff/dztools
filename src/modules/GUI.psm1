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
                Window = $window
                Controls = $controls
            }
        }
        return $window
    }
    catch {
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
        "OK" = [System.Windows.MessageBoxButton]::OK
        "OKCancel" = [System.Windows.MessageBoxButton]::OKCancel
        "YesNo" = [System.Windows.MessageBoxButton]::YesNo
        "YesNoCancel" = [System.Windows.MessageBoxButton]::YesNoCancel
    }
    $iconMap = @{
        "Information" = [System.Windows.MessageBoxImage]::Information
        "Warning" = [System.Windows.MessageBoxImage]::Warning
        "Error" = [System.Windows.MessageBoxImage]::Error
        "Question" = [System.Windows.MessageBoxImage]::Question
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
        [action]{}
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
        [action]{}
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
    }
    else {
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

Export-ModuleMember -Function `
    New-WpfWindow,
    Show-WpfMessageBox,
    New-WpfInputDialog,
    Show-WpfProgressBar,
    Update-WpfProgressBar,
    Close-WpfProgressBar,
    Set-WpfControlEnabled,
    Get-WpfPasswordBoxText,
    Add-WpfComboBoxItems,
    Show-WpfFileDialog,
    Show-WpfFolderDialog,
    Show-NewIpForm,
    Show-ProgressBar,
    Update-ProgressBar,
    Close-ProgressBar,
    Set-ControlEnabled