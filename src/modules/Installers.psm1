#requires -Version 5.0

# ============================================
# Installers.psm1 - Versión WPF
# Módulo de instalación de software
# ============================================

# ============================================
# FUNCIÓN: Check-Chocolatey
# ============================================
function Check-Chocolatey {
    <#
    .SYNOPSIS
    Verifica si Chocolatey está instalado y ofrece instalarlo

    .DESCRIPTION
    Comprueba la existencia de Chocolatey y solicita instalación si no existe

    .OUTPUTS
    Boolean - True si está instalado o instalación exitosa, False en caso contrario
    #>

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        $result = [System.Windows.MessageBox]::Show(
            "Chocolatey no está instalado. ¿Desea instalarlo ahora?",
            "Chocolatey no encontrado",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::No) {
            Write-Host "`nEl usuario canceló la instalación de Chocolatey." -ForegroundColor Red
            return $false
        }

        Write-Host "`nInstalando Chocolatey..." -ForegroundColor Cyan
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

            Write-Host "`nChocolatey se instaló correctamente." -ForegroundColor Green

            # Configurar cacheLocation
            Write-Host "`nConfigurando Chocolatey..." -ForegroundColor Yellow
            choco config set cacheLocation C:\Choco\cache

            [System.Windows.MessageBox]::Show(
                "Chocolatey se instaló correctamente y ha sido configurado. Por favor, reinicie PowerShell antes de continuar.",
                "Reinicio requerido",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            # Cerrar el programa automáticamente
            Write-Host "`nCerrando la aplicación para permitir reinicio de PowerShell..." -ForegroundColor Red
            Stop-Process -Id $PID -Force
            return $false
        } catch {
            Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
            [System.Windows.MessageBox]::Show(
                "Error al instalar Chocolatey. Por favor, inténtelo manualmente.",
                "Error de instalación",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            return $false
        }
    } else {
        Write-Host "`tChocolatey ya está instalado." -ForegroundColor Green
        return $true
    }
}

# ============================================
# FUNCIÓN: Test-ChocolateyInstalled
# ============================================
function Test-ChocolateyInstalled {
    <#
    .SYNOPSIS
    Verifica si Chocolatey está instalado sin solicitar instalación

    .OUTPUTS
    Boolean - True si está instalado, False en caso contrario
    #>

    return $null -ne (Get-Command choco -ErrorAction SilentlyContinue)
}

# ============================================
# FUNCIÓN: Install-Software
# ============================================
function Install-Software {
    <#
    .SYNOPSIS
    Instala software mediante Chocolatey

    .PARAMETER Software
    Nombre del software a instalar (SQL2014, SQL2019, SSMS, 7Zip, MegaTools)

    .PARAMETER Force
    Forzar reinstalación si ya existe

    .OUTPUTS
    Boolean - True si la instalación fue exitosa, False en caso contrario
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('SQL2014', 'SQL2019', 'SSMS', '7Zip', 'MegaTools')]
        [string]$Software,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    # Verificar Chocolatey
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Error "Chocolatey no está instalado. Use Check-Chocolatey primero."
        return $false
    }

    switch ($Software) {
        'SQL2014' {
            Write-Host "Instalando SQL Server 2014 Express..." -ForegroundColor Cyan
            $arguments = @(
                'install', 'sql-server-express',
                '-y',
                '--version=2014.0.2000.8',
                '--params', '"/SQLUSER:sa /SQLPASSWORD:National09 /INSTANCENAME:NationalSoft /FEATURES:SQL"'
            )
        }

        'SQL2019' {
            Write-Host "Instalando SQL Server 2019 Express..." -ForegroundColor Cyan
            $arguments = @(
                'install', 'sql-server-express',
                '-y',
                '--version=2019.20190106',
                '--params', '"/SQLUSER:sa /SQLPASSWORD:National09 /INSTANCENAME:SQL2019 /FEATURES:SQL"'
            )
        }

        'SSMS' {
            Write-Host "Instalando SQL Server Management Studio..." -ForegroundColor Cyan
            $arguments = @('install', 'sql-server-management-studio', '-y')
        }

        '7Zip' {
            Write-Host "Instalando 7-Zip..." -ForegroundColor Cyan
            $arguments = @('install', '7zip', '-y')
        }

        'MegaTools' {
            Write-Host "Instalando MegaTools..." -ForegroundColor Cyan
            $arguments = @('install', 'megatools', '-y')
        }
    }

    try {
        Start-Process choco -ArgumentList $arguments -NoNewWindow -Wait

        Write-Host "$Software instalado correctamente" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error instalando $Software : $_"
        return $false
    }
}

# ============================================
# FUNCIÓN: Download-File
# ============================================
function Download-File {
    <#
    .SYNOPSIS
    Descarga un archivo desde una URL

    .PARAMETER Url
    URL del archivo a descargar

    .PARAMETER OutputPath
    Ruta de destino del archivo

    .PARAMETER ShowProgress
    Mostrar barra de progreso durante la descarga

    .OUTPUTS
    Boolean - True si la descarga fue exitosa, False en caso contrario
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $false)]
        [switch]$ShowProgress
    )

    try {
        if ($ShowProgress) {
            # Descarga con barra de progreso
            $webClient = New-Object System.Net.WebClient

            # Evento para mostrar progreso
            $eventHandler = {
                $global:downloadPercentage = $_.ProgressPercentage
                Write-Progress -Activity "Descargando..." -Status "$($global:downloadPercentage)%" `
                    -PercentComplete $global:downloadPercentage
            }

            $webClient.add_DownloadProgressChanged($eventHandler)
            $webClient.DownloadFileAsync((New-Object System.Uri($Url)), $OutputPath)

            # Esperar a que termine la descarga
            while ($webClient.IsBusy) {
                Start-Sleep -Milliseconds 100
            }

            $webClient.remove_DownloadProgressChanged($eventHandler)
            Write-Progress -Activity "Descargando..." -Completed
        } else {
            # Descarga simple
            Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
        }

        Write-Verbose "Archivo descargado: $OutputPath"
        return $true
    } catch {
        Write-Error "Error descargando archivo: $_"
        return $false
    }
}

# ============================================
# FUNCIÓN: Expand-ArchiveFile
# ============================================
function Expand-ArchiveFile {
    <#
    .SYNOPSIS
    Extrae un archivo comprimido

    .PARAMETER ArchivePath
    Ruta del archivo comprimido

    .PARAMETER DestinationPath
    Ruta de destino para extraer

    .PARAMETER Force
    Forzar extracción sobrescribiendo archivos existentes

    .OUTPUTS
    Boolean - True si la extracción fue exitosa, False en caso contrario
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ArchivePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    try {
        # Usar Expand-Archive de PowerShell 5
        if ($Force -and (Test-Path $DestinationPath)) {
            Remove-Item $DestinationPath -Recurse -Force
        }

        if (-not (Test-Path $DestinationPath)) {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        }

        Expand-Archive -Path $ArchivePath -DestinationPath $DestinationPath -Force

        Write-Verbose "Archivo extraído a: $DestinationPath"
        return $true
    } catch {
        Write-Error "Error extrayendo archivo: $_"
        return $false
    }
}

# ============================================
# FUNCIÓN: Show-SSMSInstallerDialog (WPF)
# ============================================
function Show-SSMSInstallerDialog {
    <#
    .SYNOPSIS
    Muestra diálogo para seleccionar versión de SSMS a instalar

    .OUTPUTS
    String - "latest" o "ssms14", $null si se cancela
    #>

    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Instalar SSMS"
        Height="200" Width="380"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="$($theme.FormBackground)">

    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                            VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="$($theme.AccentPrimary)"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Título -->
        <TextBlock Grid.Row="0"
                  Text="Elige la versión a instalar:"
                  FontSize="13"
                  FontWeight="Bold"
                  Margin="0,0,0,15"/>

        <!-- ComboBox -->
        <ComboBox Grid.Row="1"
                 Name="cmbVersion"
                 FontSize="12"
                 Padding="8"
                 Margin="0,0,0,20">
            <ComboBoxItem Content="Último disponible" IsSelected="True"/>
            <ComboBoxItem Content="SSMS 14 (2014)"/>
        </ComboBox>

        <!-- Espacio -->
        <Grid Grid.Row="2"/>

        <!-- Botones -->
        <StackPanel Grid.Row="3"
                   Orientation="Horizontal"
                   HorizontalAlignment="Right">
            <Button Name="btnOK"
                   Content="Instalar"
                   Width="140"
                   Margin="0,0,10,0"/>
            <Button Name="btnCancel"
                   Content="Cancelar"
                   Width="140"/>
        </StackPanel>
    </Grid>
</Window>
"@

    try {
        [xml]$xaml = $stringXaml
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)


        # Obtener controles
        $cmbVersion = $window.FindName("cmbVersion")
        $btnOK = $window.FindName("btnOK")
        $btnCancel = $window.FindName("btnCancel")

        # Variable para almacenar resultado
        $script:selectedVersion = $null

        # Eventos
        $btnOK.Add_Click({
                $selectedIndex = $cmbVersion.SelectedIndex
                $script:selectedVersion = switch ($selectedIndex) {
                    0 { "latest" }
                    1 { "ssms14" }
                    default { "latest" }
                }
                $window.DialogResult = $true
                $window.Close()
            })

        $btnCancel.Add_Click({
                $script:selectedVersion = $null
                $window.DialogResult = $false
                $window.Close()
            })

        # Mostrar diálogo
        $result = $window.ShowDialog()

        if ($result) {
            return $script:selectedVersion
        }

        return $null

    } catch {
        Write-Error "Error mostrando diálogo de SSMS: $_"
        return $null
    }
}

# ============================================
# FUNCIÓN: Test-7ZipInstalled
# ============================================
function Test-7ZipInstalled {
    <#
    .SYNOPSIS
    Verifica si 7-Zip está instalado

    .OUTPUTS
    Boolean - True si está instalado, False en caso contrario
    #>

    $paths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $true
        }
    }

    return $false
}

# ============================================
# FUNCIÓN: Test-MegaToolsInstalled
# ============================================
function Test-MegaToolsInstalled {
    <#
    .SYNOPSIS
    Verifica si MegaTools está instalado

    .OUTPUTS
    Boolean - True si está instalado, False en caso contrario
    #>

    return $null -ne (Get-Command megatools -ErrorAction SilentlyContinue)
}

# ============================================
# FUNCIÓN: Get-InstalledChocoPackages
# ============================================
function Get-InstalledChocoPackages {
    <#
    .SYNOPSIS
    Obtiene lista de paquetes instalados con Chocolatey

    .OUTPUTS
    Array de objetos con Name y Version
    #>

    if (-not (Test-ChocolateyInstalled)) {
        Write-Warning "Chocolatey no está instalado"
        return @()
    }

    try {
        $output = choco list --local-only --limit-output

        $packages = @()
        foreach ($line in $output) {
            if ($line -match '^(?<name>[^\|]+)\|(?<version>.+)$') {
                $packages += [PSCustomObject]@{
                    Name    = $Matches['name']
                    Version = $Matches['version']
                }
            }
        }

        return $packages
    } catch {
        Write-Error "Error obteniendo paquetes instalados: $_"
        return @()
    }
}

# ============================================
# FUNCIÓN: Search-ChocoPackages
# ============================================
function Search-ChocoPackages {
    <#
    .SYNOPSIS
    Busca paquetes en el repositorio de Chocolatey

    .PARAMETER Query
    Término de búsqueda

    .OUTPUTS
    Array de objetos con Name, Version y Description
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query
    )

    if (-not (Test-ChocolateyInstalled)) {
        Write-Warning "Chocolatey no está instalado"
        return @()
    }

    try {
        $output = choco search $Query --page-size=20

        $packages = @()
        $pattern = '^(?<name>[A-Za-z0-9\.\+\-_]+)\s+(?<version>[0-9][A-Za-z0-9\.\-]*)\s+(?<description>.+)$'

        foreach ($line in $output) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match '^Chocolatey') { continue }
            if ($line -match 'packages?\s+found' -or $line -match 'page size') { continue }

            if ($line -match $pattern) {
                $packages += [PSCustomObject]@{
                    Name        = $Matches['name']
                    Version     = $Matches['version']
                    Description = $Matches['description'].Trim()
                }
            }
        }

        return $packages
    } catch {
        Write-Error "Error buscando paquetes: $_"
        return @()
    }
}

# ============================================
# FUNCIÓN: Install-ChocoPackage
# ============================================
function Install-ChocoPackage {
    <#
    .SYNOPSIS
    Instala un paquete de Chocolatey

    .PARAMETER PackageName
    Nombre del paquete a instalar

    .PARAMETER Version
    Versión específica (opcional)

    .OUTPUTS
    Boolean - True si la instalación fue exitosa, False en caso contrario
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageName,

        [Parameter(Mandatory = $false)]
        [string]$Version = $null
    )

    if (-not (Test-ChocolateyInstalled)) {
        Write-Error "Chocolatey no está instalado"
        return $false
    }

    try {
        $arguments = "install $PackageName -y"

        if (-not [string]::IsNullOrWhiteSpace($Version)) {
            $arguments = "install $PackageName --version=$Version -y"
        }

        Write-Host "Instalando $PackageName..." -ForegroundColor Cyan
        Start-Process choco -ArgumentList $arguments -NoNewWindow -Wait

        Write-Host "$PackageName instalado correctamente" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error instalando $PackageName : $_"
        return $false
    }
}

# ============================================
# FUNCIÓN: Uninstall-ChocoPackage
# ============================================
function Uninstall-ChocoPackage {
    <#
    .SYNOPSIS
    Desinstala un paquete de Chocolatey

    .PARAMETER PackageName
    Nombre del paquete a desinstalar

    .OUTPUTS
    Boolean - True si la desinstalación fue exitosa, False en caso contrario
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageName
    )

    if (-not (Test-ChocolateyInstalled)) {
        Write-Error "Chocolatey no está instalado"
        return $false
    }

    try {
        $arguments = "uninstall $PackageName -y"

        Write-Host "Desinstalando $PackageName..." -ForegroundColor Yellow
        Start-Process choco -ArgumentList $arguments -NoNewWindow -Wait

        Write-Host "$PackageName desinstalado correctamente" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error desinstalando $PackageName : $_"
        return $false
    }
}
function Show-ChocolateyInstallerMenu {
    <#
    .SYNOPSIS
        Menú de instalación de paquetes Chocolatey con búsqueda mejorada.
    #>

    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Instaladores Choco" Height="420" Width="520"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        Background="$($theme.FormBackground)">
    <Window.Resources>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="RowBackground" Value="$($theme.ControlBackground)"/>
            <Setter Property="AlternatingRowBackground" Value="$($theme.InfoBackground)"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="$($theme.AccentPrimary)"/>
                    <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
        </Style>
        <Style x:Key="PresetLabelStyle" TargetType="Label">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="GeneralButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonGeneralBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonGeneralForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10" Background="$($theme.FormBackground)">
        <Label Content="Buscar en Chocolatey:" HorizontalAlignment="Left" VerticalAlignment="Top"
               Margin="0,0,0,0"/>
        <TextBox Name="txtChocoSearch" HorizontalAlignment="Left" VerticalAlignment="Top"
                 Width="360" Height="25" Margin="0,25,0,0"/>
        <Button Content="Buscar" Name="btnBuscarChoco" Width="120" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="370,23,0,0"
                Style="{StaticResource GeneralButtonStyle}"/>

        <Label Content="SSMS" Name="lblPresetSSMS" Width="70" Height="25"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,65,0,0"
               Style="{StaticResource PresetLabelStyle}"/>
        <Label Content="Heidi" Name="lblPresetHeidi" Width="70" Height="25"
               HorizontalAlignment="Left" VerticalAlignment="Top" Margin="80,65,0,0"
               Style="{StaticResource PresetLabelStyle}"/>

        <Button Content="Mostrar instalados" Name="btnShowInstalledChoco" Width="150" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,100,0,0" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Content="Instalar seleccionado" Name="btnInstallSelectedChoco" Width="170" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="160,100,0,0" IsEnabled="False" Style="{StaticResource GeneralButtonStyle}"/>
        <Button Content="Desinstalar seleccionado" Name="btnUninstallSelectedChoco" Width="150" Height="32"
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="340,100,0,0" IsEnabled="False" Style="{StaticResource GeneralButtonStyle}"/>
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
                HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,355,0,0" Style="{StaticResource GeneralButtonStyle}"/>
    </Grid>
</Window>
"@

    try {
        $result = New-WpfWindow -Xaml $stringXaml -PassThru
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
function Invoke-PortableTool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ToolName,

        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$ZipPath,

        [Parameter(Mandatory)]
        [string]$ExtractPath,

        [Parameter(Mandatory)]
        [string]$ExeName,

        # UI opcional
        [System.Windows.Controls.TextBlock]$InfoTextBlock,

        # Si tuvieras un Owner WPF para MessageBox (opcional)
        [System.Windows.Window]$OwnerWindow
    )

    try {
        $exePath = Join-Path $ExtractPath $ExeName

        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] ToolName=$ToolName" ([System.ConsoleColor]::DarkGray)
        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] Url=$Url" ([System.ConsoleColor]::DarkGray)
        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] Zip=$ZipPath" ([System.ConsoleColor]::DarkGray)
        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] Extract=$ExtractPath" ([System.ConsoleColor]::DarkGray)
        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] ExePath=$exePath" ([System.ConsoleColor]::DarkGray)

        # 1) ¿Ya existe?
        if (Test-Path -LiteralPath $exePath) {

            $msg = "$ToolName ya existe en:`n$exePath`n`nSí = Abrir local`nNo = Volver a descargar`nCancelar = Cancelar operación"
            $rExist = Show-WpfMessageBox -Message $msg -Title "Ya existe" -Buttons "YesNoCancel" -Icon "Question"

            Write-DzDebug "`t[DEBUG][Invoke-PortableTool] rExist=$rExist" ([System.ConsoleColor]::Cyan)

            if ($rExist -eq [System.Windows.MessageBoxResult]::Yes) {
                Start-Process $exePath
                if ($InfoTextBlock) { $InfoTextBlock.Text = "Abriendo $ToolName existente..." }
                return $true
            }

            if ($rExist -eq [System.Windows.MessageBoxResult]::Cancel) {
                if ($InfoTextBlock) { $InfoTextBlock.Text = "Operación cancelada." }
                return $false
            }

            # No => sigue a descargar
        } else {
            $r = Show-WpfMessageBox -Message "¿Deseas descargar $ToolName?" -Title "Confirmar descarga" -Buttons "YesNo" -Icon "Question"
            Write-DzDebug "`t[DEBUG][Invoke-PortableTool] confirm=$r" ([System.ConsoleColor]::Cyan)

            if ($r -ne [System.Windows.MessageBoxResult]::Yes) {
                if ($InfoTextBlock) { $InfoTextBlock.Text = "Operación cancelada." }
                return $false
            }
        }

        # 2) Progress
        $pw = Show-WpfProgressBar -Title "Descargando $ToolName" -Message "Iniciando..."
        if ($null -eq $pw -or $null -eq $pw.ProgressBar) {
            Write-DzDebug "`t[DEBUG][Invoke-PortableTool] ERROR: No progress window" ([System.ConsoleColor]::Red)
            return $false
        }

        try {
            # Limpia zip anterior
            if (Test-Path -LiteralPath $ZipPath) {
                Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
            }

            Update-WpfProgressBar -Window $pw -Percent 0 -Message "Preparando descarga..."
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Preparando descarga..." }

            # 3) Descargar (tu función existente)
            $ok = Download-FileWithProgressWpfStream -Url $Url -OutFile $ZipPath -Window $pw -OnStatus {
                param($p, $m)
                Write-DzDebug "`t[DEBUG][Invoke-PortableTool] $p% - $m" ([System.ConsoleColor]::DarkGray)
                if ($InfoTextBlock) { $InfoTextBlock.Text = $m }
            }
            if (-not $ok) { throw "Descarga fallida." }

            Update-WpfProgressBar -Window $pw -Percent 100 -Message "Extrayendo..."
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Extrayendo..." }

            # 4) Si el exe está corriendo, no borres su carpeta
            $exeRunning = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path -ieq $exePath }

            if (-not $exeRunning) {
                if (Test-Path -LiteralPath $ExtractPath) {
                    Remove-Item -LiteralPath $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                Write-DzDebug "`t[DEBUG][Invoke-PortableTool] WARN: $ExeName está en ejecución, no se limpia carpeta" ([System.ConsoleColor]::Yellow)
            }

            Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force

            if (-not (Test-Path -LiteralPath $exePath)) {
                throw "No se encontró el ejecutable: $exePath"
            }

            Update-WpfProgressBar -Window $pw -Percent 100 -Message "Ejecutando..."
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Ejecutando..." }

            Start-Process $exePath
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: Ejecutado $ToolName" }

            Write-DzDebug "`t[DEBUG][Invoke-PortableTool] OK: Ejecutado $ToolName" ([System.ConsoleColor]::Cyan)
            return $true

        } catch {
            $err = $_.Exception.Message
            Write-DzDebug "`t[DEBUG][Invoke-PortableTool] ERROR: $err" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
            return $false
        } finally {
            Close-WpfProgressBar -Window $pw
        }

    } catch {
        Write-DzDebug "`t[DEBUG][Invoke-PortableTool] FATAL: $($_.Exception.Message)`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
        return $false
    }
}

Export-ModuleMember -Function @(
    'Check-Chocolatey',
    'Test-ChocolateyInstalled',
    'Install-Software',
    'Download-File',
    'Expand-ArchiveFile',
    'Show-SSMSInstallerDialog',
    'Test-7ZipInstalled',
    'Test-MegaToolsInstalled',
    'Get-InstalledChocoPackages',
    'Search-ChocoPackages',
    'Install-ChocoPackage',
    'Uninstall-ChocoPackage',
    'Show-ChocolateyInstallerMenu',
    'Invoke-PortableTool'
)
