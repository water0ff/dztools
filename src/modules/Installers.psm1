#requires -Version 5.0
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
        $result = Show-WpfMessageBox `
            -Message "Chocolatey no está instalado. ¿Desea instalarlo ahora?" `
            -Title "Chocolatey no encontrado" `
            -Buttons "YesNo" `
            -Icon "Question"

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
            Show-WpfMessageBox `
                -Message "Chocolatey se instaló correctamente y ha sido configurado. Por favor, reinicie PowerShell antes de continuar." `
                -Title "Reinicio requerido" `
                -Buttons "OK" `
                -Icon "Information" | Out-Null
            # Cerrar el programa automáticamente
            Write-Host "`nCerrando la aplicación para permitir reinicio de PowerShell..." -ForegroundColor Red
            Stop-Process -Id $PID -Force
            return $false
        } catch {
            Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
            Show-WpfMessageBox `
                -Message "Error al instalar Chocolatey. Por favor, inténtelo manualmente.`n`nDetalle: $($_.Exception.Message)" `
                -Title "Error de instalación" `
                -Buttons "OK" `
                -Icon "Error" | Out-Null
            return $false
        }
    } else {
        Write-Host "`tChocolatey ya está instalado." -ForegroundColor Green
        return $true
    }
}
function Test-ChocolateyInstalled {
    <#
    .SYNOPSIS
    Verifica si Chocolatey está instalado sin solicitar instalación
    .OUTPUTS
    Boolean - True si está instalado, False en caso contrario
    #>

    return $null -ne (Get-Command choco -ErrorAction SilentlyContinue)
}
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

    <!-- Texto por defecto -->
    <Style TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
    </Style>

    <!-- TextBox por defecto -->
    <Style TargetType="TextBox">
        <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="BorderThickness" Value="1"/>
    </Style>

    <!-- Botón base: SIEMPRE define Disabled -->
    <Style x:Key="BaseButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Opacity" Value="1"/>
        <Setter Property="Padding" Value="12,6"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="6"
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center"
                                          VerticalAlignment="Center"/>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>

        <!-- Aquí está el FIX: disabled legible -->
        <Style.Triggers>
            <Trigger Property="IsEnabled" Value="False">
                <!-- no dejes que WPF “lave” los colores -->
                <Setter Property="Opacity" Value="1"/>
                <Setter Property="Cursor" Value="Arrow"/>
                <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            </Trigger>
        </Style.Triggers>
    </Style>

    <!-- Acción (botones verdes/azules) -->
    <Style x:Key="ActionButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
        <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
        <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Style.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
                <!-- disabled coherente también aquí -->
                <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            </Trigger>
        </Style.Triggers>
    </Style>
        <Style x:Key="OutlineButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                    <Setter Property="BorderThickness" Value="0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

    <!-- ========================= -->
    <!-- DataGrid: estilos completos -->
    <!-- ========================= -->

    <Style TargetType="DataGrid">
        <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="RowBackground" Value="{DynamicResource PanelBg}"/>
        <Setter Property="AlternatingRowBackground" Value="{DynamicResource ControlBg}"/>
        <Setter Property="GridLinesVisibility" Value="Horizontal"/>
        <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="VerticalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
    </Style>

    <Style TargetType="DataGridColumnHeader">
        <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="BorderThickness" Value="0,0,0,1"/>
        <Setter Property="Padding" Value="8,6"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <Style TargetType="DataGridRow">
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
                <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style TargetType="DataGridCell">
        <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
        <Setter Property="Background" Value="Transparent"/>
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
        <Setter Property="BorderThickness" Value="0,0,0,1"/>
        <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
                <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
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
        Title="Instaladores (Chocolatey)"
        Height="520" Width="720"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">

    <!-- ✅ ESTO ES LO QUE TE FALTA -->
    <Window.Resources>

        <!-- Texto por defecto -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource PanelFg}"/>
        </Style>

        <!-- TextBox por defecto -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
            <Style x:Key="BaseButtonStyle" TargetType="Button">
                <Setter Property="OverridesDefaultStyle" Value="True"/>
                <Setter Property="SnapsToDevicePixels" Value="True"/>

                <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Cursor" Value="Hand"/>
                <Setter Property="Padding" Value="12,6"/>

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}"
                                    BorderBrush="{TemplateBinding BorderBrush}"
                                    BorderThickness="{TemplateBinding BorderThickness}"
                                    CornerRadius="8"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center"
                                                VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>

                <Style.Triggers>
                    <Trigger Property="IsEnabled" Value="False">
                        <Setter Property="Opacity" Value="1"/>
                        <Setter Property="Cursor" Value="Arrow"/>
                        <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                        <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
                    </Trigger>
                </Style.Triggers>
            </Style>
        <Style x:Key="ActionButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
            <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
                    <Setter Property="Foreground" Value="{DynamicResource AccentMuted}"/>
                    <Setter Property="BorderThickness" Value="1"/>
                    <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="OutlineButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>

            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentSecondary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                    <Setter Property="BorderThickness" Value="0"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- ✅ DataGrid -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource PanelBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="{DynamicResource PanelBg}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource ControlBg}"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="VerticalGridLinesBrush" Value="{DynamicResource BorderBrushColor}"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource ControlBg}"/>
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{DynamicResource AccentPrimary}"/>
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{DynamicResource ControlFg}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Foreground" Value="{DynamicResource FormFg}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

    </Window.Resources>
    <!-- ✅ FIN DE RESOURCES -->

    <Grid Background="{DynamicResource FormBg}" Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="12"
                Margin="0,0,0,10">
            <DockPanel>
                <Button DockPanel.Dock="Left"
                        Name="btnExitInstaladores"
                        Content="Salir"
                        Width="120"
                        Height="34"
                        Margin="0,0,12,0"
                        Style="{StaticResource OutlineButtonStyle}"/>

                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="Chocolatey Installer"
                               Foreground="{DynamicResource FormFg}"
                               FontSize="16"
                               FontWeight="SemiBold"/>
                    <TextBlock Text="Busca, instala y desinstala paquetes"
                               Foreground="{DynamicResource PanelFg}"
                               Margin="0,2,0,0"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- Search + presets + actions -->
        <Border Grid.Row="1"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="12"
                Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <TextBox Name="txtChocoSearch"
                             Height="34"
                             Padding="10,6"
                             VerticalContentAlignment="Center"/>

                    <Button Grid.Column="1"
                            Name="btnBuscarChoco"
                            Content="Buscar"
                            Width="140"
                            Height="34"
                            Margin="10,0,0,0"
                            Style="{StaticResource ActionButtonStyle}"/>
                </Grid>

                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                    <TextBlock Text="Presets:"
                               VerticalAlignment="Center"
                               Margin="0,0,10,0"/>

                    <Button Name="btnPresetSSMS"
                            Content="SSMS"
                            Width="90"
                            Height="30"
                            Margin="0,0,10,0"
                            Background="{DynamicResource AccentMuted}"
                            Foreground="{DynamicResource FormFg}"
                            BorderThickness="0"
                            Cursor="Hand"/>

                    <Button Name="btnPresetHeidi"
                            Content="Heidi"
                            Width="90"
                            Height="30"
                            Background="{DynamicResource AccentMuted}"
                            Foreground="{DynamicResource FormFg}"
                            BorderThickness="0"
                            Cursor="Hand"/>
                </StackPanel>

                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
                    <Button Name="btnShowInstalledChoco"
                            Content="Mostrar instalados"
                            Width="170"
                            Height="34"
                            Style="{StaticResource OutlineButtonStyle}"/>

                    <Button Name="btnInstallSelectedChoco"
                            Content="Instalar seleccionado"
                            Width="190"
                            Height="34"
                            Margin="10,0,0,0"
                            IsEnabled="False"
                            Style="{StaticResource ActionButtonStyle}"/>

                    <Button Name="btnUninstallSelectedChoco"
                            Content="Desinstalar seleccionado"
                            Width="200"
                            Height="34"
                            Margin="10,0,0,0"
                            IsEnabled="False"
                            Style="{StaticResource OutlineButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Results -->
        <Border Grid.Row="2"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="10">
            <Grid>
                <DataGrid Name="dgvChocoResults"
                          IsReadOnly="True"
                          AutoGenerateColumns="False"
                          SelectionMode="Single"
                          CanUserAddRows="False">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Paquete" Binding="{Binding Name}" Width="220"/>
                        <DataGridTextColumn Header="Versión" Binding="{Binding Version}" Width="140"/>
                        <DataGridTextColumn Header="Descripción" Binding="{Binding Description}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>
        <!-- Status -->
        <Border Grid.Row="3"
                Background="{DynamicResource PanelBg}"
                BorderBrush="{DynamicResource BorderBrushColor}"
                BorderThickness="1"
                CornerRadius="10"
                Padding="10"
                Margin="0,10,0,0">
            <DockPanel>
                <TextBlock Name="txtStatus"
                           Text="Listo."
                           VerticalAlignment="Center"/>
            </DockPanel>
        </Border>
    </Grid>
</Window>
"@

    try {
        $result = New-WpfWindow -Xaml $stringXaml -PassThru
        $window = $result.Window
        $theme = Get-DzUiTheme
        Set-DzWpfThemeResources -Window $window -Theme $theme
    } catch {
        Write-Host "Error creando ventana: $_" -ForegroundColor Red
        return
    }
    $txtChocoSearch = $window.FindName("txtChocoSearch")
    $btnBuscarChoco = $window.FindName("btnBuscarChoco")
    $btnPresetSSMS = $window.FindName("btnPresetSSMS")
    $btnPresetHeidi = $window.FindName("btnPresetHeidi")
    $btnShowInstalledChoco = $window.FindName("btnShowInstalledChoco")
    $btnInstallSelectedChoco = $window.FindName("btnInstallSelectedChoco")
    $btnUninstallSelectedChoco = $window.FindName("btnUninstallSelectedChoco")
    $dgvChocoResults = $window.FindName("dgvChocoResults")
    $btnExitInstaladores = $window.FindName("btnExitInstaladores")
    $txtStatus = $window.FindName("txtStatus")
    # Colección observable
    $chocoResultsCollection = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
    $resetChocoUi = {
        $window.Dispatcher.Invoke([action] {
                $chocoResultsCollection.Clear()
                $dgvChocoResults.SelectedItem = $null
                $txtChocoSearch.Text = ""
                $txtStatus.Text = "Listo."
                $btnInstallSelectedChoco.IsEnabled = $false
                $btnUninstallSelectedChoco.IsEnabled = $false
            }) | Out-Null
    }
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
                Show-WpfMessageBox -Message "Ingresa un término para buscar" -Title "Búsqueda" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                return
            }

            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                Show-WpfMessageBox -Message "Chocolatey no está instalado." -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
                return
            }
            $btnBuscarChoco.IsEnabled = $false
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
                    Show-WpfMessageBox -Message "No se encontraron paquetes." -Title "Sin resultados" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                }
            } catch {
                Write-Error "Error: $_"
                Show-WpfMessageBox -Message "Error durante la búsqueda:`n`n$($_.Exception.Message)" -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
            } finally {
                Close-WpfProgressBar -Window $progress
                $btnBuscarChoco.IsEnabled = $true
            }
        })
    $btnPresetSSMS.Add_Click({
            $txtChocoSearch.Text = "ssms"
            $btnBuscarChoco.RaiseEvent(
                (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))
            )
        })

    $btnPresetHeidi.Add_Click({
            $txtChocoSearch.Text = "heidi"
            $btnBuscarChoco.RaiseEvent(
                (New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))
            )
        })
    $btnShowInstalledChoco.Add_Click({
            $chocoResultsCollection.Clear()
            & $updateActionButtons

            if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
                Show-WpfMessageBox -Message "Chocolatey no está instalado." -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
                return
            }
            $btnShowInstalledChoco.IsEnabled = $false
            $progress = Show-WpfProgressBar -Title "Listando instalados" -Message "Recuperando paquetes..."
            try {
                Update-WpfProgressBar -Window $progress -Percent 30 -Message "Consultando paquetes instalados..."
                $installedOutput = & choco list --local-only --limit-output 2>&1
                Update-WpfProgressBar -Window $progress -Percent 70 -Message "Procesando resultados..."

                foreach ($line in $installedOutput) {
                    & $addChocoInstalled $line
                }
                Update-WpfProgressBar -Window $progress -Percent 100 -Message "Completado"
                Start-Sleep -Milliseconds 200

                if ($chocoResultsCollection.Count -eq 0) {
                    Show-WpfMessageBox -Message "No hay paquetes instalados (o no se pudo leer la salida)." -Title "Sin resultados" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                }
            } catch {
                Write-Error "Error: $_"
                Show-WpfMessageBox -Message "Error consultando paquetes:`n`n$($_.Exception.Message)" -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
            } finally {
                Close-WpfProgressBar -Window $progress
                $btnShowInstalledChoco.IsEnabled = $true
            }
        })
    $btnInstallSelectedChoco.Add_Click({
            Write-DzDebug "`t[DEBUG] [Install-ChocoPackage] Iniciando instalación del paquete seleccionado."
            if (-not $dgvChocoResults.SelectedItem) {
                Show-WpfMessageBox -Message "Seleccione un paquete." -Title "Instalación" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                return
            }
            $packageName = $dgvChocoResults.SelectedItem.Name
            $result = Show-WpfMessageBoxSafe -Message "¿Instalar $packageName?" -Title "Confirmar" -Buttons "YesNo" -Icon "Question" -Owner $window
            Write-DzDebug "`t[DEBUG][Install-ChocoPackage] User response: $result"
            if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
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
                    Show-WpfMessageBox -Message "Paquete instalado exitosamente." -Title "Éxito" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                    try { $resetChocoUi.Invoke() } catch { }
                } else {
                    throw "Error de instalación: código $($installProcess.ExitCode)"
                }
            } catch {
                Write-Error $_
                Show-WpfMessageBox -Message "Error: $($_.Exception.Message)" -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
            } finally {
                Close-WpfProgressBar -Window $progress
            }
        })
    # Desinstalar - MEJORADO
    $btnUninstallSelectedChoco.Add_Click({
            if (-not $dgvChocoResults.SelectedItem) {
                Show-WpfMessageBox -Message "Seleccione un paquete." -Title "Desinstalación" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                return
            }
            $packageName = $dgvChocoResults.SelectedItem.Name
            $result = Show-WpfMessageBox -Message "¿Desinstalar $packageName?" -Title "Confirmar" -Buttons "YesNo" -Icon "Warning" -Owner $window
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
                    Show-WpfMessageBox -Message "Paquete desinstalado exitosamente." -Title "Éxito" -Buttons "OK" -Icon "Information" -Owner $window | Out-Null
                    try { $resetChocoUi.Invoke() } catch { }
                } else {
                    throw "Error: código $($uninstallProcess.ExitCode)"
                }
            } catch {
                Write-Error $_
                Show-WpfMessageBox -Message "Error:`n`n$($_.Exception.Message)" -Title "Error" -Buttons "OK" -Icon "Error" -Owner $window | Out-Null
            } finally {
                Close-WpfProgressBar -Window $progress
            }
        })
    $btnExitInstaladores.Add_Click({
            & $resetChocoUi
            $window.Close()
        })
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
            $rExist = Show-WpfMessageBox -Message $msg -Title "Ya existe" -Buttons "YesNoCancel" -Icon "Question" -Owner $OwnerWindow
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
            $r = Show-WpfMessageBox -Message "¿Deseas descargar $ToolName?" -Title "Confirmar descarga" -Buttons "YesNo" -Icon "Question" -Owner $OwnerWindow
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