#requires -Version 5.0

$script:DzToolsConfigPath = "C:\Temp\dztools.ini"
$script:DzDebugEnabled = $null

function Get-DzToolsConfigPath {
    <#
    .SYNOPSIS
    Obtiene la ruta del archivo de configuración
    #>
    return $script:DzToolsConfigPath
}
function Get-DzDebugPreference {
    <#
    .SYNOPSIS
    Obtiene la preferencia de debug desde el archivo de configuración
    #>
    $configPath = Get-DzToolsConfigPath

    if (-not (Test-Path -LiteralPath $configPath)) {
        return $false
    }
    $content = Get-Content -LiteralPath $configPath -ErrorAction SilentlyContinue
    $inDevSection = $false
    foreach ($line in $content) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^\s*;') { continue }
        if ($trimmed -match '^\[desarrollo\]\s*$') {
            $inDevSection = $true
            continue
        }
        if ($inDevSection -and $trimmed -match '^\[') {
            break
        }
        if ($inDevSection -and $trimmed -match '^\s*debug\s*=\s*(.+)\s*$') {
            return ($matches[1].ToLower() -eq 'true')
        }
    }
    return $false
}
function Ensure-DzUiConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return
    }

    $content = Get-Content -LiteralPath $ConfigPath -ErrorAction SilentlyContinue
    if ($null -eq $content) { $content = @() }

    $lines = New-Object System.Collections.Generic.List[string]
    $uiFound = $false
    $modeFound = $false
    $inUiSection = $false

    foreach ($line in $content) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^\s*;') {
            $lines.Add($line)
            continue
        }
        if ($trimmed -match '^\[UI\]\s*$') {
            $uiFound = $true
            $inUiSection = $true
            $lines.Add($line)
            continue
        }
        if ($inUiSection -and $trimmed -match '^\[') {
            if (-not $modeFound) {
                $lines.Add("mode=dark")
                $modeFound = $true
            }
            $inUiSection = $false
        }
        if ($inUiSection -and $trimmed -match '^\s*mode\s*=\s*(.+)\s*$') {
            $modeFound = $true
        }
        $lines.Add($line)
    }

    if ($uiFound -and -not $modeFound) {
        $lines.Add("mode=dark")
    }
    if (-not $uiFound) {
        if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -ne "") {
            $lines.Add("")
        }
        $lines.Add("[UI]")
        $lines.Add("mode=dark")
    }

    Set-Content -LiteralPath $ConfigPath -Value $lines -Encoding UTF8
}

function Update-DzIniSetting {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Section,

        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $configPath = Get-DzToolsConfigPath

    if (-not (Test-Path -LiteralPath $configPath)) {
        Initialize-DzToolsConfig | Out-Null
    }

    $content = Get-Content -LiteralPath $configPath -ErrorAction SilentlyContinue
    if ($null -eq $content) { $content = @() }

    $lines = New-Object System.Collections.Generic.List[string]
    $sectionFound = $false
    $keyUpdated = $false
    $inTargetSection = $false

    foreach ($line in $content) {
        $trimmed = $line.Trim()

        if ($trimmed -match '^\s*;') {
            $lines.Add($line)
            continue
        }

        if ($trimmed -match '^\[(.+)\]\s*$') {
            if ($inTargetSection -and -not $keyUpdated) {
                $lines.Add("$Key=$Value")
                $keyUpdated = $true
            }

            $currentSection = $matches[1].Trim()
            $inTargetSection = ($currentSection.ToLower() -eq $Section.ToLower())
            if ($inTargetSection) { $sectionFound = $true }

            $lines.Add($line)
            continue
        }

        if ($inTargetSection -and $trimmed -match "^\s*${Key}\s*=") {
            $lines.Add("$Key=$Value")
            $keyUpdated = $true
            continue
        }

        $lines.Add($line)
    }

    if ($sectionFound) {
        if (-not $keyUpdated) {
            $lines.Add("$Key=$Value")
        }
    } else {
        if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -ne "") {
            $lines.Add("")
        }
        $lines.Add("[$Section]")
        $lines.Add("$Key=$Value")
    }

    Set-Content -LiteralPath $configPath -Value $lines -Encoding UTF8
}

function Get-DzUiMode {
    <#
    .SYNOPSIS
    Obtiene el modo de UI desde el archivo de configuración
    #>
    $configPath = Get-DzToolsConfigPath

    if (-not (Test-Path -LiteralPath $configPath)) {
        return "dark"
    }
    $content = Get-Content -LiteralPath $configPath -ErrorAction SilentlyContinue
    $inUiSection = $false
    foreach ($line in $content) {
        $trimmed = $line.Trim()
        if ($trimmed -match '^\s*;') { continue }
        if ($trimmed -match '^\[UI\]\s*$') {
            $inUiSection = $true
            continue
        }
        if ($inUiSection -and $trimmed -match '^\[') {
            break
        }
        if ($inUiSection -and $trimmed -match '^\s*mode\s*=\s*(.+)\s*$') {
            return $matches[1].ToLower()
        }
    }
    return "dark"
}

function Set-DzUiMode {
    <#
    .SYNOPSIS
    Actualiza el modo de UI en el archivo de configuración
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('dark', 'light')]
        [string]$Mode
    )

    Update-DzIniSetting -Section "UI" -Key "mode" -Value $Mode
}

function Set-DzDebugPreference {
    <#
    .SYNOPSIS
    Actualiza la preferencia de debug en el archivo de configuración
    #>
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    $value = if ($Enabled) { 'true' } else { 'false' }
    Update-DzIniSetting -Section "desarrollo" -Key "debug" -Value $value
    $script:DzDebugEnabled = $Enabled
}

function Initialize-DzToolsConfig {
    <#
    .SYNOPSIS
    Inicializa el archivo de configuración
    #>
    $configPath = Get-DzToolsConfigPath
    $configDir = Split-Path -Path $configPath -Parent
    if (-not (Test-Path -LiteralPath $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        "[desarrollo]`ndebug=false`n`n[UI]`nmode=dark" | Out-File -FilePath $configPath -Encoding UTF8 -Force
    }
    Ensure-DzUiConfig -ConfigPath $configPath
    $script:DzDebugEnabled = Get-DzDebugPreference
    return $script:DzDebugEnabled
}

function Write-DzDebug {
    <#
    .SYNOPSIS
    Escribe mensajes de debug si está habilitado
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [System.ConsoleColor]$Color = [System.ConsoleColor]::Gray
    )

    if ($null -eq $script:DzDebugEnabled) {
        $script:DzDebugEnabled = Get-DzDebugPreference
    }

    if ($script:DzDebugEnabled) {

        # ✅ Si hay una barra de progreso “inline”, ciérrala para no mezclar texto
        if (Get-Command Stop-GlobalProgress -ErrorAction SilentlyContinue) {
            Stop-GlobalProgress
        } else {
            # fallback: por si no existe, evita pegarse a un -NoNewline
            Write-Host ""
        }

        Write-Host $Message -ForegroundColor $Color
    }
}

function Test-Administrator {
    <#
    .SYNOPSIS
    Verifica si el script se ejecuta con privilegios de administrador
    #>
    [CmdletBinding()]
    param()
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Get-SystemInfo {
    <#
    .SYNOPSIS
    Obtiene información del sistema
    #>
    [CmdletBinding()]
    param()
    $info = @{
        ComputerName      = [System.Net.Dns]::GetHostName()
        OS                = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        NetAdapters       = @()
    }
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    foreach ($adapter in $adapters) {
        $adapterInfo = @{
            Name        = $adapter.Name
            Status      = $adapter.Status
            MacAddress  = $adapter.MacAddress
            IPAddresses = @()
        }
        $ipAddresses = Get-NetIPAddress -InterfaceAlias $adapter.Name -AddressFamily IPv4 |
        Where-Object { $_.IPAddress -ne '127.0.0.1' }

        foreach ($ip in $ipAddresses) {
            $adapterInfo.IPAddresses += $ip.IPAddress
        }
        $info.NetAdapters += $adapterInfo
    }
    return $info
}
function Clear-TemporaryFiles {
    <#
    .SYNOPSIS
    Limpia archivos temporales del sistema
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Paths = @("$env:TEMP", "C:\Windows\Temp")
    )
    $totalDeleted = 0
    $totalSize = 0
    foreach ($path in $Paths) {
        if (Test-Path $path) {
            try {
                $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                foreach ($item in $items) {
                    try {
                        if ($item.PSIsContainer) {
                            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        } else {
                            $totalSize += $item.Length
                            Remove-Item -Path $item.FullName -Force -ErrorAction SilentlyContinue
                        }
                        $totalDeleted++
                    } catch {
                        Write-Verbose "No se pudo eliminar: $($item.FullName)"
                    }
                }
            } catch {
                Write-Warning "Error accediendo a $path : $_"
            }
        }
    }
    return @{
        FilesDeleted = $totalDeleted
        SpaceFreedMB = [math]::Round($totalSize / 1MB, 2)
    }
}
function Test-ChocolateyInstalled {
    <#
    .SYNOPSIS
    Verifica si Chocolatey está instalado
    #>
    [CmdletBinding()]
    param()
    return [bool](Get-Command choco -ErrorAction SilentlyContinue)
}
function Install-Chocolatey {
    <#
    .SYNOPSIS
    Instala Chocolatey
    #>
    [CmdletBinding()]
    param()
    if (Test-ChocolateyInstalled) {
        Write-Verbose "Chocolatey ya está instalado"
        return $true
    }
    try {
        Write-Host "Instalando Chocolatey..." -ForegroundColor Yellow
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        choco config set cacheLocation C:\Choco\cache
        Write-Host "Chocolatey instalado correctamente" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Error instalando Chocolatey: $_"
        return $false
    }
}
function Get-AdminGroupName {
    <#
    .SYNOPSIS
    Obtiene el nombre del grupo de administradores del sistema
    #>
    $groups = net localgroup | Where-Object { $_ -match "Administrador|Administrators" }
    if ($groups -match "\bAdministradores\b") {
        return "Administradores"
    } elseif ($groups -match "\bAdministrators\b") {
        return "Administrators"
    }
    try {
        $adminGroup = Get-LocalGroup | Where-Object { $_.SID -like "S-1-5-32-544" }
        return $adminGroup.Name
    } catch {
        return "Administrators"
    }
}
function Invoke-DiskCleanup {
    <#
    .SYNOPSIS
    Ejecuta el liberador de espacio en disco de Windows
    #>
    [CmdletBinding()]
    param(
        [switch]$Configure,
        [switch]$Wait,
        [int]$TimeoutMinutes = 3,
        $ProgressWindow = $null
    )
    try {
        $cleanmgr = Join-Path $env:SystemRoot "System32\cleanmgr.exe"
        $profileId = 9999
        Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: INICIO Configure=$Configure, Wait=$Wait, TimeoutMinutes=$TimeoutMinutes"
        if ($Configure) {
            Write-Host "`n`tAbriendo configuración del Liberador de espacio..." -ForegroundColor Cyan
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: lanzando /sageset:$profileId"
            Start-Process $cleanmgr -ArgumentList "/sageset:$profileId" -Verb RunAs
            return
        }
        Write-Host "`n`tEjecutando Liberador de espacio en disco..." -ForegroundColor Cyan
        if ($Wait) {
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: lanzando /sagerun:$profileId (BLOQUEANTE con timeout)"
            Write-Host "`n`tEsperando a que termine la limpieza de disco..." -ForegroundColor Yellow
            Write-Host "`t(Timeout: $TimeoutMinutes minutos)" -ForegroundColor Yellow
            $proc = Start-Process $cleanmgr -ArgumentList "/sagerun:$profileId" -WindowStyle Hidden -PassThru
            if ($null -eq $proc) {
                throw "Invoke-DiskCleanup: Start-Process devolvió NULL (no se pudo iniciar cleanmgr)."
            }
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Proceso iniciado. PID=$($proc.Id)"
            $timeoutSeconds = $TimeoutMinutes * 60
            $script:remainingSeconds = $timeoutSeconds
            $script:cleanupCompleted = $false
            # antes del if ($ProgressWindow...)
            $timer = $null

            if ($ProgressWindow -ne $null -and $ProgressWindow.IsVisible) {

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(1)

                $timer.Add_Tick({
                        param($sender, $e)

                        try {
                            if ($script:remainingSeconds -gt 0) {
                                $script:remainingSeconds--

                                $mins = [math]::Floor($script:remainingSeconds / 60)
                                $secs = $script:remainingSeconds % 60

                                if ($ProgressWindow -and $ProgressWindow.PSObject.Properties.Name -contains 'MessageLabel') {
                                    $ProgressWindow.MessageLabel.Text = "Liberando espacio en disco...`nTiempo restante: $mins min $secs seg"
                                }
                            } else {
                                # ✅ en vez de $this.Stop()
                                $sender.Stop()
                            }
                        } catch {
                            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Tick EXCEPCIÓN: $($_.Exception.Message)" Red
                            $sender.Stop()
                        }
                    })

                $timer.Start()
                Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Timer iniciado"
            }
            $checkInterval = 500
            $elapsed = 0
            while (-not $proc.HasExited -and $elapsed -lt ($timeoutSeconds * 1000)) {
                Start-Sleep -Milliseconds $checkInterval
                $elapsed += $checkInterval
                if ($ProgressWindow -ne $null -and $ProgressWindow.IsVisible) {
                    $ProgressWindow.Dispatcher.Invoke(
                        [System.Windows.Threading.DispatcherPriority]::Background,
                        [action] {}
                    )
                }
            }
            if ($timer) {
                $timer.Stop()
            }
            if ($proc.HasExited) {
                Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Proceso completado. ExitCode=$($proc.ExitCode)"
                Write-Host "`n`tLiberador de espacio completado." -ForegroundColor Green
                if ($ProgressWindow -ne $null -and $ProgressWindow.IsVisible) {
                    if ($ProgressWindow.PSObject.Properties.Name -contains 'MessageLabel') {
                        $ProgressWindow.MessageLabel.Text = "Limpieza completada exitosamente"
                    }
                }
            } else {
                Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: TIMEOUT alcanzado" Yellow
                Write-Host "`n`tAdvertencia: El proceso excedió el tiempo límite." -ForegroundColor Yellow
                try {
                    $proc.Kill()
                    $proc.WaitForExit(5000)
                    Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Proceso terminado forzosamente"
                } catch {
                    Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Error al terminar proceso" Red
                }
            }
        } else {
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: lanzando /sagerun:$profileId (NO bloqueante)"
            $proc = Start-Process $cleanmgr `
                -ArgumentList "/sagerun:$profileId" `
                -WindowStyle Hidden `
                -PassThru
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: lanzado. PID=$($proc.Id)"
            Write-Host "`n`tLiberador de espacio iniciado." -ForegroundColor Green
        }
        Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: FIN OK"
    } catch {
        Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: EXCEPCIÓN: $($_.Exception.Message)" Red
        Write-Host "`n`tError en limpieza de disco: $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Stop-CleanmgrProcesses {
    <#
    .SYNOPSIS
    Detiene todos los procesos de cleanmgr activos
    #>
    [CmdletBinding()]
    param()
    try {
        $cleanmgrProcesses = Get-Process -Name "cleanmgr" -ErrorAction SilentlyContinue
        if ($cleanmgrProcesses) {
            Write-Host "`n`tEncontrando procesos cleanmgr activos..." -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]Stop-CleanmgrProcesses: Encontrados $($cleanmgrProcesses.Count) procesos"
            foreach ($proc in $cleanmgrProcesses) {
                Write-Host "`t  Terminando proceso cleanmgr (PID: $($proc.Id))..." -ForegroundColor Yellow
                $proc.Kill()
                Start-Sleep -Milliseconds 500
            }
            Write-Host "`tProcesos cleanmgr terminados." -ForegroundColor Green
            Write-DzDebug "`t[DEBUG]Stop-CleanmgrProcesses: Procesos terminados"
        } else {
            Write-DzDebug "`t[DEBUG]Stop-CleanmgrProcesses: No hay procesos cleanmgr activos"
        }
    } catch {
        Write-DzDebug "`t[DEBUG]Stop-CleanmgrProcesses: Error: $($_.Exception.Message)" Red
        Write-Host "`tError al terminar procesos: $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Show-SystemComponents {
    <#
    .SYNOPSIS
    Muestra los componentes del sistema detectados
    #>
    param(
        [switch]$SkipOnError
    )
    Write-Host "`n=== Componentes del sistema detectados ===" -ForegroundColor Cyan
    $os = $null
    $maxAttempts = 3
    $retryDelaySeconds = 2
    for ($attempt = 1; $attempt -le $maxAttempts -and -not $os; $attempt++) {
        try {
            $os = Get-CimInstance -ClassName CIM_OperatingSystem -ErrorAction Stop
        } catch {
            if ($attempt -lt $maxAttempts) {
                $msg = "Show-SystemComponents: ERROR intento {0}: {1}. Reintento en {2}s" -f `
                    $attempt, $_.Exception.Message, $retryDelaySeconds

                Write-Host $msg -ForegroundColor DarkYellow
                Start-Sleep -Seconds $retryDelaySeconds
            } else {
                Write-Host "`n[Windows]" -ForegroundColor Yellow
                Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red

                if (-not $SkipOnError) {
                    throw "No se pudo obtener información crítica del sistema"
                } else {
                    Write-Host "Continuando sin información del sistema..." -ForegroundColor Yellow
                    return
                }
            }
        }
    }
    if (-not $os) {
        if (-not $SkipOnError) {
            throw "No se pudo obtener información crítica del sistema"
        }
        return
    }
    Write-Host "`n[Windows]" -ForegroundColor Yellow
    Write-Host "Versión: $($os.Caption) (Build $($os.Version))" -ForegroundColor White
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_Processor..."
        $procesador = Get-CimInstance -ClassName CIM_Processor -ErrorAction Stop
        Write-Host "`n[Procesador]" -ForegroundColor Yellow
        Write-Host "Modelo: $($procesador.Name)" -ForegroundColor White
        Write-Host "Núcleos: $($procesador.NumberOfCores)" -ForegroundColor White
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo procesador" Red
        Write-Host "`n[Procesador]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_PhysicalMemory..."
        $memoria = Get-CimInstance -ClassName CIM_PhysicalMemory -ErrorAction Stop
        Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
        $memoria | ForEach-Object {
            Write-Host "Módulo: $([math]::Round($_.Capacity/1GB, 2)) GB $($_.Manufacturer) ($($_.Speed) MHz)" -ForegroundColor White
        }
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo memoria" Red
        Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    try {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Obteniendo CIM_DiskDrive..."
        $discos = Get-CimInstance -ClassName CIM_DiskDrive -ErrorAction Stop
        Write-Host "`n[Discos duros]" -ForegroundColor Yellow
        $discos | ForEach-Object {
            Write-Host "Disco: $($_.Model) ($([math]::Round($_.Size/1GB, 2)) GB)" -ForegroundColor White
        }
    } catch {
        Write-DzDebug "`t[DEBUG]Show-SystemComponents: Error leyendo discos" Red
        Write-Host "`n[Discos duros]" -ForegroundColor Yellow
        Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-DzDebug "`t[DEBUG]Show-SystemComponents: FIN"
}
function Test-SameHost {
    <#
    .SYNOPSIS
    Verifica si el servidor SQL está en el mismo host
    #>
    param(
        [string]$serverName
    )
    $machinePart = $serverName.Split('\')[0]
    $machineName = $machinePart.Split(',')[0]
    if ($machineName -eq '.') {
        $machineName = $env:COMPUTERNAME
    }
    return ($env:COMPUTERNAME -eq $machineName)
}
function Test-7ZipInstalled {
    <#
    .SYNOPSIS
    Verifica si 7-Zip está instalado
    #>
    return (Test-Path "C:\Program Files\7-Zip\7z.exe")
}
function Test-MegaToolsInstalled {
    <#
    .SYNOPSIS
    Verifica si MegaTools está instalado
    #>
    return ([bool](Get-Command megatools -ErrorAction SilentlyContinue))
}
function Check-Permissions {
    <#
    .SYNOPSIS
    Verifica y opcionalmente modifica permisos de carpeta
    #>
    param(
        [string]$folderPath = "C:\NationalSoft"
    )
    if (-not (Test-Path -LiteralPath $folderPath)) {
        Write-Host "La carpeta $folderPath no existe." -ForegroundColor Red

        [System.Windows.MessageBox]::Show(
            "La carpeta '$folderPath' no existe en este equipo.`r`nCrea la carpeta o corrige la ruta antes de continuar.",
            "Carpeta no encontrada",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }
    try {
        $acl = Get-Acl -LiteralPath $folderPath -ErrorAction Stop
    } catch {
        Write-Host "Error obteniendo ACL de $folderPath : $_" -ForegroundColor Red
        [System.Windows.MessageBox]::Show(
            "Error obteniendo permisos de '$folderPath':`r`n$($_.Exception.Message)",
            "Error al obtener permisos",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return
    }
    $permissions = @()
    $everyoneSid = New-Object System.Security.Principal.SecurityIdentifier(
        [System.Security.Principal.WellKnownSidType]::WorldSid,
        $null
    )
    $everyonePermissions = @()
    $everyoneHasFullControl = $false
    foreach ($access in $acl.Access) {
        try {
            $userSid = ($access.IdentityReference).Translate(
                [System.Security.Principal.SecurityIdentifier]
            )
        } catch {
            continue
        }
        $permissions += [PSCustomObject]@{
            Usuario = $access.IdentityReference
            Permiso = $access.FileSystemRights
            Tipo    = $access.AccessControlType
        }
        if ($userSid -eq $everyoneSid) {
            $everyonePermissions += $access.FileSystemRights
            if ($access.FileSystemRights -match "FullControl") {
                $everyoneHasFullControl = $true
            }
        }
    }
    Write-Host ""
    Write-Host "Permisos en $folderPath :" -ForegroundColor Cyan
    $permissions | ForEach-Object {
        Write-Host "`t$($_.Usuario) - $($_.Tipo) - $($_.Permiso)" -ForegroundColor Green
    }
    if ($everyonePermissions.Count -gt 0) {
        Write-Host "`tEveryone tiene: $($everyonePermissions -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "`tNo hay permisos para 'Everyone'" -ForegroundColor Red
    }
    if (-not $everyoneHasFullControl) {
        $result = [System.Windows.MessageBox]::Show(
            "El usuario 'Everyone' no tiene permisos de 'Full Control'. ¿Deseas concederlo?",
            "Permisos 'Everyone'",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                $directoryInfo = New-Object System.IO.DirectoryInfo($folderPath)
                $dirAcl = $directoryInfo.GetAccessControl()
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $everyoneSid,
                    [System.Security.AccessControl.FileSystemRights]::FullControl,
                    [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit",
                    [System.Security.AccessControl.PropagationFlags]::None,
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $dirAcl.AddAccessRule($accessRule)
                $dirAcl.SetAccessRuleProtection($false, $true)
                $directoryInfo.SetAccessControl($dirAcl)
                Write-Host "Se ha concedido 'Full Control' a 'Everyone'." -ForegroundColor Green
                [System.Windows.MessageBox]::Show(
                    "Se ha concedido 'Full Control' a 'Everyone' en '$folderPath'.",
                    "Permisos actualizados",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                ) | Out-Null
            } catch {
                Write-Host "Error aplicando permisos: $_" -ForegroundColor Red
                [System.Windows.MessageBox]::Show(
                    "Error aplicando permisos a '$folderPath':`r`n$($_.Exception.Message)",
                    "Error al aplicar permisos",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            }
        }
    }
}

function Refresh-AdapterStatus {
    try {
        if ($null -eq $global:txt_AdapterStatus) {
            Write-Host "ADVERTENCIA: El control de estado de adaptadores no está disponible." -ForegroundColor Yellow
            return
        }

        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        $adapterInfo = @()

        foreach ($adapter in $adapters) {
            $profile = Get-NetConnectionProfile -InterfaceAlias $adapter.Name -ErrorAction SilentlyContinue
            $networkType = if ($profile) {
                switch ($profile.NetworkCategory) {
                    'Private' { "Privada" }
                    'Public' { "Pública" }
                    'DomainAuthenticated' { "Dominio" }
                    default { "Desconocida" }
                }
            } else {
                "Sin perfil"
            }

            $adapterInfo += "$($adapter.Name): $networkType"
        }

        if ($adapterInfo.Count -gt 0) {
            $global:txt_AdapterStatus.Dispatcher.Invoke([action] {
                    $global:txt_AdapterStatus.Text = $adapterInfo -join "`n"
                })
        } else {
            $global:txt_AdapterStatus.Dispatcher.Invoke([action] {
                    $global:txt_AdapterStatus.Text = "Sin adaptadores activos"
                })
        }
    } catch {
        Write-Host "Error al actualizar estado de adaptadores: $_" -ForegroundColor Red
    }
}
function Get-NetworkAdapterStatus {
    <#
    .SYNOPSIS
    Obtiene el estado de los adaptadores de red
    #>
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    $profiles = Get-NetConnectionProfile
    $adapterStatus = @()
    foreach ($adapter in $adapters) {
        $profile = $profiles | Where-Object { $_.InterfaceIndex -eq $adapter.ifIndex }
        $networkCategory = if ($profile) { $profile.NetworkCategory } else { "Desconocido" }

        $adapterStatus += [PSCustomObject]@{
            AdapterName     = $adapter.Name
            NetworkCategory = $networkCategory
            InterfaceIndex  = $adapter.ifIndex
        }
    }
    return $adapterStatus
}
function Start-SystemUpdate {
    <#
    .SYNOPSIS
    Ejecuta el proceso completo de actualización del sistema
    #>
    param(
        [int]$DiskCleanupTimeoutMinutes = 3
    )
    $progressWindow = $null
    Write-DzDebug "`t[DEBUG]Start-SystemUpdate: INICIO (Timeout=$DiskCleanupTimeoutMinutes min)"
    try {
        $progressWindow = Show-WpfProgressBar -Title "Actualización del Sistema" -Message "Iniciando proceso..."
        $totalSteps = 5
        $currentStep = 0
        Write-Host "`nIniciando proceso de actualización..." -ForegroundColor Cyan
        Write-DzDebug "`t[DEBUG]Paso 1: Deteniendo servicio winmgmt..."
        Write-Host "`n[Paso 1/$totalSteps] Deteniendo servicio winmgmt..." -ForegroundColor Yellow
        $service = Get-Service -Name "winmgmt" -ErrorAction Stop
        Write-DzDebug "`t[DEBUG]Paso 1: Estado actual winmgmt=$($service.Status)"
        if ($service.Status -eq "Running") {
            Update-WpfProgressBar -Window $progressWindow -Percent 10 -Message "Deteniendo servicio WMI..."
            Stop-Service -Name "winmgmt" -Force -ErrorAction Stop
            Write-DzDebug "`t[DEBUG]Paso 1: winmgmt detenido OK"
            Write-Host "`n`tServicio detenido correctamente." -ForegroundColor Green
        }
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 20 -Message "Servicio WMI detenido"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 2: Renombrando carpeta Repository..."
        Write-Host "`n[Paso 2/$totalSteps] Renombrando carpeta Repository..." -ForegroundColor Yellow
        Update-WpfProgressBar -Window $progressWindow -Percent 30 -Message "Renombrando Repository..."
        try {
            $repoPath = Join-Path $env:windir "System32\Wbem\Repository"
            Write-DzDebug "`t[DEBUG]Paso 2: repoPath=$repoPath"

            if (Test-Path $repoPath) {
                $newName = "Repository_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Rename-Item -Path $repoPath -NewName $newName -Force -ErrorAction Stop
                Write-DzDebug "`t[DEBUG]Paso 2: Carpeta renombrada a $newName"
                Write-Host "`n`tCarpeta renombrada: $newName" -ForegroundColor Green
            }
        } catch {
            Write-DzDebug "`t[DEBUG]Paso 2: EXCEPCIÓN: $($_.Exception.Message)" Red
            Write-Host "`n`tAdvertencia: No se pudo renombrar Repository. Continuando..." -ForegroundColor Yellow
        }
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 40 -Message "Repository renovado"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 3: Reiniciando servicio winmgmt..."
        Write-Host "`n[Paso 3/$totalSteps] Reiniciando servicio winmgmt..." -ForegroundColor Yellow
        Update-WpfProgressBar -Window $progressWindow -Percent 50 -Message "Reiniciando servicio WMI..."
        net start winmgmt *>&1 | Write-Host
        Write-DzDebug "`t[DEBUG]Paso 3: winmgmt reiniciado"
        Write-Host "`n`tServicio WMI reiniciado." -ForegroundColor Green
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 60 -Message "Servicio WMI reiniciado"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 4: Limpiando archivos temporales..."
        Write-Host "`n[Paso 4/$totalSteps] Limpiando archivos temporales..." -ForegroundColor Cyan
        Update-WpfProgressBar -Window $progressWindow -Percent 70 -Message "Limpiando archivos temporales..."
        $cleanupResult = Clear-TemporaryFiles
        Write-DzDebug "`t[DEBUG]Paso 4: FilesDeleted=$($cleanupResult.FilesDeleted) SpaceFreedMB=$($cleanupResult.SpaceFreedMB)"
        Write-Host "`n`tTotal archivos eliminados: $($cleanupResult.FilesDeleted)" -ForegroundColor Green
        Write-Host "`n`tEspacio liberado: $($cleanupResult.SpaceFreedMB) MB" -ForegroundColor Green
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 80 -Message "Archivos temporales limpiados"
        Start-Sleep -Milliseconds 500
        Write-DzDebug "`t[DEBUG]Paso 5: Ejecutando Liberador de espacio..."
        Write-Host "`n[Paso 5/$totalSteps] Ejecutando Liberador de espacio..." -ForegroundColor Cyan
        Update-WpfProgressBar -Window $progressWindow -Percent 90 -Message "Preparando limpieza de disco..."
        Invoke-DiskCleanup -Wait -TimeoutMinutes $DiskCleanupTimeoutMinutes -ProgressWindow $progressWindow
        Write-DzDebug "`t[DEBUG]Paso 5: Liberador completado"
        $currentStep++
        Update-WpfProgressBar -Window $progressWindow -Percent 100 -Message "Proceso completado exitosamente"
        Start-Sleep -Seconds 1
        if ($progressWindow -ne $null -and $progressWindow.IsVisible) {
            Close-WpfProgressBar -Window $progressWindow
            $progressWindow = $null
        }
        Write-Host "`n`n============================================" -ForegroundColor Green
        Write-Host "   Proceso de actualización completado" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "`nSe recomienda REINICIAR el equipo" -ForegroundColor Yellow
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Mostrando diálogo de reinicio"
        $result = [System.Windows.MessageBox]::Show(
            "El proceso de actualización se completó exitosamente.`n`n" +
            "Se recomienda REINICIAR el equipo para completar la actualización del sistema WMI.`n`n" +
            "¿Desea reiniciar ahora?",
            "Actualización completada",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Write-Host "`n`tReiniciando equipo en 10 segundos..." -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Usuario eligió reiniciar"
            Start-Sleep -Seconds 3
            shutdown /r /t 10 /c "Reinicio para completar actualización de sistema WMI"
        } else {
            Write-Host "`n`tRecuerde reiniciar el equipo más tarde." -ForegroundColor Yellow
            Write-DzDebug "`t[DEBUG]Start-SystemUpdate: Usuario canceló reinicio"
        }
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: FIN OK"
        return $true
    } catch {
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: EXCEPCIÓN: $($_.Exception.Message)" Red
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: ScriptStackTrace: $($_.ScriptStackTrace)" Red
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red        [System.Windows.MessageBox]::Show(
            "Error durante la actualización: $($_.Exception.Message)`n`n" +
            "Revise los logs y considere reiniciar manualmente el equipo.",
            "Error en actualización",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return $false
    } finally {
        Write-DzDebug "`t[DEBUG]Start-SystemUpdate: FINALLY (cerrando progressWindow)"
        if ($progressWindow -ne $null -and $progressWindow.IsVisible) {
            Close-WpfProgressBar -Window $progressWindow
        }
    }
}
function Get-SqlPortWithDebug {
    Write-DzDebug "`n[DEBUG] === INICIANDO BÚSQUEDA DE PUERTOS SQL ==="
    Write-DzDebug "`t[DEBUG] Fecha/Hora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    $ports = @()

    # Definir rutas de registro a buscar (64-bit y 32-bit)
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server",  # 64-bit
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server"  # 32-bit
    )

    # 1. Buscar por todas las instancias instaladas
    Write-DzDebug "`t[DEBUG] 1. Buscando instancias SQL instaladas..."

    foreach ($basePath in $registryPaths) {
        try {
            Write-DzDebug "`t[DEBUG]   Examinando ruta: $basePath"

            if (-not (Test-Path $basePath)) {
                Write-DzDebug "`t[DEBUG]   ✗ La ruta no existe"
                continue
            }

            # Método 1: Obtener de InstalledInstances
            Write-DzDebug "`t[DEBUG]   Método 1: Buscando en 'InstalledInstances'..."
            $installedInstances = Get-ItemProperty -Path $basePath -Name "InstalledInstances" -ErrorAction SilentlyContinue

            if ($installedInstances -and $installedInstances.InstalledInstances) {
                Write-DzDebug "`t[DEBUG]   ✓ Instancias encontradas: $($installedInstances.InstalledInstances -join ', ')"

                foreach ($instance in $installedInstances.InstalledInstances) {
                    # Si ya procesamos esta instancia, saltar
                    if ($ports.Instance -contains $instance) {
                        Write-DzDebug "`t[DEBUG]     Instancia '$instance' ya procesada, saltando..."
                        continue
                    }

                    Write-DzDebug "`t[DEBUG]     Procesando instancia: '$instance'"

                    # Construir rutas posibles para esta instancia
                    $possiblePaths = @(
                        "$basePath\$instance\MSSQLServer\SuperSocketNetLib\Tcp",
                        "$basePath\$instance\MSSQLServer\SuperSocketNetLib\Tcp\IPAll",
                        "$basePath\MSSQLServer\$instance\SuperSocketNetLib\Tcp"
                    )

                    $portFound = $false
                    foreach ($tcpPath in $possiblePaths) {
                        Write-DzDebug "`t[DEBUG]       Probando ruta: $tcpPath"

                        if (Test-Path $tcpPath) {
                            Write-DzDebug "`t[DEBUG]       ✓ Ruta existe"

                            # Buscar puerto TCP
                            $tcpPort = Get-ItemProperty -Path $tcpPath -Name "TcpPort" -ErrorAction SilentlyContinue
                            $tcpDynamicPorts = Get-ItemProperty -Path $tcpPath -Name "TcpDynamicPorts" -ErrorAction SilentlyContinue

                            if ($tcpPort -and $tcpPort.TcpPort) {
                                $portInfo = [PSCustomObject]@{
                                    Instance = $instance
                                    Port     = $tcpPort.TcpPort
                                    Path     = $tcpPath
                                    Type     = "Static"
                                }
                                $ports += $portInfo
                                Write-DzDebug "`t[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
                                $portFound = $true
                                break
                            } elseif ($tcpDynamicPorts -and $tcpDynamicPorts.TcpDynamicPorts) {
                                $portInfo = [PSCustomObject]@{
                                    Instance = $instance
                                    Port     = $tcpDynamicPorts.TcpDynamicPorts
                                    Path     = $tcpPath
                                    Type     = "Dynamic"
                                }
                                $ports += $portInfo
                                Write-DzDebug "`t[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                                $portFound = $true
                                break
                            } else {
                                Write-DzDebug "`t[DEBUG]       ✗ No se encontró puerto en esta ruta"
                            }
                        } else {
                            Write-DzDebug "`t[DEBUG]       ✗ Ruta no existe"
                        }
                    }

                    if (-not $portFound) {
                        Write-DzDebug "`t[DEBUG]     ✗ No se encontró puerto para la instancia '$instance'"
                    }
                }
            } else {
                Write-DzDebug "`t[DEBUG]   ✗ No se encontró la clave 'InstalledInstances' en esta ruta"
            }

            # Método 2: Explorar todas las carpetas bajo SQL Server
            Write-DzDebug "`t[DEBUG]   Método 2: Explorando todas las carpetas..."

            $allSqlEntries = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer } |
            Select-Object -ExpandProperty PSChildName

            if ($allSqlEntries) {
                Write-DzDebug "`t[DEBUG]   Carpetas encontradas: $($allSqlEntries -join ', ')"

                foreach ($entry in $allSqlEntries) {
                    # Filtrar nombres que parecen instancias
                    if ($entry -match "^MSSQL\d+" -or $entry -match "^SQL" -or $entry -match "NATIONALSOFT") {
                        # Si ya procesamos esta instancia, saltar
                        if ($ports.Instance -contains $entry) {
                            Write-DzDebug "`t[DEBUG]     Instancia '$entry' ya procesada, saltando..."
                            continue
                        }

                        Write-DzDebug "`t[DEBUG]     Analizando posible instancia: '$entry'"

                        $tcpPath = "$basePath\$entry\MSSQLServer\SuperSocketNetLib\Tcp"

                        if (Test-Path $tcpPath) {
                            Write-DzDebug "`t[DEBUG]       ✓ Ruta TCP encontrada: $tcpPath"

                            $tcpPort = Get-ItemProperty -Path $tcpPath -Name "TcpPort" -ErrorAction SilentlyContinue
                            $tcpDynamicPorts = Get-ItemProperty -Path $tcpPath -Name "TcpDynamicPorts" -ErrorAction SilentlyContinue

                            if ($tcpPort -and $tcpPort.TcpPort) {
                                $portInfo = [PSCustomObject]@{
                                    Instance = $entry
                                    Port     = $tcpPort.TcpPort
                                    Path     = $tcpPath
                                    Type     = "Static"
                                }
                                $ports += $portInfo
                                Write-DzDebug "`t[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
                            } elseif ($tcpDynamicPorts -and $tcpDynamicPorts.TcpDynamicPorts) {
                                $portInfo = [PSCustomObject]@{
                                    Instance = $entry
                                    Port     = $tcpDynamicPorts.TcpDynamicPorts
                                    Path     = $tcpPath
                                    Type     = "Dynamic"
                                }
                                $ports += $portInfo
                                Write-DzDebug "`t[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                            } else {
                                Write-DzDebug "`t[DEBUG]       ✗ No se encontró puerto en esta ruta"
                            }
                        }
                    }
                }
            }

        } catch {
            Write-DzDebug "`t[DEBUG]   ERROR en búsqueda: $($_.Exception.Message)"
            Write-DzDebug "`t[DEBUG]   StackTrace: $($_.ScriptStackTrace)"
        }
    }

    # Método 3: Buscar por servicios SQL Server (común para ambos)
    Write-DzDebug "`t[DEBUG]   Método 3: Buscando servicios SQL Server..."

    $sqlServices = Get-Service -Name "*SQL*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*SQL Server (*" }

    foreach ($service in $sqlServices) {
        Write-DzDebug "`t[DEBUG]     Servicio: $($service.DisplayName)"

        # Extraer nombre de instancia del servicio
        if ($service.DisplayName -match "SQL Server \((.+)\)") {
            $instanceName = $matches[1]

            # Si ya procesamos esta instancia, saltar
            if ($ports.Instance -contains $instanceName) {
                Write-DzDebug "`t[DEBUG]       Instancia '$instanceName' ya procesada, saltando..."
                continue
            }

            Write-DzDebug "`t[DEBUG]       Posible instancia: '$instanceName'"
            # Buscar en ambas rutas del registro
            foreach ($basePath in $registryPaths) {
                $tcpPath = "$basePath\MSSQLServer\$instanceName\SuperSocketNetLib\Tcp"

                if (Test-Path $tcpPath) {
                    $tcpPort = Get-ItemProperty -Path $tcpPath -Name "TcpPort" -ErrorAction SilentlyContinue
                    $tcpDynamicPorts = Get-ItemProperty -Path $tcpPath -Name "TcpDynamicPorts" -ErrorAction SilentlyContinue

                    if ($tcpPort -and $tcpPort.TcpPort) {
                        $portInfo = [PSCustomObject]@{
                            Instance = $instanceName
                            Port     = $tcpPort.TcpPort
                            Path     = $tcpPath
                            Type     = "Static"
                        }
                        $ports += $portInfo
                        Write-DzDebug "`t[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
                        break
                    } elseif ($tcpDynamicPorts -and $tcpDynamicPorts.TcpDynamicPorts) {
                        $portInfo = [PSCustomObject]@{
                            Instance = $instanceName
                            Port     = $tcpDynamicPorts.TcpDynamicPorts
                            Path     = $tcpPath
                            Type     = "Dynamic"
                        }
                        $ports += $portInfo
                        Write-DzDebug "`t[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                        break
                    }
                }
            }
        }
    }

    # Método 4: Buscar puertos en uso por SQL Server
    Write-DzDebug "`t[DEBUG]   Método 4: Buscando puertos en uso por sqlservr.exe..."

    $sqlProcesses = Get-Process -Name "sqlservr" -ErrorAction SilentlyContinue
    if ($sqlProcesses) {
        foreach ($process in $sqlProcesses) {
            Write-DzDebug "`t[DEBUG]     Proceso sqlservr.exe encontrado (PID: $($process.Id))"

            # Obtener puertos usando netstat
            $netstatOutput = netstat -ano | Select-String ":$($process.Id)\s"

            foreach ($line in $netstatOutput) {
                if ($line -match ":(\d+)\s.*$($process.Id)$") {
                    $port = $matches[1]
                    Write-DzDebug "`t[DEBUG]       Puerto en uso: $port"

                    # Si no tenemos esta instancia en la lista, agregarla
                    if (-not ($ports.Port -contains $port)) {
                        # Intentar obtener nombre de instancia del proceso
                        $processInfo = Get-WmiObject Win32_Process -Filter "ProcessId = $($process.Id)" | Select-Object CommandLine
                        if ($processInfo.CommandLine -match "-s(.+?)\s") {
                            $instanceFromCmd = $matches[1]
                        } else {
                            $instanceFromCmd = "Unknown"
                        }

                        $portInfo = [PSCustomObject]@{
                            Instance = $instanceFromCmd
                            Port     = $port
                            Path     = "From Process"
                            Type     = "In Use"
                        }
                        $ports += $portInfo
                    }
                }
            }
        }
    } else {
        Write-DzDebug "`t[DEBUG]     ✗ No se encontraron procesos sqlservr.exe"
    }

    # Resumen final
    Write-DzDebug "`t[DEBUG] `n=== RESUMEN DE BÚSQUEDA ==="
    Write-DzDebug "`t[DEBUG] Total de instancias con puerto encontradas: $($ports.Count)"

    if ($ports.Count -gt 0) {
        foreach ($port in $ports) {
            Write-DzDebug "`t[DEBUG]   - Instancia: $($port.Instance) | Puerto: $($port.Port) | Tipo: $($port.Type)"
            Write-DzDebug "`t[DEBUG]     Ruta: $($port.Path)"
        }
    } else {
        Write-DzDebug "`t[DEBUG]   ✗ No se encontraron puertos SQL configurados"

        # Sugerencias para debugging
        Write-DzDebug "`t[DEBUG] `n=== SUGERENCIAS ==="
        Write-DzDebug "`t[DEBUG] 1. Verifica si SQL Server está instalado"
        Write-DzDebug "`t[DEBUG] 2. Revisa el Configuration Manager de SQL Server"
        Write-DzDebug "`t[DEBUG] 3. Verifica si el servicio SQL Server está ejecutándose"
        Write-DzDebug "`t[DEBUG] 4. Consulta el log de errores de SQL Server"
    }
    Write-DzDebug "`t[DEBUG] === FIN DE BÚSQUEDA ===`n"
    if ($ports.Count -gt 0) {
        # Ordenar por instancia
        $ports = $ports | Sort-Object -Property Instance
        # Formatear cada puerto con una línea por instancia
        $formattedPorts = @()
        foreach ($port in $ports) {
            # Formatear nombre de instancia
            $instanceName = if ($port.Instance -eq "MSSQLSERVER") { "Default" } else { $port.Instance }
            # Crear un nuevo objeto con todas las propiedades necesarias
            $formattedPort = [PSCustomObject]@{
                Instance       = $port.Instance
                Port           = $port.Port
                Path           = $port.Path
                Type           = $port.Type
                FormattedText  = "$instanceName`: $($port.Port)"
                SingleLineText = "$instanceName`: $($port.Port) - $($port.Type)"
            }
            $formattedPorts += $formattedPort
        }
        $ports = $formattedPorts
    }
    return $ports
}
function Show-SqlPortsInfo {
    param([array]$sqlPorts)
    if ($sqlPorts.Count -eq 0) {
        Write-Host "=== RESUMEN DE BÚSQUEDA SQL ===" -ForegroundColor Yellow
        Write-Host "No se encontraron puertos SQL ni instalaciones de SQL Server" -ForegroundColor Red
        Write-Host "=== FIN DE BÚSQUEDA ===" -ForegroundColor Yellow
        return
    }
    Write-Host "`n=== RESUMEN DE BÚSQUEDA SQL ===" -ForegroundColor Cyan
    Write-Host "Total de instancias con puerto encontradas: $($sqlPorts.Count)" -ForegroundColor White
    Write-Host ""
    $sqlPorts | ForEach-Object {
        Write-Host "  - Instancia: $($_.Instance) | Puerto: $($_.Port) | Tipo: $($_.Type)" -ForegroundColor Green
        if ($_.Method -and $global:debugEnabled) {
            Write-Host "    Método de detección: $($_.Method)" -ForegroundColor DarkGray
        }
    }
    Write-Host "`n=== FIN DE BÚSQUEDA ===" -ForegroundColor Cyan
}
function Show-ConfirmDialog {
    <#
      .SYNOPSIS
      Confirmación Yes/No con WPF MessageBox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter(Mandatory)]
        [string]$Title
    )

    Add-Type -AssemblyName PresentationFramework | Out-Null
    $result = [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    return ($result -eq [System.Windows.MessageBoxResult]::Yes)
}

function Show-InfoDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter(Mandatory)]
        [string]$Title
    )
    Add-Type -AssemblyName PresentationFramework | Out-Null
    [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

function Show-WarnDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter(Mandatory)]
        [string]$Title
    )
    Add-Type -AssemblyName PresentationFramework | Out-Null
    [System.Windows.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    ) | Out-Null
}

function Test-7ZipInstalled {
    [CmdletBinding()]
    param()

    $paths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { return $true }
    }

    # Por si existe en PATH
    return [bool](Get-Command 7z -ErrorAction SilentlyContinue)
}

function Get-7ZipPath {
    [CmdletBinding()]
    param()

    $paths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    foreach ($p in $paths) {
        if (Test-Path $p) { return $p }
    }

    $cmd = Get-Command 7z -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    return $null
}

function Install-7ZipWithChoco {
    [CmdletBinding()]
    param()

    if (-not (Test-ChocolateyInstalled)) {
        Write-DzDebug "[Install-7ZipWithChoco] Chocolatey no está instalado"
        return $false
    }

    try {
        Write-Host "Instalando 7zip con Chocolatey..." -ForegroundColor Yellow
        choco install 7zip -y --no-progress | Out-Null

        # refrescar PATH para el proceso actual (a veces choco no refresca inmediatamente)
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
        [System.Environment]::GetEnvironmentVariable("Path", "User")

        Start-Sleep -Seconds 2
        return (Test-7ZipInstalled)
    } catch {
        Write-Host "Error instalando 7zip: $_" -ForegroundColor Red
        return $false
    }
}

function Download-FileWithProgressWpfStream {
    param(
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [string]$OutFile,
        [Parameter(Mandatory)] $Window,
        [Parameter()] [ScriptBlock]$OnStatus
    )

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

    $dir = Split-Path $OutFile -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $total = $null
    try {
        $head = Invoke-WebRequest -Uri $Url -Method Head -UseBasicParsing -ErrorAction Stop
        $cl = $head.Headers["Content-Length"]
        if ($cl) { $total = [int64]$cl }
    } catch {
        $total = $null
    }

    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "GET"
    $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    $req.Accept = "*/*"
    $req.AllowAutoRedirect = $true

    $resp = $null
    $inStream = $null
    $outStream = $null

    try {
        $resp = $req.GetResponse()
        if (-not $total) {
            try { $total = [int64]$resp.ContentLength } catch { $total = $null }
        }

        $inStream = $resp.GetResponseStream()
        $outStream = New-Object System.IO.FileStream($OutFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)

        $buffer = New-Object byte[] (1024 * 128)
        [int64]$readTotal = 0
        [int64]$lastUi = 0
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($Window -and $Window.ProgressBar) {
            try {
                $Window.Dispatcher.Invoke([Action] {
                        $Window.ProgressBar.IsIndeterminate = $false
                        $Window.ProgressBar.Value = 0
                    }, [System.Windows.Threading.DispatcherPriority]::Render) | Out-Null
            } catch {}
        }

        while (($read = $inStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $outStream.Write($buffer, 0, $read)
            $readTotal += $read

            if (($readTotal - $lastUi) -ge (512KB) -or $sw.ElapsedMilliseconds -ge 200) {
                $lastUi = $readTotal
                $sw.Restart()

                $percent = 0
                if ($total -and $total -gt 0) {
                    $percent = [int][Math]::Min(100, [Math]::Floor(($readTotal * 100.0) / $total))
                }

                $mb = [Math]::Round($readTotal / 1MB, 2)
                $totalMb = if ($total) { [Math]::Round($total / 1MB, 2) } else { $null }
                $msg = if ($totalMb) { "Descargando... $mb / $totalMb MB ($percent%)" } else { "Descargando... $mb MB" }

                if ($Window) {
                    try {
                        Update-WpfProgressBar -Window $Window -Percent $percent -Message $msg
                        $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render) | Out-Null
                    } catch {}
                }

                if ($OnStatus) { try { & $OnStatus $percent $msg } catch {} }
            }
        }

        if ($Window) {
            try {
                Update-WpfProgressBar -Window $Window -Percent 100 -Message "Descarga completada."
                $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render) | Out-Null
            } catch {}
        }

        return $true
    } finally {
        try { if ($outStream) { $outStream.Flush(); $outStream.Close() } } catch {}
        try { if ($inStream) { $inStream.Close() } } catch {}
        try { if ($resp) { $resp.Close() } } catch {}
    }
}


function Show-InstallerExtractorDialog {
    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] INICIO"

    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Extractor de instalador"
        Height="360" Width="640"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="GeneralButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonGeneralBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonGeneralForeground)"/>
        </Style>
        <Style x:Key="NationalSoftButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonNationalBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonNationalForeground)"/>
        </Style>
    </Window.Resources>
    <Border Background="$($theme.FormBackground)"
            CornerRadius="10"
            BorderBrush="$($theme.ButtonNationalBackground)"
            BorderThickness="2"
            Padding="0">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
        </Border.Effect>

        <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="36"/>   <!-- Header -->
            <RowDefinition Height="Auto"/> <!-- Instrucción -->
            <RowDefinition Height="Auto"/> <!-- Label instalador -->
            <RowDefinition Height="Auto"/> <!-- picker instalador -->
            <RowDefinition Height="Auto"/> <!-- info instalador -->
            <RowDefinition Height="Auto"/> <!-- label destino -->
            <RowDefinition Height="Auto"/> <!-- picker destino -->
            <RowDefinition Height="Auto"/> <!-- botones -->
        </Grid.RowDefinitions>
            <!-- Header custom (arrastrable) -->
            <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Extractor de instalador"
                           VerticalAlignment="Center"
                           FontWeight="SemiBold"/>

                <Button Name="btnClose"
                        Grid.Column="1"
                        Content="✕"
                        Width="34" Height="26"
                        Margin="8,0,0,0"
                        ToolTip="Cerrar"
                        Background="Transparent"
                        BorderBrush="Transparent"/>
            </Grid>

            <TextBlock Grid.Row="1"
                       Text="Seleccione el instalador y el destino de extracción."
                       FontWeight="SemiBold"
                       Margin="0,0,0,12"/>

            <TextBlock Grid.Row="2" Text="Instalador (.exe)" Margin="0,0,0,6"/>
            <Grid Grid.Row="3" Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnPickInstaller"
                        Content="📁"
                        Width="36" Height="32"
                        Margin="0,0,8,0"
                        ToolTip="Seleccionar instalador"
                        Style="{StaticResource GeneralButtonStyle}"/>
                <TextBox Name="txtInstallerPath"
                         Grid.Column="1"
                         Height="32"
                         IsReadOnly="True"
                         VerticalContentAlignment="Center"
                         Text=""/>
            </Grid>

            <StackPanel Grid.Row="4" Margin="0,0,0,12">
                <TextBlock Name="lblVersionInfo" Text="Versión: -"/>
                <TextBlock Name="lblLastWrite" Text="Última modificación: -" Margin="0,4,0,0"/>
            </StackPanel>

            <TextBlock Grid.Row="5" Text="Destino" Margin="0,0,0,6"/>
            <Grid Grid.Row="6" Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnPickDestination"
                        Content="📁"
                        Width="36" Height="32"
                        Margin="0,0,8,0"
                        ToolTip="Seleccionar destino"
                        Style="{StaticResource GeneralButtonStyle}"/>
                <TextBox Name="txtDestinationPath"
                         Grid.Column="1"
                         Height="32"
                         IsReadOnly="False"
                         VerticalContentAlignment="Center"
                         Text=""/>
            </Grid>

            <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnCancel" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0" IsCancel="True" Style="{StaticResource GeneralButtonStyle}"/>
                <Button Name="btnExtract" Content="Extraer" Width="110" Height="30" Style="{StaticResource NationalSoftButtonStyle}"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@

    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        Set-DzWpfThemeResources -Window $ui.Window -Theme $theme
    } catch {
        Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana del extractor." -Title "Error" -Buttons OK -Icon Error | Out-Null
        return
    }

    $window = $ui.Window
    $c = $ui.Controls
    if ($c.ContainsKey('btnClose') -and $c['btnClose']) {
        $c['btnClose'].Add_Click({ $window.Close() })
    }
    if ($c.ContainsKey('HeaderBar') -and $c['HeaderBar']) {
        $c['HeaderBar'].Add_MouseLeftButtonDown({
                if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
                    $window.DragMove()
                }
            })
    }
    try {
        if ($Global:window -is [System.Windows.Window]) {
            $window.Owner = $Global:window
        }
    } catch {
        Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] No se pudo asignar owner: $($_.Exception.Message)" Yellow
    }

    $installerPath = $null
    $defaultFolderName = $null
    $destinationManuallySet = $false

    $updateInstallerInfo = {
        param([string]$path)

        $c['lblVersionInfo'].Text = "Versión: -"
        $c['lblLastWrite'].Text = "Última modificación: -"
        Set-Variable -Name defaultFolderName -Value $null -Scope 1

        if ([string]::IsNullOrWhiteSpace($path)) { return }

        try {
            $file = Get-Item -Path $path -ErrorAction Stop
            $fileInfo = (Get-ItemProperty $file.FullName).VersionInfo
            $creationDate = $file.LastWriteTime
            $formattedDate = $creationDate.ToString("yyMMdd")

            $versionText = if ($fileInfo.FileVersion) { $fileInfo.FileVersion } else { "N/D" }
            $c['lblVersionInfo'].Text = "Versión: $versionText"
            $c['lblLastWrite'].Text = "Última modificación: $($creationDate.ToString('dd/MM/yyyy HH:mm')) ($formattedDate)"

            $computedDefault = if ($versionText -ne "N/D") {
                "$versionText`_$formattedDate"
            } else {
                $formattedDate
            }

            Set-Variable -Name defaultFolderName -Value $computedDefault -Scope 1

            if (-not $destinationManuallySet) {
                $c['txtDestinationPath'].Text = Join-Path "C:\Temp" $computedDefault
            }
        } catch {
            Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR leyendo info del instalador: $($_.Exception.Message)" Red
            Show-WpfMessageBox -Message "No se pudo leer la información del instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
        }
    }

    $c['btnPickInstaller'].Add_Click({
            try {
                $dialog = New-Object Microsoft.Win32.OpenFileDialog
                $dialog.Filter = "Instalador (*.exe)|*.exe"
                $dialog.Title = "Seleccionar instalador"
                $dialog.Multiselect = $false

                if ($dialog.ShowDialog() -ne $true) { return }

                $selectedPath = $dialog.FileName
                if ([System.IO.Path]::GetExtension($selectedPath).ToLowerInvariant() -ne ".exe") {
                    Show-WpfMessageBox -Message "El instalador debe ser un archivo .EXE." -Title "Formato inválido" -Buttons OK -Icon Warning | Out-Null
                    return
                }

                # IMPORTANTE: guardar en el scope del function
                Set-Variable -Name installerPath -Value $selectedPath -Scope 1

                $c['txtInstallerPath'].Text = $selectedPath
                & $updateInstallerInfo $selectedPath
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR seleccionando instalador: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "Error al seleccionar el instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnPickDestination'].Add_Click({
            try {
                $initialDir = "C:\Temp"
                if (-not [string]::IsNullOrWhiteSpace($c['txtDestinationPath'].Text)) {
                    $initialDir = Split-Path -Path $c['txtDestinationPath'].Text -Parent
                }

                $selectedFolder = Show-WpfFolderDialog -Description "Seleccionar destino de extracción" -InitialDirectory $initialDir
                if (-not $selectedFolder) { return }

                Set-Variable -Name destinationManuallySet -Value $true -Scope 1

                if ($defaultFolderName) {
                    $c['txtDestinationPath'].Text = Join-Path $selectedFolder $defaultFolderName
                } else {
                    $c['txtDestinationPath'].Text = $selectedFolder
                }
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR seleccionando destino: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "Error al seleccionar el destino." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnExtract'].Add_Click({
            try {
                if (-not $installerPath) {
                    Show-WpfMessageBox -Message "Seleccione un instalador primero." -Title "Falta instalador" -Buttons OK -Icon Warning | Out-Null
                    return
                }

                if ([System.IO.Path]::GetExtension($installerPath).ToLowerInvariant() -ne ".exe") {
                    Show-WpfMessageBox -Message "El instalador debe ser un archivo .EXE." -Title "Formato inválido" -Buttons OK -Icon Warning | Out-Null
                    return
                }

                $destinationPath = $c['txtDestinationPath'].Text.Trim()
                if ([string]::IsNullOrWhiteSpace($destinationPath)) {
                    Show-WpfMessageBox -Message "Seleccione un destino válido." -Title "Falta destino" -Buttons OK -Icon Warning | Out-Null
                    return
                }

                if (-not (Test-Path -Path $destinationPath)) {
                    New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                }

                $arguments = "/extract `"$destinationPath`""
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] Ejecutando: '$installerPath' $arguments"
                $proc = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop

                if ($proc.ExitCode -ne 0) {
                    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ExitCode: $($proc.ExitCode)" Yellow
                    Show-WpfMessageBox -Message "El instalador devolvió código de salida $($proc.ExitCode)." -Title "Atención" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                Show-WpfMessageBox -Message "Extracción completada en:`n$destinationPath" -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                # Abrir la carpeta destino en el explorador
                try {
                    if (Test-Path -Path $destinationPath) {
                        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$destinationPath`""
                    }
                } catch {
                    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] No se pudo abrir Explorer: $($_.Exception.Message)" Yellow
                }
                $window.Close()
            } catch {
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] ERROR extracción: $($_.Exception.Message)" Red
                Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] Stack: $($_.ScriptStackTrace)" Red
                Show-WpfMessageBox -Message "Error al extraer el instalador." -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnCancel'].Add_Click({
            $window.Close()
        })

    $window.ShowDialog() | Out-Null
    Write-DzDebug "`t[DEBUG][Show-InstallerExtractorDialog] FIN"
}
function Show-SQLselector {
    param(
        [array]$Managers,
        [array]$SSMSVersions
    )

    function Get-ManagerBits {
        param([string]$Path)
        if ($Path -match "\\SysWOW64\\") { return "32 bits" }
        return "64 bits"
    }

    function Get-ManagerVersion {
        param([string]$Path)
        if ($Path -match "SQLServerManager(\d+)\.msc") { return $matches[1] }
        return "?"
    }

    function New-SelectorItem {
        param(
            [string]$Path,
            [string]$Display,
            [string]$DisplayShort
        )

        [PSCustomObject]@{
            Path         = $Path
            Display      = $Display
            DisplayShort = $DisplayShort
        }
    }



    if ($Managers -and $Managers.Count -gt 0) {
        $items = @()
        $unique = $Managers | Where-Object { $_ } | Select-Object -Unique

        foreach ($m in $unique) {
            $ver = Get-ManagerVersion -Path $m
            $bits = Get-ManagerBits    -Path $m

            $display = "SQLServerManager$ver  |  $bits  |  $m"
            $displayShort = "SQLServerManager$ver  |  $bits"

            $items += (New-SelectorItem -Path $m -Display $display -DisplayShort $displayShort)
        }


        $selected = Show-WpfPathSelectionDialog `
            -Title  "Seleccionar Configuration Manager" `
            -Prompt "Seleccione la versión de SQL Server Configuration Manager a ejecutar:" `
            -Items  $items `
            -ExecuteButtonText "Abrir"

        if ($selected) {
            Write-DzDebug "`t[DEBUG][Show-SQLselector] Seleccionado: $($selected.Display)"
            Start-Process -FilePath $selected.Path
        }

        return
    }

    # 2) SSMS
    if ($SSMSVersions -and $SSMSVersions.Count -gt 0) {
        $items = @()
        $unique = $SSMSVersions | Where-Object { $_ } | Select-Object -Unique

        foreach ($p in $unique) {
            try {
                $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($p)
                $prod = if ($vi.ProductName) { $vi.ProductName } else { "SSMS" }
                $ver = if ($vi.FileVersion) { $vi.FileVersion } else { "" }

                $display = "$prod  |  $ver  |  $p"
                $displayShort = "$prod  |  $ver"

                $items += (New-SelectorItem -Path $p -Display $display -DisplayShort $displayShort)
            } catch {
                $items += (New-SelectorItem -Path $p -Display "SSMS  |  $p" -DisplayShort "SSMS")
            }
        }

        $selected = Show-WpfPathSelectionDialog `
            -Title  "Seleccionar SSMS" `
            -Prompt "Seleccione la versión de SQL Server Management Studio a ejecutar:" `
            -Items  $items `
            -ExecuteButtonText "Ejecutar"

        if ($selected) {
            Write-DzDebug "`t[DEBUG][Show-SQLselector] Seleccionado: $($selected.DisplayShort)"
            Start-Process -FilePath $selected.Path
        }
        return

    }

    Write-DzDebug "`t[DEBUG][Show-SQLselector] No se recibieron rutas para Managers ni para SSMS." Yellow
}

function Show-IPConfigDialog {

    Write-Host "`n`t- - - Comenzando el proceso - - -" -ForegroundColor Gray

    $theme = Get-DzUiTheme

    # Helper: validar IPv4
    function Test-IPv4 {
        param([string]$Ip)
        if ([string]::IsNullOrWhiteSpace($Ip)) { return $false }
        $Ip = $Ip.Trim()
        return [System.Net.IPAddress]::TryParse($Ip, [ref]([System.Net.IPAddress]$null)) -and ($Ip -match '^\d{1,3}(\.\d{1,3}){3}$')
    }

    # Helper: refrescar texto de IPs
    function Get-AdapterIpsText {
        param([string]$Alias)
        try {
            $ips = Get-NetIPAddress -InterfaceAlias $Alias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty IPAddress
            if ($ips) { return "IPs asignadas: " + ($ips -join ", ") }
        } catch { }
        return "IPs asignadas: -"
    }

    # XAML
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Asignación de IPs"
        Height="250" Width="560"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="$($theme.FormBackground)"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
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
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="12" Background="$($theme.FormBackground)">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0"
                   Text="Seleccione el adaptador de red:"
                   Margin="0,0,0,8"/>

        <ComboBox Name="cmbAdapters"
                  Grid.Row="1"
                  Height="28"
                  Margin="0,0,0,10"/>

        <TextBlock Name="lblIps"
                   Grid.Row="2"
                   Text="IPs asignadas: -"
                   TextWrapping="Wrap"
                   Margin="0,0,0,12"
                   MinHeight="36"/>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnAssignIp" Content="Asignar Nueva IP" Width="140" Height="32" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnDhcp"     Content="Cambiar a DHCP"  Width="140" Height="32" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
            <Button Name="btnClose"    Content="Cerrar"         Width="110" Height="32" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $ui = New-WpfWindow -Xaml $stringXaml -PassThru
    $window = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $window -Theme $theme

    # CenterOwner real
    try {
        if ($Global:window -is [System.Windows.Window]) {
            $window.Owner = $Global:window
            $window.WindowStartupLocation = "CenterOwner"
        } else {
            $window.WindowStartupLocation = "CenterScreen"
        }
    } catch {
        $window.WindowStartupLocation = "CenterScreen"
    }

    # Cargar adaptadores Up
    $adapters = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "Up" })
    $c['cmbAdapters'].Items.Clear()
    $c['cmbAdapters'].Items.Add("Selecciona 1 adaptador de red") | Out-Null
    foreach ($a in $adapters) { $c['cmbAdapters'].Items.Add($a.Name) | Out-Null }
    $c['cmbAdapters'].SelectedIndex = 0

    # Actualizar UI según selección
    $updateUi = {
        $sel = [string]$c['cmbAdapters'].SelectedItem
        $valid = ($sel -and $sel -ne "Selecciona 1 adaptador de red")

        $c['btnAssignIp'].IsEnabled = $valid
        $c['btnDhcp'].IsEnabled = $valid

        if ($valid) {
            $c['lblIps'].Text = (Get-AdapterIpsText -Alias $sel)
        } else {
            $c['lblIps'].Text = "IPs asignadas: -"
        }
    }

    $c['cmbAdapters'].Add_SelectionChanged({ & $updateUi })
    & $updateUi

    # Botón: Asignar nueva IP
    $c['btnAssignIp'].Add_Click({
            $alias = [string]$c['cmbAdapters'].SelectedItem
            if (-not $alias -or $alias -eq "Selecciona 1 adaptador de red") {
                Show-WpfMessageBox -Message "Por favor, selecciona un adaptador de red." -Title "Error" -Buttons OK -Icon Error | Out-Null
                return
            }

            # Config actual
            $current = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $current) {
                Show-WpfMessageBox -Message "No se pudo obtener la configuración IPv4 del adaptador." -Title "Error" -Buttons OK -Icon Error | Out-Null
                return
            }

            $prefixLen = $current.PrefixLength

            # Pedir IP
            $newIp = New-WpfInputDialog -Title "Nueva IP" -Prompt "Ingrese la nueva dirección IP IPv4:" -DefaultValue ""
            if ([string]::IsNullOrWhiteSpace($newIp)) { return }

            if (-not (Test-IPv4 -Ip $newIp)) {
                Show-WpfMessageBox -Message "La IP '$newIp' no es válida." -Title "Error" -Buttons OK -Icon Error | Out-Null
                return
            }

            # Evitar duplicado
            $exists = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -eq $newIp.Trim() }
            if ($exists) {
                Show-WpfMessageBox -Message "La IP $newIp ya está asignada a $alias." -Title "Error" -Buttons OK -Icon Error | Out-Null
                return
            }

            try {
                New-NetIPAddress -IPAddress $newIp.Trim() -PrefixLength $prefixLen -InterfaceAlias $alias -ErrorAction Stop | Out-Null
                Show-WpfMessageBox -Message "Se agregó la IP $newIp al adaptador $alias." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $c['lblIps'].Text = (Get-AdapterIpsText -Alias $alias)
            } catch {
                Show-WpfMessageBox -Message "Error al agregar IP:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    # Botón: Cambiar a DHCP
    $c['btnDhcp'].Add_Click({
            $alias = [string]$c['cmbAdapters'].SelectedItem
            if (-not $alias -or $alias -eq "Selecciona 1 adaptador de red") {
                Show-WpfMessageBox -Message "Por favor, selecciona un adaptador de red." -Title "Error" -Buttons OK -Icon Error | Out-Null
                return
            }

            # Ver si ya es DHCP
            $any = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($any -and $any.PrefixOrigin -eq "Dhcp") {
                Show-WpfMessageBox -Message "El adaptador ya está en DHCP." -Title "Información" -Buttons OK -Icon Information | Out-Null
                return
            }

            $conf = Show-WpfMessageBox -Message "¿Está seguro de que desea cambiar a DHCP?" -Title "Confirmación" -Buttons YesNo -Icon Question
            if ($conf -ne [System.Windows.MessageBoxResult]::Yes) { return }

            try {
                # Quitar IPs manuales
                $manualIps = Get-NetIPAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.PrefixOrigin -eq "Manual" }
                foreach ($ip in $manualIps) {
                    Remove-NetIPAddress -IPAddress $ip.IPAddress -PrefixLength $ip.PrefixLength -InterfaceAlias $alias -Confirm:$false -ErrorAction SilentlyContinue
                }

                # Habilitar DHCP y reset DNS
                Set-NetIPInterface -InterfaceAlias $alias -Dhcp Enabled -ErrorAction Stop | Out-Null
                Set-DnsClientServerAddress -InterfaceAlias $alias -ResetServerAddresses -ErrorAction SilentlyContinue | Out-Null

                Show-WpfMessageBox -Message "Se cambió a DHCP en el adaptador $alias." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $c['lblIps'].Text = "Generando IP por DHCP. Seleccione de nuevo."
            } catch {
                Show-WpfMessageBox -Message "Error al cambiar a DHCP:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnClose'].Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}
# --- Helpers privados (no exportar) ---

function Get-OtmConfigFromSyscfg {
    param(
        [Parameter(Mandatory)][string]$SyscfgPath
    )

    # OJO: Tu código usa Get-Content (texto). Si el archivo fuera binario, esto no sería ideal.
    # Mantengo tu lógica tal cual para no cambiar comportamiento.
    $fileContent = Get-Content -LiteralPath $SyscfgPath -ErrorAction Stop

    $isSQL = ($fileContent -match "494E5354414C4C=1") -and ($fileContent -match "56455253495354454D41=3")
    $isDBF = ($fileContent -match "494E5354414C4C=2") -and ($fileContent -match "56455253495354454D41=2")

    if ($isSQL) { return "SQL" }
    if ($isDBF) { return "DBF" }
    return $null
}

function Get-OtmIniFiles {
    param(
        [Parameter(Mandatory)][string]$IniPath
    )

    $iniFiles = Get-ChildItem -LiteralPath $IniPath -Filter "*.ini" -ErrorAction Stop
    if (-not $iniFiles -or $iniFiles.Count -eq 0) { return $null }

    $iniSQLFile = $null
    $iniDBFFile = $null

    foreach ($iniFile in $iniFiles) {
        $content = Get-Content -LiteralPath $iniFile.FullName -ErrorAction SilentlyContinue

        if (-not $iniDBFFile -and ($content -match "Provider=VFPOLEDB\.1")) { $iniDBFFile = $iniFile }
        if (-not $iniSQLFile -and ($content -match "Provider=SQLOLEDB\.1")) { $iniSQLFile = $iniFile }

        if ($iniSQLFile -and $iniDBFFile) { break }
    }

    if (-not $iniSQLFile -or -not $iniDBFFile) { return $null }

    return [pscustomobject]@{
        SQL = $iniSQLFile
        DBF = $iniDBFFile
        All = $iniFiles
    }
}

function Set-OtmSyscfgConfig {
    param(
        [Parameter(Mandatory)][string]$SyscfgPath,
        [Parameter(Mandatory)][ValidateSet("SQL", "DBF")][string]$Target
    )

    if ($Target -eq "SQL") {
        Write-DzDebug "`t[DEBUG][OTM] Cambiando Syscfg a SQL" ([System.ConsoleColor]::Yellow)

        (Get-Content -LiteralPath $SyscfgPath) `
            -replace "494E5354414C4C=2", "494E5354414C4C=1" `
            -replace "56455253495354454D41=2", "56455253495354454D41=3" |
        Set-Content -LiteralPath $SyscfgPath
    } else {
        Write-DzDebug "`t[DEBUG][OTM] Cambiando Syscfg a DBF" ([System.ConsoleColor]::Yellow)

        (Get-Content -LiteralPath $SyscfgPath) `
            -replace "494E5354414C4C=1", "494E5354414C4C=2" `
            -replace "56455253495354454D41=3", "56455253495354454D41=2" |
        Set-Content -LiteralPath $SyscfgPath
    }
}

function Rename-OtmIniForTarget {
    param(
        [Parameter(Mandatory)][ValidateSet("SQL", "DBF")][string]$Target,
        [Parameter(Mandatory)]$IniSqlFile,
        [Parameter(Mandatory)]$IniDbfFile
    )

    # Para evitar conflicto si "checadorsql.ini" ya existe, renombramos con cuidado:
    $iniDir = Split-Path -Parent $IniSqlFile.FullName
    $finalIni = Join-Path $iniDir "checadorsql.ini"

    if ($Target -eq "SQL") {
        # DBF -> backup
        Rename-Item -LiteralPath $IniDbfFile.FullName -NewName "checadorsql_DBF_old.ini" -ErrorAction Stop
        # SQL -> activo
        Rename-Item -LiteralPath $IniSqlFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
    } else {
        # SQL -> backup
        Rename-Item -LiteralPath $IniSqlFile.FullName -NewName "checadorsql_SQL_old.ini" -ErrorAction Stop
        # DBF -> activo
        Rename-Item -LiteralPath $IniDbfFile.FullName -NewName "checadorsql.ini" -ErrorAction Stop
    }
}

# --- Función pública (exportar) ---
function Invoke-CambiarOTMConfig {
    [CmdletBinding()]
    param(
        [string]$SyscfgPath = "C:\Windows\SysWOW64\Syscfg45_2.0.dll",
        [string]$IniPath = "C:\NationalSoft\OnTheMinute4.5",
        [System.Windows.Controls.TextBlock]$InfoTextBlock
    )

    Write-DzDebug "`t[DEBUG][Invoke-CambiarOTMConfig] INICIO" ([System.ConsoleColor]::DarkGray)

    try {
        if (-not (Test-Path -LiteralPath $SyscfgPath)) {
            Show-WpfMessageBox -Message "El archivo de configuración no existe:`n$SyscfgPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] Syscfg no existe: $SyscfgPath" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe Syscfg." }
            return $false
        }

        if (-not (Test-Path -LiteralPath $IniPath)) {
            Show-WpfMessageBox -Message "No existe la ruta de INIs:`n$IniPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] IniPath no existe: $IniPath" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe ruta INI." }
            return $false
        }

        $current = Get-OtmConfigFromSyscfg -SyscfgPath $SyscfgPath
        if (-not $current) {
            Show-WpfMessageBox -Message "No se detectó una configuración válida de SQL o DBF en:`n$SyscfgPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] No se detectó config válida" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: config inválida." }
            return $false
        }

        $ini = Get-OtmIniFiles -IniPath $IniPath
        if (-not $ini) {
            Show-WpfMessageBox -Message "No se encontraron los archivos INI esperados en:`n$IniPath" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            Write-DzDebug "`t[DEBUG][OTM] No se encontraron INIs esperados" ([System.ConsoleColor]::Red)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: faltan INIs." }
            return $false
        }

        $new = if ($current -eq "SQL") { "DBF" } else { "SQL" }

        $msg = "Actualmente tienes configurado: $current.`n¿Quieres cambiar a $new?"
        $res = Show-WpfMessageBox -Message $msg -Title "Cambiar Configuración" -Buttons "YesNo" -Icon "Question"

        if ($res -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-DzDebug "`t[DEBUG][OTM] Usuario canceló" ([System.ConsoleColor]::Cyan)
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Operación cancelada." }
            return $false
        }

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Aplicando cambios..." }

        # 1) Cambiar syscfg
        Set-OtmSyscfgConfig -SyscfgPath $SyscfgPath -Target $new

        # 2) Renombrar INIs (activa el correcto)
        Rename-OtmIniForTarget -Target $new -IniSqlFile $ini.SQL -IniDbfFile $ini.DBF

        Show-WpfMessageBox -Message "Configuración cambiada exitosamente a $new." -Title "Éxito" -Buttons "OK" -Icon "Information" | Out-Null
        Write-DzDebug "`t[DEBUG][OTM] OK cambiado a $new" ([System.ConsoleColor]::Green)
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: cambiado a $new." }

        return $true

    } catch {
        $err = $_.Exception.Message
        Write-DzDebug "`t[DEBUG][Invoke-CambiarOTMConfig] ERROR: $err`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
        Show-WpfMessageBox -Message "Error: $err" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
        return $false
    }
}

function Invoke-CreateApk {
    [CmdletBinding()]
    param(
        [string]$DllPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\up.dll",
        [string]$InfoPath = "C:\Inetpub\wwwroot\ComanderoMovil\info\info.txt",
        [System.Windows.Controls.TextBlock]$InfoTextBlock
    )

    Write-DzDebug "`t[DEBUG][Invoke-CreateApk] INICIO" ([System.ConsoleColor]::DarkGray)

    try {
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Validando componentes..." }
        Write-Host "`nIniciando proceso de creación de APK..." -ForegroundColor Cyan

        if (-not (Test-Path -LiteralPath $DllPath)) {
            $msg = "Componente necesario no encontrado:`n$DllPath`n`nVerifique la instalación del Enlace Android."
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe up.dll" }
            return $false
        }

        if (-not (Test-Path -LiteralPath $InfoPath)) {
            $msg = "Archivo de configuración no encontrado:`n$InfoPath`n`nVerifique la instalación del Enlace Android."
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: no existe info.txt" }
            return $false
        }

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Leyendo versión..." }

        $jsonContent = Get-Content -LiteralPath $InfoPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $versionApp = [string]$jsonContent.versionApp

        if ([string]::IsNullOrWhiteSpace($versionApp)) {
            $msg = "No se pudo leer 'versionApp' desde:`n$InfoPath"
            Write-Host $msg -ForegroundColor Red
            Show-WpfMessageBox -Message $msg -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: versionApp vacío" }
            return $false
        }

        Write-Host "Versión detectada: $versionApp" -ForegroundColor Green

        $confirmMsg = "Se creará el APK versión: $versionApp`n¿Desea continuar?"
        $confirmation = Show-WpfMessageBox -Message $confirmMsg -Title "Confirmación" -Buttons "YesNo" -Icon "Question"

        if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Host "Proceso cancelado por el usuario" -ForegroundColor Yellow
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Cancelado." }
            return $false
        }

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Selecciona dónde guardar..." }

        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "Archivo APK (*.apk)|*.apk"
        $saveDialog.FileName = "SRM_$versionApp.apk"
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

        if ($saveDialog.ShowDialog() -ne $true) {
            Write-Host "Guardado cancelado por el usuario" -ForegroundColor Yellow
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Guardado cancelado." }
            return $false
        }

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Copiando APK..." }

        Copy-Item -LiteralPath $DllPath -Destination $saveDialog.FileName -Force -ErrorAction Stop

        $okMsg = "APK generado exitosamente en:`n$($saveDialog.FileName)"
        Write-Host $okMsg -ForegroundColor Green
        Show-WpfMessageBox -Message "APK creado correctamente!" -Title "Éxito" -Buttons "OK" -Icon "Information" | Out-Null
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: APK creado." }

        return $true

    } catch {
        $err = $_.Exception.Message
        Write-Host "Error durante el proceso: $err" -ForegroundColor Red
        Write-DzDebug "`t[DEBUG][Invoke-CreateApk] ERROR: $err`n$($_.ScriptStackTrace)" ([System.ConsoleColor]::Magenta)
        Show-WpfMessageBox -Message "Error durante la creación del APK:`n$err" -Title "Error" -Buttons "OK" -Icon "Error" | Out-Null
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }
        return $false
    }
}
function Get-NSIniConnectionInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FilePath
    )

    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }

    try {
        $content = Get-Content -LiteralPath $FilePath -ErrorAction Stop

        $dataSource = ($content | Select-String -Pattern '^DataSource=(.*)' -ErrorAction SilentlyContinue | Select-Object -First 1).Matches.Groups[1].Value
        $catalog = ($content | Select-String -Pattern '^Catalog=(.*)'    -ErrorAction SilentlyContinue | Select-Object -First 1).Matches.Groups[1].Value
        $authType = ($content | Select-String -Pattern '^autenticacion=(\d+)' -ErrorAction SilentlyContinue | Select-Object -First 1).Matches.Groups[1].Value

        $authUser = if ($authType -eq "2") { "sa" } elseif ($authType -eq "1") { "Windows" } else { "Desconocido" }

        return [pscustomobject]@{
            DataSource = $dataSource
            Catalog    = $catalog
            Usuario    = $authUser
        }
    } catch {
        Write-DzDebug "`t[DEBUG][Get-NSIniConnectionInfo] ERROR: $($_.Exception.Message)" ([System.ConsoleColor]::Yellow)
        return $null
    }
}

function Get-NSApplicationsIniReport {
    [CmdletBinding()]
    param(
        [hashtable[]]$PathsToCheck = @(
            @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
            @{ Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
            @{ Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
            @{ Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
            @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
            @{ Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
        ),
        [string]$RestCardPath = "C:\NationalSoft\Restcard\RestCard.ini"
    )

    $resultados = New-Object System.Collections.Generic.List[object]

    foreach ($entry in $PathsToCheck) {
        $basePath = $entry.Path
        $mainIni = Join-Path $basePath $entry.INI
        $appName = $entry.Nombre

        if (Test-Path -LiteralPath $mainIni) {
            $iniData = Get-NSIniConnectionInfo -FilePath $mainIni
            if ($iniData) {
                $resultados.Add([pscustomobject]@{
                        Aplicacion = $appName
                        INI        = $entry.INI
                        DataSource = $iniData.DataSource
                        Catalog    = $iniData.Catalog
                        Usuario    = $iniData.Usuario
                    })
            } else {
                $resultados.Add([pscustomobject]@{
                        Aplicacion = $appName
                        INI        = $entry.INI
                        DataSource = "NA"
                        Catalog    = "NA"
                        Usuario    = "NA"
                    })
            }
        } else {
            $resultados.Add([pscustomobject]@{
                    Aplicacion = $appName
                    INI        = "No encontrado"
                    DataSource = "NA"
                    Catalog    = "NA"
                    Usuario    = "NA"
                })
        }

        # Leer INIS\*.ini adicionales
        $inisFolder = Join-Path $basePath "INIS"
        if (Test-Path -LiteralPath $inisFolder) {
            $iniFiles = Get-ChildItem -LiteralPath $inisFolder -Filter "*.ini" -ErrorAction SilentlyContinue

            if ($appName -eq "OnTheMinute") {
                # Tu lógica: solo si hay más de 1
                if ($iniFiles -and $iniFiles.Count -gt 1) {
                    foreach ($iniFile in $iniFiles) {
                        $iniData = Get-NSIniConnectionInfo -FilePath $iniFile.FullName
                        if ($iniData) {
                            $resultados.Add([pscustomobject]@{
                                    Aplicacion = $appName
                                    INI        = $iniFile.Name
                                    DataSource = $iniData.DataSource
                                    Catalog    = $iniData.Catalog
                                    Usuario    = $iniData.Usuario
                                })
                        }
                    }
                }
            } else {
                foreach ($iniFile in $iniFiles) {
                    $iniData = Get-NSIniConnectionInfo -FilePath $iniFile.FullName
                    if ($iniData) {
                        $resultados.Add([pscustomobject]@{
                                Aplicacion = $appName
                                INI        = $iniFile.Name
                                DataSource = $iniData.DataSource
                                Catalog    = $iniData.Catalog
                                Usuario    = $iniData.Usuario
                            })
                    }
                }
            }
        }
    }

    # Restcard (tal cual tu regla)
    if (Test-Path -LiteralPath $RestCardPath) {
        $resultados.Add([pscustomobject]@{
                Aplicacion = "Restcard"
                INI        = "RestCard.ini"
                DataSource = "existe"
                Catalog    = "existe"
                Usuario    = "existe"
            })
    } else {
        $resultados.Add([pscustomobject]@{
                Aplicacion = "Restcard"
                INI        = "No encontrado"
                DataSource = "NA"
                Catalog    = "NA"
                Usuario    = "NA"
            })
    }

    return $resultados
}

function Show-NSApplicationsIniReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Resultados
    )

    $columnas = @("Aplicacion", "INI", "DataSource", "Catalog", "Usuario")
    $anchos = @{}

    foreach ($col in $columnas) { $anchos[$col] = $col.Length }

    foreach ($res in $Resultados) {
        foreach ($col in $columnas) {
            $val = [string]$res.$col
            if ($val.Length -gt $anchos[$col]) { $anchos[$col] = $val.Length }
        }
    }

    $titulos = $columnas | ForEach-Object { $_.PadRight($anchos[$_] + 2) }
    Write-Host ($titulos -join "") -ForegroundColor Cyan

    $separador = $columnas | ForEach-Object { ("-" * $anchos[$_]).PadRight($anchos[$_] + 2) }
    Write-Host ($separador -join "") -ForegroundColor Cyan

    foreach ($res in $Resultados) {
        $fila = $columnas | ForEach-Object { ([string]$res.$_).PadRight($anchos[$_] + 2) }

        if ($res.INI -eq "No encontrado") {
            Write-Host ($fila -join "") -ForegroundColor Red
        } else {
            Write-Host ($fila -join "")
        }
    }
}
function Get-NSPrinters {
    [CmdletBinding()]
    param()

    try {
        # Get-CimInstance es más moderno; mantiene compatibilidad con Win10/11 en PS5
        $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | ForEach-Object {
            $p = $_
            [PSCustomObject]@{
                Name       = $p.Name
                PortName   = $p.PortName
                DriverName = $p.DriverName
                Shared     = [bool]$p.Shared
            }
        }
        return $printers
    } catch {
        Write-DzDebug "`t[DEBUG][Get-NSPrinters] ERROR: $($_.Exception.Message)" ([System.ConsoleColor]::Yellow)
        return @()
    }
}

function Show-NSPrinters {
    [CmdletBinding()]
    param()

    Write-Host "`nImpresoras disponibles en el sistema:"

    $printers = Get-NSPrinters
    if (-not $printers -or $printers.Count -eq 0) {
        Write-Host "`nNo se encontraron impresoras."
        return
    }

    # Truncado como tu código original
    $view = $printers | ForEach-Object {
        [PSCustomObject]@{
            Name       = ([string]$_.Name).Substring(0, [Math]::Min(24, ([string]$_.Name).Length))
            PortName   = ([string]$_.PortName).Substring(0, [Math]::Min(19, ([string]$_.PortName).Length))
            DriverName = ([string]$_.DriverName).Substring(0, [Math]::Min(19, ([string]$_.DriverName).Length))
            IsShared   = if ($_.Shared) { "Sí" } else { "No" }
        }
    }

    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "Nombre", "Puerto", "Driver", "Compartida")
    Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f "------", "------", "------", "---------")
    $view | ForEach-Object {
        Write-Host ("{0,-25} {1,-20} {2,-20} {3,-10}" -f $_.Name, $_.PortName, $_.DriverName, $_.IsShared)
    }
}

function Invoke-ClearPrintJobs {
    [CmdletBinding()]
    param(
        [System.Windows.Controls.TextBlock]$InfoTextBlock
    )

    try {
        if (-not (Test-Administrator)) {
            Show-WpfMessageBox `
                -Message "Esta acción requiere permisos de administrador.`nPor favor, ejecuta Gerardo Zermeño Tools como administrador." `
                -Title "Permisos insuficientes" `
                -Buttons "OK" `
                -Icon "Warning" | Out-Null
            return $false
        }

        $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
        if (-not $spooler) {
            Show-WpfMessageBox `
                -Message "No se encontró el servicio 'Cola de impresión (Spooler)' en este equipo." `
                -Title "Servicio no encontrado" `
                -Buttons "OK" `
                -Icon "Error" | Out-Null
            return $false
        }

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Limpiando trabajos de impresión..." }

        # 1) Intento “limpio” con PrintManagement (si existe)
        try {
            Get-Printer -ErrorAction Stop | ForEach-Object {
                try {
                    Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
                } catch {
                    Write-Host "`tNo se pudieron limpiar trabajos de '$($_.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "`tNo se pudieron enumerar impresoras (Get-Printer). ¿Está instalado el módulo PrintManagement?" -ForegroundColor Yellow
        }

        # 2) Reinicio del Spooler
        if ($spooler.Status -eq 'Running') {
            Write-Host "`tDeteniendo servicio Spooler..." -ForegroundColor DarkYellow
            Stop-Service -Name Spooler -Force -ErrorAction Stop
        } else {
            Write-Host "`tSpooler no está en 'Running' (estado actual: $($spooler.Status))." -ForegroundColor DarkYellow
        }

        $spooler.Refresh()

        if ($spooler.StartType -eq 'Disabled') {
            Show-WpfMessageBox `
                -Message "El servicio 'Cola de impresión (Spooler)' está DESHABILITADO.`nHabilítalo manualmente desde services.msc para poder iniciarlo." `
                -Title "Spooler deshabilitado" `
                -Buttons "OK" `
                -Icon "Warning" | Out-Null
            if ($InfoTextBlock) { $InfoTextBlock.Text = "Spooler deshabilitado." }
            return $false
        }

        Write-Host "`tIniciando servicio Spooler..." -ForegroundColor DarkYellow
        Start-Service -Name Spooler -ErrorAction Stop

        if ($InfoTextBlock) { $InfoTextBlock.Text = "Listo: cola de impresión reiniciada." }

        Show-WpfMessageBox `
            -Message "Los trabajos de impresión han sido eliminados y el servicio de cola de impresión se reinició correctamente." `
            -Title "Operación exitosa" `
            -Buttons "OK" `
            -Icon "Information" | Out-Null

        return $true

    } catch {
        $err = $_.Exception.Message
        Write-Host "`n[ERROR Invoke-ClearPrintJobs] $err" -ForegroundColor Red
        if ($InfoTextBlock) { $InfoTextBlock.Text = "Error: $err" }

        Show-WpfMessageBox `
            -Message "Ocurrió un error al intentar limpiar impresoras o reiniciar el servicio:`n$err" `
            -Title "Error" `
            -Buttons "OK" `
            -Icon "Error" | Out-Null

        return $false
    }
}
function Show-WpfPathSelectionDialog {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][array] $Items,  # PSCustomObject { Path, Display }
        [string]$ExecuteButtonText = "Ejecutar"
    )

    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title"
        Height="420" Width="780"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="$($theme.ControlBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="$($theme.AccentPrimary)"/>
                    <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Border Background="$($theme.FormBackground)"
            CornerRadius="10"
            BorderBrush="$($theme.ButtonNationalBackground)"
            BorderThickness="2"
            Padding="0">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
        </Border.Effect>

        <Grid Margin="16">
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="250"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Header -->
            <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Name="txtHeader"
                           Text="$Title"
                           VerticalAlignment="Center"
                           FontWeight="SemiBold"/>

                <Button Name="btnClose"
                        Grid.Column="1"
                        Content="✕"
                        Width="34" Height="26"
                        Margin="8,0,0,0"
                        ToolTip="Cerrar"
                        Background="Transparent"
                        BorderBrush="Transparent"/>
            </Grid>

            <TextBlock Grid.Row="1"
                       Name="lblPrompt"
                       Text="$Prompt"
                       FontWeight="SemiBold"
                       Margin="0,0,0,10"/>

            <ListBox Grid.Row="2"
                     Name="lstItems"
                     DisplayMemberPath="Display"
                     SelectedValuePath="Path" />

            <!-- Footer: versión seleccionada + botones a la derecha -->
            <Grid Grid.Row="3" Margin="0,10,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0">
                    <TextBlock Text="Versión seleccionada:" Margin="0,0,0,2"/>
                    <TextBlock Name="lblSelectedDisplay"
                            Text=""
                            FontFamily="{DynamicResource CodeFontFamily}"
                            TextWrapping="NoWrap"
                            TextTrimming="CharacterEllipsis"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom">
                    <Button Name="btnCancel" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
                    <Button Name="btnExecute" Content="$ExecuteButtonText" Width="110" Height="30" Style="{StaticResource SystemButtonStyle}" IsDefault="True"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@


    $ui = New-WpfWindow -Xaml $stringXaml -PassThru
    $w = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme

    # Owner / centrado
    try { if (Get-Command Set-WpfDialogOwner -ErrorAction SilentlyContinue) { Set-WpfDialogOwner -Dialog $w } } catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }

    # Cerrar + Drag
    $c['btnClose'].Add_Click({ $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({
            if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() }
        })

    # Cargar items
    $c['lstItems'].ItemsSource = $Items
    if ($Items.Count -gt 0) { $c['lstItems'].SelectedIndex = 0 }

    $updateSelected = {
        $it = $c['lstItems'].SelectedItem
        if ($it) {
            $short = $null
            if ($it.PSObject.Properties.Match('DisplayShort').Count -gt 0) {
                $short = [string]$it.DisplayShort
            }
            $c['lblSelectedDisplay'].Text = if (-not [string]::IsNullOrWhiteSpace($short)) { $short } else { [string]$it.Display }
        } else {
            $c['lblSelectedDisplay'].Text = ""
        }
    }


    & $updateSelected
    $c['lstItems'].Add_SelectionChanged({ & $updateSelected })
    $script:_selectedItem = $null

    $c['btnExecute'].Add_Click({
            $it = $c['lstItems'].SelectedItem
            if (-not $it) { return }
            $script:_selectedItem = $it   # <- devolvemos el objeto completo
            $w.DialogResult = $true
            $w.Close()
        })

    $c['btnCancel'].Add_Click({
            $w.DialogResult = $false
            $w.Close()
        })

    $ok = $w.ShowDialog()
    if ($ok) { return $script:_selectedItem }
    return $null
}
function Show-LZMADialog {
    param(
        [array]$Instaladores
    )

    Write-DzDebug "`t[DEBUG][Show-LZMADialog] INICIO"
    $theme = Get-DzUiTheme

    if (-not $Instaladores -or $Instaladores.Count -eq 0) {
        $LZMAregistryPath = "HKLM:\SOFTWARE\WOW6432Node\Caphyon\Advanced Installer\LZMA"

        if (-not (Test-Path $LZMAregistryPath)) {
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] No existe la clave LZMA: $LZMAregistryPath" Yellow
            Show-WpfMessageBox -Message "No se encontró Advanced Installer (LZMA) en este equipo.`n`nRuta no existe:`n$LZMAregistryPath" `
                -Title "Sin instaladores" -Buttons OK -Icon Information | Out-Null
            return
        }

        try {
            $carpetasPrincipales = Get-ChildItem -Path $LZMAregistryPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }
            if (-not $carpetasPrincipales -or $carpetasPrincipales.Count -lt 1) {
                Write-DzDebug "`t[DEBUG][Show-LZMADialog] No se encontraron carpetas principales." Yellow
                Show-WpfMessageBox -Message "No se encontraron carpetas principales en la ruta del registro." `
                    -Title "Sin resultados" -Buttons OK -Icon Information | Out-Null
                return
            }

            $tmp = @()
            foreach ($carpeta in $carpetasPrincipales) {
                $subdirs = Get-ChildItem -Path $carpeta.PSPath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
                foreach ($sd in $subdirs) {
                    $tmp += [PSCustomObject]@{
                        Name = $sd.PSChildName
                        Path = $sd.PSPath
                    }
                }
            }

            if (-not $tmp -or $tmp.Count -lt 1) {
                Write-DzDebug "`t[DEBUG][Show-LZMADialog] No se encontraron subcarpetas." Yellow
                Show-WpfMessageBox -Message "No se encontraron instaladores (subcarpetas) en la ruta del registro." `
                    -Title "Sin resultados" -Buttons OK -Icon Information | Out-Null
                return
            }

            $Instaladores = $tmp | Sort-Object Name -Descending
        } catch {
            Write-DzDebug "`t[DEBUG][Show-LZMADialog] Error accediendo al registro: $($_.Exception.Message)" Red
            Show-WpfMessageBox -Message "Error accediendo al registro:`n$($_.Exception.Message)" `
                -Title "Error" -Buttons OK -Icon Error | Out-Null
            return
        }
    }

    # Construir items: lista muestra nombre + ruta, pero guardamos Name/Path separadas
    $items = foreach ($i in $Instaladores) {
        if (-not $i) { continue }
        [PSCustomObject]@{
            Name    = [string]$i.Name
            Path    = [string]$i.Path
            Display = ("{0}  |  {1}" -f $i.Name, $i.Path) # LISTA con ruta
        }
    }

    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Carpetas LZMA"
        Height="290" Width="760"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
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
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
            <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
        </Style>
    </Window.Resources>
    <Border Background="$($theme.FormBackground)"
            CornerRadius="10"
            BorderBrush="$($theme.ButtonNationalBackground)"
            BorderThickness="2"
            Padding="0">
        <Border.Effect>
            <DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/>
        </Border.Effect>

        <Grid Margin="16">
            <Grid.RowDefinitions>
                <RowDefinition Height="36"/>   <!-- Header -->
                <RowDefinition Height="Auto"/> <!-- Instrucción -->
                <RowDefinition Height="Auto"/> <!-- Combo -->
                <RowDefinition Height="*"/>    <!-- AI_ExePath -->
                <RowDefinition Height="Auto"/> <!-- Botones -->
            </Grid.RowDefinitions>

            <!-- Header -->
            <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Carpetas LZMA"
                           VerticalAlignment="Center"
                           FontWeight="SemiBold"/>

                <Button Name="btnClose"
                        Grid.Column="1"
                        Content="✕"
                        Width="34" Height="26"
                        Margin="8,0,0,0"
                        ToolTip="Cerrar"
                        Background="Transparent"
                        BorderBrush="Transparent"/>
            </Grid>

            <TextBlock Grid.Row="1"
                       Text="Seleccione el instalador (registro) que desea renombrar."
                       FontWeight="SemiBold"
                       Margin="0,0,0,10"/>

            <ComboBox Name="cmbInstallers"
                      Grid.Row="2"
                      Height="30"
                      Margin="0,0,0,10"
                      DisplayMemberPath="Display"
                      SelectedValuePath="Path"/>

            <!-- AI_ExePath -->
            <Border Grid.Row="3"
                    Background="$($theme.ControlBackground)"
                    CornerRadius="8"
                    Padding="10"
                    Margin="0,0,0,8"
                    MinHeight="78">
                <StackPanel>
                    <TextBlock Text="AI_ExePath:"
                               Margin="0,0,0,4"/>
                    <TextBlock Name="lblExePath"
                               Text="-"
                               Foreground="#B00020"
                               TextWrapping="Wrap"
                               TextTrimming="None"
                               MaxHeight="48"/>
                </StackPanel>
            </Border>

            <!-- Botones -->
            <StackPanel Grid.Row="4"
                        Orientation="Horizontal"
                        HorizontalAlignment="Right"
                        Margin="0,0,0,0">
                <Button Name="btnRename" Content="Renombrar" Width="110" Height="30" Margin="0,0,10,0" IsEnabled="False" Style="{StaticResource SystemButtonStyle}"/>
                <Button Name="btnExit" Content="Salir" Width="110" Height="30" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
            </StackPanel>

        </Grid>
    </Border>
</Window>
"@

    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
    } catch {
        Write-DzDebug "`t[DEBUG][Show-LZMADialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana LZMA." -Title "Error" -Buttons OK -Icon Error | Out-Null
        return
    }

    $w = $ui.Window
    $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme

    # Owner / centrado
    try { Set-WpfDialogOwner -Dialog $w } catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }

    # Header drag + close
    $c['btnClose'].Add_Click({ $w.DialogResult = $false; $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({
            if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() }
        })

    # Poblar combo
    $placeholder = [PSCustomObject]@{ Name = ""; Path = ""; Display = "Selecciona instalador a renombrar" }
    $c['cmbInstallers'].ItemsSource = @($placeholder) + @($items)
    $c['cmbInstallers'].SelectedIndex = 0

    $updateUi = {
        $idx = $c['cmbInstallers'].SelectedIndex
        $c['btnRename'].IsEnabled = ($idx -gt 0)

        if ($idx -le 0) {
            $c['lblExePath'].Text = "-"
            return
        }

        $it = $c['cmbInstallers'].SelectedItem
        if (-not $it -or [string]::IsNullOrWhiteSpace($it.Path)) {
            $c['lblExePath'].Text = "No encontrado"
            return
        }

        try {
            $prop = Get-ItemProperty -Path $it.Path -Name "AI_ExePath" -ErrorAction SilentlyContinue
            if ($prop -and $prop.AI_ExePath) {
                $pathTxt = [string]$prop.AI_ExePath
                $pathTxt = $pathTxt -replace "\\", "\ "
                $c['lblExePath'].Text = $pathTxt

            } else {
                $c['lblExePath'].Text = "No encontrado"
            }
        } catch {
            $c['lblExePath'].Text = "Error leyendo AI_ExePath"
        }
    }

    $c['cmbInstallers'].Add_SelectionChanged({ & $updateUi })
    & $updateUi

    $c['btnRename'].Add_Click({
            $idx = $c['cmbInstallers'].SelectedIndex
            if ($idx -le 0) { return }

            $it = $c['cmbInstallers'].SelectedItem
            if (-not $it -or [string]::IsNullOrWhiteSpace($it.Path)) { return }

            $rutaVieja = [string]$it.Path
            $nombre = [string]$it.Name
            $nuevoNombre = "$nombre.backup"

            $msg = "¿Está seguro de renombrar el registro?`n`n$rutaVieja`n`nA:`n$nuevoNombre"
            $conf = Show-WpfMessageBox -Message $msg -Title "Confirmar renombrado" -Buttons YesNo -Icon Warning

            if ($conf -ne [System.Windows.MessageBoxResult]::Yes) { return }

            try {
                Rename-Item -Path $rutaVieja -NewName $nuevoNombre -ErrorAction Stop
                Show-WpfMessageBox -Message "Registro renombrado correctamente." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $w.DialogResult = $true
                $w.Close()
            } catch {
                Show-WpfMessageBox -Message "Error al renombrar:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })

    $c['btnExit'].Add_Click({
            $w.DialogResult = $false
            $w.Close()
        })

    $w.ShowDialog() | Out-Null
    Write-DzDebug "`t[DEBUG][Show-LZMADialog] FIN"
}

function Show-AddUserDialog {
    Write-DzDebug "`t[DEBUG][Show-AddUserDialog] INICIO"
    $theme = Get-DzUiTheme
    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Crear Usuario de Windows"
        Height="420" Width="640"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        FontFamily="{DynamicResource UiFontFamily}"
        FontSize="{DynamicResource UiFontSize}">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="$($theme.ControlBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="PasswordBox">
      <Setter Property="Background" Value="$($theme.ControlBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
    </Style>
    <Style TargetType="ToggleButton">
      <Setter Property="Background" Value="$($theme.ControlBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ControlForeground)"/>
      <Setter Property="BorderBrush" Value="$($theme.BorderColor)"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Style.Triggers>
        <Trigger Property="IsChecked" Value="True">
          <Setter Property="Background" Value="$($theme.AccentPrimary)"/>
          <Setter Property="Foreground" Value="$($theme.FormForeground)"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="SystemButtonStyle" TargetType="Button">
      <Setter Property="Background" Value="$($theme.ButtonSystemBackground)"/>
      <Setter Property="Foreground" Value="$($theme.ButtonSystemForeground)"/>
    </Style>
  </Window.Resources>
  <Border Background="$($theme.FormBackground)" CornerRadius="10" BorderBrush="$($theme.ButtonNationalBackground)" BorderThickness="2" Padding="0">
    <Border.Effect><DropShadowEffect Color="Black" Direction="270" ShadowDepth="4" BlurRadius="12" Opacity="0.25"/></Border.Effect>
    <Grid Margin="16">
      <Grid.RowDefinitions>
        <RowDefinition Height="36"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <Grid Grid.Row="0" Name="HeaderBar" Background="Transparent">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Text="Crear Usuario de Windows" VerticalAlignment="Center" FontWeight="SemiBold"/>
        <Button Name="btnClose" Grid.Column="1" Content="✕" Width="34" Height="26" Margin="8,0,0,0" ToolTip="Cerrar" Background="Transparent" BorderBrush="Transparent"/>
      </Grid>
      <TextBlock Grid.Row="1" Text="Crea un usuario local y asígnalo al grupo correspondiente." FontWeight="SemiBold" Margin="0,0,0,12"/>
      <Grid Grid.Row="2" Margin="0,0,0,12">
        <Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Nombre de usuario" VerticalAlignment="Center" Margin="0,0,10,8"/>
        <TextBox Name="txtUsername" Grid.Row="0" Grid.Column="1" Height="32" VerticalContentAlignment="Center" Margin="0,0,0,8"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Contraseña" VerticalAlignment="Center" Margin="0,0,10,8"/>
        <Grid Grid.Row="1" Grid.Column="1" Margin="0,0,0,8">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <PasswordBox Name="pwdPassword" Grid.Column="0" Height="32" Padding="6,0,6,0"/>
          <TextBox Name="txtPasswordVisible" Grid.Column="0" Height="32" VerticalContentAlignment="Center" Visibility="Collapsed"/>
          <ToggleButton Name="tglShowPassword" Grid.Column="1" Content="👁" Width="40" Height="32" Margin="8,0,0,0" ToolTip="Mostrar/Ocultar contraseña"/>
        </Grid>
        <TextBlock Grid.Row="2" Grid.Column="0" Text="Tipo de usuario" VerticalAlignment="Center"/>
        <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <RadioButton Name="rbStandard" Content="Usuario estándar" IsChecked="True" Margin="0,0,12,0"/>
          <RadioButton Name="rbAdmin" Content="Administrador"/>
        </StackPanel>
        <Button Name="btnShowUsers" Grid.Row="2" Grid.Column="2" Content="Ver usuarios" Width="110" Height="30" Margin="12,0,0,0" Style="{StaticResource SystemButtonStyle}"/>
      </Grid>
      <Border Grid.Row="3" Background="$($theme.ControlBackground)" CornerRadius="8" Padding="12">
        <StackPanel>
          <TextBlock Text="Requisitos:" FontWeight="SemiBold" Margin="0,0,0,6"/>
          <TextBlock Text="• Nombre: sin espacios (ej. soporte01)" Margin="0,0,0,2"/>
          <TextBlock Text="• Contraseña: mínimo 8 caracteres" Margin="0,0,0,2"/>
          <TextBlock Text="• Administrador: úsalo solo si es necesario"/>
        </StackPanel>
      </Border>
      <Grid Grid.Row="4" Margin="0,12,0,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Name="lblStatus" Grid.Column="0" Text="Listo." Foreground="#2E7D32" VerticalAlignment="Center" TextWrapping="Wrap" Margin="0,0,10,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
          <Button Name="btnCancel" Content="Cancelar" Width="110" Height="30" Margin="0,0,10,0" IsCancel="True" Style="{StaticResource SystemButtonStyle}"/>
          <Button Name="btnCreate" Content="Crear usuario" Width="130" Height="30" Style="{StaticResource SystemButtonStyle}" IsEnabled="False" IsDefault="True"/>
        </StackPanel>
      </Grid>
    </Grid>
  </Border>
</Window>
"@
    try { $ui = New-WpfWindow -Xaml $stringXaml -PassThru }catch { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR creando ventana: $($_.Exception.Message)" Red; Show-WpfMessageBox -Message "No se pudo crear la ventana de usuario." -Title "Error" -Buttons OK -Icon Error | Out-Null; return }
    $w = $ui.Window; $c = $ui.Controls
    Set-DzWpfThemeResources -Window $w -Theme $theme
    try { Set-WpfDialogOwner -Dialog $w }catch {}
    if (-not $w.Owner) { $w.WindowStartupLocation = "CenterScreen" }
    $c['btnClose'].Add_Click({ $w.DialogResult = $false; $w.Close() })
    $c['HeaderBar'].Add_MouseLeftButtonDown({ if ($_.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) { $w.DragMove() } })
    try { $adminGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-544').Name; $userGroup = (Get-LocalGroup | Where-Object SID -EQ 'S-1-5-32-545').Name }catch { Show-WpfMessageBox -Message "No se pudieron obtener los grupos locales (requiere permisos).`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null; $w.Close(); return }
    function Set-Status { param([string]$Text, [string]$Level = "Ok"); switch ($Level) { "Ok" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::ForestGreen }"Warn" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::DarkGoldenrod }"Error" { $c['lblStatus'].Foreground = [System.Windows.Media.Brushes]::Firebrick } }; $c['lblStatus'].Text = $Text }
    function Get-PasswordText { if ($c['txtPasswordVisible'].Visibility -eq 'Visible') { return [string]$c['txtPasswordVisible'].Text }; return [string]$c['pwdPassword'].Password }
    function Validate-Form {
        $username = ([string]$c['txtUsername'].Text).Trim()
        $pass = (Get-PasswordText).Trim()
        Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Validate username='$username' passLen=$($pass.Length)"
        if ([string]::IsNullOrWhiteSpace($username)) { Set-Status "Escriba un nombre de usuario." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ($username -match "\s") { Set-Status "El nombre no debe contener espacios." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ([string]::IsNullOrWhiteSpace($pass)) { Set-Status "Escriba una contraseña." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        if ($pass.Length -lt 8) { Set-Status "La contraseña debe tener al menos 8 caracteres." "Warn"; $c['btnCreate'].IsEnabled = $false; return }
        try {
            $exists = Get-LocalUser -Name $username -ErrorAction SilentlyContinue
            if ($exists) { Set-Status "El usuario '$username' ya existe." "Error"; $c['btnCreate'].IsEnabled = $false; return }
        } catch {
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Validate existencia falló: $($_.Exception.Message)" Yellow
            Set-Status "Aviso: no se pudo validar si el usuario ya existe (permisos)." "Warn"
        }
        Set-Status "Listo para crear usuario." "Ok"
        $c['btnCreate'].IsEnabled = $true
    }
    $c['tglShowPassword'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowPassword ON"; $c['txtPasswordVisible'].Text = [string]$c['pwdPassword'].Password; $c['pwdPassword'].Visibility = 'Collapsed'; $c['txtPasswordVisible'].Visibility = 'Visible'; Validate-Form })
    $c['tglShowPassword'].Add_Unchecked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowPassword OFF"; $c['pwdPassword'].Password = [string]$c['txtPasswordVisible'].Text; $c['txtPasswordVisible'].Visibility = 'Collapsed'; $c['pwdPassword'].Visibility = 'Visible'; Validate-Form })
    $c['txtUsername'].Add_TextChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] txtUsername changed"; Validate-Form })
    $c['pwdPassword'].Add_PasswordChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] pwdPassword changed"; Validate-Form })
    $c['txtPasswordVisible'].Add_TextChanged({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] txtPasswordVisible changed"; Validate-Form })
    $c['rbStandard'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Tipo=Standard"; Validate-Form })
    $c['rbAdmin'].Add_Checked({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Tipo=Admin"; Validate-Form })
    $c['btnShowUsers'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnShowUsers click"
            try {
                $users = Get-LocalUser | Select-Object Name, Enabled
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Usuarios obtenidos: $($users.Count)"
                $lines = foreach ($u in $users) {
                    $estado = if ($u.Enabled) { "Habilitado" }else { "Deshabilitado" }
                    "{0,-20}  {1}" -f $u.Name, $estado
                }
                $msg = $lines -join "`n"
                if ([string]::IsNullOrWhiteSpace($msg)) { $msg = "(Sin resultados)" }
                Show-WpfMessageBox -Message $msg -Title "Usuarios locales" -Buttons OK -Icon Information | Out-Null
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnShowUsers mostrado OK"
            } catch {
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR btnShowUsers: $($_.Exception.Message)" Red
                Show-WpfMessageBox -Message "No se pudieron listar usuarios:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnCreate'].Add_Click({
            Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnCreate click"
            $username = ([string]$c['txtUsername'].Text).Trim()
            $password = Get-PasswordText
            $isAdmin = $false; try { $isAdmin = [bool]$c['rbAdmin'].IsChecked }catch {}
            $tipo = if ($isAdmin) { "Administrador" }else { "Usuario estándar" }
            $group = if ($isAdmin) { $adminGroup }else { $userGroup }
            $confirmMsg = "Se creará el usuario:`n`n$username`n`nTipo: $tipo`nGrupo: $group"
            $conf = Show-WpfMessageBox -Message $confirmMsg -Title "Confirmar" -Buttons YesNo -Icon Question
            if ($conf -ne [System.Windows.MessageBoxResult]::Yes) { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Creación cancelada"; return }
            try {
                if (Get-LocalUser -Name $username -ErrorAction SilentlyContinue) { Set-Status "El usuario '$username' ya existe." "Error"; Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Ya existe: $username" Yellow; return }
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                New-LocalUser -Name $username -Password $securePassword -AccountNeverExpires -PasswordNeverExpires | Out-Null
                Add-LocalGroupMember -Group $group -Member $username
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] Usuario creado: $username Grupo: $group"
                Show-WpfMessageBox -Message "Usuario '$username' creado y agregado al grupo '$group'." -Title "Éxito" -Buttons OK -Icon Information | Out-Null
                $w.DialogResult = $true; $w.Close()
            } catch {
                Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR creando usuario: $($_.Exception.Message)" Red
                Set-Status "Error: $($_.Exception.Message)" "Error"
                Show-WpfMessageBox -Message "Error al crear usuario:`n$($_.Exception.Message)" -Title "Error" -Buttons OK -Icon Error | Out-Null
            }
        })
    $c['btnCancel'].Add_Click({ Write-DzDebug "`t[DEBUG][Show-AddUserDialog] btnCancel"; $w.DialogResult = $false; $w.Close() })
    Validate-Form
    try { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ShowDialog()"; $w.ShowDialog() | Out-Null }catch { Write-DzDebug "`t[DEBUG][Show-AddUserDialog] ERROR ShowDialog: $($_.Exception.Message)" Red; throw }
    Write-DzDebug "`t[DEBUG][Show-AddUserDialog] FIN"
}

Export-ModuleMember -Function @(
    'Get-DzToolsConfigPath',
    'Get-DzDebugPreference',
    'Get-DzUiMode',
    'Set-DzUiMode',
    'Set-DzDebugPreference',
    'Initialize-DzToolsConfig',
    'Write-DzDebug',
    'Test-Administrator',
    'Get-SystemInfo',
    'Clear-TemporaryFiles',
    'Test-ChocolateyInstalled',
    'Install-Chocolatey',
    'Get-AdminGroupName',
    'Invoke-DiskCleanup',
    'Stop-CleanmgrProcesses',
    'Show-SystemComponents',
    'Test-SameHost',
    'Test-7ZipInstalled',
    'Test-MegaToolsInstalled',
    'Check-Permissions',
    'Download-FileWithProgressWpfStream',
    'Refresh-AdapterStatus',
    'Get-NetworkAdapterStatus',
    'Start-SystemUpdate',
    'Get-SqlPortWithDebug',
    'Show-SqlPortsInfo',
    'Show-ConfirmDialog',
    'Show-InfoDialog',
    'Show-WarnDialog',
    'Get-7ZipPath',
    'Install-7ZipWithChoco',
    'Show-LZMADialog',
    'Show-InstallerExtractorDialog',
    'Show-SQLselector',
    'Show-IPConfigDialog',
    'Invoke-CambiarOTMConfig',
    'Invoke-CreateApk',
    'get-NSApplicationsIniReport',
    'Show-NSApplicationsIniReport',
    'show-NSPrinters',
    'Invoke-ClearPrintJobs',
    'Show-AddUserDialog'
)
