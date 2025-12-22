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
        "[desarrollo]`ndebug=false" | Out-File -FilePath $configPath -Encoding UTF8 -Force
    }
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
            $proc = Start-Process $cleanmgr `
                -ArgumentList "/sagerun:$profileId" `
                -WindowStyle Hidden `
                -PassThru
            Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Proceso iniciado. PID=$($proc.Id)"
            $timeoutSeconds = $TimeoutMinutes * 60
            $script:remainingSeconds = $timeoutSeconds
            $script:cleanupCompleted = $false
            if ($ProgressWindow -ne $null -and $ProgressWindow.IsVisible) {
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(1)
                $timer.Add_Tick({
                        if ($script:remainingSeconds -gt 0) {
                            $script:remainingSeconds--
                            $mins = [math]::Floor($script:remainingSeconds / 60)
                            $secs = $script:remainingSeconds % 60
                            if ($ProgressWindow.PSObject.Properties.Name -contains 'MessageLabel') {
                                $ProgressWindow.MessageLabel.Text = "Liberando espacio en disco...`nTiempo restante: $mins min $secs seg"
                            }
                        } else {
                            $this.Stop()
                        }
                    }.GetNewClosure())
                $timer.Start()
                Write-DzDebug "`t[DEBUG]Invoke-DiskCleanup: Timer iniciado"
            }
            $checkInterval = 500
            $elapsed = 0
            while (-not $proc.HasExited -and $elapsed -lt ($timeoutSeconds * 1000)) {
                Start-Sleep -Milliseconds $checkInterval
                $elapsed += $checkInterval
                if ($ProgressWindow -ne $null -and $ProgressWindow.IsVisible) {
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
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
function DownloadAndRun {
    <#
    .SYNOPSIS
    Descarga y ejecuta una aplicación desde URL
    #>
    param(
        [string]$url,
        [string]$zipPath,
        [string]$extractPath,
        [string]$exeName,
        [string]$validationPath
    )
    if (!(Test-Path -Path $validationPath)) {
        $response = [System.Windows.MessageBox]::Show(
            "El archivo o aplicación no se encontró en '$validationPath'. ¿Desea descargarlo?",
            "Archivo no encontrado",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($response -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red
            return
        }
    }
    if (Test-Path -Path $zipPath) {
        $response = [System.Windows.MessageBox]::Show(
            "Archivo encontrado. ¿Lo desea eliminar y volver a descargar?",
            "Archivo ya descargado",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($response -eq [System.Windows.MessageBoxResult]::Yes) {
            Remove-Item -Path $zipPath -Force
            Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "`tEliminando archivos anteriores..."
        } elseif ($response -eq [System.Windows.MessageBoxResult]::No) {
            $exePath = Join-Path -Path $extractPath -ChildPath $exeName
            if (Test-Path -Path $exePath) {
                Write-Host "`tEjecutando el archivo ya descargado..."
                Start-Process -FilePath $exePath
                Write-Host "`t$exeName se está ejecutando."
                return
            } else {
                Write-Host "`tNo se pudo encontrar el archivo ejecutable." -ForegroundColor Red
                return
            }
        } elseif ($response -eq [System.Windows.MessageBoxResult]::Cancel) {
            Write-Host "`tEl usuario canceló la operación." -ForegroundColor Red
            return
        }
    }
    Write-Host "`tDescargando desde: $url"
    try {
        $response = Invoke-WebRequest -Uri $url -Method Head
        $totalSize = $response.Headers["Content-Length"]
        $totalSizeKB = [math]::round($totalSize / 1KB, 2)
        Write-Host "`tTamaño total: $totalSizeKB KB" -ForegroundColor Yellow
        Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
        Write-Host "`tDescarga completada." -ForegroundColor Green
    } catch {
        Write-Host "`tError en descarga: $_" -ForegroundColor Red
        return
    }
    if (!(Test-Path -Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }
    Write-Host "`tExtrayendo archivos..."
    try {
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Host "`tArchivos extraídos correctamente." -ForegroundColor Green
    } catch {
        Write-Host "`tError al descomprimir: $_" -ForegroundColor Red
        return
    }
    $exePath = Join-Path -Path $extractPath -ChildPath $exeName
    if (Test-Path -Path $exePath) {
        Write-Host "`tEjecutando $exeName..."
        Start-Process -FilePath $exePath
        Write-Host "`n$exeName se está ejecutando."
    } else {
        Write-Host "`nNo se pudo encontrar el archivo ejecutable." -ForegroundColor Red
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
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.MessageBox]::Show(
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
    Write-DzDebug "[DEBUG] Fecha/Hora: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    $ports = @()

    # Definir rutas de registro a buscar (64-bit y 32-bit)
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server",  # 64-bit
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server"  # 32-bit
    )

    # 1. Buscar por todas las instancias instaladas
    Write-DzDebug "[DEBUG] 1. Buscando instancias SQL instaladas..."

    foreach ($basePath in $registryPaths) {
        try {
            Write-DzDebug "[DEBUG]   Examinando ruta: $basePath"

            if (-not (Test-Path $basePath)) {
                Write-DzDebug "[DEBUG]   ✗ La ruta no existe"
                continue
            }

            # Método 1: Obtener de InstalledInstances
            Write-DzDebug "[DEBUG]   Método 1: Buscando en 'InstalledInstances'..."
            $installedInstances = Get-ItemProperty -Path $basePath -Name "InstalledInstances" -ErrorAction SilentlyContinue

            if ($installedInstances -and $installedInstances.InstalledInstances) {
                Write-DzDebug "[DEBUG]   ✓ Instancias encontradas: $($installedInstances.InstalledInstances -join ', ')"

                foreach ($instance in $installedInstances.InstalledInstances) {
                    # Si ya procesamos esta instancia, saltar
                    if ($ports.Instance -contains $instance) {
                        Write-DzDebug "[DEBUG]     Instancia '$instance' ya procesada, saltando..."
                        continue
                    }

                    Write-DzDebug "[DEBUG]     Procesando instancia: '$instance'"

                    # Construir rutas posibles para esta instancia
                    $possiblePaths = @(
                        "$basePath\$instance\MSSQLServer\SuperSocketNetLib\Tcp",
                        "$basePath\$instance\MSSQLServer\SuperSocketNetLib\Tcp\IPAll",
                        "$basePath\MSSQLServer\$instance\SuperSocketNetLib\Tcp"
                    )

                    $portFound = $false
                    foreach ($tcpPath in $possiblePaths) {
                        Write-DzDebug "[DEBUG]       Probando ruta: $tcpPath"

                        if (Test-Path $tcpPath) {
                            Write-DzDebug "[DEBUG]       ✓ Ruta existe"

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
                                Write-DzDebug "[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
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
                                Write-DzDebug "[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                                $portFound = $true
                                break
                            } else {
                                Write-DzDebug "[DEBUG]       ✗ No se encontró puerto en esta ruta"
                            }
                        } else {
                            Write-DzDebug "[DEBUG]       ✗ Ruta no existe"
                        }
                    }

                    if (-not $portFound) {
                        Write-DzDebug "[DEBUG]     ✗ No se encontró puerto para la instancia '$instance'"
                    }
                }
            } else {
                Write-DzDebug "[DEBUG]   ✗ No se encontró la clave 'InstalledInstances' en esta ruta"
            }

            # Método 2: Explorar todas las carpetas bajo SQL Server
            Write-DzDebug "[DEBUG]   Método 2: Explorando todas las carpetas..."

            $allSqlEntries = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue |
            Where-Object { $_.PSIsContainer } |
            Select-Object -ExpandProperty PSChildName

            if ($allSqlEntries) {
                Write-DzDebug "[DEBUG]   Carpetas encontradas: $($allSqlEntries -join ', ')"

                foreach ($entry in $allSqlEntries) {
                    # Filtrar nombres que parecen instancias
                    if ($entry -match "^MSSQL\d+" -or $entry -match "^SQL" -or $entry -match "NATIONALSOFT") {
                        # Si ya procesamos esta instancia, saltar
                        if ($ports.Instance -contains $entry) {
                            Write-DzDebug "[DEBUG]     Instancia '$entry' ya procesada, saltando..."
                            continue
                        }

                        Write-DzDebug "[DEBUG]     Analizando posible instancia: '$entry'"

                        $tcpPath = "$basePath\$entry\MSSQLServer\SuperSocketNetLib\Tcp"

                        if (Test-Path $tcpPath) {
                            Write-DzDebug "[DEBUG]       ✓ Ruta TCP encontrada: $tcpPath"

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
                                Write-DzDebug "[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
                            } elseif ($tcpDynamicPorts -and $tcpDynamicPorts.TcpDynamicPorts) {
                                $portInfo = [PSCustomObject]@{
                                    Instance = $entry
                                    Port     = $tcpDynamicPorts.TcpDynamicPorts
                                    Path     = $tcpPath
                                    Type     = "Dynamic"
                                }
                                $ports += $portInfo
                                Write-DzDebug "[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                            } else {
                                Write-DzDebug "[DEBUG]       ✗ No se encontró puerto en esta ruta"
                            }
                        }
                    }
                }
            }

        } catch {
            Write-DzDebug "[DEBUG]   ERROR en búsqueda: $($_.Exception.Message)"
            Write-DzDebug "[DEBUG]   StackTrace: $($_.ScriptStackTrace)"
        }
    }

    # Método 3: Buscar por servicios SQL Server (común para ambos)
    Write-DzDebug "[DEBUG]   Método 3: Buscando servicios SQL Server..."

    $sqlServices = Get-Service -Name "*SQL*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*SQL Server (*" }

    foreach ($service in $sqlServices) {
        Write-DzDebug "[DEBUG]     Servicio: $($service.DisplayName)"

        # Extraer nombre de instancia del servicio
        if ($service.DisplayName -match "SQL Server \((.+)\)") {
            $instanceName = $matches[1]

            # Si ya procesamos esta instancia, saltar
            if ($ports.Instance -contains $instanceName) {
                Write-DzDebug "[DEBUG]       Instancia '$instanceName' ya procesada, saltando..."
                continue
            }

            Write-DzDebug "[DEBUG]       Posible instancia: '$instanceName'"
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
                        Write-DzDebug "[DEBUG]       ✓ Puerto estático encontrado: $($tcpPort.TcpPort)"
                        break
                    } elseif ($tcpDynamicPorts -and $tcpDynamicPorts.TcpDynamicPorts) {
                        $portInfo = [PSCustomObject]@{
                            Instance = $instanceName
                            Port     = $tcpDynamicPorts.TcpDynamicPorts
                            Path     = $tcpPath
                            Type     = "Dynamic"
                        }
                        $ports += $portInfo
                        Write-DzDebug "[DEBUG]       ✓ Puerto dinámico encontrado: $($tcpDynamicPorts.TcpDynamicPorts)"
                        break
                    }
                }
            }
        }
    }

    # Método 4: Buscar puertos en uso por SQL Server
    Write-DzDebug "[DEBUG]   Método 4: Buscando puertos en uso por sqlservr.exe..."

    $sqlProcesses = Get-Process -Name "sqlservr" -ErrorAction SilentlyContinue
    if ($sqlProcesses) {
        foreach ($process in $sqlProcesses) {
            Write-DzDebug "[DEBUG]     Proceso sqlservr.exe encontrado (PID: $($process.Id))"

            # Obtener puertos usando netstat
            $netstatOutput = netstat -ano | Select-String ":$($process.Id)\s"

            foreach ($line in $netstatOutput) {
                if ($line -match ":(\d+)\s.*$($process.Id)$") {
                    $port = $matches[1]
                    Write-DzDebug "[DEBUG]       Puerto en uso: $port"

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
        Write-DzDebug "[DEBUG]     ✗ No se encontraron procesos sqlservr.exe"
    }

    # Resumen final
    Write-DzDebug "[DEBUG] `n=== RESUMEN DE BÚSQUEDA ==="
    Write-DzDebug "[DEBUG] Total de instancias con puerto encontradas: $($ports.Count)"

    if ($ports.Count -gt 0) {
        foreach ($port in $ports) {
            Write-DzDebug "[DEBUG]   - Instancia: $($port.Instance) | Puerto: $($port.Port) | Tipo: $($port.Type)"
            Write-DzDebug "[DEBUG]     Ruta: $($port.Path)"
        }
    } else {
        Write-DzDebug "[DEBUG]   ✗ No se encontraron puertos SQL configurados"

        # Sugerencias para debugging
        Write-DzDebug "[DEBUG] `n=== SUGERENCIAS ==="
        Write-DzDebug "[DEBUG] 1. Verifica si SQL Server está instalado"
        Write-DzDebug "[DEBUG] 2. Revisa el Configuration Manager de SQL Server"
        Write-DzDebug "[DEBUG] 3. Verifica si el servicio SQL Server está ejecutándose"
        Write-DzDebug "[DEBUG] 4. Consulta el log de errores de SQL Server"
    }

    Write-DzDebug "[DEBUG] === FIN DE BÚSQUEDA ===`n"

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
Export-ModuleMember -Function @(
    'Get-DzToolsConfigPath',
    'Get-DzDebugPreference',
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
    'DownloadAndRun',
    'Refresh-AdapterStatus',
    'Get-NetworkAdapterStatus',
    'Start-SystemUpdate',
    'Get-SqlPortWithDebug',
    'Show-SqlPortsInfo'
)