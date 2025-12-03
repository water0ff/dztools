#requires -Version 5.0

function Test-Administrator {
    [CmdletBinding()]
    param()
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Get-SystemInfo {
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
    [CmdletBinding()]
    param()
    return [bool](Get-Command choco -ErrorAction SilentlyContinue)
}
function Install-Chocolatey {
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
        return "Administrators" # Valor por defecto
    }
}
function Invoke-DiskCleanup {
    try {
        Write-Host "`nEjecutando Liberador de espacio en disco..." -ForegroundColor Cyan
        $cleanmgr = "$env:SystemDrive\Windows\System32\cleanmgr.exe"
        $sagerun = "9999"
        Start-Process $cleanmgr -ArgumentList "/sageset:$sagerun" -Wait
        Start-Process $cleanmgr -ArgumentList "/sagerun:$sagerun" -Wait

        Write-Host "Limpieza de disco completada correctamente" -ForegroundColor Green
    } catch {
        Write-Host "Error en limpieza de disco: $($_.Exception.Message)" -ForegroundColor Red
    }
}
function Show-SystemComponents {
    $criticalError = $false

    Write-Host "`n=== Componentes del sistema detectados ===" -ForegroundColor Cyan

    # Versión de Windows (componente crítico)
    try {
        $os = Get-CimInstance -ClassName CIM_OperatingSystem -ErrorAction Stop
        Write-Host "`n[Windows]" -ForegroundColor Yellow
        Write-Host "Versión: $($os.Caption) (Build $($os.Version))" -ForegroundColor White
    } catch {
        $criticalError = $true
        Write-Host "`n[Windows]" -ForegroundColor Yellow
        Write-Host "ERROR CRÍTICO: $($_.Exception.Message)" -ForegroundColor Red
        throw "No se pudo obtener información crítica del sistema"
    }

    # Resto de componentes (no críticos)
    if (-not $criticalError) {
        # Procesador
        try {
            $procesador = Get-CimInstance -ClassName CIM_Processor -ErrorAction Stop
            Write-Host "`n[Procesador]" -ForegroundColor Yellow
            Write-Host "Modelo: $($procesador.Name)" -ForegroundColor White
            Write-Host "Núcleos: $($procesador.NumberOfCores)" -ForegroundColor White
        } catch {
            Write-Host "`n[Procesador]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Memoria RAM
        try {
            $memoria = Get-CimInstance -ClassName CIM_PhysicalMemory -ErrorAction Stop
            Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
            $memoria | ForEach-Object {
                Write-Host "Módulo: $([math]::Round($_.Capacity/1GB, 2)) GB $($_.Manufacturer) ($($_.Speed) MHz)" -ForegroundColor White
            }
        } catch {
            Write-Host "`n[Memoria RAM]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Discos duros
        try {
            $discos = Get-CimInstance -ClassName CIM_DiskDrive -ErrorAction Stop
            Write-Host "`n[Discos duros]" -ForegroundColor Yellow
            $discos | ForEach-Object {
                Write-Host "Disco: $($_.Model) ($([math]::Round($_.Size/1GB, 2)) GB)" -ForegroundColor White
            }
        } catch {
            Write-Host "`n[Discos duros]" -ForegroundColor Yellow
            Write-Host "Error de lectura: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
function Test-SameHost {
    param(
        [string]$serverName
    )
    $machinePart = $serverName.Split('\')[0]
    $machineName = $machinePart.Split(',')[0]
    if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }
    return ($env:COMPUTERNAME -eq $machineName)
}
function Test-7ZipInstalled {
    return (Test-Path "C:\Program Files\7-Zip\7z.exe")
}
function Test-MegaToolsInstalled {
    return ([bool](Get-Command megatools -ErrorAction SilentlyContinue))
}
function Check-Permissions {
    $folderPath = "C:\NationalSoft"
    try {
        $directoryInfo = New-Object System.IO.DirectoryInfo($folderPath)
        $acl = $directoryInfo.GetAccessControl()  # Devuelve DirectorySecurity
    } catch {
        Write-Host ("Error obteniendo ACL de {0}: {1}" -f $folderPath, $_.Exception.Message) -ForegroundColor Red
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
    $permissions | ForEach-Object {
        Write-Host "`t$($_.Usuario) - $($_.Tipo) - " -NoNewline
        Write-Host "` $($_.Permiso)" -ForegroundColor Green
    }
    if ($everyonePermissions.Count -gt 0) {
        Write-Host "`tEveryone tiene los siguientes permisos:" -NoNewline -ForegroundColor Yellow
        Write-Host " $($everyonePermissions -join ', ')" -ForegroundColor Green
    } else {
        Write-Host "`tNo hay permisos para 'Everyone'" -ForegroundColor Red
    }
    if (-not $everyoneHasFullControl) {
        $message = "El usuario 'Everyone' no tiene permisos de 'Full Control'. ¿Deseas concederlo?"
        $title = "Permisos 'Everyone'"
        $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
        $icon = [System.Windows.Forms.MessageBoxIcon]::Question
        $result = [System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $everyoneSid,
                    [System.Security.AccessControl.FileSystemRights]::FullControl,
                    [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit",
                    [System.Security.AccessControl.PropagationFlags]::None,
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                $acl.AddAccessRule($accessRule)
                $acl.SetAccessRuleProtection($false, $true)
                $directoryInfo.SetAccessControl($acl)
                Write-Host "Se ha concedido 'Full Control' a 'Everyone'." -ForegroundColor Green
            } catch {
                Write-Host ("Error aplicando permisos a {0}: {1}" -f $folderPath, $_.Exception.Message) -ForegroundColor Red
            }
        }
    }
}
function DownloadAndRun($url, $zipPath, $extractPath, $exeName, $validationPath) {
    # Validar si el archivo o aplicación ya existe
    if (!(Test-Path -Path $validationPath)) {
        $response = [System.Windows.Forms.MessageBox]::Show(
            "El archivo o aplicación no se encontró en '$validationPath'. ¿Desea descargarlo?",
            "Archivo no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        # Si el usuario selecciona "No", salir de la función
        if ($response -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
            return
        }
    }
    # Verificar si el archivo ZIP ya existe
    if (Test-Path -Path $zipPath) {
        $response = [System.Windows.Forms.MessageBox]::Show(
            "Archivo encontrado. ¿Lo desea eliminar y volver a descargar?",
            "Archivo ya descargado",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
        )
        if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-Item -Path $zipPath -Force
            Remove-Item -Path $extractPath -Recurse -Force
            Write-Host "`tEliminando archivos anteriores..."
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::No) {
            # Si selecciona "No", abrir el programa sin eliminar archivos
            $exePath = Join-Path -Path $extractPath -ChildPath $exeName
            if (Test-Path -Path $exePath) {
                Write-Host "`tEjecutando el archivo ya descargado..."
                Start-Process -FilePath $exePath #-Wait   # Se quitó para ver si se usaban múltiples apps.
                Write-Host "`t$exeName se está ejecutando."
                return
            } else {
                Write-Host "`tNo se pudo encontrar el archivo ejecutable."  -ForegroundColor Red
                return
            }
        } elseif ($response -eq [System.Windows.Forms.DialogResult]::Cancel) {
            # Si selecciona "Cancelar", no hacer nada y decir que el usuario canceló
            Write-Host "`tEl usuario canceló la operación."  -ForegroundColor Red
            return  # Aquí se termina la ejecución si el usuario cancela
        }
    }
    # Proceder con la descarga si no fue cancelada
    Write-Host "`tDescargando desde: $url"
    # Obtener el tamaño total del archivo antes de la descarga
    $response = Invoke-WebRequest -Uri $url -Method Head
    $totalSize = $response.Headers["Content-Length"]
    $totalSizeKB = [math]::round($totalSize / 1KB, 2)
    Write-Host "`tTamaño total: $totalSizeKB KB" -ForegroundColor Yellow
    # Descargar el archivo con barra de progreso
    $downloaded = 0
    $request = Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
    foreach ($chunk in $request.Content) {
        $downloaded += $chunk.Length
        $downloadedKB = [math]::round($downloaded / 1KB, 2)
        $progress = [math]::round(($downloaded / $totalSize) * 100, 2)
        Write-Progress -PercentComplete $progress -Status "Descargando..." -Activity "Progreso de la descarga" -CurrentOperation "$downloadedKB KB de $totalSizeKB KB descargados"
    }
    Write-Host "`tDescarga completada."  -ForegroundColor Green
    # Crear directorio de extracción si no existe
    if (!(Test-Path -Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }
    Write-Host "`tExtrayendo archivos..."
    try {
        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
        Write-Host "`tArchivos extraídos correctamente."  -ForegroundColor Green
    } catch {
        Write-Host "`tError al descomprimir el archivo: $_"   -ForegroundColor Red
    }
    $exePath = Join-Path -Path $extractPath -ChildPath $exeName
    if (Test-Path -Path $exePath) {
        Write-Host "`tEjecutando $exeName..."
        Start-Process -FilePath $exePath #-Wait
        Write-Host "`n$exeName se está ejecutando."
    } else {
        Write-Host "`nNo se pudo encontrar el archivo ejecutable."  -ForegroundColor Red
    }
}
function Refresh-AdapterStatus {
    $statuses = Get-NetworkAdapterStatus
    if ($statuses.Count -gt 0) {
        $lines = $statuses | ForEach-Object {
            "- $($_.AdapterName) - $($_.NetworkCategory)"
        }
        $txt_AdapterStatus.Text = $lines -join "`r`n"
    } else {
        $txt_AdapterStatus.Text = "No se encontraron adaptadores activos."
    }
}
function Get-NetworkAdapterStatus {
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    $profiles = Get-NetConnectionProfile
    $adapterStatus = @()
    foreach ($adapter in $adapters) {
        $profile = $profiles | Where-Object { $_.InterfaceIndex -eq $adapter.ifIndex }
        $networkCategory = if ($profile) { $profile.NetworkCategory } else { "Desconocido" }
        $adapterStatus += [PSCustomObject]@{
            AdapterName     = $adapter.Name
            NetworkCategory = $networkCategory
            InterfaceIndex  = $adapter.ifIndex  # Guardar el InterfaceIndex para identificar el adaptador
        }
    }
    return $adapterStatus
}
function Start-SystemUpdate {
    $progressForm = $null
    try {
        $progressForm = Show-ProgressBar
        $totalSteps = 6
        $currentStep = 0
        Write-Host "`nIniciando proceso de actualización..." -ForegroundColor Cyan
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 1/$totalSteps] Deteniendo servicio winmgmt..." -ForegroundColor Yellow
        $service = Get-Service -Name "winmgmt" -ErrorAction Stop
        if ($service.Status -eq "Running") {
            Stop-Service -Name "winmgmt" -Force -ErrorAction Stop
            Write-Host "Servicio detenido correctamente." -ForegroundColor Green
        }
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 2/$totalSteps] Renombrando carpeta Repository..." -ForegroundColor Yellow
        try {
            $repoPath = Join-Path $env:windir "System32\Wbem\Repository"
            if (Test-Path $repoPath) {
                $newName = "Repository_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Rename-Item -Path $repoPath -NewName $newName -Force -ErrorAction Stop
                Write-Host "Carpeta renombrada: $newName" -ForegroundColor Green
            }
        } catch {
            Write-Host "Advertencia: No se pudo renombrar la carpeta Repository. Continuando..." -ForegroundColor Yellow
        }
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 3/$totalSteps] Reiniciando servicio winmgmt..." -ForegroundColor Yellow
        net start winmgmt *>&1 | Write-Host
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 4/$totalSteps] Limpiando archivos temporales (ignorar si hay errores)..." -ForegroundColor Cyan
        $totalDeleted = 0
        $totalDeleted += Clear-TemporaryFiles -folderPath $env:TEMP
        $totalDeleted += Clear-TemporaryFiles -folderPath "$env:SystemDrive\Windows\Temp"
        Write-Host "Total archivos eliminados: $totalDeleted" -ForegroundColor Green
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 5/$totalSteps] Ejecutando Liberador de espacio..." -ForegroundColor Cyan
        Invoke-DiskCleanup
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`n[Paso 6/$totalSteps] Obteniendo información del sistema..." -ForegroundColor Cyan
        Show-SystemComponents
        $currentStep++
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentStep -TotalSteps $totalSteps
        Write-Host "`nProceso completado con éxito" -ForegroundColor Green
    } catch {
        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Detalles: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        [System.Windows.Forms.MessageBox]::Show(
            "Error: $($_.Exception.Message)`nRevise los logs antes de reiniciar.",
            "Error crítico",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    } finally {
        if ($progressForm -ne $null -and -not $progressForm.IsDisposed) {
            Close-ProgressBar $progressForm
        }
    }
}
#requires -Version 5.0
Export-ModuleMember -Function Test-Administrator, Get-SystemInfo, Clear-TemporaryFiles,
Test-ChocolateyInstalled, Install-Chocolatey, Get-AdminGroupName, Invoke-DiskCleanup,
Show-SystemComponents, Test-SameHost, Test-7ZipInstalled, Test-MegaToolsInstalled,
Refresh-AdapterStatus, Get-NetworkAdapterStatus, DownloadAndRun, Start-SystemUpdate, Check-Permissions