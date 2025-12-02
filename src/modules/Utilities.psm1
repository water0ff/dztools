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
        ComputerName = [System.Net.Dns]::GetHostName()
        OS = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        NetAdapters = @()
    }
    
    # Obtener adaptadores de red
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    foreach ($adapter in $adapters) {
        $adapterInfo = @{
            Name = $adapter.Name
            Status = $adapter.Status
            MacAddress = $adapter.MacAddress
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
                    }
                    catch {
                        Write-Verbose "No se pudo eliminar: $($item.FullName)"
                    }
                }
            }
            catch {
                Write-Warning "Error accediendo a $path : $_"
            }
        }
    }
    
    return @{
        FilesDeleted = $totalDeleted
        SpaceFreedMB = [math]::Round($totalSize / 1MB, 2)
    }
}

function Get-IniConnections {
    [CmdletBinding()]
    param()
    
    $connections = @()
    $pathsToCheck = @(
        @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
        @{ Path = "C:\NationalSoft\Softrestaurant10.0";    INI = "restaurant.ini"; Nombre = "SR10" },
        @{ Path = "C:\NationalSoft\Softrestaurant11.0";    INI = "restaurant.ini"; Nombre = "SR11" },
        @{ Path = "C:\NationalSoft\Softrestaurant12.0";    INI = "restaurant.ini"; Nombre = "SR12" },
        @{ Path = "C:\Program Files (x86)\NsBackOffice1.0";    INI = "DbConfig.ini"; Nombre = "NSBackOffice" },
        @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0";INI = "nshoteles.ini";   Nombre = "Hoteles" },
        @{ Path = "C:\NationalSoft\OnTheMinute4.5";        INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
    )
    
    function Get-IniValue {
        param([string]$FilePath, [string]$Key)
        
        if (Test-Path $FilePath) {
            $line = Get-Content $FilePath | Where-Object { $_ -match "^$Key\s*=" }
            if ($line) {
                return $line.Split('=')[1].Trim()
            }
        }
        return $null
    }
    
    foreach ($entry in $pathsToCheck) {
        $mainIni = Join-Path $entry.Path $entry.INI
        if (Test-Path $mainIni) {
            $dataSource = Get-IniValue -FilePath $mainIni -Key "DataSource"
            if ($dataSource -and $dataSource -notin $connections) {
                $connections += $dataSource
            }
        }
        
        $inisFolder = Join-Path $entry.Path "INIS"
        if (Test-Path $inisFolder) {
            $iniFiles = Get-ChildItem -Path $inisFolder -Filter "*.ini"
            foreach ($iniFile in $iniFiles) {
                $dataSource = Get-IniValue -FilePath $iniFile.FullName -Key "DataSource"
                if ($dataSource -and $dataSource -notin $connections) {
                    $connections += $dataSource
                }
            }
        }
    }
    
    return $connections | Sort-Object
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
        
        # Configurar cache location
        choco config set cacheLocation C:\Choco\cache
        
        Write-Host "Chocolatey instalado correctamente" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Error instalando Chocolatey: $_"
        return $false
    }
}

Export-ModuleMember -Function Test-Administrator, Get-SystemInfo, Clear-TemporaryFiles, 
    Get-IniConnections, Test-ChocolateyInstalled, Install-Chocolatey
