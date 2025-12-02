#requires -Version 5.0

function Install-Software {
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
        Write-Error "Chocolatey no está instalado. Use Install-Chocolatey primero."
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
    }
    catch {
        Write-Error "Error instalando $Software : $_"
        return $false
    }
}

function Download-File {
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
        }
        else {
            # Descarga simple
            Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
        }
        
        Write-Verbose "Archivo descargado: $OutputPath"
        return $true
    }
    catch {
        Write-Error "Error descargando archivo: $_"
        return $false
    }
}

function Expand-ArchiveFile {
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
    }
    catch {
        Write-Error "Error extrayendo archivo: $_"
        return $false
    }
}

Export-ModuleMember -Function Install-Software, Download-File, Expand-ArchiveFile
