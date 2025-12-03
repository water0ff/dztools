function Initialize-DanielToolsEnvironment {
    <#
        .SYNOPSIS
            Prepara rutas temporales y ensambla dependencias de Windows Forms.
        .PARAMETER TempPath
            Ruta base donde se almacenarán archivos temporales y carpetas auxiliares.
        .PARAMETER IconDirectory
            Carpeta específica para íconos utilizados por la aplicación.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TempPath = "C:\\Temp",

        [Parameter()]
        [string]$IconDirectory = (Join-Path $TempPath "icos")
    )

    if (-not (Test-Path -Path $TempPath)) {
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
        Write-Host "Carpeta '$TempPath' creada correctamente."
    }

    if (-not (Test-Path -Path $IconDirectory)) {
        New-Item -ItemType Directory -Path $IconDirectory -Force | Out-Null
        Write-Host "Carpeta de íconos creada: $IconDirectory"
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    return @{
        TempPath     = $TempPath
        IconDirectory = $IconDirectory
    }
}

Export-ModuleMember -Function Initialize-DanielToolsEnvironment
