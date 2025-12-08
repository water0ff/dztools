#requires -Version 5.0
function Check-Chocolatey {
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        $response = [System.Windows.Forms.MessageBox]::Show(
            "Chocolatey no está instalado. ¿Desea instalarlo ahora?",
            "Chocolatey no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($response -eq [System.Windows.Forms.DialogResult]::No) {
            Write-Host "`nEl usuario canceló la instalación de Chocolatey." -ForegroundColor Red
            return $false  # Retorna falso si el usuario cancela
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

            [System.Windows.Forms.MessageBox]::Show(
                "Chocolatey se instaló correctamente y ha sido configurado. Por favor, reinicie PowerShell antes de continuar.",
                "Reinicio requerido",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            # Cerrar el programa automáticamente
            Write-Host "`nCerrando la aplicación para permitir reinicio de PowerShell..." -ForegroundColor Red
            Stop-Process -Id $PID -Force
            return $false # Retorna falso para indicar que se debe reiniciar
        } catch {
            Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "Error al instalar Chocolatey. Por favor, inténtelo manualmente.",
                "Error de instalación",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false # Retorna falso en caso de error
        }
    } else {
        Write-Host "`tChocolatey ya está instalado." -ForegroundColor Green
        return $true # Retorna verdadero si Chocolatey ya está instalado
    }
}
function Invoke-ChocoCommandWithProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Arguments,
        [Parameter(Mandatory = $true)][string]$OperationTitle
    )
    $progressForm = $null
    $process = $null
    try {
        Write-DzDebug ("`t[DEBUG] Invoke-ChocoCommandWithProgress: argumentos='{0}'" -f $Arguments)
        $progressForm = Show-ProgressBar
        if ($null -ne $progressForm -and -not $progressForm.IsDisposed) {
            $progressForm.HeaderLabel.Text = $OperationTitle
        }
        $currentPercent = 0
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentPercent -TotalSteps 100 -Status "Preparando comando de Chocolatey..."
        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'choco'
        $psi.Arguments = "$Arguments --verbose"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $outputHandler = {
            param($sender, $eventArgs)
            if (-not [string]::IsNullOrWhiteSpace($eventArgs.Data)) {
                $queue.Enqueue($eventArgs.Data)
            }
        }
        $process.add_OutputDataReceived($outputHandler)
        $process.add_ErrorDataReceived($outputHandler)
        if (-not $process.Start()) {
            throw "No se pudo iniciar el proceso de Chocolatey."
        }
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        $line = $null
        while (-not $process.HasExited) {
            while ($queue.TryDequeue([ref]$line)) {
                $currentPercent = [math]::Min(95, $currentPercent + 1)
                Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentPercent -TotalSteps 100 -Status $line
                Write-DzDebug ("`t[DEBUG] choco> {0}" -f $line)
            }
            Start-Sleep -Milliseconds 150
        }
        $process.WaitForExit()
        while ($queue.TryDequeue([ref]$line)) {
            $currentPercent = [math]::Min(95, $currentPercent + 1)
            Update-ProgressBar -ProgressForm $progressForm -CurrentStep $currentPercent -TotalSteps 100 -Status $line
            Write-DzDebug ("`t[DEBUG] choco> {0}" -f $line)
        }
        $exitCode = $process.ExitCode
        $finalStatus = "Chocolatey finalizó con código $exitCode"
        Update-ProgressBar -ProgressForm $progressForm -CurrentStep 100 -TotalSteps 100 -Status $finalStatus
        Write-DzDebug ("`t[DEBUG] Invoke-ChocoCommandWithProgress: código de salida {0}" -f $exitCode)
        return $exitCode
    } catch {
        Write-DzDebug ("`t[DEBUG] Invoke-ChocoCommandWithProgress - Error: {0}" -f $_)
        [System.Windows.Forms.MessageBox]::Show(
            "Ocurrió un error al ejecutar Chocolatey: $($_.Exception.Message)",
            "Error en instalación/desinstalación",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return -1
    } finally {
        if ($null -ne $process) {
            try {
                $process.Dispose()
            } catch {
                Write-DzDebug "`t[DEBUG] No se pudo liberar el proceso de Chocolatey" -Color DarkYellow
            }
        }
    }
}
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
    } catch {
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
function Show-SSMSInstallerDialog {
    $form = Create-Form -Title "Instalar SSMS" `
        -Size (New-Object System.Drawing.Size(360, 180)) `
        -StartPosition ([System.Windows.Forms.FormStartPosition]::CenterScreen) `
        -FormBorderStyle ([System.Windows.Forms.FormBorderStyle]::FixedDialog) `
        -MaximizeBox $false -MinimizeBox $false -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))

    $lbl = Create-Label -Text "Elige la versión a instalar:" -Location (New-Object System.Drawing.Point(10, 15)) -Size (New-Object System.Drawing.Size(320, 20))
    $cmb = Create-ComboBox -Location (New-Object System.Drawing.Point(10, 40)) -Size (New-Object System.Drawing.Size(320, 22)) -DropDownStyle DropDownList
    $null = $cmb.Items.Add("Último disponible.")
    $null = $cmb.Items.Add("SSMS 14 (2014)")
    $cmb.SelectedIndex = 0

    $btnOK = Create-Button -Text "Instalar" -Location (New-Object System.Drawing.Point(10, 80)) -Size (New-Object System.Drawing.Size(140, 30))
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = Create-Button -Text "Cancelar" -Location (New-Object System.Drawing.Point(190, 80)) -Size (New-Object System.Drawing.Size(140, 30))
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($lbl, $cmb, $btnOK, $btnCancel))

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    switch ($cmb.SelectedIndex) {
        0 { return "latest" }
        1 { return "ssms14" }
    }
}
Export-ModuleMember -Function Install-Software, Download-File, Expand-ArchiveFile, Check-Chocolatey, Show-SSMSInstallerDialog, Invoke-ChocoCommandWithProgress