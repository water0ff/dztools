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
            return $false
        }
        Write-Host "`nInstalando Chocolatey..." -ForegroundColor Cyan
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            Write-Host "`nChocolatey se instaló correctamente." -ForegroundColor Green
            Write-Host "`nConfigurando Chocolatey..." -ForegroundColor Yellow
            choco config set cacheLocation C:\Choco\cache
            [System.Windows.Forms.MessageBox]::Show(
                "Chocolatey se instaló correctamente. La aplicación se cerrará. Por favor, reiníciela para continuar.",
                "Reinicio requerido",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            [System.Environment]::Exit(0)
        } catch {
            Write-Host "`nError al instalar Chocolatey: $_" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show(
                "Error al instalar Chocolatey: $($_.Exception.Message)`n`nPor favor, inténtelo manualmente.",
                "Error de instalación",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    } else {
        Write-Host "`tChocolatey ya está instalado." -ForegroundColor Green
        return $true
    }
}
function Invoke-ChocoCommandWithProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Arguments,
        [Parameter(Mandatory = $true)][string]$OperationTitle
    )

    Write-DzDebug "`n========== INICIO Invoke-ChocoCommandWithProgress =========="
    Write-DzDebug ("`t[DEBUG] Argumentos: {0}" -f $Arguments)
    Write-DzDebug ("`t[DEBUG] Título: {0}" -f $OperationTitle)

    $process = $null
    $exitCode = -1
    $statusMessage = "Procesando (revise la consola)..."

    try {
        # Verificar Chocolatey
        Write-DzDebug "`t[DEBUG] Verificando Chocolatey..."
        $chocoCmd = Get-Command choco -ErrorAction SilentlyContinue

        if (-not $chocoCmd) {
            Write-DzDebug "`t[ERROR] Chocolatey no encontrado en PATH" -Color Red
            throw "Chocolatey no está instalado o no está en el PATH"
        }

        Write-DzDebug ("`t[DEBUG] Chocolatey encontrado en: {0}" -f $chocoCmd.Source)

        if ($OperationTitle -match 'Desinstalando') {
            $statusMessage = "Desinstalando (revise la consola)..."
        } elseif ($OperationTitle -match 'Instalando') {
            $statusMessage = "Instalando (revise la consola)..."
        }

        # Progreso deshabilitado temporalmente por inestabilidad con Chocolatey
        Write-DzDebug "`t[DEBUG] Barra de progreso deshabilitada para comandos de Chocolatey"
        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'choco'
        $psi.Arguments = "$Arguments --verbose"
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
        $psi.CreateNoWindow = $true
        $psi.WorkingDirectory = [System.IO.Path]::GetTempPath()

        Write-DzDebug ("`t[DEBUG] WorkingDirectory: {0}" -f $psi.WorkingDirectory)

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $outputHandler = {
            param($sender, $eventArgs)
            try {
                if ($eventArgs -and -not [string]::IsNullOrWhiteSpace($eventArgs.Data)) {
                    $queue.Enqueue($eventArgs.Data)
                }
            } catch {
                Write-DzDebug "[ERROR handler] $($_.Exception.Message)" -Color Red
            }
        }
        if ($process -eq $null) {
            throw "El objeto Process es NULL antes de iniciar Chocolatey."
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
                Write-DzDebug ("`t[DEBUG] choco> {0}" -f $line)
            }
            Start-Sleep -Milliseconds 150
        }
        $process.WaitForExit()
        # Drenar lo que quede en la cola
        while ($queue.TryDequeue([ref]$line)) {
            Write-DzDebug ("`t[DEBUG] choco> {0}" -f $line)
        }
        $exitCode = $process.ExitCode
        $finalStatus = "Chocolatey finalizó con código $exitCode"
        $finalStatus = if ($exitCode -eq 0) { "Completado exitosamente" } else { "Finalizado con código $exitCode" }

        Write-DzDebug ("`t[DEBUG] Retornando código de salida: {0}" -f $exitCode)
        Write-DzDebug "========== FIN Invoke-ChocoCommandWithProgress ==========`n"

        return $exitCode

    } catch {
        $errorMsg = $_.Exception.Message
        $errorType = $_.Exception.GetType().FullName

        Write-DzDebug "`n========== ERROR EN Invoke-ChocoCommandWithProgress ==========" -Color Red
        Write-DzDebug ("`t[ERROR] Tipo: {0}" -f $errorType) -Color Red
        Write-DzDebug ("`t[ERROR] Mensaje: {0}" -f $errorMsg) -Color Red

        if ($_.InvocationInfo) {
            Write-DzDebug ("`t[ERROR] Archivo: {0}" -f $_.InvocationInfo.ScriptName) -Color DarkYellow
            Write-DzDebug ("`t[ERROR] Línea: {0}" -f $_.InvocationInfo.ScriptLineNumber) -Color DarkYellow
            Write-DzDebug ("`t[ERROR] Comando: {0}" -f $_.InvocationInfo.Line) -Color DarkYellow
        }

        if ($_.ScriptStackTrace) {
            Write-DzDebug "`t[ERROR] Stack:" -Color DarkGray
            Write-DzDebug $_.ScriptStackTrace -Color DarkGray
        }

        Write-DzDebug "========== FIN ERROR ==========`n" -Color Red

        # Mostrar error al usuario
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "Error ejecutando Chocolatey:`n`n$errorMsg`n`nTipo: $errorType",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        } catch {
            Write-DzDebug ("`t[ERROR] No se pudo mostrar MessageBox: {0}" -f $_) -Color Red
        }

        return -1

    } finally {
        Write-DzDebug "`t[DEBUG] === Entrando a bloque finally ==="
        # Asegurar que el proceso principal no herede códigos de salida de Chocolatey
        try {
            if ($global:LASTEXITCODE -ne 0) {
                Write-DzDebug ("`t[DEBUG] Restableciendo LASTEXITCODE de {0} a 0" -f $global:LASTEXITCODE)
                $global:LASTEXITCODE = 0
            }
        } catch {
            Write-DzDebug ("`t[WARN] No se pudo restablecer LASTEXITCODE: {0}" -f $_) -Color DarkYellow
        }
        # Limpiar proceso
        if ($null -ne $process) {
            try {
                Write-DzDebug "`t[DEBUG] Limpiando proceso..."

                if (-not $process.HasExited) {
                    Write-DzDebug "`t[DEBUG] Proceso aún vivo, terminando forzadamente..."
                    try {
                        $process.Kill()
                        $process.WaitForExit(2000)
                    } catch {
                        Write-DzDebug ("`t[WARN] No se pudo terminar proceso: {0}" -f $_) -Color DarkYellow
                    }
                }

                try {
                    $process.Dispose()
                    Write-DzDebug "`t[DEBUG] Proceso disposed"
                } catch {
                    Write-DzDebug ("`t[WARN] Error en Dispose: {0}" -f $_) -Color DarkYellow
                }

            } catch {
                Write-DzDebug ("`t[WARN] Error general limpiando proceso: {0}" -f $_) -Color DarkYellow
            }
        }

        Write-DzDebug "`t[DEBUG] === Finally completado ==="
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