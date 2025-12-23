#requires -Version 5.0
function Process-SqlProgressMessage {
    param([string]$Message)

    # Extraer porcentaje de mensajes de backup
    if ($Message -match '(\d+) percent processed') {
        $percent = [int]$Matches[1]
        Write-Output "Progreso: $percent%"
    } elseif ($Message -match 'BACKUP DATABASE successfully processed') {
        Write-Output "Backup completado exitosamente"
    }
}
function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $true)]
        [string]$Database,
        [Parameter(Mandatory = $true)]
        [string]$Query,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)]
        [scriptblock]$InfoMessageCallback
    )

    $connection = $null
    Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] INICIO"
    Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Server='$Server' Database='$Database' User='$($Credential.UserName)'"

    try {
        # Convertir secure string a password
        $passwordBstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringUni($passwordBstr)

        # Crear connection string
        $connectionString = "Server=$Server;Database=$Database;User Id=$($Credential.UserName);Password=$plainPassword;MultipleActiveResultSets=True"

        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Creando SqlConnection..."
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)

        # Configurar InfoMessage si se proporciona callback
        if ($InfoMessageCallback) {
            $connection.add_InfoMessage({
                    param($sender, $e)
                    try {
                        & $InfoMessageCallback $e.Message
                    } catch {
                        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Error en InfoMessageCallback: $_"
                    }
                })
            $connection.FireInfoMessageEventOnUserErrors = $true
        }

        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Abriendo conexión..."
        $connection.Open()

        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Estado conexión tras Open(): $($connection.State)"

        # Crear y ejecutar comando
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 0

        # Determinar tipo de consulta
        $returnsResultSet = $query -match "(?si)^\s*(SELECT|WITH)" -or $query -match "(?si)\bOUTPUT\b"

        if ($returnsResultSet) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Ejecutando consulta tipo SELECT/WITH"
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)

            return @{
                Success   = $true
                DataTable = $dataTable
                Type      = "Query"
            }
        } else {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Ejecutando consulta tipo NonQuery"
            $rowsAffected = $command.ExecuteNonQuery()

            return @{
                Success      = $true
                RowsAffected = $rowsAffected
                Type         = "NonQuery"
            }
        }

    } catch {
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] CATCH: $($_.Exception.Message)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Tipo de excepción: $($_.Exception.GetType().FullName)"
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Stack: $($_.ScriptStackTrace)"

        return @{
            Success      = $false
            ErrorMessage = $_.Exception.Message
            Type         = "Error"
        }
    } finally {
        # Limpiar contraseña en memoria
        if ($plainPassword) {
            $plainPassword = $null
        }
        if ($passwordBstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBstr)
        }

        # Cerrar conexión
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] FINALLY: connection = $($connection)"
        if ($null -ne $connection) {
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Estado antes de cerrar: $($connection.State)"
            if ($connection.State -eq [System.Data.ConnectionState]::Open) {
                Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Cerrando conexión..."
                $connection.Close()
            }
            Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] Disposing conexión..."
            $connection.Dispose()
        }
        Write-DzDebug "`t[DEBUG][Invoke-SqlQuery] FIN"
    }
}
function Remove-SqlComments {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Query
    )
    $query = $Query -replace '(?s)/\*.*?\*/', ''
    $query = $query -replace '(?m)^\s*--.*\n?', ''
    $query = $query -replace '(?<!\w)--.*$', ''
    return $query.Trim()
}

function Get-SqlDatabases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    $query = @"
SELECT name
FROM sys.databases
WHERE name NOT IN ('tempdb','model','msdb')
  AND state_desc = 'ONLINE'
ORDER BY CASE WHEN name = 'master' THEN 0 ELSE 1 END, name
"@
    $result = Invoke-SqlQuery -Server $Server -Database "master" -Query $query -Credential $Credential
    if (-not $result.Success) {
        throw "Error obteniendo bases de datos: $($result.ErrorMessage)"
    }
    $databases = @()
    foreach ($row in $result.DataTable.Rows) {
        $databases += $row["name"]
    }
    return $databases
}

function Backup-Database {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $true)]
        [string]$Database,
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        [Parameter(Mandatory = $false)]
        [scriptblock]$ProgressCallback
    )

    try {
        Write-DzDebug "[Backup-Database] Iniciando backup de '$Database' a '$BackupPath'"

        # Crear la consulta de backup con STATS = 1 para más actualizaciones
        $backupQuery = @"
BACKUP DATABASE [$Database]
TO DISK='$BackupPath'
WITH
    CHECKSUM,
    STATS = 1,
    FORMAT,
    INIT
"@

        # Configurar callback para mensajes de SQL
        $infoCallback = {
            param($message)
            if ($ProgressCallback) {
                & $ProgressCallback $message
            }
        }

        # Ejecutar el backup
        $result = Invoke-SqlQuery -Server $Server -Database "master" `
            -Query $backupQuery -Credential $Credential `
            -InfoMessageCallback $infoCallback

        if ($result.Success) {
            Write-DzDebug "[Backup-Database] Backup completado exitosamente"
            return @{
                Success    = $true
                BackupPath = $BackupPath
            }
        } else {
            Write-DzDebug "[Backup-Database] Error en backup: $($result.ErrorMessage)"
            return @{
                Success      = $false
                ErrorMessage = $result.ErrorMessage
            }
        }
    } catch {
        Write-DzDebug "[Backup-Database] Excepción: $($_.Exception.Message)"
        return @{
            Success      = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}
function Execute-SqlQuery {
    param (
        [string]$server,
        [string]$database,
        [string]$query
    )
    try {
        $connectionString = "Server=$server;Database=$database;User Id=$global:user;Password=$global:password;MultipleActiveResultSets=True"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $infoMessages = New-Object System.Collections.ArrayList
        $connection.add_InfoMessage({
                param($sender, $e)
                $infoMessages.Add($e.Message) | Out-Null
            })
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $returnsResultSet = $query -match "(?si)^\s*(SELECT|WITH)" -or $query -match "(?si)\bOUTPUT\b"
        if ($returnsResultSet) {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            $adapter.Fill($dataTable) | Out-Null
            return @{
                DataTable = $dataTable
                Messages  = $infoMessages
            }
        } else {
            $rowsAffected = $command.ExecuteNonQuery()
            return @{
                RowsAffected = $rowsAffected
                Messages     = $infoMessages
            }
        }
    } catch {
        throw $_
    } finally {
        if ($null -ne $connection) {
            if ($connection.State -ne [System.Data.ConnectionState]::Closed) {
                $connection.Close()
            } else {
                $connection.Dispose()
            }
        }
    }
}

function Show-ResultsConsole {
    param (
        [string]$query
    )
    try {
        $results = Execute-SqlQuery -server $global:server -database $global:database -query $query
        if ($results.GetType().Name -eq 'Hashtable') {
            $consoleData = $results.ConsoleData
            if ($consoleData.Count -gt 0) {
                $columns = $consoleData[0].Keys
                $columnWidths = @{}
                foreach ($col in $columns) {
                    $columnWidths[$col] = $col.Length
                }
                Write-Host ""
                $header = ""
                foreach ($col in $columns) {
                    $header += $col.PadRight($columnWidths[$col] + 4)
                }
                Write-Host $header
                Write-Host ("-" * $header.Length)
                foreach ($row in $consoleData) {
                    $rowText = ""
                    foreach ($col in $columns) {
                        $rowText += ($row[$col].ToString()).PadRight($columnWidths[$col] + 4)
                    }
                    Write-Host $rowText
                }
            } else {
                Write-Host "`nNo se encontraron resultados." -ForegroundColor Yellow
            }
        } else {
            Write-Host "`nFilas afectadas: $results" -ForegroundColor Green
        }
    } catch {
        Write-Host "`nError al ejecutar la consulta: $_" -ForegroundColor Red
    }
}

function Get-IniConnections {
    $connections = @()
    $pathsToCheck = @(
        @{ Path = "C:\NationalSoft\Softrestaurant9.5.0Pro"; INI = "restaurant.ini"; Nombre = "SR9.5" },
        @{ Path = "C:\NationalSoft\Softrestaurant10.0"; INI = "restaurant.ini"; Nombre = "SR10" },
        @{ Path = "C:\NationalSoft\Softrestaurant11.0"; INI = "restaurant.ini"; Nombre = "SR11" },
        @{ Path = "C:\NationalSoft\Softrestaurant12.0"; INI = "restaurant.ini"; Nombre = "SR12" },
        @{ Path = "C:\Program Files (x86)\NsBackOffice1.0"; INI = "DbConfig.ini"; Nombre = "NSBackOffice" },
        @{ Path = "C:\NationalSoft\NationalSoftHoteles3.0"; INI = "nshoteles.ini"; Nombre = "Hoteles" },
        @{ Path = "C:\NationalSoft\OnTheMinute4.5"; INI = "checadorsql.ini"; Nombre = "OnTheMinute" }
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

function Load-IniConnectionsToComboBox {
    param(
        [Parameter(Mandatory = $true)]
        $Combo
    )
    $connections = Get-IniConnections
    $Combo.Items.Clear()
    if ($connections.Count -gt 0) {
        foreach ($connection in $connections) {
            $Combo.Items.Add($connection) | Out-Null
        }
    } else {
        Write-Host "`tNo se encontraron conexiones en archivos INI" -ForegroundColor Yellow
    }
    if ($Combo -is [System.Windows.Controls.ComboBox]) {
        $Combo.Text = ".\NationalSoft"
    } else {
        $Combo.Text = ".\NationalSoft"
    }
}

function ConvertTo-DataTable {
    param($InputObject)
    $dt = New-Object System.Data.DataTable
    if (-not $InputObject) { return $dt }
    $columns = $InputObject[0].Keys
    foreach ($col in $columns) {
        $dt.Columns.Add($col) | Out-Null
    }
    foreach ($row in $InputObject) {
        $dr = $dt.NewRow()
        foreach ($col in $columns) {
            $dr[$col] = $row[$col]
        }
        $dt.Rows.Add($dr)
    }
    return $dt
}
function Show-BackupDialog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [string]$Password,
        [Parameter(Mandatory = $true)]
        [string]$Database
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Opciones de Respaldo" Height="500" Width="600"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Respaldar -->
        <CheckBox x:Name="chkRespaldo" Grid.Row="0" IsChecked="True" IsEnabled="False"
                  Margin="0,0,0,10">
            <TextBlock Text="Respaldar" FontWeight="Bold"/>
        </CheckBox>

        <!-- Nombre del respaldo -->
        <Label Grid.Row="1" Content="Nombre del respaldo:"/>
        <TextBox x:Name="txtNombre" Grid.Row="2" Height="25" Margin="0,5,0,10"/>

        <!-- Comprimir -->
        <CheckBox x:Name="chkComprimir" Grid.Row="3" Margin="0,0,0,10">
            <TextBlock Text="Comprimir (requiere Chocolatey)" FontWeight="Bold"/>
        </CheckBox>

        <!-- Contraseña ZIP -->
        <Label x:Name="lblPassword" Grid.Row="4" Content="Contraseña (opcional) para ZIP:"/>
        <Grid Grid.Row="5" Margin="0,5,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <PasswordBox x:Name="txtPassword" Grid.Column="0" Height="25"/>
            <Button x:Name="btnTogglePassword" Grid.Column="1" Content="👁" Width="30" Margin="5,0,0,0"/>
        </Grid>

        <!-- Subir a Mega.nz -->
        <CheckBox x:Name="chkSubir" Grid.Row="6" Margin="0,0,0,20">
            <TextBlock Text="Subir a Mega.nz (requiere MegaTools y Chocolatey)" FontWeight="Bold"/>
        </CheckBox>

        <!-- Área de progreso -->
        <GroupBox Grid.Row="7" Header="Progreso" Margin="0,0,0,10">
            <StackPanel>
                <ProgressBar x:Name="pbBackup" Height="20" Margin="5"
                           Minimum="0" Maximum="100" Value="0"/>
                <TextBlock x:Name="txtProgress" Text="Esperando..."
                          Margin="5,5,5,10" TextWrapping="Wrap"/>
                <TextBlock x:Name="txtTimeElapsed" Text="Tiempo: 00:00:00"
                          Margin="5,0,5,5" Foreground="Gray"/>
            </StackPanel>
        </GroupBox>

        <!-- Log -->
        <GroupBox Grid.Row="8" Header="Log">
            <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Auto" Height="160"/>
        </GroupBox>

        <!-- Botones -->
        <StackPanel Grid.Row="9" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnAceptar" Content="Iniciar Respaldo" Width="120" Height="30" Margin="5,0"/>
            <Button x:Name="btnAbrirCarpeta" Content="Abrir Carpeta" Width="100" Height="30" Margin="5,0"/>
            <Button x:Name="btnCerrar" Content="Cerrar" Width="80" Height="30" Margin="5,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader] $xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)
    $chkRespaldo = $window.FindName("chkRespaldo")
    $txtNombre = $window.FindName("txtNombre")
    $chkComprimir = $window.FindName("chkComprimir")
    $txtPassword = $window.FindName("txtPassword")
    $lblPassword = $window.FindName("lblPassword")
    $chkSubir = $window.FindName("chkSubir")
    $pbBackup = $window.FindName("pbBackup")
    $txtProgress = $window.FindName("txtProgress")
    $txtTimeElapsed = $window.FindName("txtTimeElapsed")
    $txtLog = $window.FindName("txtLog")
    $btnAceptar = $window.FindName("btnAceptar")
    $btnAbrirCarpeta = $window.FindName("btnAbrirCarpeta")
    $btnCerrar = $window.FindName("btnCerrar")
    $btnTogglePassword = $window.FindName("btnTogglePassword")

    # Configurar valores iniciales
    $timestampsDefault = Get-Date -Format 'yyyyMMdd-HHmmss'
    if ($Database) {
        $txtNombre.Text = "$Database-$timestampsDefault.bak"
    } else {
        $txtNombre.Text = "Backup-$timestampsDefault.bak"
    }

    # Deshabilitar opciones avanzadas por defecto
    $txtPassword.IsEnabled = $false
    $lblPassword.IsEnabled = $false

    # Evento para CheckBox Comprimir
    $chkComprimir.Add_Checked({
            $txtPassword.IsEnabled = $true
            $lblPassword.IsEnabled = $true
        })

    $chkComprimir.Add_Unchecked({
            $txtPassword.IsEnabled = $false
            $lblPassword.IsEnabled = $false
            $txtPassword.Password = ""
            $chkSubir.IsChecked = $false
        })
    function Add-Log {
        param([string]$Message)
        $timestamp = Get-Date -Format 'HH:mm:ss'
        # Insertar al principio en lugar de al final
        $txtLog.Text = "$timestamp $Message`n" + $txtLog.Text
        # Mantener el scroll arriba
        $txtLog.ScrollToLine(0)
    }
    function Update-Progress {
        param([int]$Percent, [string]$Message)
        # Usar Dispatcher para actualizar elementos de UI desde otro hilo
        $window.Dispatcher.Invoke([Action] {
                $pbBackup.Value = $Percent
                $txtProgress.Text = $Message
            })
    }
    function New-SafeCredential {
        param([string]$Username, [string]$PlainPassword)

        # Convertir string a SecureString manualmente
        $securePassword = New-Object System.Security.SecureString
        $chars = $PlainPassword.ToCharArray()
        foreach ($char in $chars) {
            $securePassword.AppendChar($char)
        }
        $securePassword.MakeReadOnly()

        # Crear credencial
        New-Object System.Management.Automation.PSCredential($Username, $securePassword)
    }
    $btnAceptar.Add_Click({
            try {
                Write-DzDebug "[btnAceptar_Click] Inicio del clic"
                $btnAceptar.IsEnabled = $false
                $btnAceptar.Content = "Procesando..."
                $txtLog.Text = ""
                Add-Log -Message "Iniciando proceso de backup..."

                # Reiniciar barra de progreso
                $pbBackup.Value = 0
                $txtProgress.Text = "Esperando..."

                # Iniciar temporizador
                $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                $timer = [System.Windows.Threading.DispatcherTimer]::new()
                $timer.Interval = [TimeSpan]::FromSeconds(1)
                $timer.Add_Tick({
                        $elapsed = $stopwatch.Elapsed
                        $txtTimeElapsed.Text = "Tiempo: $($elapsed.ToString('hh\:mm\:ss'))"
                    })
                $timer.Start()

                # Mostrar parámetros
                Add-Log -Message "Servidor: $Server"
                Add-Log -Message "Base de datos: $Database"
                Add-Log -Message "Usuario: $User"

                # Determinar ruta LOCAL (no usar red para el archivo .bak)
                $machinePart = $Server.Split('\')[0]
                $machineName = $machinePart.Split(',')[0]
                if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }
                $backupFolder = "C:\Temp\SQLBackups"
                $backupFileName = $txtNombre.Text
                $backupPath = Join-Path $backupFolder $backupFileName

                Add-Log -Message "Ruta de backup: $backupPath"
                Add-Log -Message "Creando carpeta si no existe..."
                if (-not (Test-Path $backupFolder)) {
                    New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null
                }

                # Crear credencial (usando tu función segura)
                Add-Log -Message "Creando credenciales..."
                $credential = New-SafeCredential -Username $User -PlainPassword $Password
                Add-Log -Message "✓ Credenciales listas"

                # ========== EJECUCIÓN DIRECTA SIN JOB ==========
                Add-Log -Message "Ejecutando comando BACKUP directamente..."
                Update-Progress -Percent 5 -Message "Conectando a SQL Server..."

                # 1. Construir la consulta de backup
                $backupQuery = @"
BACKUP DATABASE [$Database]
TO DISK = '$backupPath'
WITH
    CHECKSUM,
    STATS = 1,  -- Cambiado a 1 para más actualizaciones frecuentes
    FORMAT,
    INIT
"@

                # Definir el callback para manejar mensajes de progreso
                $progressCallback = {
                    param([string]$Message)

                    # Procesar mensajes de progreso
                    if ($Message -match '(\d+) percent processed') {
                        $percent = [int]$Matches[1]
                        # Actualizar la barra de progreso
                        $window.Dispatcher.Invoke([Action] {
                                $pbBackup.Value = $percent
                                $txtProgress.Text = "Progreso: $percent%"
                                Add-Log -Message "Progreso backup: $percent%"
                            })
                    } elseif ($Message -match 'BACKUP DATABASE successfully processed') {
                        $window.Dispatcher.Invoke([Action] {
                                Update-Progress -Percent 100 -Message "¡Backup completado!"
                                Add-Log -Message "✅ BACKUP DATABASE successfully processed"
                            })
                    } elseif ($Message -match '^Processed') {
                        # Solo log, sin actualizar progreso
                        $window.Dispatcher.Invoke([Action] {
                                Add-Log -Message $Message
                            })
                    }
                }

                # 2. Ejecutar DIRECTAMENTE usando Invoke-SqlQuery
                Add-Log -Message "Enviando comando a SQL Server..."
                Update-Progress -Percent 10 -Message "Iniciando backup..."

                $result = Invoke-SqlQuery -Server $Server -Database "master" `
                    -Query $backupQuery -Credential $credential `
                    -InfoMessageCallback $progressCallback

                # 3. Procesar resultado
                if ($result.Success) {
                    Update-Progress -Percent 100 -Message "¡Backup completado!"

                    # Detener temporizador
                    $timer.Stop()
                    $stopwatch.Stop()
                    $txtTimeElapsed.Text = "Tiempo: $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"

                    Add-Log -Message "✅ BACKUP DATABASE successfully processed"

                    # Verificar archivo creado
                    Start-Sleep -Milliseconds 500  # Esperar un poco para que el archivo se escriba
                    if (Test-Path $backupPath) {
                        $sizeMB = [math]::Round((Get-Item $backupPath).Length / 1MB, 2)
                        Add-Log -Message "📊 Tamaño del archivo: $sizeMB MB"
                        Add-Log -Message "📁 Ubicación: $backupPath"
                    } else {
                        Add-Log -Message "⚠️ Comando exitoso pero no se encontró el archivo en la ruta esperada."
                    }
                } else {
                    Update-Progress -Percent 0 -Message "Error en backup"
                    Add-Log -Message "❌ Error de SQL: $($result.ErrorMessage)"
                }

            } catch {
                Add-Log -Message "❌ Error inesperado: $($_.Exception.Message)"
                Add-Log -Message "Detalle: $($_.ScriptStackTrace)"
                Update-Progress -Percent 0 -Message "Error"
            } finally {
                $btnAceptar.IsEnabled = $true
                $btnAceptar.Content = "Iniciar Respaldo"
                # Detener temporizador si aún está corriendo
                if ($timer -and $timer.IsEnabled) {
                    $timer.Stop()
                }
            }
        })
    $btnAbrirCarpeta.Add_Click({
            $machinePart = $Server.Split('\')[0]
            $machineName = $machinePart.Split(',')[0]
            if ($machineName -eq '.') { $machineName = $env:COMPUTERNAME }

            $backupFolder = "\\$machineName\C$\Temp\SQLBackups"

            if (Test-Path $backupFolder) {
                Start-Process explorer.exe $backupFolder
            } else {
                [System.Windows.MessageBox]::Show(
                    "La carpeta de respaldos no existe todavía.`n`nRuta: $backupFolder",
                    "Atención",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
            }
        })

    $btnCerrar.Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })

    # Mostrar la ventana
    $null = $window.ShowDialog()
}
Export-ModuleMember -Function @(
    'Invoke-SqlQuery',
    'Remove-SqlComments',
    'Get-SqlDatabases',
    'Backup-Database',
    'Execute-SqlQuery',
    'Show-ResultsConsole',
    'Get-IniConnections',
    'Load-IniConnectionsToComboBox',
    'ConvertTo-DataTable',
    'Show-BackupDialog'
)
