#requires -Version 5.0
function Show-DllRegistrationDialog {
    [CmdletBinding()]
    param()

    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] INICIO"
    $theme = Get-DzUiTheme
    $defaultList = @(
        "C:\Windows\SysWOW64\slicensing.dll"
        "C:\Windows\SysWOW64\slicensingr1.dll"
        "C:\Windows\SysWOW64\sservices1_0.dll"
        "C:\Windows\SysWOW64\sservices1_0r1.dll"
    ) -join "`r`n"

    $stringXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Registro de DLLs"
        Height="520" Width="760"
        WindowStartupLocation="CenterOwner">
    <Window.Resources>
        <Style x:Key="SystemButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource AccentBlue}"/>
            <Setter Property="Foreground" Value="{DynamicResource OnAccentFg}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrushColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    <Grid Background="{DynamicResource FormBg}" Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0"
                   Text="Escribe las rutas a registrar o deregistrar (una por línea)."
                   Margin="0,0,0,6"
                   FontWeight="SemiBold"/>
        <TextBlock Grid.Row="1"
                   Text="Nota: si un renglón empieza con #, se omitirá."
                   Foreground="{DynamicResource AccentMuted}"
                   Margin="0,0,0,8"/>
        <RichTextBox Grid.Row="2" Name="rtbDlls"
                     VerticalScrollBarVisibility="Auto"
                     Background="{DynamicResource ControlBg}"
                     Foreground="{DynamicResource ControlFg}"
                     BorderBrush="{DynamicResource BorderBrushColor}"
                     BorderThickness="1"/>
        <Grid Grid.Row="3" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <RadioButton Name="rbRegister" Content="Registrar (regsvr32)" IsChecked="True" Margin="0,0,16,0"/>
                <RadioButton Name="rbUnregister" Content="Deregistrar (regsvr32 /u)"/>
            </StackPanel>
            <Button Grid.Column="1" Name="btnRun" Content="Ejecutar" Width="120" Height="32" Margin="0,0,10,0" Style="{StaticResource SystemButtonStyle}"/>
            <Button Grid.Column="2" Name="btnClose" Content="Cerrar" Width="110" Height="32" Style="{StaticResource SystemButtonStyle}"/>
        </Grid>
    </Grid>
</Window>
"@
    try {
        $ui = New-WpfWindow -Xaml $stringXaml -PassThru
        $w = $ui.Window
        $c = $ui.Controls
        Set-DzWpfThemeResources -Window $w -Theme $theme
        try { Set-WpfDialogOwner -Dialog $w } catch {}

        $rtb = $c['rtbDlls']
        $textRange = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
        $textRange.Text = $defaultList
        $script:dllFormatting = $false

        function Get-TextPointerFromOffset {
            param([System.Windows.Documents.FlowDocument]$Document, [int]$Offset)
            $navigator = $Document.ContentStart
            $count = 0
            while ($navigator -ne $null) {
                $context = $navigator.GetPointerContext([System.Windows.Documents.LogicalDirection]::Forward)
                if ($context -eq [System.Windows.Documents.TextPointerContext]::Text) {
                    $text = $navigator.GetTextInRun([System.Windows.Documents.LogicalDirection]::Forward)
                    if (($count + $text.Length) -ge $Offset) {
                        return $navigator.GetPositionAtOffset($Offset - $count)
                    }
                    $count += $text.Length
                    $navigator = $navigator.GetPositionAtOffset($text.Length)
                } else {
                    $navigator = $navigator.GetNextContextPosition([System.Windows.Documents.LogicalDirection]::Forward)
                }
            }
            return $Document.ContentEnd
        }

        $applyHighlight = {
            if ($script:dllFormatting) { return }
            $script:dllFormatting = $true
            try {
                $caretOffset = $rtb.Document.ContentStart.GetOffsetToPosition($rtb.CaretPosition)
                $rawRange = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
                $rawText = ($rawRange.Text -replace "`r", "")
                $rtb.Document.Blocks.Clear()
                foreach ($line in $rawText -split "`n", -1) {
                    $paragraph = New-Object System.Windows.Documents.Paragraph
                    $paragraph.Margin = "0"
                    $run = New-Object System.Windows.Documents.Run($line)
                    if ($line.TrimStart().StartsWith("#")) { $run.Foreground = [System.Windows.Media.Brushes]::Gray }
                    $paragraph.Inlines.Add($run)
                    $rtb.Document.Blocks.Add($paragraph)
                }
                $pointer = Get-TextPointerFromOffset -Document $rtb.Document -Offset $caretOffset
                if ($pointer) { $rtb.CaretPosition = $pointer }
            } finally {
                $script:dllFormatting = $false
            }
        }.GetNewClosure()
        & $applyHighlight
        $rtb.Add_TextChanged({ & $applyHighlight })

        $c['btnRun'].Add_Click({
                Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] Ejecutando registro"
                if (-not (Test-Administrator)) {
                    Show-WpfMessageBox -Message "Esta acción requiere permisos de administrador." -Title "Permisos requeridos" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                $range = New-Object System.Windows.Documents.TextRange($rtb.Document.ContentStart, $rtb.Document.ContentEnd)
                $rawText = ($range.Text -replace "`r", "")
                $lines = $rawText -split "`n"
                $paths = @()
                $currentDir = $null
                foreach ($line in $lines) {
                    $trim = $line.Trim()
                    if ([string]::IsNullOrWhiteSpace($trim)) { continue }
                    if ($trim.StartsWith("#")) { continue }
                    if (Test-Path -LiteralPath $trim -PathType Container) {
                        $currentDir = $trim
                        continue
                    }
                    $candidate = $trim
                    if (-not [System.IO.Path]::IsPathRooted($candidate) -and $currentDir) {
                        $candidate = Join-Path $currentDir $candidate
                    }
                    $paths += $candidate
                }
                if (-not $paths -or $paths.Count -eq 0) {
                    Show-WpfMessageBox -Message "No se encontraron rutas válidas para procesar." -Title "Sin datos" -Buttons OK -Icon Warning | Out-Null
                    return
                }
                $isUnregister = ($c['rbUnregister'].IsChecked -eq $true)
                $success = @()
                $errors = @()
                foreach ($path in $paths) {
                    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                        $errors += "No existe: $path"
                        continue
                    }
                    $args = if ($isUnregister) { "/s /u `"$path`"" } else { "/s `"$path`"" }
                    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] regsvr32 $args"
                    try {
                        $proc = Start-Process -FilePath "regsvr32.exe" -ArgumentList $args -NoNewWindow -Wait -PassThru -WindowStyle Hidden
                        if ($proc.ExitCode -eq 0) {
                            $success += $path
                        } else {
                            $errors += "Error ($($proc.ExitCode)) en: $path"
                        }
                    } catch {
                        $errors += "Error en $path`: $($_.Exception.Message)"
                    }
                }
                $summary = @()
                if ($success.Count -gt 0) { $summary += "Procesados correctamente: $($success.Count)" }
                if ($errors.Count -gt 0) { $summary += "Errores: $($errors.Count)" }
                $detail = ($summary -join "`n")
                if ($errors.Count -gt 0) { $detail += "`n`nDetalles:`n$($errors -join "`n")" }
                $title = if ($errors.Count -gt 0) { "Proceso con errores" } else { "Proceso completado" }
                $icon = if ($errors.Count -gt 0) { "Warning" } else { "Information" }
                Show-WpfMessageBox -Message $detail -Title $title -Buttons OK -Icon $icon | Out-Null
            })
        $c['btnClose'].Add_Click({ $w.Close() })
        $w.ShowDialog() | Out-Null
    } catch {
        Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] ERROR creando ventana: $($_.Exception.Message)" Red
        Show-WpfMessageBox -Message "No se pudo crear la ventana de registro de DLLs." -Title "Error" -Buttons OK -Icon Error | Out-Null
    }
    Write-DzDebug "`t[DEBUG][Show-DllRegistrationDialog] FIN"
}
Export-ModuleMember -Function @('Show-DllRegistrationDialog')
