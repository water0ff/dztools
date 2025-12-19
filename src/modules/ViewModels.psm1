#requires -Version 5.0

class OperationResult {
    [bool]$Success
    [string]$Message
    [string]$Level

    OperationResult([bool]$Success, [string]$Message, [string]$Level) {
        $this.Success = $Success
        $this.Message = $Message
        $this.Level = $Level
    }

    static [OperationResult] Success([string]$Message) {
        return [OperationResult]::new($true, $Message, 'Info')
    }

    static [OperationResult] Warning([string]$Message) {
        return [OperationResult]::new($false, $Message, 'Warning')
    }

    static [OperationResult] Error([string]$Message) {
        return [OperationResult]::new($false, $Message, 'Error')
    }
}

class SystemInfoViewModel {
    [IClipboardService]$ClipboardService
    [INetworkProfileService]$NetworkProfileService

    SystemInfoViewModel(
        [IClipboardService]$ClipboardService,
        [INetworkProfileService]$NetworkProfileService
    ) {
        $this.ClipboardService = $ClipboardService
        $this.NetworkProfileService = $NetworkProfileService
    }

    [OperationResult] CopyPortFromLabelText([string]$LabelText) {
        if ([string]::IsNullOrWhiteSpace($LabelText)) {
            return [OperationResult]::Warning("El texto del Label del puerto está vacío.")
        }
        $port = [regex]::Match($LabelText, '\d+').Value
        if ([string]::IsNullOrWhiteSpace($port)) {
            return [OperationResult]::Warning("El texto del Label del puerto no contiene un número válido para copiar.")
        }
        $this.ClipboardService.SetText($port)
        return [OperationResult]::Success("Puerto copiado al portapapeles: $port")
    }

    [OperationResult] CopyHostname([string]$HostnameText) {
        if ([string]::IsNullOrWhiteSpace($HostnameText)) {
            return [OperationResult]::Warning("El nombre de equipo está vacío, no hay nada que copiar.")
        }
        $this.ClipboardService.SetText($HostnameText)
        return [OperationResult]::Success("Nombre del equipo copiado al portapapeles: $HostnameText")
    }

    [OperationResult] CopyIpAddresses([string]$IpText) {
        if ([string]::IsNullOrWhiteSpace($IpText)) {
            return [OperationResult]::Warning("No hay IPs para copiar.")
        }
        $this.ClipboardService.SetText($IpText)
        return [OperationResult]::Success("IP's copiadas al portapapeles: $IpText")
    }

    [OperationResult] SetNetworksPrivate() {
        try {
            $this.NetworkProfileService.SetAllNetworksPrivate()
            return [OperationResult]::Success("Todas las redes se han establecido como Privadas.")
        } catch {
            return [OperationResult]::Error("Error al establecer redes privadas: $_")
        }
    }
}
