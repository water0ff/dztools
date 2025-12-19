#requires -Version 5.0

interface IClipboardService {
    [void] SetText([string]$Text)
}

interface IMessageService {
    [void] ShowInfo([string]$Message, [string]$Title)
    [void] ShowWarning([string]$Message, [string]$Title)
}

interface INetworkProfileService {
    [void] SetAllNetworksPrivate()
}

class ClipboardService : IClipboardService {
    [void] SetText([string]$Text) {
        [System.Windows.Forms.Clipboard]::SetText($Text)
    }
}

class MessageBoxService : IMessageService {
    [void] ShowInfo([string]$Message, [string]$Title) {
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }

    [void] ShowWarning([string]$Message, [string]$Title) {
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

class NetworkProfileService : INetworkProfileService {
    [void] SetAllNetworksPrivate() {
        Get-NetConnectionProfile |
        Where-Object { $_.NetworkCategory -ne 'Private' } |
        ForEach-Object {
            Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private
        }
    }
}
