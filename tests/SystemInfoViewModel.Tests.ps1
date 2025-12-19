# tests/SystemInfoViewModel.Tests.ps1

Import-Module "$PSScriptRoot/../src/modules/Services.psm1" -Force
Import-Module "$PSScriptRoot/../src/modules/ViewModels.psm1" -Force

class FakeClipboardService : IClipboardService {
    [string]$Text
    [void] SetText([string]$Text) {
        $this.Text = $Text
    }
}

class FakeNetworkProfileService : INetworkProfileService {
    [int]$CallCount = 0
    [void] SetAllNetworksPrivate() {
        $this.CallCount++
    }
}

Describe "SystemInfoViewModel" {
    Context "Copiar información del sistema" {
        It "Copia el puerto del texto de etiqueta" {
            $clipboard = [FakeClipboardService]::new()
            $network = [FakeNetworkProfileService]::new()
            $viewModel = [SystemInfoViewModel]::new($clipboard, $network)

            $result = $viewModel.CopyPortFromLabelText("Puerto SQL NationalSoft: 1433")

            $result.Success | Should -BeTrue
            $clipboard.Text | Should -Be "1433"
        }

        It "Devuelve advertencia si el puerto no está presente" {
            $clipboard = [FakeClipboardService]::new()
            $network = [FakeNetworkProfileService]::new()
            $viewModel = [SystemInfoViewModel]::new($clipboard, $network)

            $result = $viewModel.CopyPortFromLabelText("Sin puerto")

            $result.Success | Should -BeFalse
            $result.Level | Should -Be "Warning"
            $clipboard.Text | Should -BeNullOrEmpty
        }

        It "Copia hostname si hay texto" {
            $clipboard = [FakeClipboardService]::new()
            $network = [FakeNetworkProfileService]::new()
            $viewModel = [SystemInfoViewModel]::new($clipboard, $network)

            $result = $viewModel.CopyHostname("EQUIPO-TEST")

            $result.Success | Should -BeTrue
            $clipboard.Text | Should -Be "EQUIPO-TEST"
        }
    }

    Context "Redes" {
        It "Llama al servicio para marcar redes privadas" {
            $clipboard = [FakeClipboardService]::new()
            $network = [FakeNetworkProfileService]::new()
            $viewModel = [SystemInfoViewModel]::new($clipboard, $network)

            $result = $viewModel.SetNetworksPrivate()

            $result.Success | Should -BeTrue
            $network.CallCount | Should -Be 1
        }
    }
}
