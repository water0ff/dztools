# tests/Basic.Tests.ps1

# Raíz del proyecto (carpeta donde está src/)
$projectRoot = Resolve-Path "$PSScriptRoot/.."

Describe "Pruebas básicas del proyecto" {

    Context "Verificación de archivos" {

        It "Debe existir el archivo main.ps1" {
            $path = Join-Path $projectRoot "src/main.ps1"
            (Test-Path $path) | Should -BeTrue
        }

        It "Deben existir los módulos principales" {
            $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Installers.psm1")

            foreach ($module in $modules) {
                $modulePath = Join-Path $projectRoot "src/modules/$module"
                (Test-Path $modulePath) | Should -BeTrue
            }
        }

        It "Debe existir el archivo version.json" {
            $versionPath = Join-Path $projectRoot "src/version.json"
            (Test-Path $versionPath) | Should -BeTrue
        }
    }

    Context "Verificación de formato" {

        $versionPath = Join-Path $projectRoot "src/version.json"

        It "version.json debe ser JSON válido" {
            { Get-Content $versionPath -Raw | ConvertFrom-Json } | Should -Not -Throw
        }

        It "version.json debe tener propiedad Version" {
            $json = Get-Content $versionPath -Raw | ConvertFrom-Json
            $json.Version | Should -Not -BeNullOrEmpty
        }
    }

    Context "Compatibilidad PowerShell 5" {

        It "Los archivos no deben usar características de PS6+" {
            # Buscar #requires -Version 6 o superior en los scripts de src/
            $matches = Get-ChildItem -Path (Join-Path $projectRoot "src") -Recurse -Filter "*.ps*1" |
            Select-String -Pattern '#requires.*-Version.*([6-9]|\d{2,})'

            $matches | Should -BeNullOrEmpty
        }
    }
}