# tests/Basic.Tests.ps1

Describe "Pruebas básicas del proyecto" {

    Context "Verificación de archivos" {

        It "Debe existir el archivo main.ps1" {
            Test-Path ".\src\main.ps1" | Should -BeTrue
        }

        It "Deben existir los módulos principales" {
            $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Installers.psm1")

            foreach ($module in $modules) {
                Test-Path ".\src\modules\$module" | Should -BeTrue
            }
        }

        It "Debe existir el archivo version.json" {
            Test-Path ".\src\version.json" | Should -BeTrue
        }
    }

    Context "Verificación de formato" {

        It "version.json debe ser JSON válido" {
            { Get-Content ".\src\version.json" -Raw | ConvertFrom-Json } | Should -Not -Throw
        }

        It "version.json debe tener propiedad Version" {
            $json = Get-Content ".\src\version.json" -Raw | ConvertFrom-Json
            $json.Version | Should -Not -BeNullOrEmpty
        }
    }

    Context "Compatibilidad PowerShell 5" {

        It "Los archivos no deben usar características de PS6+" {
            # Busca '#requires -Version 6+' en los scripts dentro de src
            $matches = Get-ChildItem -Path ".\src" -Recurse -Filter "*.ps*1" |
            Select-String -Pattern '#requires.*-Version.*([6-9]|\d{2,})'

            $matches | Should -BeNullOrEmpty
        }
    }
}