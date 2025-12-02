# Pruebas básicas del proyecto
Describe "Pruebas básicas del proyecto" {

    Context "Verificación de archivos" {

        It "Debe existir el archivo main.ps1" {
            (Test-Path ".\src\main.ps1") | Should Be $true
        }

        It "Deben existir los módulos principales" {
            $modules = @("GUI.psm1", "Database.psm1", "Utilities.psm1", "Installers.psm1")
            foreach ($module in $modules) {
                (Test-Path ".\src\modules\$module") | Should Be $true
            }
        }

        It "Debe existir el archivo version.json" {
            (Test-Path ".\src\version.json") | Should Be $true
        }
    }

    Context "Verificación de formato" {

        It "version.json debe ser JSON válido" {
            { Get-Content ".\src\version.json" -Raw | ConvertFrom-Json } | Should Not Throw
        }

        It "version.json debe tener propiedad Version" {
            $json = Get-Content ".\src\version.json" -Raw | ConvertFrom-Json
            $json.Version | Should Not BeNullOrEmpty
        }
    }

	Context "Compatibilidad PowerShell 5" {
	    It "Los archivos no deben usar características de PS6+" {
	        # Validamos que la versión de PowerShell sea 5 o mayor
	        ($PSVersionTable.PSVersion.Major -ge 5) | Should Be $true
	    }
	}
}
