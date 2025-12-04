@{
    # Versión mínima de PowerShell pensada para DESARROLLO
    # (Tu código lo hacemos compatible con 5.0, pero puedes probar en algo más nuevo)
    PowerShellVersion = '5.0'

    # Módulos necesarios para trabajar el proyecto
    Modules           = @(
        @{
            Name           = 'PSScriptAnalyzer'
            MinimumVersion = '1.22.0'
        }
        @{
            Name           = 'Pester'
            MinimumVersion = '5.7.1'
        }
    )
}