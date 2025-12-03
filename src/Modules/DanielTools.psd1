@{
    RootModule        = ''
    ModuleVersion     = '1.0.0'
    GUID              = '5f6b419d-7b77-4c3f-9b12-3b2e36c4176f'
    Author            = 'Daniel Tools'
    CompanyName       = 'Daniel Tools'
    Copyright         = '(c) Daniel Tools. All rights reserved.'
    Description       = 'Módulos auxiliares para inicialización, UI y lógica de base de datos de Daniel Tools.'
    PowerShellVersion = '5.0'
    RequiredModules   = @()
    ScriptsToProcess  = @()
    NestedModules     = @(
        'Init.psm1',
        'UiFactory.psm1',
        'Database.psm1'
    )
    FunctionsToExport = @()
    CmdletsToExport   = @()
    VariablesToExport = '*'
    AliasesToExport   = '*'
    FileList          = @('Init.psm1','UiFactory.psm1','Database.psm1')
}
