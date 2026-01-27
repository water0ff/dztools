$script:SqlEditorAssemblyLoaded = $false
$script:SqlEditorHighlighting = $null
function Get-SqlEditorPaths {
    $moduleRoot = Split-Path -Parent $PSScriptRoot
    $assemblyPath = Join-Path $moduleRoot "lib" "AvalonEdit.dll"
    $highlightingPath = Join-Path $moduleRoot "resources" "SQL.xshd"
    return [pscustomobject]@{
        AssemblyPath     = $assemblyPath
        HighlightingPath = $highlightingPath
    }
}
function Import-AvalonEditAssembly {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AssemblyPath
    )
    if ($script:SqlEditorAssemblyLoaded) { return }
    if (-not (Test-Path -LiteralPath $AssemblyPath)) {
        throw "No se encontró AvalonEdit.dll en '$AssemblyPath'."
    }
    Add-Type -Path $AssemblyPath
    $script:SqlEditorAssemblyLoaded = $true
}
function Get-SqlEditorHighlighting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HighlightingPath
    )
    if ($script:SqlEditorHighlighting) { return $script:SqlEditorHighlighting }
    if (-not (Test-Path -LiteralPath $HighlightingPath)) { return $null }
    $reader = [System.Xml.XmlReader]::Create($HighlightingPath)
    try {
        $script:SqlEditorHighlighting = [ICSharpCode.AvalonEdit.Highlighting.Xshd.HighlightingLoader]::Load(
            $reader,
            [ICSharpCode.AvalonEdit.Highlighting.HighlightingManager]::Instance
        )
    } finally {
        $reader.Close()
    }
    return $script:SqlEditorHighlighting
}
function New-SqlEditor {
    [CmdletBinding()]
    param(
        [Parameter()][System.Windows.Controls.Border]$Container,
        [string]$FontFamily = "Consolas",
        [int]$FontSize = 12
    )
    $paths = Get-SqlEditorPaths
    Import-AvalonEditAssembly -AssemblyPath $paths.AssemblyPath
    $editor = New-Object ICSharpCode.AvalonEdit.TextEditor
    $editor.ShowLineNumbers = $true
    $editor.FontFamily = $FontFamily
    $editor.FontSize = $FontSize
    $editor.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $editor.Options.ConvertTabsToSpaces = $false
    $editor.SyntaxHighlighting = Get-SqlEditorHighlighting -HighlightingPath $paths.HighlightingPath
    if ($Container) {
        $Container.Child = $editor
    }
    return $editor
}
function Set-SqlEditorText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Editor,
        [Parameter(Mandatory)][string]$Text
    )
    if (-not $Editor) { return }
    $Editor.Text = $Text
}
function Get-SqlEditorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return "" }
    return [string]$Editor.Text
}
function Clear-SqlEditorText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return }
    $Editor.Clear()
}
function Insert-SqlEditorText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Editor,
        [Parameter(Mandatory)][string]$Text
    )
    if (-not $Editor) { return }
    $offset = $Editor.CaretOffset
    $Editor.Document.Insert($offset, $Text)
    $Editor.CaretOffset = $offset + $Text.Length
}
function Get-SqlEditorSelectedText {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Editor)
    if (-not $Editor) { return "" }
    return [string]$Editor.SelectedText
}
Export-ModuleMember -Function @(
    'Get-SqlEditorPaths',
    'Import-AvalonEditAssembly',
    'Get-SqlEditorHighlighting',
    'New-SqlEditor',
    'Set-SqlEditorText',
    'Get-SqlEditorText',
    'Clear-SqlEditorText',
    'Insert-SqlEditorText',
    'Get-SqlEditorSelectedText'
)
