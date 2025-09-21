param(
    [Parameter(Mandatory = $true)][string]$Folder,
    [string]$Filter = '*.msi',
    [string]$VariableName = 'revisionVersion',
    [switch]$SetBuildNumber
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Folder)) {
    throw "MSI folder not found: $Folder"
}

$msi = Get-ChildItem -LiteralPath $Folder -Filter $Filter -File |
       Sort-Object LastWriteTime -Descending |
       Select-Object -First 1

if (-not $msi) {
    throw "No MSI matching '$Filter' found in '$Folder'."
}

# Use Windows Installer COM to read ProductVersion
$installer = New-Object -ComObject WindowsInstaller.Installer
$db = $installer.OpenDatabase($msi.FullName, 0)
$view = $db.OpenView("SELECT Value FROM Property WHERE Property='ProductVersion'")
$view.Execute()
$rec = $view.Fetch()
if (-not $rec) { throw "ProductVersion not found in MSI: $($msi.FullName)" }
$version = $rec.StringData(1)

Write-Host "Found MSI: $($msi.Name)"
Write-Host "ProductVersion: $version"

# Set variables for pipeline use
Write-Host "##vso[task.setvariable variable=$VariableName]$version"
Write-Host "##vso[task.setvariable variable=$VariableName;isOutput=true]$version"

if ($SetBuildNumber) {
    Write-Host "##vso[build.updatebuildnumber]$($msi.BaseName) $version"
}
