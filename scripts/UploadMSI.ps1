param(
    [Parameter(Mandatory = $true)][string]$SourceRoot,       # UNC root, no trailing '\'
    [Parameter(Mandatory = $true)][string]$DestinationRoot,  # e.g. \\kbcfgmgresource\apps$\Workstation\Pipeline
    [Parameter(Mandatory = $true)]
    [ValidateSet('KB','BB')]
    [string]$Brand,                                          # KB or BB
    [switch]$KeepExisting                                    # if set, don't clean old MSIs in dest
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceRoot)) {
    throw "Source path not found: $SourceRoot"
}

# Pick brand pattern
$pattern = if ($Brand -eq 'KB') { 'ActivateKB*.msi' } else { 'ActivateBB*.msi' }

# Select newest matching MSI from source
$msi = Get-ChildItem -LiteralPath $SourceRoot -Filter $pattern -File |
       Sort-Object LastWriteTime -Descending |
       Select-Object -First 1

if (-not $msi) { throw "No MSI found in '$SourceRoot' matching '$pattern'." }

# Brand subfolder under destination root
$brandFolderName = "Activate$Brand"
$destFolder = Join-Path -Path $DestinationRoot -ChildPath $brandFolderName
New-Item -ItemType Directory -Path $destFolder -Force | Out-Null

if (-not $KeepExisting) {
    Get-ChildItem -LiteralPath $destFolder -Filter '*.msi' -File -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

Copy-Item -LiteralPath $msi.FullName -Destination $destFolder -Force

Write-Host "Copied: $($msi.FullName)  →  $destFolder"

# Expose outputs for downstream stages if needed
Write-Host "##vso[task.setvariable variable=BrandFolder;isOutput=true]$brandFolderName"
Write-Host "##vso[task.setvariable variable=EffectiveMsiFolder;isOutput=true]$destFolder"
Write-Host "##vso[task.setvariable variable=CopiedMsiName;isOutput=true]$($msi.Name)"
