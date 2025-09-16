[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$SourceRepo,            # UNC root
  [Parameter(Mandatory)][string]$OutputDir,             # e.g. $(Build.SourcesDirectory)\drop
  [Parameter()][string]$MsiFileName = '',               # exact name (optional)
  [Parameter()][string]$MsiPattern = 'Activate.*\.msi', # regex-ish
  [Parameter()][ValidateSet('x86','x64','any')][string]$PreferArch = 'any'
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($SourceRepo)) { throw "SourceRepo is empty." }
if (-not (Test-Path $SourceRepo)) { throw "SourceRepo not found: $SourceRepo" }

$archRegex = switch ($PreferArch) {
  'x64' { '(?i)(x64|amd64|win64|64-bit)' }
  'x86' { '(?i)(x86|win32|32-bit)' }
  default { '' }
}

$items = Get-ChildItem -Path $SourceRepo -Filter *.msi -Recurse
if ($MsiPattern) { $items = $items | Where-Object { $_.Name -match $MsiPattern } }
if ($archRegex -and -not $MsiFileName) { $items = $items | Where-Object { $_.Name -match $archRegex } }

$msi = $null
if ($MsiFileName) {
  $cand = Join-Path $SourceRepo $MsiFileName
  if (-not (Test-Path $cand)) { throw "Specified MsiFileName not found: $cand" }
  $msi = Get-Item $cand
} else {
  $msi = $items | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

if (-not $msi) { throw "No MSI found after filtering (pattern='$MsiPattern', arch='$PreferArch')." }

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
Copy-Item -Path $msi.FullName -Destination $OutputDir -Force

$dest = Join-Path $OutputDir $msi.Name
Write-Host "Selected MSI: $($msi.FullName)"
Write-Host "Copied to: $dest"
Write-Host "##vso[task.setvariable variable=MsiLocalPath;isOutput=true]$dest"
Write-Host "##vso[task.setvariable variable=MsiFileName;isOutput=true]$($msi.Name)"
