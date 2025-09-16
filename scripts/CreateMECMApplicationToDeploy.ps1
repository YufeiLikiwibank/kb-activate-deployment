param(
  [switch]$DryRun,
  [Parameter(Mandatory)][ValidateSet('KB','BB')]$ReleaseTo,
  [Parameter(Mandatory)][string]$SoftwareVersion,
  [Parameter(Mandatory)][string]$ReleaseRing,
  [Parameter(Mandatory)][string]$SCCMSiteServer,
  [Parameter(Mandatory)][string]$SCCMSiteCode,
  [Parameter(Mandatory)][string]$ApplicationFolderPath,    # \\...\Applications\Workstation
  [Parameter(Mandatory)][string]$MsiFolderLocation,        # "$(Pipeline.Workspace)\drop"
  [Parameter(Mandatory)][string]$IconRepo,
  [Parameter(Mandatory)][string]$IconFileName,
  [Parameter(Mandatory)][string]$ContentDistributionPoints  # "Corp;VPN"
)

$ErrorActionPreference = 'Stop'

function Enter-CMSite {
  param([string]$Server,[string]$Code)
  $mod = Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) 'ConfigurationManager.psd1'
  if (!(Test-Path $mod)) { throw "ConfigMgr console not found on agent." }
  Import-Module $mod -Force
  if (-not (Get-PSDrive -Name $Code -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $Code -PSProvider CMSite -Root $Server | Out-Null
  }
  Set-Location ($Code + ':')
}

$productBase = if ($ReleaseTo -eq 'KB') { 'Activate KB' } else { 'Activate BB' }
$appName     = "{0} - {1} ({2})" -f $productBase, $SoftwareVersion, $ReleaseRing
$msi         = Get-ChildItem (Join-Path $MsiFolderLocation '*.msi') | Select-Object -First 1
if (-not $msi) { throw "MSI not found in $MsiFolderLocation." }

$destFolder  = Join-Path $ApplicationFolderPath $appName
$iconPath    = Join-Path $IconRepo $IconFileName
$dpGroups    = $ContentDistributionPoints.Split(';') | Where-Object { $_ -and $_.Trim() }

Write-Host "AppName: $appName"
Write-Host "DestFolder: $destFolder"
Write-Host "DP Groups: $($dpGroups -join ', ')"
if ($DryRun) { return }

# Copy MSI into repo folder structure
if (!(Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }
Copy-Item $msi.FullName -Destination $destFolder -Force
$msiInRepo = Join-Path $destFolder $msi.Name

# Connect to CM site
Enter-CMSite -Server $SCCMSiteServer -Code $SCCMSiteCode

# Create / reuse application and MSI deployment type
$app = Get-CMApplication -Name $appName -Fast -ErrorAction SilentlyContinue
if (-not $app) {
  New-CMApplication -Name $appName -Publisher 'Kiwibank' -SoftwareVersion $SoftwareVersion | Out-Null
  Add-CMMsiDeploymentType -ApplicationName $appName -DeploymentTypeName "$appName - MSI (System)" `
    -InstallationBehaviorType InstallForSystem -MsiInstaller -ContentLocation $destFolder -MsiFilePath $msiInRepo `
    -LogonRequirementType WhetherOrNotUserLoggedOn -UserInteractionMode Hidden | Out-Null
}

# Software Center branding
if (Test-Path $iconPath) {
  Set-CMApplication -Name $appName -LocalizedApplicationName $appName -IconLocationFile $iconPath | Out-Null
} else {
  Set-CMApplication -Name $appName -LocalizedApplicationName $appName | Out-Null
}

# Distribute content
foreach ($g in $dpGroups) {
  Start-CMContentDistribution -ApplicationName $appName -DistributionPointGroupName $g | Out-Null
}
