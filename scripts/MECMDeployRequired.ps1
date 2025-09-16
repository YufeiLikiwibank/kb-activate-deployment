param(
  [switch]$DryRun,
  [Parameter(Mandatory)][ValidateSet('KB','BB')]$ReleaseTo,
  [Parameter(Mandatory)][string]$SoftwareVersion,
  [Parameter(Mandatory)][string]$ReleaseRing,
  [Parameter(Mandatory)][string]$SCCMSiteServer,
  [Parameter(Mandatory)][string]$SCCMSiteCode,
  [Parameter(Mandatory)][string]$TargetCollection,   # scheduled device collection
  [Parameter(Mandatory)][string]$AvailableDateTime,  # 'yyyy-MM-dd HH:mm'
  [Parameter(Mandatory)][string]$DeadlineDateTime    # 'yyyy-MM-dd HH:mm'
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
Write-Host "Required deploy: $appName -> $TargetCollection (Avail=$AvailableDateTime, Deadline=$DeadlineDateTime)"
if ($DryRun) { return }

Enter-CMSite -Server $SCCMSiteServer -Code $SCCMSiteCode
Start-CMApplicationDeployment `
  -CollectionName $TargetCollection `
  -ApplicationName $appName `
  -DeployAction Install `
  -DeployPurpose Required `
  -UserNotification HideAll `
  -AvailableDateTime ([datetime]::Parse($AvailableDateTime)) `
  -DeadlineDateTime ([datetime]::Parse($DeadlineDateTime)) `
  -TimeBaseOn LocalTime `
  -EnableMomAlert:$false -GenerateScomAlert:$false -UseMeteredNetwork:$false | Out-Null
