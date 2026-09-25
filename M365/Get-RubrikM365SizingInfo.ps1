<#
.SYNOPSIS
    Get-RubrikM365SizingInfo.ps1 returns M365 usage information for a subscription.
.DESCRIPTION
    Get-RubrikM365SizingInfo.ps1 returns M365 usage information for a subscprtion.
    Data is gathered using the Microsoft Graph APIs and Exchange module.

    The usage data should be similar to the metrics in the Admin Center under each
    workload's "Usage Reports" data. There could be some discrepency between the
    "Total Users" shown in the Usage Charts and what the script gathers because the
    script uses the detailed user .CSV and summarizes that information.

    The M365 Usage Reports do not contain information on Exchange In Place Archives.
    By default, the script will try to gather this information by looping through
    each user that has an In Place Archive and gathering that info directly. Unfortunately,
    if there are a lot of users, this may time out the script. If that is the case,
    you can use the flag to skip gathering in place archive data and try to provide
    an estimate.

    Each In Place Archive mailbox is retried a few times when a transient error such as
    throttling or token expiry is hit. Mailboxes that still cannot be read are counted and
    written to an ArchiveFailures CSV, and the report states that the archive totals are
    incomplete, rather than dropping those mailboxes from the totals silently. Progress is
    written to a checkpoint file as it is gathered so an interrupted run can be resumed
    with -ResumeArchive instead of starting over.

    If the aggregate archive statistics call still fails after exhausting its retries, the
    script falls back to rolling the mailbox's size up from its per-folder statistics instead -
    a different EXO API path that can succeed when the aggregate call keeps timing out or being
    throttled. Use -SkipArchiveFallback $true to disable this and only ever use the aggregate call.

    The M365 Usage Reports do not contain information on Exchange Recoverable Items Folder.
    By default, the script will try to gather this information by looping through
    every user and gathering that info directly. Unfortunately,
    if there are a lot of users, this may time out the script. If that is the case,
    you can use the flag to skip gathering Recoverable Items data and try to provide
    an estimate.

    Each Recoverable Items mailbox is retried a few times when a transient error such as
    throttling or token expiry is hit. Mailboxes that still cannot be read are counted and
    written to a RIFFailures CSV, and the report states that the Recoverable Items totals
    are incomplete, rather than dropping those mailboxes from the totals silently. Progress
    is written to a checkpoint file as it is gathered so an interrupted run can be resumed
    with -ResumeRIF instead of starting over.

.EXAMPLE
    PS C:\> .\Get-RubrikM365SizingInfo.ps1
    Opens a browser window to authenticate to M365 Graph APIs and Microsoft Exchange
    Module to pull usage information.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -UseAppAccess $true
    Prompts the user to enter tenant ID, app client ID and app client secret to
    authenticate to M365 Graph APIs and Microsoft Exchange Module to pull usage
    information.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -AnnualGrowth 35
    Calculate sizing to include a 35% annual growth rate

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -SkipArchiveMailbox $true
    Skip gathering In Place Archive mailboxes.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -SkipSharedMailbox $true
    Skip gathering Shared mailboxes.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -SkipRecoverableItems $true
    Skip gathering Recoverable Items hierarchy.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -ResumeArchive $true
    Resume In Place Archive gathering from the checkpoint file left behind by an
    interrupted run, instead of gathering every mailbox again.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -ArchiveMaxAttempts 5 -ArchiveRetryDelaySeconds 10
    Retry each In Place Archive mailbox up to 5 times, starting with a 10 second backoff.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -SkipArchiveFallback $true
    Only ever use Get-EXOMailboxStatistics -Archive; never fall back to a per-folder rollup.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -ResumeRIF $true
    Resume Recoverable Items gathering from the checkpoint file left behind by an
    interrupted run, instead of gathering every mailbox again.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -RIFMaxAttempts 5 -RIFRetryDelaySeconds 10
    Retry each Recoverable Items mailbox up to 5 times, starting with a 10 second backoff.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -ADGroup <ad_group_name>
    Gather user info for only the AD Group specified.
.NOTES
    Author:         Chris Lumnah
    Created Date:   6/17/2021
    Updated: 7/9/24
    By: Steven Tong
    Updated: 26/08/24
    By: Sameer Arora
    Updated: 25/09/26
    By: In Place Archive gathering made resilient - retry with backoff, failure
        accounting, and checkpoint/resume.
    Updated: 25/09/26
    By: Recoverable Items gathering made resilient the same way - retry with backoff,
        failure accounting, and checkpoint/resume.
    Updated: 25/09/26
    By: In Place Archive gathering falls back to a per-folder statistics rollup when the
        aggregate archive statistics call still fails after exhausting its retries.
#>

[CmdletBinding()]
param (
    [Parameter()]
    [bool]$EnableDebug = $false,
    # Provide your estimated annual growth rate, eg '10' for 10%
    [Parameter(HelpMessage="Estimated annual growth rate, eg 10 for 10%")]
    [int]$AnnualGrowth,
    # Provide AD Group name if you only want to gather data for a specific AD Group
    [Parameter()]
    [String]$ADGroup,
    # Provide AD Group name to exclude if you want to exclude it from sizing
    [Parameter()]
    [String]$ExcludeADGroup,
    # If gathering AD Group, CSV file to output AD Group membership info
    [Parameter()]
    [String]$ADGroupCSVFilename = './adgrouplist.csv',
    # Whether or not to skip gathering archived mailboxes, which can timeout
    [Parameter()]
    [bool]$SkipArchiveMailbox = $false,
    # Whether or not to skip gathering shared mailboxes
    [Parameter()]
    [bool]$SkipSharedMailbox = $false,
    # Whether or not to skip gathering Recoverable Items fodler items, which can timeout
    [Parameter()]
    [bool]$SkipRecoverableItems = $true,
    # Number of attempts per mailbox when gathering In Place Archive stats (1 means no retry)
    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$ArchiveMaxAttempts = 3,
    # Seconds to wait before the first In Place Archive retry; doubles each attempt, capped at 60
    [Parameter()]
    [ValidateRange(0, 300)]
    [int]$ArchiveRetryDelaySeconds = 5,
    # Skip the per-folder statistics rollup fallback and only ever use Get-EXOMailboxStatistics -Archive
    [Parameter()]
    [bool]$SkipArchiveFallback = $false,
    # Number of attempts for the per-folder rollup fallback, used only after -ArchiveMaxAttempts
    # is exhausted on the aggregate call (1 means no retry)
    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$ArchiveFallbackMaxAttempts = 2,
    # Resume In Place Archive gathering from the checkpoint file left by an interrupted run
    [Parameter()]
    [bool]$ResumeArchive = $false,
    # Checkpoint file written as In Place Archive stats are gathered, and read back by -ResumeArchive
    [Parameter()]
    [String]$ArchiveCheckpointFilename = './archive-checkpoint.csv',
    # Number of attempts per mailbox when gathering Recoverable Items stats (1 means no retry)
    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$RIFMaxAttempts = 3,
    # Seconds to wait before the first Recoverable Items retry; doubles each attempt, capped at 60
    [Parameter()]
    [ValidateRange(0, 300)]
    [int]$RIFRetryDelaySeconds = 5,
    # Resume Recoverable Items gathering from the checkpoint file left by an interrupted run
    [Parameter()]
    [bool]$ResumeRIF = $false,
    # Checkpoint file written as Recoverable Items stats are gathered, and read back by -ResumeRIF
    [Parameter()]
    [String]$RIFCheckpointFilename = './rif-checkpoint.csv',
    # Number of days to get historical stats for: 7, 30, 90, 180
    [Parameter()]
    [Int]$Period = 180,
    # UseAppAccess indicates that the user wants to access Exchange through App access
    # instead of delegated user access.
    [Parameter(HelpMessage="Set this to true when you want to access Exchange through
    App access instead of delegated user access.")]
    [String]$UseAppAccess = $false

)

$date = Get-Date
$dateString = $date.ToString("yyyy-MM-dd")
$dateStringHH = $date.ToString("yyyy-MM-dd_HHmm")

# Filename to export the html report to
$outFilename = "./Rubrik-M365-Sizing-$dateStringHH.html"

# Folder to export CSVs to
$ExportFolder = '.'

$Version = "6.5"

$ProgressPreference = 'SilentlyContinue'

# Define the capacity metric conversions
$GB = 1000000000
$GiB = 1073741824
$TB = 1000000000000
$TiB = 1099511627776

# Set which capacity metric to use
$capacityMetric = $GB
$capacityDisplay = 'GB'

Write-Host "Starting the Rubrik Microsoft 365 sizing script ($Version)."

$ExchangeHTMLTitle = "User"

if ($AnnualGrowth -eq '') {
  $AnnualGrowth = 30
}

# Function to use Graph APIs to download report info
Function Get-MgReport {
  [CmdletBinding()]
  param (
    # MS Graph API report name
    [Parameter(Mandatory)]
    [String]$ReportName,
    # Report Period (Days)
    [Parameter()]
    [ValidateSet("7", "30", "90", "180")]
    [String]$Period
  )
  try {
    if ($reportName -eq 'getMailboxUsageDetail' -or $reportName -eq 'getMailboxUsageStorage') {
      $graphApiVersion = "Beta"
    } else {
      $graphApiVersion = "Beta"
    }
    if ($Period -ne '') {
      Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$($graphApiVersion)/reports/$($ReportName)(period=`'D$($Period)`')" -OutputFilePath "$ExportFolder\$ReportName.csv"
    } else {
      Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$($graphApiVersion)/reports/$($ReportName)" -OutputFilePath "$ExportFolder\$ReportName.csv"
    }
    return "$ExportFolder\$ReportName.csv"
  }
  catch {
    $errorMessage = $_.Exception | Out-String
    if ($errorMessage.Contains('Response status code does not indicate success: Forbidden (Forbidden)')) {
      Disconnect-MgGraph
      throw "The user account used for authentication must have permissions covered by Reports Reader admin role."
    }
    throw $_.Exception
  }
}

# Function to calculate annual growth based on historical report
# This may not be a good representation of actual growth in an environment
# if many changes are happening - for example, a migration event which would
# inflate steady state growth numbers.
function Measure-AverageGrowth {
  param (
    [Parameter(Mandatory)]
    [string]$ReportCSV,
    [Parameter(Mandatory)]
    [string]$ReportName
  )
  $UsageReport = Import-Csv -Path $ReportCSV | Sort-Object -Property "Report Date" -Descending
  $ReportDays = $UsageReport[0].'Report Period'
  $LatestUsageGB = [math]::Round($UsageReport[0].'Storage Used (Byte)' / $capacityMetric, 2)
  $EarliestUsageGB = [math]::Round($UsageReport[-1].'Storage Used (Byte)' / $capacityMetric, 2)
  $GrowthOverPeriod = [math]::Round($LatestUsageGB - $EarliestUsageGB, 2)
  $AvgGrowthPerDay = $GrowthOverPeriod / $ReportDays
  $GrowthPerYearGB = [math]::Round($AvgGrowthPerDay * 365, 2)
  $GrowthPerYearPct = [math]::Round($GrowthPerYearGB / $LatestUsageGB, 2)
  Write-Host "$ReportName historical storage usage:"
  Write-Host "  - Usage on $($UsageReport[0].'Report Date'): $LatestUsageGB $capacityDisplay"
  Write-Host "  - Usage on $($UsageReport[-1].'Report Date'): $EarliestUsageGB $capacityDisplay"
  Write-Host "  - Growth over $ReportDays days: $GrowthOverPeriod $capacityDisplay"
  Write-Host "  - Growth annualized per year: $GrowthPerYearGB $capacityDisplay, $($GrowthPerYearPct * 100)%"
  return $GrowthPerYearPct
}


# Function to solve for licenses required
# Query M365Licsolver Azure Function
function Solve-License {
  param (
    [Parameter(Mandatory)]
    [int]$userLicense,
    [Parameter(Mandatory)]
    [int]$storageGB
  )
  # If less than 76GB Average per user then query the azure function that calculates the best mix of subscription types. If more than 76 then Unlimited is the best option.
  if (($storageGB) / $userLicense -le 76) {
    # Query the M365Licsolver Azure Function
    $SolverQuery = '{"users":"' + $userLicense + '","data":"' + $storageGB + '"}'
    try {
      $APIReturn = ConvertFrom-JSON (Invoke-WebRequest 'https://m365licsolver-azure.azurewebsites.net:/api/httpexample' -ContentType "application/json" -Body $SolverQuery -UseBasicParsing -Method 'POST')
    }
    catch {
      $errorMessage = $_.Exception | Out-String
      if ($errorMessage.Contains('Response status code does not indicate success: 404')) {
        Write-Host "[Info] Unable to calculate license recommendations."
      }
    }
    $FiveGBUsers = $APIReturn.FiveGBSubscriptions * 10
    $TwentyGBUsers = $APIReturn.TwentyGBSubscriptions * 10
    $FiftyGBUsers = $APIReturn.FiftyGBSubscriptions * 10
    $UnlimitedGBUsers = 0
  } else {
    $FiveGBUsers = 0
    $TwentyGBUsers = 0
    $FiftyGBUsers = 0
    $UnlimitedGBPacks = [math]::ceiling($userLicense / 10)
    $UnlimitedGBUsers = $UnlimitedGBPacks * 10
  }
  $licenseRequired = [PSCustomObject] @{
    "FiveGBUsers" = $FiveGBUsers
    "TwentyGBUsers" = $TwentyGBUsers
    "FiftyGBUsers" = $FiftyGBUsers
    "UnlimitedGBUsers" = $UnlimitedGBUsers
  }
  return $licenseRequired
}

# Get Recoverable Items folder statistics (Primary, and In-Place Archive if requested) for one
# mailbox. Thin wrapper: only the EXO calls and the guarded per-folder parse live here, the
# retry/backoff/reconnect loop is shared via Invoke-MailboxStatsWithRetry (defined further below,
# alongside Get-ArchiveMailboxStats - function bodies resolve names at call time, not definition
# order, so the forward reference is safe).
function Get-RIFStatsForMailbox {
  param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true, HelpMessage = "Enter whether to include In-Place archive mailbox Recoverable Items folder or not.")]
    [bool]$IncludeArchiveMailbox,

    [Parameter()]
    [int]$MaxAttempts = 3,

    [Parameter()]
    [int]$RetryDelaySeconds = 5,

    [Parameter()]
    [bool]$AllowReconnect = $true,

    [Parameter()]
    [bool]$ShowRetryDetail = $false
  )

  $RecoverableItemsSpecialFolders = @(
    "/Deletions",
    "/Purges",
    "/Versions",
    "/DiscoveryHolds"
  )

  # A scriptblock literal is lexically scoped to where it's DEFINED (here, inside
  # Get-RIFStatsForMailbox), not where it's later invoked from (Invoke-MailboxStatsWithRetry) -
  # so $IncludeArchiveMailbox and $RecoverableItemsSpecialFolders resolve correctly without
  # .GetNewClosure(). That method isn't just a variable snapshot: it binds the scriptblock to a
  # brand-new, isolated session state, which can fail to resolve commands that aren't part of
  # that fresh state - never needed here anyway, since this scriptblock is freshly created on
  # every call rather than being reused across loop iterations (the case GetNewClosure is for).
  $FetchStats = {
    param($UserPrincipalName)

    # -ErrorAction Stop is required: without it a non-terminating error skips the catch and
    # leaves $PrimaryStats null, which used to fall through into the size parse below.
    # The @(...) wraps the PIPELINE, not the resulting variable: a zero-match Where-Object gives
    # a true empty array this way, instead of collapsing to $null and then wrapping into a
    # one-element array containing that $null (@($null) is length 1, not 0).
    $PrimaryStats = @(Get-MailboxFolderStatistics -Identity $UserPrincipalName -FolderScope RecoverableItems -ErrorAction Stop |
      Where-Object { $RecoverableItemsSpecialFolders -contains $_.FolderPath })

    $ArchiveFolderStats = @()
    if ($IncludeArchiveMailbox) {
      try {
        $ArchiveFolderStats = @(Get-MailboxFolderStatistics -Identity $UserPrincipalName -FolderScope RecoverableItems -Archive -ErrorAction Stop |
          Where-Object { $RecoverableItemsSpecialFolders -contains $_.FolderPath })
      } catch {
        # A transient failure here must not be silently absorbed: doing so would combine
        # possibly-incomplete archive data with the primary result under an overall 'success',
        # with no retry - the exact silent under-report this whole rewrite exists to prevent.
        # Rethrow so Invoke-MailboxStatsWithRetry's transient-error handling retries the WHOLE
        # mailbox. A permanent archive error (e.g. the archive was disabled after 'Has Archive'
        # was reported TRUE for this mailbox) means the archive genuinely has nothing to
        # contribute, so the mailbox is still counted using its primary Recoverable Items alone,
        # instead of failing the mailbox entirely and reporting 0 for data already in hand.
        if (Test-ExoTransientError -ErrorRecord $_) { throw }
      }
    }

    $FolderStats = $PrimaryStats + $ArchiveFolderStats
    $RIFSize = [long]0
    foreach ($Stats in $FolderStats) {
      $ParsedSize = ConvertFrom-ExoByteSizeString -SizeString ([string]$Stats.FolderSize)
      if ($null -eq $ParsedSize) {
        return @{
          Success = $false
          ErrorType = 'UnparsableSize'
          ErrorMessage = "Could not parse FolderSize for folder $($Stats.FolderPath)"
        }
      }
      $RIFSize += $ParsedSize
      if ($ShowRetryDetail) {
        Write-Host "[INFO] $UserPrincipalName folder $($Stats.FolderPath): $ParsedSize bytes, $($Stats.ItemsInFolder) items"
      }
    }

    $TotalItems = [long]($FolderStats | Measure-Object -Property 'ItemsInFolder' -Sum).Sum
    if ($ShowRetryDetail) {
      Write-Host "[INFO] $UserPrincipalName total: $RIFSize bytes, $TotalItems items across $($FolderStats.Count) folder(s)"
    }

    return @{
      Success = $true
      Fields = @{
        RIFSize = $RIFSize
        RIFItems = $TotalItems
      }
    }
  }

  return Invoke-MailboxStatsWithRetry -UserPrincipalName $UserPrincipalName -FetchStats $FetchStats `
    -EmptyFields ([ordered]@{ RIFSize = [long]0; RIFItems = [long]0 }) `
    -MaxAttempts $MaxAttempts -RetryDelaySeconds $RetryDelaySeconds -AllowReconnect $AllowReconnect -ShowRetryDetail $ShowRetryDetail
}

function Get-AccessToken {
    param (
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$TenantId,
        [string]$Scope = "https://outlook.office365.com/.default"
    )

    $body = @{
        client_id     = $ClientId
        scope         = $Scope
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    $url = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    try {
        $response = Invoke-RestMethod -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body
        return $response.access_token
    } catch {
        Write-Error "Failed to get access token: $_"
        return $null
    }
}

# Refresh the EXO connection so the access token does not age out during the
# multi-day archive enumeration on very large tenants (SPARK-887257).
function Reset-ExoConnection {
  try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { }
  if ($script:ExoAuthMode -eq 'App') {
    $token = Get-AccessToken -clientId $script:ExoClientId -clientSecret $script:ExoClientSecret -tenantId $script:ExoTenantId
    Connect-ExchangeOnlineForSizing -AccessToken $token -Organization $script:ExoTenantId
  } elseif ($script:ExoAuthMode -eq 'Delegate') {
    Connect-ExchangeOnlineForSizing -UserPrincipalName $script:ExoUserPrincipalName
  } else {
    Connect-ExchangeOnlineForSizing
  }
}

# Test whether an exception is the MSAL broker assembly mismatch that surfaces when
# an older Microsoft.Identity.Client.dll is loaded into the PowerShell session
# (typically via stale AzureAD / Microsoft.Graph / EXO modules). See SPARK-887253.
function Test-ExoBrokerAssemblyConflict {
  param ([string]$Message)
  if (-not $Message) { return $false }
  return ($Message -match 'WithBroker' -or
          $Message -match 'BrokerExtension' -or
          ($Message -match 'Method not found' -and $Message -match 'Identity\.Client'))
}

function Write-ExoBrokerRemediation {
  Write-Host ""
  Write-Host "[ERROR] Exchange Online connect failed due to a stale Microsoft.Identity.Client assembly." -ForegroundColor Red
  Write-Host "        An older Microsoft.Identity.Client.dll is already loaded into this PowerShell session" -ForegroundColor Red
  Write-Host "        (commonly from the legacy AzureAD/AzureADPreview modules or an outdated Microsoft.Graph)." -ForegroundColor Red
  Write-Host "        PowerShell cannot unload it, so this session is unrecoverable." -ForegroundColor Red
  Write-Host ""
  Write-Host "To fix: open a FRESH PowerShell session and run:" -ForegroundColor Yellow
  Write-Host "  Uninstall-Module AzureAD -AllVersions -ErrorAction SilentlyContinue" -ForegroundColor Yellow
  Write-Host "  Uninstall-Module AzureADPreview -AllVersions -ErrorAction SilentlyContinue" -ForegroundColor Yellow
  Write-Host "  Update-Module ExchangeOnlineManagement -Force" -ForegroundColor Yellow
  Write-Host "  Update-Module Microsoft.Graph -Force" -ForegroundColor Yellow
  Write-Host "Then close ALL PowerShell windows, open a new one, and re-run this script." -ForegroundColor Yellow
  Write-Host ""
}

# Connect to Exchange Online with retry. If the broker MSAL conflict is hit during
# an interactive (delegate) connect, fall back to device-code (-Device) which bypasses
# WAM/broker entirely. App-auth bypasses the broker, so retries print remediation only.
function Connect-ExchangeOnlineForSizing {
  param (
    [string]$AccessToken,
    [string]$Organization,
    [string]$UserPrincipalName
  )
  try {
    if ($AccessToken) {
      Connect-ExchangeOnline -AccessToken $AccessToken -Organization $Organization -ShowBanner:$false
    } elseif ($UserPrincipalName) {
      Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false
    } else {
      Connect-ExchangeOnline -ShowBanner:$false
    }
  } catch {
    $msg = $_.Exception.Message
    if (Test-ExoBrokerAssemblyConflict -Message $msg) {
      if (-not $AccessToken) {
        Write-Host "[WARN] Broker MSAL conflict detected; retrying with device-code authentication (-Device)." -ForegroundColor Yellow
        try {
          if ($UserPrincipalName) {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -Device -ShowBanner:$false
          } else {
            Connect-ExchangeOnline -Device -ShowBanner:$false
          }
          return
        } catch {
          $msg = $_.Exception.Message
        }
      }
      Write-ExoBrokerRemediation
    }
    throw
  }
}

# Validate that Period (days) for historical reports is valid
# Must be: 7, 30, 90, or 180
$PeriodValues = @(7, 30, 90, 180)
if ($Period -in $PeriodValues) {
} else {
    throw "Error: Period (days) needs to be: 7, 30, 90, or 180"
}

# Validate the required 'Microsoft.Graph.Reports' is installed
# and provide a user friendly message when it's not.
if (Get-Module -ListAvailable -Name Microsoft.Graph.Reports) {
} else {
  throw "The 'Microsoft.Graph.Reports' module is required for this script. Run the follow command to install: Install-Module Microsoft.Graph.Reports"
}

# Validate the required 'ExchangeOnlineManagement' is installed and provide a user friendly message when it's not.
# Require >= 3.4.0 to avoid the MSAL broker mismatch
# ("Method not found ... BrokerExtension.WithBroker(..., BrokerOptions)") seen when an older
# Microsoft.Identity.Client.dll is loaded into the session via stale modules (SPARK-887253).
$MinEXOModuleVersion = [version]'3.4.0'
$exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
  Sort-Object Version -Descending | Select-Object -First 1
if (-not $exoModule) {
  throw "The 'ExchangeOnlineManagement' module is required for this script. Run the follow command to install: Install-Module ExchangeOnlineManagement -MinimumVersion $MinEXOModuleVersion"
}
if ($exoModule.Version -lt $MinEXOModuleVersion) {
  throw "The 'ExchangeOnlineManagement' module is version $($exoModule.Version), but $MinEXOModuleVersion or later is required to avoid an MSAL broker mismatch. Run: Update-Module ExchangeOnlineManagement -Force, then start a fresh PowerShell session."
}

$AzureAdRequired = $PSBoundParameters.ContainsKey('ADGroup') -or $PSBoundParameters.ContainsKey('ExcludeADGroup')

if ($AzureAdRequired) {
  # Validate the required 'Azure.Graph.Authentication' is installed
  # and provide a user friendly message when it's not.
  if (Get-Module -ListAvailable -Name Microsoft.Graph.Groups) {
  } else {
    throw "The 'Microsoft.Graph.Groups' module is required for filtering by a specific Azure AD Group. Run the follow command to install: Install-Module Microsoft.Graph.Groups"
  }
}

if ($UseAppAccess -eq $true) {
    # Prompt the user to enter the Client ID, Client Secret, and Tenant ID
    Write-Host "You have chosen to authenticate through App access for accessing Exchange Online."
    Write-Host "[INFO] Please enter app info to authenticate for Graph Access that has the following permissions: "
    Write-Host "'Reports.Read.All', 'User.Read.All', and 'Group.Read.All'."
    $tenantId = Read-Host -Prompt "Enter your Tenant ID"
    $clientId = Read-Host -Prompt "Enter your Azure AD Application Client ID"
    $ClientSecretCredential = Get-Credential -Credential $clientId

    try {
        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientSecretCredential
    }
    catch {
        $errorException = $_.Exception
        $errorMessage = $errorException.Message
        Write-Host "[ERROR] Unable to Connect to the Microsoft Graph PowerShell Module: $errorMessage"
        return
    }

    if ($SkipArchiveMailbox -eq $false -or $SkipRecoverableItems -eq $false) {
        Write-Host "Connecting to the Microsoft Exchange Online Module to gather per-mailbox In Place Archive stats."
        try {
            $clientSecretSecure = Read-Host -Prompt "Password for user $clientId" -AsSecureString
            $clientSecret = ConvertFrom-SecureString -SecureString $clientSecretSecure -AsPlainText   # legit:ignore
            $token = Get-AccessToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenantId
            Connect-ExchangeOnlineForSizing -AccessToken $token -Organization $tenantId
            # Cache credentials so long-running loops can refresh the EXO token without re-prompting.
            $script:ExoAuthMode = 'App'
            $script:ExoTenantId = $tenantId
            $script:ExoClientId = $clientId
            $script:ExoClientSecret = $clientSecret
        } catch {
            if (Test-ExoBrokerAssemblyConflict -Message $_.Exception.Message) {
                Write-ExoBrokerRemediation
            }
            $errorMessage = $_.Exception.Message
            Write-Host "[ERROR] Unable to Connect to the Microsoft Exchange PowerShell Module: $errorMessage"
            return
        }
    }
} else {
    Write-Host "You have chosen to authenticate through delegate user access for accessing Exchange Online."
    Write-Host "[INFO] Please authenticate through an user with the following permissions:"
    Write-Host "'Reports.Read.All', 'User.Read.All', and 'Group.Read.All'."
    try {
        Connect-MgGraph -Scopes "Reports.Read.All", "User.Read.All", "Group.Read.All"  | Out-Null
    }
    catch {
        $errorException = $_.Exception
        $errorMessage = $errorException.Message
        Write-Host "[ERROR] Unable to Connect to the Microsoft Graph PowerShell Module: $errorMessage"
        return
    }

    if ($SkipArchiveMailbox -eq $false -or $SkipRecoverableItems -eq $false) {
        Write-Host "Connecting to the Microsoft Exchange Online Module to gather per-mailbox In Place Archive stats."
        try {
            Connect-ExchangeOnlineForSizing
            $script:ExoAuthMode = 'Delegate'
            $script:ExoUserPrincipalName = (Get-ConnectionInformation).UserPrincipalName
        } catch {
            if (Test-ExoBrokerAssemblyConflict -Message $_.Exception.Message) {
                Write-ExoBrokerRemediation
            }
            $errorMessage = $_.Exception.Message
            Write-Host "[ERROR] Unable to Connect to the Microsoft Exchange PowerShell Module: $errorMessage"
            return
        }
    }
}

# If AD Group is provided, get the AD Group membership info
if ($AzureAdRequired) {
  Write-Host "Looking up AD Group users in: $ADGroup" -foregroundcolor green
  $AzureAdGroupDetails = Get-MgGroup -Filter "DisplayName eq '$ADGroup'"
  if ($AzureAdGroupDetails.Count -eq 0) {
    throw "The Azure AD Group '$ADGroup' does not exist. Exiting script."
  }
  $AzureAdGroupMembersById = Get-MgGroupTransitiveMember -GroupId $AzureAdGroupDetails.Id -All
  if ($EnableDebug) {
    Write-Host "[DEBUG] Azure AD Group Members Size: $($AzureAdGroupMembersById.Count)"
  }
  $AzureAdGroupMembersByUserPrincipalName = @()
  $AzureAdGroupMembersById | Foreach-Object {
    if ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user") {
      $AzureAdGroupMembersByUserPrincipalName += $_.AdditionalProperties["userPrincipalName"]
    }
  }
  if ($AzureAdGroupMembersByUserPrincipalName.Count -eq 0) {
    throw "The Azure AD Group '$ADGroup' does not contain any User Principal Names."
  }
  Write-Host "# of Azure AD Group members found: $($AzureAdGroupMembersByUserPrincipalName.Count)" -foregroundcolor green
  Write-Host "AD Group user principal names exported to CSV: $ADGroupCSVFilename" -foregroundcolor green
  $AzureAdGroupMembersByUserPrincipalName | Out-File -Path $ADGroupCSVFilename
}

if ($EnableDebug) {
  try {
    $user = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me"
    $permissions = Get-MgUserOauth2PermissionGrant -UserId $user.id
    Write-Host "[DEBUG] The authenticated user account has the following permissions:$($permissions.Scope)"
  }
  catch {
    $errorMessage = $_.Exception | Out-String
    throw $_.Exception
  }
}

Write-Host ""
Write-Host "The data gathered here is the same as in the M365 Admin Center" -foreground green
Write-Host "under Reports -> Usage -> <Workload> -> Usage Reports"
Write-Host "Microsoft metrics could differ a bit depending on what you are looking at"
Write-Host ""

### Exchange - Get reports for Exchange and process them
Write-Host "*** Retrieving usage info for: Exchange ***" -foregroundcolor green
Write-Host "Data will be gathered from the chart data and per-mailbox usage report" -foregroundcolor green

Write-Host "Getting chart data - this should match the chart for the Users drop-down and Total mailboxes" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageMailboxCounts' -Period $Period
Write-Host "Exchange chart CSV saved to: $ReportCSV" -foregroundcolor green
$ExchangeChartData = Import-Csv -Path $ReportCSV
Write-Host "Total # of User (non-shared mailboxes) - chart data: $($ExchangeChartData[0].total)" -foregroundcolor green
Write-Host ""

# Get Exchange per user usage report
Write-Host "Getting the per-mailbox usage report - this should match the details in the Admin Center reports" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageDetail' -Period $Period
Write-Host "Exchange per-mailbox usage CSV saved to: $ReportCSV" -foregroundcolor green

$ExchangeUsageReport = Import-Csv -Path $ReportCSV
$ExchangeTotalUsersCount = $ExchangeUsageReport.count
$ExchangeNonDeletedUsersCount = $($ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$ExchangeDeletedUsersCount = $($ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of mailboxes - from usage report: $ExchangeTotalUsersCount"
Write-Host "Total # of deleted mailboxes - from usage report: $ExchangeDeletedUsersCount"
Write-Host "Total # of active (non-deleted) mailboxes - from usage report: $ExchangeNonDeletedUsersCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

# List of all active (non-deleted) user mailboxes
$ExchangeUsageReportUsers = $ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' -and
  $_.'Recipient Type' -ne 'Shared'}

# List of all active (non-deleted) shared mailboxes
if ($SkipSharedMailbox -eq $false) {
  $ExchangeUsageReportShared = $ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' -and
    $_.'Recipient Type' -eq 'Shared'}
  $ExchangeActiveUsers = $ExchangeUsageReportUsers + $ExchangeUsageReportShared
} else {
  $ExchangeUsageReportShared = @()
  $ExchangeActiveUsers = $ExchangeUsageReportUsers
}

if ($AzureAdRequired) {
  if ($ADGroup -ne '') {
    Write-Host "Filtering user mailboxes by Azure AD Group: $ADGroup"
  } else {

  }
  $FilterByField = "User Principal Name"
  $ExchangeUsageReportUsers = $ExchangeUsageReportUsers | Where-Object { $_.$FilterByField -in $AzureAdGroupMembersByUserPrincipalName }
  # If we didn't get any usage for the Azure AD group users, it might be because the reports are masking User IDs
  if ($ExchangeUsageReportUsers.count -eq 0) {
    Write-Host "[ERROR] Did not match any Azure AD Group users to the usage reports" -foregroundcolor red
    Write-Host "[ERROR] Check the Azure AD group csv ($ADGroupCSVFilename) to see what users are part of the Azure AD Group" -foregroundcolor red
    Write-Host "[ERROR] Check the mailbox csv ($reportCSV) to see if 'User Principal Name' are being masked" -foregroundcolor red
    Write-Host "[ERROR] Un-mask by going to M365 Admin Center -> Settings -> Org Settings -> Services"  -foregroundcolor red
    Write-Host "[ERROR] Then click on Reports and clear: Display concealed user, group, and site names in all reports, and then select Save"  -foregroundcolor red
    Write-Host "[ERROR] See: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/miscellaneous/reports-show-anonymous-user-name" -foregroundcolor red
    throw "Error running script with Azure AD group option - could not find any matching users. Exiting script."
  }
  Write-Host "Matched $($ExchangeUsageReportUsers.count) users in the provided Azure AD Group" -foregroundcolor green
}

# Calculate metrics for user mailboxes
$userMailboxStorageSum = $ExchangeUsageReportUsers | Measure-Object -Property 'Storage Used (Byte)' -Sum
$userMailboxStorageSumDisplay = [math]::Round($userMailboxStorageSum.Sum / $capacityMetric, 2)
$userMailboxItems = $ExchangeUsageReportUsers | Measure-Object -Property 'Item Count' -Sum

# Calculate metrics for shared mailboxes
if ($SkipSharedMailbox -eq $false) {
  $sharedMailboxStorageSum = $ExchangeUsageReportShared | Measure-Object -Property 'Storage Used (Byte)' -Sum
  $sharedMailboxItems = $ExchangeUsageReportShared | Measure-Object -Property 'Item Count' -Sum
} else {
  Write-Host "Skipping gathering Shared mailbox usage" -foregroundcolor green
  $sharedMailboxStorageSum = [PSCustomObject]@{ Sum = 0 }
  $sharedMailboxItems = [PSCustomObject]@{ Sum = 0 }
}
$sharedMailboxStorageSumDisplay = [math]::Round($sharedMailboxStorageSum.Sum / $capacityMetric, 2)


Write-Host "Total # of active user mailboxes - from usage report: $($ExchangeUsageReportUsers.count)" -foregroundcolor green
Write-Host "Active users storage used (not including in-place archve, Recoverable Items Folder): $userMailboxStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Active users item count (not including in-place archve, Recoverable Items Folder): $($userMailboxItems.sum)" -foregroundcolor green
Write-Host "Total # of active shared mailboxes - from usage report: $($ExchangeUsageReportShared.count)" -foregroundcolor green
Write-Host "Shared mailboxes storage used: $sharedMailboxStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Shared mailboxes item count (not including in-place archve, Recoverable Items Folder): $($sharedMailboxItems.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical Exchange storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$ExchangeDetails = [PSCustomObject] @{
  "Chart User Mailboxes" = $ExchangeChartData[0].total
  "Total Mailboxes" = $($ExchangeUsageReportUsers.count + $ExchangeUsageReportShared.count)
  "User Mailboxes" = $ExchangeUsageReportUsers.count
  "User Storage Used No Archive" = $userMailboxStorageSum.sum
  "User Items No Archive" = $userMailboxItems.sum
  "Shared Mailboxes" = if ($SkipSharedMailbox -eq $false) { $ExchangeUsageReportShared.count } else { "Skipped" }
  "Shared Storage Used No Archive" = if ($SkipSharedMailbox -eq $false) { $sharedMailboxStorageSum.sum } else { '_' }
  "Shared Items No Archive" = if ($SkipSharedMailbox -eq $false) { $sharedMailboxItems.sum } else { '_' }
  "Total Storage Used" = $($userMailboxStorageSum.sum + $sharedMailboxStorageSum.sum)
  "Total Items" = $($userMailboxItems.sum + $sharedMailboxItems.sum)
  "Calculated Growth %" = $CalculatedGrowth
}

### OneDrive - Get reports for OneDrive and process them
Write-Host "*** Retrieving usage info for: OneDrive ***" -foregroundcolor green
Write-Host "Data will be gathered from the chart data and per-account usage report" -foregroundcolor green

Write-Host "Getting chart data - this should match the chart for Total Accounts" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageAccountCounts' -Period $Period
Write-Host "OneDrive account chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartAccount = Import-Csv -Path $ReportCSV
Write-Host "Total # of OneDrive accounts - chart data: $($OneDriveChartAccount[0].total)" -foregroundcolor green
Write-Host ""

Write-Host "Getting chart data - this should match the chart for Total Storage" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageStorage' -Period $Period
Write-Host "OneDrive storage chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartStorage = Import-Csv -Path $ReportCSV
$OneDriveChartStorageUsed = [math]::round($OneDriveChartStorage[0].'Storage Used (Byte)' / $capacityMetric, 2)
Write-Host "Total # of OneDrive storage - chart data: $oneDriveChartStorageUsed $capacityDisplay" -foregroundcolor green
Write-Host ""

Write-Host "Getting chart data - this should match the chart for Total Files" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageFileCounts' -Period $Period
Write-Host "OneDrive files chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartFiles = Import-Csv -Path $ReportCSV
Write-Host "Total # of OneDrive files - chart data: $($OneDriveChartFiles[0].total)" -foregroundcolor green
Write-Host ""


Write-Host "Getting the per-account usage report - this should match the details in the Admin Center reports" -foregroundcolor green
Write-Host "However, this count may not match the Total count in the chart" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageAccountDetail' -Period $Period
Write-Host "OneDrive per-account usage CSV saved to: $ReportCSV" -foregroundcolor green

$OneDriveUsageReport = Import-Csv -Path $ReportCSV
$OneDriveTotalUsersCount = $OneDriveUsageReport.count
$OneDriveNonDeletedUsersCount = $($OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$OneDriveDeletedUsersCount = $($OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of accounts - from usage report: $($OneDriveUsageReportAccounts.count)"
Write-Host "Total # of deleted accounts - from usage report: $OneDriveDeletedUsersCount"
Write-Host "Total # of active (non-deleted) accounts - from usage report: $OneDriveNonDeletedUsersCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

$OneDriveUsageReportAccounts = $OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }

if ($AzureAdRequired) {
  Write-Host "Filtering user accounts by Azure AD Group: $ADGroup"
  $FilterByField = "Owner Principal Name"
  $OneDriveUsageReportAccounts = $OneDriveUsageReportAccounts | Where-Object { $_.$FilterByField -in $AzureAdGroupMembersByUserPrincipalName }
  # If we didn't get any usage for the Azure AD group users, it might be because the reports are masking User IDs
  if ($OneDriveUsageReportAccounts.count -eq 0) {
    Write-Host "[ERROR] Did not match any Azure AD Group users to the usage reports" -foregroundcolor red
    Write-Host "[ERROR] Check the Azure AD group csv ($ADGroupCSVFilename) to see what users are part of the Azure AD Group" -foregroundcolor red
    Write-Host "[ERROR] Check the OneDrive csv ($reportCSV) to see if 'Owner Principal Name' are being masked" -foregroundcolor red
    Write-Host "[ERROR] Un-mask by going to M365 Admin Center -> Settings -> Org Settings -> Services"  -foregroundcolor red
    Write-Host "[ERROR] Then click on Reports and clear: Display concealed user, group, and site names in all reports, and then select Save"  -foregroundcolor red
    Write-Host "[ERROR] See: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/miscellaneous/reports-show-anonymous-user-name" -foregroundcolor red
    throw "Error running script with Azure AD group option - could not find any matching users. Exiting script."
  }
  Write-Host "Matched $($OneDriveUsageReportAccounts.count) users in the provided Azure AD Group" -foregroundcolor green
}

# Calculate metrics for OneDrive accounts
$oneDriveStorageSum = $OneDriveUsageReportAccounts | Measure-Object -Property 'Storage Used (Byte)' -Sum
$oneDriveStorageSumDisplay = [math]::Round($oneDriveStorageSum.Sum / $capacityMetric, 2)
$oneDriveFiles = $OneDriveUsageReportAccounts | Measure-Object -Property 'File Count' -Sum

Write-Host "Total # of Accounts - from usage report: $($OneDriveUsageReportAccounts.count)" -foregroundcolor green
Write-Host "Accounts storage used: $oneDriveStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Accounts file count: $($oneDriveFiles.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical OneDrive storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$OneDriveDetails = [PSCustomObject] @{
  "Chart Accounts" = $OneDriveChartAccount[0].total
  "Chart Storage Used" = $OneDriveChartStorage[0].'Storage Used (Byte)'
  "Chart Files" = $OneDriveChartFiles[0].total
  "Usage Accounts" = $OneDriveUsageReportAccounts.count
  "Usage Account Storage Used" = $oneDriveStorageSum.sum
  "Usage Account Files" = $oneDriveFiles.sum
  "Calculated Growth %" = $CalculatedGrowth
}

# If AD Group is specified, assume each user is licensed so use that count
if ($AzureAdRequired) {
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Accounts' -Value $OneDriveDetails.'Usage Accounts'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Storage Used' -Value $OneDriveDetails.'Usage Account Storage Used'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Total Files' -Value $OneDriveDetails.'Usage Account Files'
} else {
  # Otherwise, use the counts from the Chart
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Accounts' -Value $OneDriveDetails.'Chart Accounts'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Storage Used' -Value $OneDriveDetails.'Chart Storage Used'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Total Files' -Value $OneDriveDetails.'Chart Files'
}

### SharePoint - Get reports for SharePoint and process them
Write-Host "*** Retrieving usage info for: SharePoint ***" -foregroundcolor green
Write-Host "Data will be gathered site usage report" -foregroundcolor green

Write-Host "Getting site usage report" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getSharePointSiteUsageDetail' -Period $Period
Write-Host "SharePoint site usage CSV saved to: $ReportCSV" -foregroundcolor green
$SharePointUsageReport = Import-Csv -Path $ReportCSV

# Calculate metrics for SharePoint Sites
$SharePointSitesCount = $SharePointUsageReport.count
$SharePointNonDeletedSitesCount = $($SharePointUsageReportSites | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$SharePointDeletedSitesCount = $($SharePointUsageReportSites | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of sites - from usage report: $SharePointSitesCount"
Write-Host "Total # of deleted mailboxes - from usage report: $SharePointDeletedSitesCount"
Write-Host "Total # of active (non-deleted) mailboxes - from usage report: $SharePointNonDeletedSitesCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

$SharePointUsageReportSites = $SharePointUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }

# Calculate metrics for SharePoint sites
$sharePointStorageSum = $SharePointUsageReportSites | Measure-Object -Property 'Storage Used (Byte)' -Sum
$sharePointStorageSumDisplay = [math]::Round($sharePointStorageSum.Sum / $capacityMetric, 2)
$sharePointFiles = $SharePointUsageReportSites | Measure-Object -Property 'File Count' -Sum

Write-Host "Total # of SharePoint Sites - from usage report: $($SharePointUsageReportSites.count)" -foregroundcolor green
Write-Host "SharePoint Sites storage used: $sharePointStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "SharePoint Sites file count: $($sharePointFiles.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical OneDrive storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getSharePointSiteUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$SharePointDetails = [PSCustomObject] @{
  "Sites" = $SharePointUsageReportSites.count
  "Sites Storage Used" = $sharePointStorageSum.sum
  "Account Files" = $sharePointFiles.sum
  "Calculated Growth %" = $CalculatedGrowth
}

Write-Host "[INFO] Disconnecting from the Microsoft Graph API."
Disconnect-MgGraph

# The Microsoft Exchange Reports do not contain In-Place Archive sizing information.DESCRIPTION
# We need to connect to the Exchange Online module to get this information

# @(...) wraps the PIPELINE, not the resulting variable: a tenant with zero In-Place Archive
# mailboxes makes Where-Object emit nothing, which collapses to $null (not an empty array) when
# assigned directly. Wrapping guarantees a real (possibly empty) array reaches
# Invoke-MailboxKindGathering's -Population parameter downstream regardless of population size.
$ArchiveMailboxes = @($ExchangeUsageReportUsers | Where-Object { $_.'Has Archive' -eq 'TRUE' })
$ArchiveMailboxesCount = $ArchiveMailboxes.Count

# Parse an EXO "size" string of the form "... (N bytes)" - used by both
# Get-MailboxFolderStatistics's FolderSize and Get-EXOMailboxStatistics's TotalItemSize - into a
# byte count. Returns $null on no match rather than leaving it to the caller to inspect $Matches,
# which a bare '-match' would leave holding the PREVIOUS successful parse's capture on failure,
# silently crediting the wrong byte count. Shared by Get-ArchiveMailboxStats and
# Get-RIFStatsForMailbox so this regex and its guard exist exactly once.
function ConvertFrom-ExoByteSizeString {
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$SizeString
  )

  if ($SizeString -notmatch '\(([^)]+) bytes\)') {
    return $null
  }
  return [long]($Matches[1] -replace ',', '')
}

# Classify an Exchange Online error as transient (worth retrying) or permanent. Permanent
# per-mailbox errors must never be retried: across tens of thousands of mailboxes the backoff
# sleeps alone would add hours of wall clock to a run that is already measured in days.
function Test-ExoTransientError {
  param (
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  $Msg = [string]$ErrorRecord.Exception.Message
  $Fqid = [string]$ErrorRecord.FullyQualifiedErrorId

  # Permanent: retrying these can never succeed, so fail fast and keep the run moving.
  # Apostrophes are matched as '.' so both the straight and typographic forms are caught.
  if ($Fqid -match 'ManagementObjectNotFoundException') { return $false }
  if ($Msg -match "couldn.t be found on") { return $false }
  if ($Msg -match '(?i)archive' -and $Msg -match "(?i)(isn.t enabled|is not enabled|doesn.t have|does not have|not enabled for)") { return $false }
  if ($Msg -match "(?i)not a valid (SmtpAddress|value)" -or $Fqid -match 'ParameterBindingValidationException') { return $false }
  if ($Msg -match "(?i)(insufficient access rights|you don.t have permission|isn.t assigned to any management roles)") { return $false }

  # Transient by exception type. More reliable than message text for the network family.
  $TypeNames = @()
  $Ex = $ErrorRecord.Exception
  if ($null -ne $Ex) {
    $TypeNames += $Ex.GetType().FullName
    if ($null -ne $Ex.InnerException) { $TypeNames += $Ex.InnerException.GetType().FullName }
  }
  foreach ($TypeName in $TypeNames) {
    if ($TypeName -match '(HttpRequestException|TaskCanceledException|TimeoutException|WebException|SocketException|IOException)') { return $true }
  }

  # Transient by message.
  if ($Msg -match '(?i)(throttl|429|too many requests)') { return $true }
  if ($Msg -match '(?i)(micro delay|budget.*exceeded|exceeded.*budget)') { return $true }
  if ($Msg -match '(?i)(token.*expired|expired.*token|AADSTS700082|AADSTS50173|unauthorized|\b401\b|authentication failed)') { return $true }
  if ($Msg -match '(?i)(timed out|timeout|task was canceled|operation has timed out)') { return $true }
  if ($Msg -match '(?i)(connection was closed|connection reset|unable to connect|remote server returned an error|remote name could not be resolved)') { return $true }
  if ($Msg -match '(?i)(ServiceUnavailable|InternalServerError|BadGateway|service is unavailable|try again later|temporary)') { return $true }
  if ($Msg -match '(?i)(starting a command on the remote server failed|cmdlet not found|pipeline.*broken)') { return $true }

  # Unknown: treat as transient. An unknown-but-transient error silently corrupts the total,
  # while an unknown-but-permanent one only costs time - and that cost is bounded by the
  # consecutive-failure breaker in the gathering loop below.
  return $true
}

# A session error is a transient error that a sleep alone will never fix: the EXO connection or
# access token is dead and has to be re-established before the next attempt. Deliberately a
# strict subset of Test-ExoTransientError - throttling and 5xx are excluded, because reconnecting
# while being throttled makes the throttling worse.
function Test-ExoSessionError {
  param (
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  $Msg = [string]$ErrorRecord.Exception.Message

  if ($Msg -match '(?i)(token.*expired|expired.*token|AADSTS700082|AADSTS50173|unauthorized|\b401\b|authentication failed)') { return $true }
  if ($Msg -match '(?i)(starting a command on the remote server failed|cmdlet not found|pipeline.*broken)') { return $true }
  if ($Msg -match '(?i)(connection was closed|connection reset|unable to connect)') { return $true }

  return $false
}

# Bounded-retry driver for a single mailbox's stats. Never throws: always returns a result
# object with UserPrincipalName/Status/Attempts/ErrorType/ErrorMessage plus whatever fields
# $EmptyFields declares, merged in from $FetchStats's success outcome. Shared by
# Get-ArchiveMailboxStats and Get-RIFStatsForMailbox so the retry/backoff/reconnect/
# exception-classification logic exists exactly once.
#
# $FetchStats is called once per attempt as `& $FetchStats $UserPrincipalName` and must either:
#   - throw (a live EXO error - classified via Test-ExoTransientError/Test-ExoSessionError), or
#   - return @{ Success = $true; Fields = @{ <one entry per $EmptyFields key> } }, or
#   - return @{ Success = $false; ErrorType = '...'; ErrorMessage = '...' } for a permanent,
#     non-retryable failure that isn't a thrown exception (e.g. unparsable data).
# $EmptyFields must be an [ordered] hashtable so checkpoint/CSV column order stays stable.
function Invoke-MailboxStatsWithRetry {
  param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $true)]
    [scriptblock]$FetchStats,

    [Parameter(Mandatory = $true)]
    [System.Collections.Specialized.OrderedDictionary]$EmptyFields,

    [Parameter()]
    [int]$MaxAttempts = 3,

    [Parameter()]
    [int]$RetryDelaySeconds = 5,

    [Parameter()]
    [bool]$AllowReconnect = $true,

    [Parameter()]
    [bool]$ShowRetryDetail = $false
  )

  $ResultProps = [ordered]@{ "UserPrincipalName" = $UserPrincipalName }
  foreach ($Key in $EmptyFields.Keys) { $ResultProps[$Key] = $EmptyFields[$Key] }
  $ResultProps["Status"] = 'Failed'
  $ResultProps["Attempts"] = 0
  $ResultProps["ErrorType"] = ''
  $ResultProps["ErrorMessage"] = ''
  $Result = [PSCustomObject]$ResultProps

  for ($Attempt = 1; $Attempt -le $MaxAttempts; $Attempt++) {
    $Result.Attempts = $Attempt
    try {
      $FetchOutcome = & $FetchStats $UserPrincipalName

      if (-not $FetchOutcome.Success) {
        $Result.Status = 'Failed'
        $Result.ErrorType = $FetchOutcome.ErrorType
        $Result.ErrorMessage = ([string]$FetchOutcome.ErrorMessage) -replace '[,"\r\n]', ' '
        return $Result
      }

      foreach ($Key in $EmptyFields.Keys) { $Result.$Key = $FetchOutcome.Fields[$Key] }
      $Result.Status = 'OK'
      $Result.ErrorType = ''
      $Result.ErrorMessage = ''
      return $Result
    } catch {
      $ErrorRecord = $_
      $ExceptionName = $ErrorRecord.Exception.GetType().Name
      $Result.ErrorMessage = ([string]$ErrorRecord.Exception.Message) -replace '[,"\r\n]', ' '
      if ($Result.ErrorMessage.Length -gt 200) { $Result.ErrorMessage = $Result.ErrorMessage.Substring(0, 200) }

      if (-not (Test-ExoTransientError -ErrorRecord $ErrorRecord)) {
        $Result.Status = 'Failed'
        $Result.ErrorType = "Permanent:$ExceptionName"
        return $Result
      }

      $Result.Status = 'Failed'
      $Result.ErrorType = $ExceptionName

      if ($Attempt -ge $MaxAttempts) { return $Result }

      if ($ShowRetryDetail) {
        Write-Host "[INFO] Retry $Attempt/$MaxAttempts for $UserPrincipalName after transient error: $($Result.ErrorMessage)"
      }

      # A dead token or session is never fixed by sleeping, so reconnect for those only.
      if ($AllowReconnect -and (Test-ExoSessionError -ErrorRecord $ErrorRecord)) {
        try { Reset-ExoConnection } catch {
          Write-Host "[WARN] EXO reconnect failed: $($_.Exception.Message). Will retry on next interval."
        }
      }

      Start-Sleep -Seconds ([math]::Min(60, $RetryDelaySeconds * [math]::Pow(2, $Attempt - 1)))
    }
  }

  return $Result
}

# Get In-Place Archive statistics for one mailbox. Primary path: a single aggregate
# Get-EXOMailboxStatistics -Archive call, retried via Invoke-MailboxStatsWithRetry exactly as
# before. Fallback path: when the aggregate call still fails after exhausting -MaxAttempts, and
# the failure isn't one where retrying anything is pointless (the archive genuinely doesn't exist
# or isn't accessible), roll the size up from Get-MailboxFolderStatistics -Archive instead - one
# row per archive folder rather than a single aggregate row, so heavier, but a different EXO API
# path that can succeed when the aggregate call keeps timing out or being throttled. Reuses
# ConvertFrom-ExoByteSizeString and Invoke-MailboxStatsWithRetry so neither the byte-size parsing
# nor the retry/backoff/reconnect logic is duplicated here.
function Get-ArchiveMailboxStats {
  param (
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [Parameter()]
    [int]$MaxAttempts = 3,

    [Parameter()]
    [int]$RetryDelaySeconds = 5,

    [Parameter()]
    [bool]$AllowReconnect = $true,

    [Parameter()]
    [bool]$ShowRetryDetail = $false,

    [Parameter()]
    [bool]$EnableFolderRollupFallback = $true,

    [Parameter()]
    [int]$FallbackMaxAttempts = 2
  )

  $EmptyFields = [ordered]@{ ArchiveSize = [long]0; ArchiveItems = [long]0 }

  $FetchStats = {
    param($UserPrincipalName)

    # -ErrorAction Stop is required: without it a non-terminating error skips the catch and
    # leaves $ArchiveMailboxStats null, which used to fall through into the size parse below.
    $ArchiveMailboxStats = Get-EXOMailboxStatistics -Archive -Identity $UserPrincipalName -ErrorAction Stop

    $ParsedSize = ConvertFrom-ExoByteSizeString -SizeString ([string]$ArchiveMailboxStats.TotalItemSize)
    if ($null -eq $ParsedSize) {
      return @{
        Success = $false
        ErrorType = 'UnparsableSize'
        ErrorMessage = "Could not parse TotalItemSize: $($ArchiveMailboxStats.TotalItemSize)"
      }
    }

    return @{
      Success = $true
      Fields = @{
        ArchiveSize = $ParsedSize
        ArchiveItems = [long]$ArchiveMailboxStats.ItemCount
      }
    }
  }

  $PrimaryResult = Invoke-MailboxStatsWithRetry -UserPrincipalName $UserPrincipalName -FetchStats $FetchStats `
    -EmptyFields $EmptyFields -MaxAttempts $MaxAttempts -RetryDelaySeconds $RetryDelaySeconds `
    -AllowReconnect $AllowReconnect -ShowRetryDetail $ShowRetryDetail

  if ($PrimaryResult.Status -eq 'OK' -or -not $EnableFolderRollupFallback) { return $PrimaryResult }

  # A permanent failure here (e.g. "archive isn't enabled") means the aggregate call already
  # reached the mailbox and got a definitive answer - the folder rollup below would hit the exact
  # same archive mailbox and fail identically, just burning an extra EXO call for no chance of a
  # different outcome. Only fall back for failures that might be specific to the aggregate
  # statistics path itself: timeouts, throttling, and other transient errors exhausted across
  # every retry.
  if ($PrimaryResult.ErrorType -match '^Permanent:') { return $PrimaryResult }

  if ($ShowRetryDetail) {
    Write-Host "[INFO] $UserPrincipalName archive stats failed after $($PrimaryResult.Attempts) attempt(s) ($($PrimaryResult.ErrorType)); falling back to per-folder rollup."
  }

  $FolderFetchStats = {
    param($UserPrincipalName)

    # No -FolderScope: unlike Get-RIFStatsForMailbox's Recoverable Items rollup, every folder in
    # the archive mailbox has to be summed to reconstruct the mailbox's total size. Each row's
    # FolderSize is that folder's own content only, not its subfolders', so summing every row
    # double-counts nothing.
    $FolderStats = @(Get-MailboxFolderStatistics -Identity $UserPrincipalName -Archive -ErrorAction Stop)

    $ArchiveSize = [long]0
    foreach ($Stats in $FolderStats) {
      $ParsedSize = ConvertFrom-ExoByteSizeString -SizeString ([string]$Stats.FolderSize)
      if ($null -eq $ParsedSize) {
        return @{
          Success = $false
          ErrorType = 'UnparsableSize'
          ErrorMessage = "Could not parse FolderSize for folder $($Stats.FolderPath)"
        }
      }
      $ArchiveSize += $ParsedSize
    }

    return @{
      Success = $true
      Fields = @{
        ArchiveSize = $ArchiveSize
        ArchiveItems = [long]($FolderStats | Measure-Object -Property 'ItemsInFolder' -Sum).Sum
      }
    }
  }

  $FallbackResult = Invoke-MailboxStatsWithRetry -UserPrincipalName $UserPrincipalName -FetchStats $FolderFetchStats `
    -EmptyFields $EmptyFields -MaxAttempts $FallbackMaxAttempts -RetryDelaySeconds $RetryDelaySeconds `
    -AllowReconnect $AllowReconnect -ShowRetryDetail $ShowRetryDetail

  if ($FallbackResult.Status -eq 'OK') {
    if ($ShowRetryDetail) {
      Write-Host "[INFO] $UserPrincipalName recovered archive size via per-folder rollup after the aggregate call failed."
    }
    return $FallbackResult
  }

  # Both paths failed. ErrorType is left as the aggregate call's own value (not folded together
  # with the fallback's) because ArchiveFailures summarizes with 'Group-Object ErrorType' -
  # combining the two into one string would fragment that grouping into a separate bucket per
  # (aggregate error, fallback error) pair instead of the plain exception name operators already
  # know how to read. The fallback outcome still shows up in ErrorMessage, which isn't grouped.
  # ErrorMessage can't contain a comma: it's explicitly stripped of ',"\r\n' below.
  $PrimaryResult.ErrorMessage = ("Aggregate: $($PrimaryResult.ErrorMessage) | FolderRollup also failed ($($FallbackResult.ErrorType)): $($FallbackResult.ErrorMessage)" -replace '[,"\r\n]', ' ')
  if ($PrimaryResult.ErrorMessage.Length -gt 200) { $PrimaryResult.ErrorMessage = $PrimaryResult.ErrorMessage.Substring(0, 200) }
  return $PrimaryResult
}

# Move an existing checkpoint out of the way instead of truncating or appending to it, so a
# mistake never costs the recovery data from a multi-day run. The suffix loop matters because
# $dateStringHH only has minute resolution and Move-Item -Force would overwrite the very backup
# this is meant to protect. Used for both the Archive and Recoverable Items checkpoints.
function Move-CheckpointAside {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$Stamp,

    [Parameter()]
    [string]$Reason = ''
  )

  $BackupPath = "$Path.bak-$Stamp"
  $Suffix = 1
  while (Test-Path -Path $BackupPath) {
    $BackupPath = "$Path.bak-$Stamp-$Suffix"
    $Suffix += 1
  }
  Move-Item -Path $Path -Destination $BackupPath
  if ($Reason) {
    Write-Host "[INFO] $Reason Existing file moved to $BackupPath."
  } else {
    Write-Host "[INFO] Existing checkpoint moved to $BackupPath."
  }
}

# Load a previous run's checkpoint. Returns the completed mailboxes as a hashtable keyed by
# UserPrincipalName so resume lookups are O(1) - a Where-Object scan per mailbox would
# reintroduce the O(N^2) cost that SPARK-887257 removed from this code path - plus a Valid flag
# telling the caller whether the file on disk is safe to keep appending to.
# Generalized over both the Archive and Recoverable Items checkpoints: HeaderTag identifies the
# file format, SizeField/ItemsField are the value columns specific to each kind.
function Read-Checkpoint {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$HeaderTag,

    [Parameter(Mandatory = $true)]
    [string]$SizeField,

    [Parameter(Mandatory = $true)]
    [string]$ItemsField,

    [Parameter(Mandatory = $true)]
    [string]$Kind,

    [Parameter(Mandatory = $true)]
    [string]$FilenameParamName,

    [Parameter(Mandatory = $true)]
    [string]$ResumeParamName,

    [Parameter()]
    [string]$TenantKey = 'unknown',

    [Parameter()]
    [int]$Population = 0
  )

  $Map = @{}

  if (-not (Test-Path -Path $Path)) {
    Write-Host "[WARN] No checkpoint file found at $Path. Gathering all $Kind mailboxes from the start."
    # Nothing on disk to be unsafe about: the caller will create the file with a fresh header.
    return [PSCustomObject] @{ "Valid" = $true; "Completed" = $Map }
  }

  # -Encoding UTF8 is explicit because the checkpoint is written as UTF-8 by StreamWriter, while
  # Get-Content on Windows PowerShell 5.1 defaults to the ANSI code page. Without this, a UPN
  # containing a non-ASCII character would not round-trip and that mailbox would be re-gathered.
  $Lines = Get-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
  if ($Lines.Count -lt 2 -or [string]$Lines[0] -notmatch "^#$HeaderTag\|1\|") {
    Write-Host "[WARN] $Path is not a Rubrik $Kind checkpoint, or uses a newer format. Nothing can be resumed from it." -foregroundcolor yellow
    return [PSCustomObject] @{ "Valid" = $false; "Completed" = $Map }
  }

  # Tenant is the one hard gate. Resuming across tenants would produce a confidently wrong total,
  # so refuse rather than silently gather from scratch over someone else's data.
  if ([string]$Lines[0] -match 'TenantId=([^|]*)') {
    $CheckpointTenant = $Matches[1]
    if ($CheckpointTenant -ne 'unknown' -and $TenantKey -ne 'unknown' -and $CheckpointTenant -ne $TenantKey) {
      Write-Host "[ERROR] Checkpoint $Path was written for tenant '$CheckpointTenant' but this run is connected to '$TenantKey'." -foregroundcolor red
      Write-Host "[ERROR] Refusing to resume. Use a different -$FilenameParamName, or re-run without -$ResumeParamName." -foregroundcolor red
      throw "$Kind checkpoint tenant mismatch: '$CheckpointTenant' != '$TenantKey'."
    }
  }

  if ([string]$Lines[0] -match 'Created=(\d{4}-\d{2}-\d{2})_') {
    # ParseExact against the format this script writes, so the check does not depend on the
    # culture of whichever machine runs it.
    $CreatedDate = $null
    try {
      $CreatedDate = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
      $CreatedDate = $null
    }
    if ($null -ne $CreatedDate) {
      $AgeDays = [int]((Get-Date) - $CreatedDate).TotalDays
      if ($AgeDays -gt 7) {
        Write-Host "[WARN] Checkpoint $Path is $AgeDays days old. $Kind sizes gathered then may no longer be current." -foregroundcolor yellow
      }
    }
  }

  if ([string]$Lines[0] -match 'Population=(\d+)') {
    $CheckpointPopulation = [int]$Matches[1]
    if ($Population -gt 0 -and $CheckpointPopulation -ne $Population) {
      Write-Host "[INFO] Checkpoint was written for $CheckpointPopulation mailboxes; this run has $Population. Resuming by user principal name; the difference will be gathered."
    }
  }

  $SkippedRows = 0
  $Rows = $Lines | Select-Object -Skip 1 | ConvertFrom-Csv
  foreach ($Row in $Rows) {
    if ([string]::IsNullOrWhiteSpace($Row.UserPrincipalName)) { $SkippedRows += 1; continue }
    # Failed rows are deliberately not loaded, so they are retried on resume. A resume is usually
    # a fresh session with a fresh token, and a genuinely permanent failure re-fails in milliseconds.
    if ($Row.Status -ne 'OK') { continue }
    try {
      # The [long] casts are load-bearing. ConvertFrom-Csv yields strings, and Measure-Object -Sum
      # over string properties is unreliable on PowerShell 5.1, so without these a resumed run
      # would report a SMALLER total than an uninterrupted one.
      $CompletedProps = [ordered]@{ "UserPrincipalName" = $Row.UserPrincipalName }
      $CompletedProps[$SizeField] = [long]$Row.($SizeField)
      $CompletedProps[$ItemsField] = [long]$Row.($ItemsField)
      $Map[$Row.UserPrincipalName] = [PSCustomObject]$CompletedProps
    } catch {
      $SkippedRows += 1
    }
  }

  if ($SkippedRows -gt 0) {
    Write-Host "[WARN] Skipped $SkippedRows unreadable checkpoint row(s). Those mailboxes will be gathered again." -foregroundcolor yellow
  }

  return [PSCustomObject] @{ "Valid" = $true; "Completed" = $Map }
}

# Rewrite the checkpoint down to one row per mailbox before a resume reopens it for append.
# Without this, a mailbox that fails on every resume gets a brand new row each time - growth
# is failures x resumes rather than the (fixed) mailbox population - since failed rows are
# retried and re-appended but never removed. Only $CompletedMap survives (its rows are already
# unique per UserPrincipalName), so failed rows are dropped here; the retry this run will
# either succeed and add a fresh OK row, or fail once more.
function Compact-Checkpoint {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [hashtable]$CompletedMap,

    [Parameter(Mandatory = $true)]
    [string]$SizeField,

    [Parameter(Mandatory = $true)]
    [string]$ItemsField
  )

  $HeaderLine = Get-Content -Path $Path -Encoding UTF8 -TotalCount 1
  $Writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.UTF8Encoding]::new($false))
  try {
    $Writer.WriteLine($HeaderLine)
    $Writer.WriteLine("UserPrincipalName,$SizeField,$ItemsField,Status,ErrorType")
    foreach ($Completed in $CompletedMap.Values) {
      $Writer.WriteLine("`"$($Completed.UserPrincipalName)`",$($Completed.($SizeField)),$($Completed.($ItemsField)),OK,")
    }
  } finally {
    $Writer.Flush()
    $Writer.Dispose()
  }
}

# Drives per-mailbox stats gathering over a population: checkpoint read/resume/compact, a
# do/while loop with periodic EXO reconnect, per-mailbox retry via $GetStatsForMailbox, a
# consecutive-failure circuit breaker, a failures CSV, and the summary Write-Host block. Shared
# by the In Place Archive and Recoverable Items gathering blocks below so this ~100 line driver
# exists exactly once instead of being maintained twice in parallel.
#
# $GetStatsForMailbox is called as `& $GetStatsForMailbox $Mailbox $AttemptsForThisMailbox` for
# each not-yet-completed mailbox in $Population, and must return a bounded-retry result object
# (Status/Attempts/ErrorType/ErrorMessage plus the $SizeField/$ItemsField fields) - see
# Get-ArchiveMailboxStats / Get-RIFStatsForMailbox, which close over their own extra per-mailbox
# state (e.g. whether to include the archive folder) before being handed in here.
function Invoke-MailboxStatsGathering {
  param (
    # Not Mandatory, deliberately: a Mandatory [array] parameter in PowerShell rejects not only
    # $null but ALSO an empty array ("Cannot bind argument... because it is an empty collection"),
    # and a tenant with zero mailboxes in this population is a legitimate, expected value here -
    # not missing input that should fail parameter binding.
    [array]$Population,

    [Parameter(Mandatory = $true)]
    [scriptblock]$GetStatsForMailbox,

    [Parameter(Mandatory = $true)]
    [string]$Kind,

    [Parameter(Mandatory = $true)]
    [string]$SizeField,

    [Parameter(Mandatory = $true)]
    [string]$ItemsField,

    [Parameter(Mandatory = $true)]
    [string]$CheckpointHeaderTag,

    [Parameter(Mandatory = $true)]
    [string]$CheckpointFullPath,

    [Parameter(Mandatory = $true)]
    [string]$CheckpointFilenameParamName,

    [Parameter(Mandatory = $true)]
    [string]$ResumeParamName,

    [Parameter(Mandatory = $true)]
    [bool]$Resume,

    [Parameter(Mandatory = $true)]
    [int]$MaxAttempts,

    [Parameter(Mandatory = $true)]
    [string]$FailureCsvPrefix,

    [Parameter(Mandatory = $true)]
    [string]$ExportFolder,

    [Parameter(Mandatory = $true)]
    [string]$DateStamp,

    [Parameter(Mandatory = $true)]
    [string]$ScriptVersion,

    [Parameter(Mandatory = $true)]
    [double]$CapacityMetric,

    [Parameter(Mandatory = $true)]
    [string]$CapacityDisplay,

    [Parameter()]
    [int]$ConsecutiveFailureLimit = 25,

    [Parameter()]
    [int]$ReconnectInterval = 500
  )

  $MailboxList = [System.Collections.Generic.List[object]]::new()
  $FailedList = [System.Collections.Generic.List[object]]::new()
  $PopulationCount = $Population.Count
  $CurrentMailboxNum = 0
  $ProcessedCount = 0
  $ConsecutiveFailures = 0
  $RetryEnabled = $true
  $CompletedMap = @{}
  $CheckpointWriter = $null

  Write-Host "Found $PopulationCount mailboxes with $Kind" -foregroundcolor green

  # Get-ConnectionInformation returns an array when more than one session is open.
  $ExoConnection = @(Get-ConnectionInformation)[0]
  if ($ExoConnection.TenantId) {
    $TenantKey = [string]$ExoConnection.TenantId
  } elseif ($ExoConnection.Organization) {
    $TenantKey = [string]$ExoConnection.Organization
  } else {
    $TenantKey = 'unknown'
  }

  if ($Resume -eq $true) {
    $Checkpoint = Read-Checkpoint -Path $CheckpointFullPath -HeaderTag $CheckpointHeaderTag -SizeField $SizeField -ItemsField $ItemsField -Kind $Kind -FilenameParamName $CheckpointFilenameParamName -ResumeParamName $ResumeParamName -TenantKey $TenantKey -Population $PopulationCount
    $CompletedMap = $Checkpoint.Completed
    # An unreadable checkpoint must be moved aside, not appended to. Otherwise this run writes
    # data rows onto a file with no valid header, and the NEXT resume rejects the header again
    # and discards everything gathered in between.
    if (-not $Checkpoint.Valid) {
      Move-CheckpointAside -Path $CheckpointFullPath -Stamp $DateStamp -Reason "Cannot resume from $CheckpointFullPath."
      Write-Host "[WARN] Starting a fresh checkpoint. All $Kind mailboxes will be gathered." -foregroundcolor yellow
    } else {
      # Iterate this run's population, not the checkpoint, so mailboxes no longer present cannot
      # inflate the total and new mailboxes are simply gathered.
      foreach ($Mailbox in $Population) {
        $ResumeUser = $Mailbox.'User Principal Name'
        if ($CompletedMap.ContainsKey($ResumeUser)) { [void]$MailboxList.Add($CompletedMap[$ResumeUser]) }
      }
      Write-Host "[INFO] Resuming from checkpoint: $($MailboxList.Count) of $PopulationCount mailboxes already gathered."
      Compact-Checkpoint -Path $CheckpointFullPath -CompletedMap $CompletedMap -SizeField $SizeField -ItemsField $ItemsField
    }
  } elseif (Test-Path -Path $CheckpointFullPath) {
    Move-CheckpointAside -Path $CheckpointFullPath -Stamp $DateStamp -Reason "Starting a new run. Use -$ResumeParamName `$true to continue a previous one instead."
  }

  $CheckpointIsNew = -not (Test-Path -Path $CheckpointFullPath)
  # Explicit UTF-8 without a BOM, matched by the -Encoding UTF8 on the Get-Content that reads it.
  $CheckpointWriter = [System.IO.StreamWriter]::new($CheckpointFullPath, $true, [System.Text.UTF8Encoding]::new($false))
  if ($CheckpointIsNew) {
    $CheckpointWriter.WriteLine("#$CheckpointHeaderTag|1|TenantId=$TenantKey|Population=$PopulationCount|Created=$DateStamp|ScriptVersion=$ScriptVersion")
    $CheckpointWriter.WriteLine("UserPrincipalName,$SizeField,$ItemsField,Status,ErrorType")
    $CheckpointWriter.Flush()
  }
  Write-Host "Checkpoint file for this run: $CheckpointFullPath" -foregroundcolor green

  # The do/while below runs once even on an empty collection, indexing [0] and querying a null
  # identity. That is reachable whenever a tenant has no mailboxes in this population at all.
  if ($PopulationCount -gt 0) {
  try {
    do {
      if ( ($CurrentMailboxNum % 10) -eq 0 ) {
        Write-Host "[$CurrentMailboxNum / $PopulationCount] Processing mailboxes ..."
      }
      $CurrentMailbox = $Population[$CurrentMailboxNum]
      $CurrentUser = $CurrentMailbox.'User Principal Name'
      # An if-block rather than 'continue': inside do/while, continue re-tests the loop
      # condition and is an easy thing to misread in review.
      if (-not $CompletedMap.ContainsKey($CurrentUser)) {
        $AttemptsForThisMailbox = if ($RetryEnabled) { $MaxAttempts } else { 1 }
        $Stats = & $GetStatsForMailbox $CurrentMailbox $AttemptsForThisMailbox

        if ($Stats.Status -eq 'OK') {
          [void]$MailboxList.Add($Stats)
          $ConsecutiveFailures = 0
          if (-not $RetryEnabled) {
            Write-Host "[INFO] Mailbox stats are succeeding again; re-enabling retries."
            $RetryEnabled = $true
          }
        } else {
          [void]$FailedList.Add($Stats)
          Write-Host "[WARN] Could not get $Kind stats for $CurrentUser after $($Stats.Attempts) attempt(s): $($Stats.ErrorType)"
          $ConsecutiveFailures += 1
          # Bound the damage when something systemic is being misread as transient: without this,
          # a tenant-wide outage would spend 15s of backoff on every one of tens of thousands of
          # mailboxes before finishing.
          if ($RetryEnabled -and $ConsecutiveFailures -ge $ConsecutiveFailureLimit) {
            Write-Host "[WARN] $ConsecutiveFailures mailboxes failed in a row. Disabling retries to avoid stalling the run; failures are still counted." -foregroundcolor yellow
            $RetryEnabled = $false
            # With retries off, each mailbox gets a single attempt and so never reaches the
            # session-recovery reconnect inside $GetStatsForMailbox. A run of failures this long
            # is most often a dead session, so reconnect once here.
            Write-Host "[INFO] Refreshing the Exchange Online connection before continuing."
            try { Reset-ExoConnection } catch {
              Write-Host "[WARN] EXO reconnect failed: $($_.Exception.Message). Will retry on next interval."
            }
          }
        }

        $CheckpointWriter.WriteLine("`"$($Stats.UserPrincipalName)`",$($Stats.($SizeField)),$($Stats.($ItemsField)),$($Stats.Status),$($Stats.ErrorType)")
        $ProcessedCount += 1
        # Flush periodically rather than per mailbox: the durability window is 25 mailboxes,
        # against tens of thousands of fewer disk flushes on a large tenant.
        if ( ($ProcessedCount % 25) -eq 0 ) { $CheckpointWriter.Flush() }

        # Refresh the EXO connection periodically so the access token doesn't age out mid-run on
        # very large tenants. Counting processed mailboxes rather than the loop index means a
        # resume that skips thousands of completed users does not fire pointless reconnects.
        if ( ($ProcessedCount % $ReconnectInterval) -eq 0 ) {
          Write-Host "[INFO] Refreshing Exchange Online connection at $CurrentMailboxNum / $PopulationCount to avoid token timeout."
          try { Reset-ExoConnection } catch {
            Write-Host "[WARN] EXO reconnect failed: $($_.Exception.Message). Will retry on next interval."
          }
        }
      }
      $CurrentMailboxNum += 1
    } while ($CurrentMailboxNum -lt $PopulationCount)
  } finally {
    if ($null -ne $CheckpointWriter) {
      $CheckpointWriter.Flush()
      $CheckpointWriter.Dispose()
      $CheckpointWriter = $null
    }
  }
  } else {
    if ($null -ne $CheckpointWriter) {
      $CheckpointWriter.Flush()
      $CheckpointWriter.Dispose()
      $CheckpointWriter = $null
    }
  }

  $MeasurementSize = $MailboxList | Measure-Object -Property $SizeField -Sum -Average
  $MeasurementItems = $MailboxList | Measure-Object -Property $ItemsField -Sum -Average
  $ConvertedSize = [math]::Round($($MeasurementSize.Sum / $CapacityMetric), 2)
  $TotalItems = $MeasurementItems.Sum
  $SucceededCount = $MailboxList.Count
  $FailedCount = $FailedList.Count

  Write-Host "Finished gathering stats on mailboxes with $Kind" -foregroundcolor green
  Write-Host "Total # of mailboxes with ${Kind}: $PopulationCount" -foregroundcolor green
  Write-Host "Successfully gathered: $SucceededCount of $PopulationCount" -foregroundcolor green
  Write-Host "Total size of mailboxes with ${Kind}: $ConvertedSize $CapacityDisplay" -foregroundcolor green
  Write-Host "Total # of items of mailboxes with ${Kind}: $TotalItems" -foregroundcolor green

  if ($FailedCount -gt 0) {
    Write-Host ""
    Write-Host "[WARN] $FailedCount of $PopulationCount $Kind mailboxes could not be read." -foregroundcolor yellow
    Write-Host "[WARN] The $Kind and Total rows above are INCOMPLETE and UNDER-REPORT this tenant." -foregroundcolor yellow
    $FailedList | Group-Object ErrorType | Sort-Object Count -Descending | ForEach-Object {
      Write-Host "         $($_.Count) x $($_.Name)"
    }
    $FailureCSV = "$ExportFolder\$FailureCsvPrefix-$DateStamp.csv"
    $FailedList | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath $FailureCSV -Encoding UTF8
    Write-Host "[WARN] Full failure list written to: $FailureCSV" -foregroundcolor yellow
    Write-Host "[WARN] Re-run with -$ResumeParamName `$true to retry just the mailboxes that failed." -foregroundcolor yellow
  }

  return [PSCustomObject] @{
    "PopulationCount" = $PopulationCount
    "SucceededCount" = $SucceededCount
    "FailedCount" = $FailedCount
    "SizeSum" = $MeasurementSize.Sum
    "ItemsSum" = $TotalItems
    "ConvertedSize" = $ConvertedSize
  }
}

# Drives the skip/gather branch for one "kind" of per-mailbox stats collection (In Place Archive or
# Recoverable Items): the skip early-return, banner logging, checkpoint path resolution, the call
# into Invoke-MailboxStatsGathering, and the three ExchangeDetails properties common to every kind
# (succeeded count, population Found, Failed, and a boolean Skipped flag). $KindConfig groups the
# static per-kind literals (names, checkpoint header tag, etc) in one place so a rename only has to
# be made once instead of kept in sync across every named argument at each call site - this and the
# 13-ish parameters below replace what used to be ~55 lines of near-identical boilerplate PLUS a
# 15-argument Invoke-MailboxStatsGathering call duplicated at each of the two call sites.
#
# The caller remains responsible for the kind-specific storage/item properties, since their names
# and units differ historically between 'Archive Storage Used'/'Archive Items' (raw bytes) and
# 'Recoverable Items Used'/'Recoverable Items Count' (already GB-converted) - and for folding the
# gathered totals into the running Exchange-wide totals. Returns $null when skipped.
function Invoke-MailboxKindGathering {
  param (
    # Hashtable with: Kind, PropertyPrefix, SizeField, ItemsField, CheckpointHeaderTag,
    # CheckpointFilenameParamName, ResumeParamName, FailureCsvPrefix, SkipFlagParamName.
    [Parameter(Mandatory = $true)]
    [hashtable]$KindConfig,

    [Parameter(Mandatory = $true)]
    [bool]$SkipFlag,

    # Not Mandatory, deliberately: a Mandatory [array] parameter in PowerShell rejects not only
    # $null but ALSO an empty array ("Cannot bind argument... because it is an empty collection"),
    # and a tenant with zero mailboxes in this population is a legitimate, expected value here -
    # not missing input that should fail parameter binding.
    [array]$Population,

    [Parameter(Mandatory = $true)]
    [scriptblock]$GetStatsForMailbox,

    [Parameter(Mandatory = $true)]
    [string]$CheckpointFilename,

    [Parameter(Mandatory = $true)]
    [bool]$Resume,

    [Parameter(Mandatory = $true)]
    [int]$MaxAttempts,

    [Parameter(Mandatory = $true)]
    [string]$ExportFolder,

    [Parameter(Mandatory = $true)]
    [string]$DateStamp,

    [Parameter(Mandatory = $true)]
    [string]$ScriptVersion,

    [Parameter(Mandatory = $true)]
    [double]$CapacityMetric,

    [Parameter(Mandatory = $true)]
    [string]$CapacityDisplay,

    [Parameter(Mandatory = $true)]
    [PSCustomObject]$ExchangeDetails
  )

  $Kind = $KindConfig.Kind
  $PropertyPrefix = $KindConfig.PropertyPrefix

  if ($SkipFlag -eq $true) {
    Write-Host "Skipping gathering $Kind usage" -foregroundcolor green
    $ExchangeDetails | Add-Member -MemberType NoteProperty -Name $PropertyPrefix -Value "Skipped ($($Population.Count))"
    $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Found" -Value $Population.Count
    $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Failed" -Value 0
    $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Skipped" -Value $true
    return $null
  }

  Write-Host "Now gathering $Kind usage" -foregroundcolor green
  Write-Host "This may take awhile since stats need to be gathered per user" -foregroundcolor green
  Write-Host "Progress will be written as they are gathered" -foregroundcolor green
  Write-Host "If this keeps timing out, run script with -$($KindConfig.SkipFlagParamName) `$true option" -foregroundcolor green
  Write-Host "[INFO] Retrieving all Exchange Mailbox $Kind sizing"

  # StreamWriter resolves relative paths against the .NET process working directory, not the
  # PowerShell location, so the path has to be made absolute here or the checkpoint silently
  # lands somewhere the resume will not look.
  $CheckpointFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CheckpointFilename)

  $Gathering = Invoke-MailboxStatsGathering -Population $Population -GetStatsForMailbox $GetStatsForMailbox `
    -Kind $Kind -SizeField $KindConfig.SizeField -ItemsField $KindConfig.ItemsField `
    -CheckpointHeaderTag $KindConfig.CheckpointHeaderTag -CheckpointFullPath $CheckpointFullPath `
    -CheckpointFilenameParamName $KindConfig.CheckpointFilenameParamName -ResumeParamName $KindConfig.ResumeParamName -Resume $Resume `
    -MaxAttempts $MaxAttempts -FailureCsvPrefix $KindConfig.FailureCsvPrefix `
    -ExportFolder $ExportFolder -DateStamp $DateStamp -ScriptVersion $ScriptVersion `
    -CapacityMetric $CapacityMetric -CapacityDisplay $CapacityDisplay

  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name $PropertyPrefix -Value $Gathering.SucceededCount
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Found" -Value $Gathering.PopulationCount
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Failed" -Value $Gathering.FailedCount
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name "$PropertyPrefix Skipped" -Value $false

  return $Gathering
}

# No .GetNewClosure() needed: this scriptblock is created fresh right here (not reused across
# loop iterations), so it resolves $ArchiveRetryDelaySeconds/$EnableDebug/$SkipArchiveFallback/
# $ArchiveFallbackMaxAttempts via normal lexical scoping to this block. GetNewClosure() binds to a
# new, isolated session state, which can fail to resolve commands that aren't part of that fresh
# state.
$GetArchiveStatsForMailbox = {
  param($Mailbox, $AttemptsForThisMailbox)
  # Capped at $AttemptsForThisMailbox, not just $ArchiveFallbackMaxAttempts: when
  # Invoke-MailboxStatsGathering's consecutive-failure breaker has tripped, it collapses
  # $AttemptsForThisMailbox to 1 specifically so a tenant-wide outage gets one fast attempt per
  # mailbox instead of a full backoff sequence. Without this cap the fallback would still run its
  # own multi-attempt retry-with-backoff underneath that single-attempt primary call, reintroducing
  # the exact per-mailbox stall the breaker exists to avoid.
  $FallbackAttemptsForThisMailbox = [math]::Min($ArchiveFallbackMaxAttempts, $AttemptsForThisMailbox)
  Get-ArchiveMailboxStats -UserPrincipalName $Mailbox.'User Principal Name' -MaxAttempts $AttemptsForThisMailbox -RetryDelaySeconds $ArchiveRetryDelaySeconds -ShowRetryDetail $EnableDebug `
    -EnableFolderRollupFallback (-not $SkipArchiveFallback) -FallbackMaxAttempts $FallbackAttemptsForThisMailbox
}

$ArchiveKindConfig = @{
  Kind = 'In Place Archive'
  PropertyPrefix = 'Archive Mailboxes'
  SizeField = 'ArchiveSize'
  ItemsField = 'ArchiveItems'
  CheckpointHeaderTag = 'RubrikM365ArchiveCheckpoint'
  CheckpointFilenameParamName = 'ArchiveCheckpointFilename'
  ResumeParamName = 'ResumeArchive'
  FailureCsvPrefix = 'ArchiveFailures'
  SkipFlagParamName = 'SkipArchiveMailbox'
}

$ArchiveGathering = Invoke-MailboxKindGathering -KindConfig $ArchiveKindConfig -SkipFlag $SkipArchiveMailbox `
  -Population $ArchiveMailboxes -GetStatsForMailbox $GetArchiveStatsForMailbox `
  -CheckpointFilename $ArchiveCheckpointFilename -Resume $ResumeArchive -MaxAttempts $ArchiveMaxAttempts `
  -ExportFolder $ExportFolder -DateStamp $dateStringHH -ScriptVersion $Version `
  -CapacityMetric $capacityMetric -CapacityDisplay $capacityDisplay -ExchangeDetails $ExchangeDetails

if ($null -ne $ArchiveGathering) {
  # 'Archive Mailboxes' is the count actually read, not the population: it is the denominator of
  # the per-mailbox average in the HTML report, and dividing a sum over successes by the full
  # population understates every archive. The two are identical on a clean run.
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Storage Used' -Value $ArchiveGathering.SizeSum
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Items' -Value $ArchiveGathering.ItemsSum
  $ExchangeTotalStorage = $ExchangeDetails.'Total Storage Used' + $ArchiveGathering.SizeSum
  $ExchangeDetails.'Total Storage Used' = $ExchangeTotalStorage
  $ExchangeTotalItems = $ExchangeDetails.'Total Items' + $ArchiveGathering.ItemsSum
  $ExchangeDetails.'Total Items' = $ExchangeTotalItems
} else {
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Storage Used' -Value '-'
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Items' -Value '-'
}

# The Microsoft Exchange Reports do not contain Recoverable Items sizing information.
# We need to connect to the Exchange Online module to get this information

# A single pass over every active user (shared + non-shared), deciding per-user whether to also
# read the In-Place Archive Recoverable Items folder from 'Has Archive', instead of the old
# two-population approach (Get-RIFMailboxStats called once per population) - so checkpoint/resume
# and connection-refresh state are shared across the whole run.
# @(...) guarantees a real, possibly-empty array reaches Invoke-MailboxKindGathering's
# -Population parameter even if $ExchangeActiveUsers ever collapsed to $null (e.g. an empty tenant).
$RIFMailboxes = @($ExchangeActiveUsers)

# No .GetNewClosure() needed: this scriptblock is created fresh right here (not reused across
# loop iterations), so it resolves $RIFRetryDelaySeconds/$EnableDebug via normal lexical scoping
# to this block. GetNewClosure() binds to a new, isolated session state, which can fail to
# resolve commands that aren't part of that fresh state.
$GetRIFStatsForMailbox = {
  param($Mailbox, $AttemptsForThisMailbox)
  $IncludeArchiveMailbox = ($Mailbox.'Has Archive' -eq 'TRUE')
  Get-RIFStatsForMailbox -UserPrincipalName $Mailbox.'User Principal Name' -IncludeArchiveMailbox $IncludeArchiveMailbox -MaxAttempts $AttemptsForThisMailbox -RetryDelaySeconds $RIFRetryDelaySeconds -ShowRetryDetail $EnableDebug
}

$RIFKindConfig = @{
  Kind = 'Recoverable Items'
  PropertyPrefix = 'Recoverable Items'
  SizeField = 'RIFSize'
  ItemsField = 'RIFItems'
  CheckpointHeaderTag = 'RubrikM365RIFCheckpoint'
  CheckpointFilenameParamName = 'RIFCheckpointFilename'
  ResumeParamName = 'ResumeRIF'
  FailureCsvPrefix = 'RIFFailures'
  SkipFlagParamName = 'SkipRecoverableItems'
}

$RIFGathering = Invoke-MailboxKindGathering -KindConfig $RIFKindConfig -SkipFlag $SkipRecoverableItems `
  -Population $RIFMailboxes -GetStatsForMailbox $GetRIFStatsForMailbox `
  -CheckpointFilename $RIFCheckpointFilename -Resume $ResumeRIF -MaxAttempts $RIFMaxAttempts `
  -ExportFolder $ExportFolder -DateStamp $dateStringHH -ScriptVersion $Version `
  -CapacityMetric $capacityMetric -CapacityDisplay $capacityDisplay -ExchangeDetails $ExchangeDetails

if ($null -ne $RIFGathering) {
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Recoverable Items Used' -Value $RIFGathering.ConvertedSize
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Recoverable Items Count' -Value $RIFGathering.ItemsSum
  $ExchangeTotalStorage = $ExchangeDetails.'Total Storage Used' + $RIFGathering.SizeSum
  $ExchangeDetails.'Total Storage Used' = $ExchangeTotalStorage
  $ExchangeTotalItems = $ExchangeDetails.'Total Items' + $RIFGathering.ItemsSum
  $ExchangeDetails.'Total Items' = $ExchangeTotalItems
} else {
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Recoverable Items Used' -Value '-'
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Recoverable Items Count' -Value '-'
}

Write-Host "Calculating # of license needed:"
Write-Host "Exchange user mailboxes: $($ExchangeDetails.'User Mailboxes')"
Write-Host "Exchange shared mailboxes: $($ExchangeDetails.'Shared Mailboxes')"
Write-Host "OneDrive chart accounts: $($OneDriveDetails.'Chart Accounts')"
[int]$UserLicensesRequired = $($ExchangeDetails.'User Mailboxes')
if ($SkipSharedMailbox -eq $false -and [int]$ExchangeDetails.'Shared Mailboxes' -gt $UserLicensesRequired) {
  $UserLicensesRequired = $ExchangeDetails.'Shared Mailboxes'
}
if ([int]$OneDriveDetails.'Accounts' -gt $UserLicensesRequired) {
  $UserLicensesRequired = $OneDriveDetails.'Accounts'
}
Write-Host "# of licenses required: $UserLicensesRequired" -foregroundcolor green

$totalStorage = $ExchangeDetails.'Total Storage Used' + $OneDriveDetails.'Storage Used' +
  $SharePointDetails.'Sites Storage Used'
$totalItems = $ExchangeDetails.'Total Items' + $OneDriveDetails.'Total Files' +
  $SharePointDetails.'Account Files'

$totalStorageGB = [math]::round($totalStorage / $capacityMetric, 2)

#region HTML Code for Output
$HTML_CODE = @"
<!DOCTYPE html>

<html>
<!---->
<!---->
<link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">

<head>
    <style>
        body {
            background-color: #f4f4f4
        }

        .card-container {
            display: flex;
            width: 100%;
            align-items: center;
            justify-content: center;
            padding-bottom: 20px;
        }

        .card-header {
            display: flex;
        }

        .card-header-logo {
            flex-grow: 1;
        }

        .rubrik-snowflake {
            padding-top: 15px;
        }

        .card-header-text {
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.9rem;
            line-height: 2.4rem;
        }

        .navigation-bar {
            display: flex;
            background-color: #060745;
            width: 100%;
            top: 0;
            left: 0;
            position: fixed;
            max-height: 82px;
        }

        .logo {
            padding-top: 20px;
            padding-left: 10px;
            display: block;
            max-width: 150px;

        }

        .nav-bar-text {
            padding-top: 12px;
            flex-grow: 1;
            display: flex;
            color: white;
            align-items: center;
            justify-content: center;
            font-size: 1.9rem;
            line-height: 2.4rem;
            margin-bottom: 20px;
            margin-top: 0;

        }

        .rubrik-logo path {
            fill: white;
        }

        .margin {
            padding-bottom: 130px;
        }

        .card {
            box-shadow: 0 4px 10px 0 rgb(0 0 0 / 20%), 0 4px 20px 0 rgb(0 0 0 / 19%);
            width: 98%;
            padding: 0.01em 16px;

        }

        .styled-table {
            margin: 25px 0;
            width: 100%;

        }

        .styled-table thead tr {
            text-align: left;
        }

        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
        }
    </style>
</head>

<body>
    <div class="navigation-bar">
        <div class="logo">
            <svg class="rubrik-logo" width=auto height="82">
                <defs>
                    <style>
                        .cls-1 {
                            fill: #fff
                        }

                        .cls-1,
                        .cls-2 {
                            fill-rule: evenodd
                        }
                    </style>
                    <mask id="mask" x="13.3" y="0" width="12.35" height="12.27" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-2">
                                <path id="path-1" class="cls-1"
                                    d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-2-2" x="13.3" y="26.53" width="12.35" height="12.25" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-4">
                                <path id="path-3" class="cls-1"
                                    d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-3" x="26.6" y="13.22" width="12.35" height="12.32" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-6">
                                <path id="path-5" class="cls-1"
                                    d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-4-2" x="9.63" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-8">
                                <path id="path-7" class="cls-1"
                                    d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-5" x="26.15" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-10">
                                <path id="path-9" class="cls-1"
                                    d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-6-2" x="26.15" y="26.04" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-12">
                                <path id="path-11" class="cls-1"
                                    d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-7" x="33.32" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-14">
                                <path id="path-13" class="cls-1"
                                    d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-8-2" x="26.15" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-16">
                                <path id="path-15" class="cls-1"
                                    d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-9" x="1.05" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-18">
                                <path id="path-17" class="cls-1"
                                    d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-10-2" x="6.31" y="6.25" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-20">
                                <path id="path-19" class="cls-1"
                                    d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-11" x="9.63" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-22">
                                <path id="path-21" class="cls-1"
                                    d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-12-2" x="33.32" y="26.04" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-24">
                                <path id="path-23" class="cls-1"
                                    d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-13" x="6.31" y="26.04" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-26">
                                <path id="path-25" class="cls-1"
                                    d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-14-2" x="1.05" y="26.04" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-28">
                                <path id="path-27" class="cls-1"
                                    d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-15" x="26.15" y="6.25" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-30">
                                <path id="path-29" class="cls-1"
                                    d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-16-2" x="0" y="13.22" width="12.35" height="12.32" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-32">
                                <path id="path-31" class="cls-1"
                                    d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24">
                                </path>
                            </g>
                        </g>
                    </mask>
                </defs>
                <g id="Symbols">
                    <g class="svgName">
                        <path class="name r" id="Fill-57"
                            d="M58 12.6c-1.58 0-2.29.43-3.74 2.16V14c0-.91-.12-1-1-1h-.74c-.91 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.74c.91 0 1-.12 1-1V20.7a8.24 8.24 0 0 1 .63-3.93A3.06 3.06 0 0 1 58 15.31a3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .54.54 0 0 0 .24-.21 4.5 4.5 0 0 0 .39-.67l.23-.45a2.24 2.24 0 0 0 .28-.67c0-.51-1-.94-2.24-.94"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name u" id="Fill-59"
                            d="M66.09 22.5a6.61 6.61 0 0 0 .51 3.07 3.87 3.87 0 0 0 6.34 0 6.61 6.61 0 0 0 .51-3.07V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1v8.8c0 2.39-.39 3.69-1.49 4.91a7.1 7.1 0 0 1-10 0c-1.1-1.22-1.49-2.52-1.49-4.91V14c0-.91.11-1 1-1h.75c.9 0 1 .12 1 1z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name b" id="Fill-61"
                            d="M83.42 21.13c0 3.61 2.24 6.09 5.47 6.09s5.35-2.6 5.35-6.17a5.54 5.54 0 0 0-5.39-5.86c-3.19 0-5.43 2.44-5.43 5.94zm.2-5.82a7 7 0 0 1 5.7-2.67c4.49 0 7.79 3.58 7.79 8.49s-3.34 8.64-7.87 8.64a6.89 6.89 0 0 1-5.62-2.71v1.22c0 .9-.12 1-1 1h-.74c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.74c.91 0 1 .12 1 1z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name r" id="Fill-55"
                            d="M107.72 12.6c-1.57 0-2.28.43-3.74 2.16V14c0-.91-.12-1-1-1h-.75c-.9 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.77c.9 0 1-.12 1-1V20.7a8.37 8.37 0 0 1 .63-3.93 3.07 3.07 0 0 1 3.1-1.46 3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .63.63 0 0 0 .25-.21 4.44 4.44 0 0 0 .38-.67l.24-.45a2.42 2.42 0 0 0 .27-.67c0-.51-1-.94-2.24-.94"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name i" id="Fill-63"
                            d="M116.4 28.28c0 .9-.12 1-1 1h-.75c-.9 0-1-.12-1-1V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1zm.6-21.45a2 2 0 1 1-2-2 2 2 0 0 1 2 2z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name k" id="Fill-65"
                            d="M129.84 13.47c.47-.48.47-.48 1.14-.48h1.22c.71 0 1 .2 1 .63 0 .16-.15.4-.47.71L127.08 20l7.13 8c.27.36.43.59.43.75 0 .39-.32.59-1 .59h-1.24c-.71 0-.71 0-1.14-.51l-6.14-6.92-.71.71v5.7c0 .9-.12 1-1 1h-.75c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.75c.9 0 1 .12 1 1V19z"
                            transform="translate(-.31 -.22)"></path>
                    </g>
                    <g class="svgLogo">
                        <g mask="url(#mask)">
                            <path id="Fill-68" class="cls-2"
                                d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-2-2)">
                            <path id="Fill-71" class="cls-2"
                                d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-3)">
                            <path id="Fill-74" class="cls-2"
                                d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-4-2)">
                            <path id="Fill-77" class="cls-2"
                                d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-5)">
                            <path id="Fill-80" class="cls-2"
                                d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-6-2)">
                            <path id="Fill-83" class="cls-2"
                                d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-7)">
                            <path id="Fill-86" class="cls-2"
                                d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-8-2)">
                            <path id="Fill-89" class="cls-2"
                                d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-9)">
                            <path id="Fill-92" class="cls-2"
                                d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-10-2)">
                            <path id="Fill-95" class="cls-2"
                                d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-11)">
                            <path id="Fill-98" class="cls-2"
                                d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-12-2)">
                            <path id="Fill-101" class="cls-2"
                                d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-13)">
                            <path id="Fill-104" class="cls-2"
                                d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-14-2)">
                            <path id="Fill-107" class="cls-2"
                                d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-15)">
                            <path id="Fill-110" class="cls-2"
                                d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-16-2)">
                            <path id="Fill-113" class="cls-2"
                                d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                    </g>
                    <path id="Fill-116" class="cls-2"
                        d="M134.82 13.78h.09c.1 0 .18 0 .18-.12s-.05-.12-.17-.12h-.1zm0 .45h-.18v-.8h.3a.46.46 0 0 1 .28.06.21.21 0 0 1 .08.17.21.21 0 0 1-.17.19c.08 0 .12.08.15.19a.51.51 0 0 0 .06.2h-.2a.47.47 0 0 1-.06-.2.15.15 0 0 0-.17-.12h-.09zm-.49-.42a.62.62 0 0 0 .62.64.64.64 0 0 0 0-1.27.62.62 0 0 0-.62.63zm1.43 0a.8.8 0 0 1-.81.81.81.81 0 0 1-.82-.81.8.8 0 0 1 .82-.79.79.79 0 0 1 .81.79z"
                        transform="translate(-.31 -.22)"></path>
                </g>
            </svg>
        </div>
        <div class="nav-bar-text">Microsoft 365 Sizing ($dateString)</div>
    </div>
    <div class="margin"></div>

    <!-- Exchange Mailbox -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="62" width="auto" viewBox="-8.24997 -12 71.49974 72">
                        <path fill="#28a8ea"
                            d="M51.5095 0h-12.207a3.4884 3.4884 0 00-2.4677 1.0225L8.0222 29.835a3.4884 3.4884 0 00-1.0224 2.4677v12.207A3.49 3.49 0 0010.49 48h12.207a3.4884 3.4884 0 002.4678-1.0225l28.813-28.8125a3.49 3.49 0 001.022-2.4677V3.4903A3.49 3.49 0 0051.5095 0z" />
                        <path fill="#0078d4"
                            d="M51.5098 48H39.3025a3.49 3.49 0 01-2.4678-1.0222l-5.835-5.835V30.24a6.24 6.24 0 016.24-6.24h10.903l5.8349 5.835a3.49 3.49 0 011.0222 2.4678V44.51a3.49 3.49 0 01-3.49 3.49z" />
                        <path fill="#50d9ff"
                            d="M10.4898 0H22.697a3.49 3.49 0 012.4678 1.0222l5.835 5.835V17.76a6.24 6.24 0 01-6.24 6.24H13.8569l-5.835-5.835a3.49 3.49 0 01-1.0221-2.4677V3.49a3.49 3.49 0 013.49-3.49z" />
                        <path opacity=".2"
                            d="M28.9998 12.33v26.34a1.7344 1.7344 0 01-.04.3998A2.3138 2.3138 0 0126.6697 41h-19.67V10h19.67a2.326 2.326 0 012.33 2.33z" />
                        <path opacity=".1"
                            d="M29.9998 12.33v24.34A3.3617 3.3617 0 0126.6697 40h-19.67V9h19.67a3.3418 3.3418 0 013.33 3.33z" />
                        <path opacity=".2"
                            d="M28.9998 12.33v24.34A2.326 2.326 0 0126.6697 39h-19.67V10h19.67a2.326 2.326 0 012.33 2.33z" />
                        <path opacity=".1"
                            d="M27.9998 12.33v24.34A2.326 2.326 0 0125.6697 39h-18.67V10h18.67a2.326 2.326 0 012.33 2.33z" />
                        <rect fill="#0078d4" rx="2.3333" height="28" width="28" y="10" />
                        <path fill="#fff"
                            d="M18.5851 18.8812H12.038v3.8286h6.1454v2.4537H12.038v3.9766h6.8961v2.4434H9.066v-15.167h9.5191z" />
                    </svg>
                </div>
                <div class="card-header-text">
                    Exchange Online (AD Group: $ADGroup)
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Mailbox Type</th>
                        <th>Count</th>
                        <th>Size ($capacityDisplay)</th>
                        <th>Items</th>
                        <th>Per Mailbox Size ($capacityDisplay)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>User Mailboxes</td>
                        <td>$($ExchangeDetails.'User Mailboxes')</td>
                        <td>$([math]::round($ExchangeDetails.'User Storage Used No Archive' / $capacityMetric, 2))</td>
                        <td>$($ExchangeDetails.'User Items No Archive')</td>
                        <td>$([math]::round($ExchangeDetails.'User Storage Used No Archive' / $capacityMetric / $ExchangeDetails.'User Mailboxes', 2))</td>
                    </tr>
                    <tr>
                        <td>Shared Mailboxes</td>
                        <td>$($ExchangeDetails.'Shared Mailboxes')</td>
                        <td>$(if ($ExchangeDetails.'Shared Mailboxes' -eq 'Skipped') { 'Skipped' } else { [math]::round($ExchangeDetails.'Shared Storage Used No Archive' / $capacityMetric, 2) })</td>
                        <td>$(if ($ExchangeDetails.'Shared Mailboxes' -eq 'Skipped') { 'Skipped' } else { $ExchangeDetails.'Shared Items No Archive' })</td>
                        <td>$(if ($ExchangeDetails.'Shared Mailboxes' -eq 'Skipped') { 'Skipped' } else { [math]::round($ExchangeDetails.'Shared Storage Used No Archive' / $capacityMetric / $ExchangeDetails.'Shared Mailboxes', 2) })</td>
                    </tr>
                    <tr>
                        <td>Archive Mailboxes</td>
                        <td>$(if ($ExchangeDetails.'Archive Mailboxes Skipped') { 'Skipped' } elseif ([int]$ExchangeDetails.'Archive Mailboxes Failed' -gt 0) { "$($ExchangeDetails.'Archive Mailboxes') of $($ExchangeDetails.'Archive Mailboxes Found') found" } else { $ExchangeDetails.'Archive Mailboxes' })</td>
                        <td>$(if ($ExchangeDetails.'Archive Mailboxes Skipped') { 'Skipped' } else { [math]::round($ExchangeDetails.'Archive Storage Used' / $capacityMetric, 2) })</td>
                        <td>$(if ($ExchangeDetails.'Archive Mailboxes Skipped') { 'Skipped' } else { $ExchangeDetails.'Archive Items' })</td>
                        <td>$(if ($ExchangeDetails.'Archive Mailboxes Skipped') { 'Skipped' } elseif ([int]$ExchangeDetails.'Archive Mailboxes' -le 0) { 0 } else { [math]::round($ExchangeDetails.'Archive Storage Used' / $capacityMetric / $ExchangeDetails.'Archive Mailboxes', 2) })</td>
                    </tr>
                    <tr>
                        <td>Recoverable Items</td>
                        <td>$(if ($ExchangeDetails.'Recoverable Items Skipped') { 'Skipped' } elseif ([int]$ExchangeDetails.'Recoverable Items Failed' -gt 0) { "$($ExchangeDetails.'Recoverable Items') of $($ExchangeDetails.'Recoverable Items Found') found" } else { $ExchangeDetails.'Recoverable Items' })</td>
                        <td>$(if ($ExchangeDetails.'Recoverable Items Skipped') { 'Skipped' } else { $ExchangeDetails.'Recoverable Items Used' })</td>
                        <td>$(if ($ExchangeDetails.'Recoverable Items Skipped') { 'Skipped' } else { $ExchangeDetails.'Recoverable Items Count' })</td>
                        <td>$(if ($ExchangeDetails.'Recoverable Items Skipped') { 'Skipped' } elseif ([int]$ExchangeDetails.'Recoverable Items' -le 0) { 0 } else { [math]::round($ExchangeDetails.'Recoverable Items Used' / $ExchangeDetails.'Recoverable Items', 2) })</td>
                    </tr>
                    <tr style="font-weight: bold; background-color: #f2f2f2;">
                        <td>Total</td>
                        <td>$($ExchangeDetails.'Total Mailboxes')</td>
                        <td>$([math]::round($ExchangeDetails.'Total Storage Used' / $capacityMetric, 2))</td>
                        <td>$($ExchangeDetails.'Total Items')</td>
                        <td>$([math]::round($ExchangeDetails.'Total Storage Used' / $capacityMetric / $ExchangeDetails.'Total Mailboxes', 2))</td>
                    </tr>
                </tbody>
            </table>
            $(if ([int]$ExchangeDetails.'Archive Mailboxes Failed' -gt 0) { "<p style='color:#b00020;'>Note: $($ExchangeDetails.'Archive Mailboxes Failed') of $($ExchangeDetails.'Archive Mailboxes Found') In-Place Archive mailboxes could not be read. The Archive and Total rows above under-report actual usage. See the ArchiveFailures CSV produced alongside this report.</p>" })
            $(if ([int]$ExchangeDetails.'Recoverable Items Failed' -gt 0) { "<p style='color:#b00020;'>Note: $($ExchangeDetails.'Recoverable Items Failed') of $($ExchangeDetails.'Recoverable Items Found') Recoverable Items mailboxes could not be read. The Recoverable Items and Total rows above under-report actual usage. See the RIFFailures CSV produced alongside this report.</p>" })
        </div>
    </div>

    <!-- OneDrive -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="62" width="auto"
                        viewBox="-154.5063 -164.9805 1339.0546 989.883">
                        <path
                            d="M622.292 445.338l212.613-203.327C790.741 69.804 615.338-33.996 443.13 10.168a321.9 321.9 0 00-188.92 134.837c3.29-.083 368.082 300.333 368.082 300.333z"
                            fill="#0364B8" />
                        <path
                            d="M392.776 183.283l-.01.035A256.233 256.233 0 00257.5 144.921c-1.104 0-2.189.07-3.29.083C112.063 146.765-1.74 263.424.02 405.567a257.389 257.389 0 0046.244 144.04l318.528-39.894 244.21-196.915z"
                            fill="#0078D4" />
                        <path
                            d="M834.905 242.012c-4.674-.312-9.37-.528-14.123-.528a208.464 208.464 0 00-82.93 17.117l-.006-.022-128.844 54.22 142.041 175.456 253.934 61.728c54.8-101.732 16.752-228.625-84.98-283.424a209.23 209.23 0 00-85.09-24.546z"
                            fill="#1490DF" />
                        <path
                            d="M46.264 549.607C94.36 618.757 173.27 659.967 257.5 659.922h563.281c76.946.022 147.691-42.202 184.195-109.937L609.001 312.798z"
                            fill="#28A8EA" />
                    </svg>
                </div>
                <div class="card-header-text">
                    OneDrive (AD Group: $ADGroup)
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Number of User OneDrives</th>
                        <th>Total Size ($capacityDisplay)</th>
                        <th>Total # of Files</th>
                        <th>Per Account Size ($capacityDisplay)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($OneDriveDetails.'Accounts')</td>
                        <td>$([math]::round($OneDriveDetails.'Storage Used' / $capacityMetric, 2))</td>
                        <td>$($OneDriveDetails.'Total Files')</td>
                        <td>$([math]::round($OneDriveDetails.'Storage Used' / $capacityMetric / $OneDriveDetails.'Accounts', 2) )</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- SharePoint -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="82" width="auto"
                        viewBox="-298.8501 -486.5 2590.0342 2919">
                        <circle r="556" cy="556" cx="1019.333" fill="#036C70" />
                        <circle r="509.667" cy="1065.667" cx="1482.667" fill="#1A9BA1" />
                        <circle r="393.833" cy="1552.167" cx="1088.833" fill="#37C6D0" />
                        <path
                            d="M1112 501.79v988.753c-.23 34.357-21.05 65.222-52.82 78.303a82.12 82.12 0 01-31.97 6.487H695.463c-.463-7.877-.463-15.29-.463-23.167-.154-7.734.155-15.47.927-23.167 8.48-148.106 99.721-278.782 235.837-337.77v-86.18c-302.932-48.005-509.592-332.495-461.587-635.427.333-2.098.677-4.195 1.034-6.289a391.8 391.8 0 019.73-46.333h546.27c46.753.178 84.611 38.036 84.789 84.79z"
                            opacity=".1" />
                        <path
                            d="M980.877 463.333H471.21c-51.486 302.386 151.908 589.256 454.293 640.742a555.466 555.466 0 0027.573 3.986c-143.633 68.11-248.3 261.552-257.196 420.938a193.737 193.737 0 00-.927 23.167c0 7.877 0 15.29.463 23.167a309.212 309.212 0 006.023 46.333h279.39c34.357-.23 65.222-21.05 78.303-52.82a82.098 82.098 0 006.487-31.97V548.123c-.176-46.736-38.006-84.586-84.742-84.79z"
                            opacity=".2" />
                        <path
                            d="M980.877 463.333H471.21c-51.475 302.414 151.95 589.297 454.364 640.773a556.017 556.017 0 0018.607 2.844c-139 73.021-239.543 266-248.254 422.05h284.95c46.681-.353 84.437-38.109 84.79-84.79V548.123c-.178-46.754-38.036-84.612-84.79-84.79z"
                            opacity=".2" />
                        <path
                            d="M934.543 463.333H471.21c-48.606 285.482 130.279 560.404 410.977 631.616A765.521 765.521 0 00695.927 1529h238.617c46.754-.178 84.612-38.036 84.79-84.79V548.123c-.026-46.817-37.973-84.764-84.791-84.79z"
                            opacity=".2" />
                        <linearGradient gradientTransform="matrix(1 0 0 -1 0 1948)" y2="398.972" x2="842.255"
                            y1="1551.028" x1="177.079" gradientUnits="userSpaceOnUse" id="a">
                            <stop offset="0" stop-color="#058f92" />
                            <stop offset=".5" stop-color="#038489" />
                            <stop offset="1" stop-color="#026d71" />
                        </linearGradient>
                        <path
                            d="M84.929 463.333h849.475c46.905 0 84.929 38.024 84.929 84.929v849.475c0 46.905-38.024 84.929-84.929 84.929H84.929c-46.905 0-84.929-38.024-84.929-84.929V548.262c0-46.905 38.024-84.929 84.929-84.929z"
                            fill="url(#a)" />
                        <path
                            d="M379.331 962.621a156.785 156.785 0 01-48.604-51.384 139.837 139.837 0 01-16.912-70.288 135.25 135.25 0 0131.46-91.045 185.847 185.847 0 0183.678-54.581 353.459 353.459 0 01114.304-17.699 435.148 435.148 0 01150.583 21.082v106.567a235.031 235.031 0 00-68.11-27.8 331.709 331.709 0 00-79.647-9.545 172.314 172.314 0 00-81.871 17.329 53.7 53.7 0 00-32.433 49.206 49.853 49.853 0 0013.9 34.843 124.638 124.638 0 0037.067 26.503c15.444 7.691 38.611 17.916 69.5 30.673a70.322 70.322 0 019.915 3.985 571.842 571.842 0 0187.663 43.229 156.935 156.935 0 0151.801 52.171 151.223 151.223 0 0118.533 78.767 146.506 146.506 0 01-29.468 94.798 164.803 164.803 0 01-78.767 53.005 357.22 357.22 0 01-112.312 16.309 594.113 594.113 0 01-101.933-8.34 349.057 349.057 0 01-82.612-24.279v-112.358a266.237 266.237 0 0083.4 39.847 326.268 326.268 0 0092.018 14.734 158.463 158.463 0 0083.4-17.699 55.971 55.971 0 0028.449-49.994 53.284 53.284 0 00-15.753-38.271 158.715 158.715 0 00-43.414-30.256c-18.533-9.267-45.824-21.483-81.871-36.65a465.328 465.328 0 01-81.964-42.859z"
                            fill="#FFF" />
                    </svg>
                </div>
                <div class="card-header-text">
                    SharePoint
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Total Sites</th>
                        <th>Total Size ($capacityDisplay)</th>
                        <th>Total # of Files</th>
                        <th>Per Site Size ($capacityDisplay)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($SharePointDetails.'Sites')</td>
                        <td>$([math]::round($SharePointDetails.'Sites Storage Used' / $capacityMetric, 2))</td>
                        <td>$($SharePointDetails.'Account Files')</td>
                        <td>$([math]::round($SharePointDetails.'Sites Storage Used' / $capacityMetric / $SharePointDetails.'Sites', 2) )</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Total Data Needed -->
    <div class="card-container">
        <div class="card">
            <div class="card-header ">
                <div class="M365">
                    <svg xmlns="http://www.w3.org/2000/svg" height="72" width="72" viewBox="-8 -35000 278050 403334" shape-rendering="geometricPrecision" text-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd" clip-rule="evenodd">
                    <path fill="#ea3e23" d="M278050 305556l-29-16V28627L178807 0 448 66971l-448 87 22 200227 60865-23821V80555l117920-28193-17 239519L122 267285l178668 65976v73l99231-27462v-316z"/></svg>
                </div>
                <div class="card-header-text">
                    Discovery Summary
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Total users or accounts</th>
                        <th>Total Size ($capacityDisplay)</th>
                        <th>Total # of Items & Files</th>
                        <th>Per User/Account Size ($capacityDisplay)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$UserLicensesRequired</td>
                        <td>$totalStorageGB</td>
                        <td>$totalItems</td>
                        <td>$([math]::round($totalStorageGB / $UserLicensesRequired, 2))</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <footer>
        <p style="color:#D3D3D3;text-align:right;padding-right: 10px;"<td>$CurrentDate $Version</td>
    </footer>
</body>
</html>
"@
#endregion

# Remove any previously created files
Remove-Item -Path $outFilename -ErrorAction SilentlyContinue
Write-Output $HTML_CODE | Format-Table -AutoSize | Out-File -FilePath $outFilename -Append

Write-Host "`n`nM365 Sizing information has been written to $((Get-ChildItem $outFilename).FullName)`n`n" -foregroundcolor green
Write-Host "Thank you for running this script. Please send the .html file to Rubrik." -foregroundcolor green
