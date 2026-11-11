<#
.SYNOPSIS
    Offline test harness for Get-AzureAdGroupMemberUpn (-ADGroup / -ExcludeADGroup resolution).
.DESCRIPTION
    Needs no M365 tenant and no modules. It parses Get-RubrikM365SizingInfo.ps1 (without running
    it), extracts Get-AzureAdGroupMemberUpn, stubs the two Graph cmdlets it calls, and checks:
      - include group -> returns the user UPNs and writes them to -CSVFilename
      - non-user members (groups, devices, service principals) are skipped
      - exclude group with no users (-AllowEmpty) -> warns, writes an empty CSV, returns nothing
      - include group with no users -> throws
      - unknown group -> throws
      - every call to Get-AzureAdGroupMemberUpn in the script binds to declared parameters
      - no call to a built-in cmdlet relies on a parameter alias that Windows PowerShell 5.1
        does not have (the original bug: Out-File -Path, a PowerShell 7-only alias)
      - the script is pure ASCII, so 5.1 reads it correctly without a BOM
    Out-File is wrapped by a shim that accepts only the Windows PowerShell 5.1 parameter set,
    so the behavioural checks fail on PowerShell 7 the same way the real script failed on 5.1.
    Exit code is 0 when all checks pass, 1 otherwise.
.EXAMPLE
    pwsh -NoProfile -File ./tests/Test-ADGroupMemberUpn.ps1
.EXAMPLE
    powershell.exe -NoProfile -File .\tests\Test-ADGroupMemberUpn.ps1
#>
[CmdletBinding()]
param (
  [string]$ScriptPath = (Join-Path (Join-Path $PSScriptRoot '..') 'Get-RubrikM365SizingInfo.ps1')
)

$ErrorActionPreference = 'Stop'
$Failures = 0
function Assert-Equal($Name, $Actual, $Expected) {
  if ($Actual -ceq $Expected) { Write-Host "[PASS] $Name" -ForegroundColor Green }
  else { $script:Failures++; Write-Host "[FAIL] $Name`n   expected: $Expected`n   actual:   $Actual" -ForegroundColor Red }
}
function Assert-True($Name, $Condition) {
  if ($Condition) { Write-Host "[PASS] $Name" -ForegroundColor Green }
  else { $script:Failures++; Write-Host "[FAIL] $Name" -ForegroundColor Red }
}
function Assert-Throws($Name, [scriptblock]$Block, $Pattern) {
  try { & $Block | Out-Null; $script:Failures++; Write-Host "[FAIL] $Name (did not throw)" -ForegroundColor Red }
  catch {
    if ($_.Exception.Message -match $Pattern) { Write-Host "[PASS] $Name" -ForegroundColor Green }
    else { $script:Failures++; Write-Host "[FAIL] $Name`n   threw: $($_.Exception.Message)" -ForegroundColor Red }
  }
}

$Tokens = $null; $Errors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $ScriptPath), [ref]$Tokens, [ref]$Errors)
Assert-True 'Script parses without syntax errors' ($Errors.Count -eq 0)
$Errors | ForEach-Object { Write-Host "   $($_.Message) (line $($_.Extent.StartLineNumber))" -ForegroundColor Red }

$Fn = $Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-AzureAdGroupMemberUpn' }, $true)
Assert-True 'Get-AzureAdGroupMemberUpn is defined' ($null -ne $Fn)
if (-not $Fn) { exit 1 }
. ([scriptblock]::Create($Fn.Extent.Text))

# --- Windows PowerShell 5.1 Out-File shim -------------------------------------------------------
# Exactly the 5.1 parameters (no aliases on -FilePath), forwarded to the real cmdlet.
function Out-File {
  [CmdletBinding(DefaultParameterSetName = 'ByPath')]
  param (
    [Parameter(ParameterSetName = 'ByPath', Mandatory, Position = 0)] [string]$FilePath,
    [Parameter(ParameterSetName = 'ByLiteralPath', Mandatory)] [Alias('PSPath')] [string]$LiteralPath,
    [Parameter(Position = 1)] [string]$Encoding,
    [switch]$Append, [switch]$Force, [Alias('NoOverwrite')] [switch]$NoClobber,
    [int]$Width, [switch]$NoNewline,
    [Parameter(ValueFromPipeline)] [psobject]$InputObject
  )
  begin { $Items = [System.Collections.Generic.List[object]]::new() }
  process { if ($PSBoundParameters.ContainsKey('InputObject')) { $Items.Add($InputObject) } }
  end {
    $Target = if ($FilePath) { $FilePath } else { $LiteralPath }
    $Items | Microsoft.PowerShell.Utility\Out-File -LiteralPath $Target
  }
}
# Sanity check on the shim itself: if it accepted -Path, the checks below would prove nothing.
Assert-Throws 'shim rejects Out-File -Path, as Windows PowerShell 5.1 does' {
  'x' | Out-File -Path (Join-Path ([IO.Path]::GetTempPath()) 'shim-probe.txt')
} "parameter name 'Path'"

# --- Graph stubs ---------------------------------------------------------------------------------
function New-Member($Type, $Upn) {
  $Props = @{ '@odata.type' = "#microsoft.graph.$Type" }
  if ($Upn) { $Props['userPrincipalName'] = $Upn }
  [pscustomobject]@{ AdditionalProperties = $Props }
}
$script:Groups = @{
  'Sales'       = @{ Id = 'g-sales'; Members = @(
      (New-Member 'user' 'alice@contoso.com'),
      (New-Member 'group' $null),
      (New-Member 'user' 'bob_fabrikam.com#EXT#@contoso.onmicrosoft.com'),
      (New-Member 'device' $null),
      (New-Member 'servicePrincipal' $null)) }
  'EmptyGroup'  = @{ Id = 'g-empty'; Members = @((New-Member 'group' $null)) }
}
function Get-MgGroup {
  param ([string]$Filter)
  if ($Filter -notmatch "^DisplayName eq '(.*)'$") { throw "unexpected filter: $Filter" }
  $G = $script:Groups[$Matches[1]]
  if ($G) { @([pscustomobject]@{ Id = $G.Id; DisplayName = $Matches[1] }) } else { @() }
}
function Get-MgGroupTransitiveMember {
  param ([string]$GroupId, [switch]$All)
  if (-not $All) { throw 'Get-MgGroupTransitiveMember called without -All (results would be paged)' }
  ($script:Groups.Values | Where-Object { $_.Id -eq $GroupId }).Members
}

$TempDir = Join-Path ([IO.Path]::GetTempPath()) ('adgroup-upn-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $TempDir | Out-Null
try {
  # --- include group ---
  $Csv = Join-Path $TempDir 'adgrouplist.csv'
  try { $Result = @(Get-AzureAdGroupMemberUpn -GroupName 'Sales' -CSVFilename $Csv -Label 'AD Group' 6>$null) }
  catch { $Result = @(); Assert-True "include group: runs without error ($($_.Exception.Message))" $false }
  Assert-Equal 'include group: user count (non-users skipped)' $Result.Count 2
  Assert-Equal 'include group: UPNs returned in order' ($Result -join ',') 'alice@contoso.com,bob_fabrikam.com#EXT#@contoso.onmicrosoft.com'
  Assert-True 'include group: CSV written' (Test-Path -LiteralPath $Csv)
  if (Test-Path -LiteralPath $Csv) {
    Assert-Equal 'include group: CSV holds the UPNs' ((@(Get-Content -LiteralPath $Csv) | Where-Object { $_ }) -join ',') ($Result -join ',')
  }

  # --- exclude group, no users ---
  $ExCsv = Join-Path $TempDir 'adgroupexcludelist.csv'
  try { $Ex = @(Get-AzureAdGroupMemberUpn -GroupName 'EmptyGroup' -CSVFilename $ExCsv -Label 'excluded AD Group' -AllowEmpty 6>$null) }
  catch { $Ex = @(); Assert-True "empty exclude group: runs without error ($($_.Exception.Message))" $false }
  Assert-Equal 'empty exclude group: returns nothing' $Ex.Count 0
  Assert-True 'empty exclude group: empty CSV still written' (Test-Path -LiteralPath $ExCsv)

  # --- failures ---
  Assert-Throws 'empty include group: throws' {
    Get-AzureAdGroupMemberUpn -GroupName 'EmptyGroup' -CSVFilename (Join-Path $TempDir 'x.csv') 6>$null
  } 'does not contain any User Principal Names'
  Assert-Throws 'unknown group: throws' {
    Get-AzureAdGroupMemberUpn -GroupName 'NoSuchGroup' -CSVFilename (Join-Path $TempDir 'y.csv') 6>$null
  } 'does not exist'
}
finally {
  Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue
}

# --- static binding checks -----------------------------------------------------------------------
$CommonParams = [System.Management.Automation.Cmdlet]::CommonParameters + [System.Management.Automation.Cmdlet]::OptionalCommonParameters
function Get-NamedParameters($CommandAst) {
  $CommandAst.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } |
    ForEach-Object { $_.ParameterName }
}

# 1. Calls to Get-AzureAdGroupMemberUpn use only parameters it declares (prefix match, as
#    PowerShell binds them).
$Declared = @($Fn.Body.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }) + $CommonParams
$FnCalls = @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-AzureAdGroupMemberUpn' }, $true))
Assert-Equal 'Get-AzureAdGroupMemberUpn is called twice (-ADGroup, -ExcludeADGroup)' $FnCalls.Count 2
foreach ($Call in $FnCalls) {
  $Bad = @(Get-NamedParameters $Call | Where-Object { $P = $_; -not ($Declared | Where-Object { $_ -like "$P*" }) })
  Assert-True ("line {0}: call binds to declared parameters{1}" -f $Call.Extent.StartLineNumber, $(if ($Bad) { ' (unknown: -' + ($Bad -join ', -') + ')' })) ($Bad.Count -eq 0)
}

# 2. No built-in cmdlet is called with a parameter name that only binds through an alias.
#    PowerShell 7 added aliases (eg Out-File -FilePath gained 'Path') that 5.1 lacks, so a name
#    that is not a prefix of a canonical parameter name is a 5.1 risk. Aliases known to exist in
#    5.1 too are allowed.
$Known51Aliases = @{ }   # 'Cmdlet' = @('alias', ...) - add entries here only after checking 5.1
$AliasOnly = @()
$Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object {
  $Name = $_.GetCommandName()
  if (-not $Name) { return }
  $Cmd = Get-Command -Name $Name -CommandType Cmdlet -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $Cmd -or $Cmd.ModuleName -notlike 'Microsoft.PowerShell.*') { return }
  $Canonical = @($Cmd.Parameters.Keys)
  foreach ($P in (Get-NamedParameters $_)) {
    if ($Canonical | Where-Object { $_ -like "$P*" }) { continue }
    if ($Known51Aliases[$Cmd.Name] -contains $P) { continue }
    $AliasOnly += ('line {0}: {1} -{2}' -f $_.Extent.StartLineNumber, $Cmd.Name, $P)
  }
}
Assert-True 'no built-in cmdlet call depends on a parameter alias (PowerShell 5.1 safety)' ($AliasOnly.Count -eq 0)
$AliasOnly | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }

# 3. The script must stay pure ASCII: it has no BOM, and Windows PowerShell 5.1 (including
#    Get-Content -Raw, which customers use to work around execution policy) reads a BOM-less
#    file in the ANSI code page, so any UTF-8 character would be mangled or break parsing.
$NonAscii = @(Select-String -LiteralPath $ScriptPath -Pattern '[^\x00-\x7F]' | ForEach-Object { 'line {0}' -f $_.LineNumber })
Assert-True 'script is pure ASCII (safe to read as ANSI on Windows PowerShell 5.1)' ($NonAscii.Count -eq 0)
$NonAscii | Select-Object -First 10 | ForEach-Object { Write-Host "   $_" -ForegroundColor Red }

if ($Failures -gt 0) { Write-Host "`n$Failures check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host "`nAll checks passed." -ForegroundColor Green
exit 0
