<#
.SYNOPSIS
    Offline test harness for the AD group filter subtitle in the HTML report.
.DESCRIPTION
    Needs no M365 tenant and no modules. It parses Get-RubrikM365SizingInfo.ps1 (without running
    it), extracts Get-ADGroupFilterReportNote, and checks:
      - no filter          -> empty output (nothing printed in the report)
      - -ADGroup only      -> "Included AD group" line only
      - -ExcludeADGroup    -> "Excluded AD group" line only
      - both               -> both, included first
      - HTML special chars in group names are encoded
      - whitespace/null names count as "not in use"
      - the script no longer contains the old "(AD Group: $ADGroup)" card titles
      - both the Exchange Online and OneDrive cards call the helper
    Exit code is 0 when all checks pass, 1 otherwise.
.EXAMPLE
    pwsh ./tests/Test-ADGroupFilterReportNote.ps1
#>
[CmdletBinding()]
param (
  [string]$ScriptPath = (Join-Path $PSScriptRoot '..' 'Get-RubrikM365SizingInfo.ps1')
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

$Tokens = $null; $Errors = $null
$Ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $ScriptPath), [ref]$Tokens, [ref]$Errors)
Assert-True 'Script parses without syntax errors' ($Errors.Count -eq 0)
$Errors | ForEach-Object { Write-Host "   $($_.Message) (line $($_.Extent.StartLineNumber))" -ForegroundColor Red }

$Fn = $Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-ADGroupFilterReportNote' }, $true)
Assert-True 'Get-ADGroupFilterReportNote is defined' ($null -ne $Fn)
if (-not $Fn) { exit 1 }
. ([scriptblock]::Create($Fn.Extent.Text))

$Sub = { param($t) "<div class=`"card-header-subtitle`">$t</div>" }

Assert-Equal 'no filter (both omitted) -> empty' (Get-ADGroupFilterReportNote) ''
Assert-Equal 'no filter (both null) -> empty' (Get-ADGroupFilterReportNote -IncludeGroup $null -ExcludeGroup $null) ''
Assert-Equal 'whitespace names -> empty' (Get-ADGroupFilterReportNote -IncludeGroup '  ' -ExcludeGroup '') ''
Assert-Equal 'include only' (Get-ADGroupFilterReportNote -IncludeGroup 'Sales') (& $Sub 'Included AD group: Sales')
Assert-Equal 'exclude only' (Get-ADGroupFilterReportNote -ExcludeGroup 'Contractors') (& $Sub 'Excluded AD group: Contractors')
Assert-Equal 'both' (Get-ADGroupFilterReportNote -IncludeGroup 'Sales' -ExcludeGroup 'Contractors') (& $Sub 'Included AD group: Sales | Excluded AD group: Contractors')
Assert-Equal 'HTML encoding' (Get-ADGroupFilterReportNote -IncludeGroup 'R&D <team>') (& $Sub 'Included AD group: R&amp;D &lt;team&gt;')

$Source = Get-Content -Raw $ScriptPath
Assert-True 'old "(AD Group: $ADGroup)" card titles are gone' (-not $Source.Contains('(AD Group: $ADGroup)'))
$Calls = ([regex]::Matches($Source, [regex]::Escape('$(Get-ADGroupFilterReportNote -IncludeGroup $ADGroup -ExcludeGroup $ExcludeADGroup)'))).Count
Assert-Equal 'helper called on both cards (Exchange Online + OneDrive)' $Calls 2

if ($Failures -gt 0) { Write-Host "`n$Failures check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host "`nAll checks passed." -ForegroundColor Green
exit 0
