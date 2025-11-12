<#
.SYNOPSIS
    Apply token lifetime policy for tenant.

.DESCRIPTION
    This script connects to Microsoft Graph and creates or updates a TokenLifetimePolicy
    with the specified AccessTokenLifetime, making it organization default (or not).
#>

# TODO logging to file.
# TODO removing and unassigning a policy
# TODO Dry Run Mode
# TODO Parameters for automation
# TODO [CmdletBinding()] for functions

# Connect to Graph
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
  Write-Verbose "Microsoft.Graph not found. Installing latest version..."
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
#Import-Module Microsoft.Graph -ErrorAction Stop

Connect-MgGraph -NoWelcome -Scopes "Policy.ReadWrite.ApplicationConfiguration", "Policy.Read.All", "Application.ReadWrite.All"
# TODO Check for an existing connection before reconnecting.
Write-Host "Welcome to MS Graph!"

$tokenLifetimePolicyId = $null
$ApplicationId = $null
$paramsLTP = $null


function New-TokenLTP {
  do {
    $NewTokenLifetime = Read-Host "Enter token lifetime in seconds (1-86399)"
  } while ($NewTokenLifetime -notin 1..86399) 

  $lifeTimeFormatted = [timespan]::fromseconds($NewTokenLifetime)
  $NewTokenLifetimeName = Read-Host "Enter a display name (e.g. WebPolicyScenarioXYZ)"

  $params = @{
    Definition            = @(
      @{
        TokenLifetimePolicy = @{
          Version             = 1
          AccessTokenLifetime = $lifeTimeFormatted.ToString("hh\:mm\:ss")
        }
      } | ConvertTo-Json -Compress
    )
    DisplayName           = $NewTokenLifetimeName
    IsOrganizationDefault = $false
  }
  Write-Verbose "`n--- Policy Preview ---`n"
  $params | Format-List

  if ((Read-Host "Create this policy? (Y/N)") -match '^[Yy]$') {
    try {
      $policy = New-MgPolicyTokenLifetimePolicy -BodyParameter $params -ErrorAction Stop
      Write-Host "Created policy '$($policy.DisplayName)' (Id: $($policy.Id))"
    }
    catch {
      Write-Error "Failed to create policy: $_"
    }
  }
  else {
    Write-Host "Aborted."
  }
}

function Set-TokenLTP {
  Get-TokenLTP
  $script:tokenLifetimePolicyId = Read-Host "Please enter the tokenLifetimePolicyId you want to choose"
  $script:paramsLTP = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies/$tokenLifetimePolicyId"
  }
}

function Set-AppID {
  Get-Applications
  $script:ApplicationId = Read-Host "Please enter the Id (NOT AppID!)"
}

function Set-AppTokenLTP {
  Set-TokenLTP
  Set-AppID
  $AppName = (Get-MgApplication -ApplicationId $script:ApplicationId).DisplayName
  $CurrentLTPName = (Get-MgApplicationTokenLifetimePolicy -ApplicationId $script:ApplicationId).DisplayName
  $NewLTPName = (Get-MgPolicyTokenLifetimePolicy -TokenLifetimePolicyId $script:tokenLifetimePolicyId).DisplayName
  Write-Host "You chose" $AppName
  if ($null -ne $CurrentLTPName) { Write-Host "The current TokenLifetimePolicy is"  $CurrentLTPName }
  else { Write-Host "There is no custom lifetime policy for" $AppName }

  if ($CurrentLTPName -eq $NewLTPName){
    Write-Error "The chosen policy is already bound to the application."
    return
  }

  Write-Host "You want to apply the following TokenLifetimePolicy:" $NewLTPName

  if ((Read-Host "Assign this policy? (Y/N)") -match '^[Yy]$') {
    try {
      New-MgApplicationTokenLifetimePolicyByRef -ApplicationId $script:ApplicationId -BodyParameter $script:paramsLTP
      Write-Host "Assigned policy '$NewLTPName' to '$AppName'"
    }
    catch {
      Write-Error "Failed to assign policy: '$NewLTPName' to '$AppName'"
    }
  }
  else {
    Write-Host "Aborted."
  }
}

function Get-TokenLTP { 
  Write-Verbose "`n--- Existing Token Lifetime Policies ---`n"
  try {
    Get-MgPolicyTokenLifetimePolicy |
    Select-Object DisplayName, Definition, Id |
    Format-List # Make an explicit call to Format-List, because an implicit call waits 300ms
  }
  catch {
    Write-Error "Failed to retrieve policies: $_"
  }
}

function Get-Applications { 
  Write-Verbose "`n--- Existing Applications ---`n"
  # Make an explicit call to Format-Table, because an implicit call waits 300ms 
  Get-MgApplication -All | 
  Select-Object DisplayName, Id | 
  Sort-Object DisplayName | 
  Format-Table
}

function Show-Menu { 
  $menuOptions = @{
    1 = "View existing Token Lifetime Policies"
    2 = "View exisiting Applications"
    3 = "Create new Token Lifetime Policy"
    4 = "Assign Policy to an Application"
    5 = "Exit"
  }

  do {
    Write-Host "`n--- Token Lifetime Policy Manager ---"
    foreach ($key in $menuOptions.Keys) {
      Write-Host "$key. $($menuOptions[$key])"
    }

    $choice = Read-Host "Enter your choice"
    switch ($choice) {
      1 { Get-TokenLTP }
      2 { Get-Applications }
      3 { New-TokenLTP }
      4 { Set-AppTokenLTP }
      5 { break }
      default { Write-Host "Invalid choice. Try again." }
    }
  } while ($true)

  Write-Host "Goodbye!"

}

do {Show-Menu} while ($true)


