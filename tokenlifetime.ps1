if (Get-Module -ListAvailable -Name Microsoft.Graph) {
  Write-Host "Microsoft.Graph Module exists. Skipping installation";
} 
else {
  Write-Host "Microsoft.Graph Module does not exist. Installing..";
  Install-Module Microsoft.Graph
}

Connect-MgGraph -NoWelcome -Scopes "Policy.ReadWrite.ApplicationConfiguration", "Policy.Read.All", "Application.ReadWrite.All"
Write-Host "Welcome to MS Graph!"

$tokenLifetimePolicyId = $null
$ApplicationId = $null

function New-TokenLTP {
  $NewTokenLifetime = 0
  while ($NewTokenLifetime -notin 1..86399) {
    Write-Host "The value has to be between 1s and 86399 (1 day)"
    $NewTokenLifetime = Read-Host "Please enter the token lifetime in seconds"
  }
  $lifeTimeFormatted = [timespan]::fromseconds($NewTokenLifetime)
  Write-Host "---------------------------------"
  Write-Host $lifeTimeFormatted
  Write-Host "---------------------------------"
  Write-Host "Name in the format of WebPolicyScenario XYZ"
  $NewTokenLifetimeName = Read-Host "Please enter the name for the new Token lifetime"
  $NewTokenLifetimeName.Length

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
  Write-Host "---------------------------------"
  Write-Host "You are creating a new TokenLifetimePolicy with the following parameteres:"
  Write-Host "---------------------------------"
  Write-Host "Definition" + $params.Definition
  Write-Host "DisplayName" + $params.DisplayName
  Write-Host "IsOrganizationDefault" + $params.IsOrganizationDefault
  $Sure = Read-Host "Are you sure? (y/n)"
  if ($Sure -eq "y") { 
    $script:tokenLifetimePolicyId = (New-MgPolicyTokenLifetimePolicy -BodyParameter $params).Id 
  }
  else { Write-Host "Aborted" }
}

function Set-TokenLTP {
  $script:tokenLifetimePolicyId = Read-Host "Please enter the tokenLifetimePolicyId you want to choose"
  $script:paramsLTP = @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/policies/tokenLifetimePolicies/$tokenLifetimePolicyId"
  }
}

function Set-AppID {
  Get-MgApplication -All | 
  Select-Object DisplayName, Id | 
  Sort-Object DisplayName | 
  Format-Table
  $script:ApplicationId = Read-Host "Please enter the Id (NOT AppID!)"
}

function Set-AppTokenLTP {
  Get-TokenLTP
  Set-TokenLTP
  Set-AppID
  Clear-Host
  $CurrentApp = Get-MgApplication -ApplicationId $ApplicationId
  Write-Host "You chose" $CurrentApp.DisplayName
  $CurrentLTP = Get-MgApplicationTokenLifetimePolicy -ApplicationId $ApplicationId
  if ($null -ne $CurrentLTP){Write-Host "The current TokenLifetimePolicy is" $CurrentLTP.DisplayName}
  else {Write-Host "There is no custom lifetime policy for" $CurrentApp.DisplayName}
  Write-Host "You want to apply the following TokenLifetimePolicy:" $(Get-MgPolicyTokenLifetimePolicy -TokenLifetimePolicyId $tokenLifetimePolicyId).displayname
  $Sure = Read-Host "Are you sure? (y/n)"
  if ($Sure -eq "y") { Write-Host "Do Something!" }
  else { Write-Host "Aborted" }

}

function Get-TokenLTP { 

  Write-Host ""
  Write-Host "---------------------------------"
  Write-Host "Exisiting Token Lifetime Policies"
  Write-Host "---------------------------------"
  # Make an explicit call to Format-List, because an implicit call waits 300ms 
  Get-MgPolicyTokenLifetimePolicy | 
  Select-Object displayname, definition, id |
  Format-List
}


#Get-TokenLTP
#New-TokenLTP
Set-AppTokenLTP
