# Get-CPResult.ps1
# Create a report of Policy settings retrieved from Cloud Policy, the Cloud Policy version of gpresult
# https://learn.microsoft.com/en-us/deployoffice/admincenter/overview-cloud-policy
# V0.1 20-Jun-2024 Initial version
# Author: Andres Gorzelany
# More information: https://github.com/get-itips/cpresult
# Blog post: TBD

#Variable definition
$Path="HKCU:\Software\Policies\Microsoft\Cloud\Office\16.0\"
$Policies=import-csv "CloudPoliciesFull.csv"



    # Check if the registry path exists
	Write-Host "Checking if Cloud Policy registry keys exists..." -ForegroundColor Yellow
    if (Test-Path $Path) {
		Write-Host "Cloud Policy registry keys found, now processing..." -ForegroundColor Yellow
        # Get all subkeys
        $subKeys = Get-ChildItem -Path $Path -Recurse
        $PoliciesResult = [System.Collections.Generic.List[Object]]::new()
		
        foreach ($key in $subKeys) {
			
            #Write-Host "VERBOSE - Browsing Key: $($key.PSPath)"
            # Get all values in the current key
            $values = Get-ItemProperty -Path $key.PSPath | select-object -ExcludeProperty "PS*" -Property *
            foreach ($value in $values.PSObject.Properties) {
                #Write-Host "`tValue Name: $($value.Name) - Value Data: $($value.Value)"
				#Let's look for policy name
				$policy=$policies | where-object registrySubPath -eq $value.Name
				if($policy){
					#Policy found
                    #More than one policy with that name can be found TODO: check also registryKeyPath to a match
                    if($policy.count -eq 1)
                    {
                        Write-Host $policy.Name
                        #Write-Host $value.Value
                        $PolicyItem = [PSCustomObject]@{
                            Application = $policy.Id.Split(";")[0]
                            Policy = $policy.Name
                            "Configuration Setting" = $value.Value
                        }
                    }
                    else {
                        Write-Host $policy[0].Name
                        #Write-Host $value.Value
                        $PolicyItem = [PSCustomObject]@{
                            Application = $policy[0].Id.Split(";")[0]
                            Policy = $policy[0].Name
                            "Configuration Setting" = $value.Value
                        }
                    }
					$PoliciesResult.add($PolicyItem)
					
				}
                else{
                    #Policy not found in our policies file, let's show it like it is
                    Write-Host $value.Name
					#Write-Host $value.Value
					$PolicyItem = [PSCustomObject]@{
						Policy = $value.Name
						"Configuration Setting" = $value.Value
					}
					$PoliciesResult.add($PolicyItem)
                }
				
            }
        }
		Write-Host "Done..." -ForegroundColor Yellow
		$PoliciesResult | Out-GridView
    } else {
        Write-Host "The Cloud Policies Registry path does not exist."  -ForegroundColor Red
    }
