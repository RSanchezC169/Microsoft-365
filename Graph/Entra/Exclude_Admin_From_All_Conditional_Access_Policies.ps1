##################################################################################################################################################################
##################################################################################################################################################################
<#
                 _..--+~/@-~--.
             _-=~      (  .   "}
          _-~     _.--=.\ \""""
        _~      _-       \ \_\
       =      _=          '--'
      '      =                             .
     :      :       ____                   '=_. ___
___  |      ;                            ____ '~--.~.
     ;      ;                               _____  } |
  ___=       \ ___ __     __..-...__           ___/__/__
     :        =_     _.-~~          ~~--.__
_____ \         ~-+-~                   ___~=_______
     ~@#~~ == ...______ __ ___ _--~~--_
                                                    =
██████╗ ███████╗ █████╗ ███╗   ██╗ ██████╗██╗  ██╗███████╗███████╗ ██████╗ ██╗ ██████╗ █████╗ 
██╔══██╗██╔════╝██╔══██╗████╗  ██║██╔════╝██║  ██║██╔════╝╚══███╔╝██╔════╝███║██╔════╝██╔══██╗
██████╔╝███████╗███████║██╔██╗ ██║██║     ███████║█████╗    ███╔╝ ██║     ╚██║███████╗╚██████║
██╔══██╗╚════██║██╔══██║██║╚██╗██║██║     ██╔══██║██╔══╝   ███╔╝  ██║      ██║██╔═══██╗╚═══██║
██║  ██║███████║██║  ██║██║ ╚████║╚██████╗██║  ██║███████╗███████╗╚██████╗ ██║╚██████╔╝█████╔╝
╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚══════╝ ╚═════╝ ╚═╝ ╚═════╝ ╚════╝ 
Script Version: 1
OS Version Script was written on: Microsoft Windows 11 Pro : 10.0.25100 Build 26100
PSVersion 5.1.26100.2161 : PSEdition Desktop : Build Version 10.0.26100.2161
Description of Script: 
  This script is used to add your new global admin account to all current conditional access policies exclusion list and retains all current exclusions.
NOTE:
  Before running this function replace -ExclusionInput "ADMIN@DOMAIN.COM" with your smpt address of your global admin account.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function Add-AdminToCAPolicyExclusion {
  param (
        [string]$ExclusionInput # Accepts either an email address or a domain name
    )
	
  #Clear Host
	Clear-Host
     
      TRY{
	          Set-ExecutionPolicy -Scope "Process" -ExecutionPolicy "Unrestricted" -Force  -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
     }CATCH{
            Write-Warning "Could not set execution policy"
    }
    try {
        # Ensure the Microsoft Graph SDK is installed
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }

        # Connect to Microsoft Graph with the required permissions
        Connect-MgGraph -Scopes "Policy.ReadWrite.ConditionalAccess"

        # Retrieve all conditional access policies
        Write-Output "Retrieving all conditional access policies..."
        $policies = Get-MgIdentityConditionalAccessPolicy

        if (-not $policies -or $policies.Count -eq 0) {
            Write-Output "No conditional access policies found in the tenant."
            return
        }

        # Loop through each policy and update exclusions
        foreach ($policy in $policies) {
            Write-Output "Updating policy: $($policy.DisplayName)"
            try {
                # Retrieve the current exclusions
                $currentExclusions = @()
				
                if ($policy.Conditions.Users.ExcludeUsers) {
                    $currentExclusions = $policy.Conditions.Users.ExcludeUsers
					#$currentExclusions
                }
				
                # Input is an email address
                $newExclusion = (Get-MgUser -UserId $ExclusionInput).Id
				#$newExclusion
				
				# Add the new exclusion
                $updatedExclusions = @()
                
				# Combine existing exclusions with the new exclusion
				FOREACH ( $User IN $currentExclusions){
					$updatedExclusions += $User
				}
				$updatedExclusions += $newExclusion
				#$updatedExclusions

                # Update the policy to include the updated exclusions
                $bodyParameters = @{
                    conditions = @{
                        users = @{
                            excludeusers = $updatedExclusions
                        }
                    }
                }
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $bodyParameters

                Write-Output "Policy '$($policy.DisplayName)' has been updated with the exclusion."
            } catch {
                Write-Warning "Failed to update policy '$($policy.DisplayName)'. Error: $_"
            }
        }

        Write-Output "All applicable policies have been updated with the specified exclusion."

    } catch {
        Write-Error "An error occurred: $_"
    }
}

# Example usage of the function
Add-AdminToCAPolicyExclusion -ExclusionInput "ADMIN@DOMAIN.COM"
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
