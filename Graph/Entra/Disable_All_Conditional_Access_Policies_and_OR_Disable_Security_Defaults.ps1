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
Description of Script: This script is used to disable all conditional access policies in the tenant. 
If there are no conditional access policies found it will check if security defaults are enabled and if it is it disables it.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function Disable-AllCAPolicies {
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
             Write-OutPut "Checking if security defaults is enabled and if it is disable it"
             #Check if security defaults is enabled and if it is disable it
             IF ( (Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy).IsEnabled -EQ $TRUE){
                  $params = @{
                  	isEnabled = $false
                  }
                Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -BodyParameter $params
             } 
             return
         }

         # Loop through each policy and disable it
         foreach ($policy in $policies) {
             Write-Output "Disabling policy: $($policy.DisplayName)"
             try {
                 Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter @{state = "disabled"}
                 Write-Output "Policy '$($policy.DisplayName)' has been disabled."
             } catch {
                 Write-Warning "Failed to disable policy '$($policy.DisplayName)'. Error: $_"
             }
         }

         Write-Output "All conditional access policies have been disabled. You can re-enable them manually as needed."

     } catch {
         Write-Error "An error occurred: $_"
     }
 }

 # Example usage of the function
 Disable-AllCAPolicies
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
