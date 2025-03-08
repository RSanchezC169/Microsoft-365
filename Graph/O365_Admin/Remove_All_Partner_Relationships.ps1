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
This script will remove all partner relationships in the tenant.
#>
##################################################################################################################################################################
#==============================Beginning==========================================================================================================================
##################################################################################################################################################################
function Remove-AllPartnerRelationships {
    try {
        # Ensure the Microsoft Graph SDK is installed
        if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
        }

        # Connect to Microsoft Graph with the required permissions
        Connect-MgGraph -Scopes "Directory.ReadWrite.All"

        # Retrieve all service principals associated with partner relationships
        Write-Output "Retrieving all partner relationships..."
        $partnerRelationships = Get-MgServicePrincipal | Where-Object { $_.AppDisplayName -like "*Partner*" }

        if (-not $partnerRelationships -or $partnerRelationships.Count -eq 0) {
            Write-Output "No partner relationships were found in your tenant."
            return
        }

        # Display the partner relationships found
        Write-Output "Found the following partner relationships:"
        $partnerRelationships | ForEach-Object { Write-Output "AppDisplayName: $($_.AppDisplayName), ObjectId: $($_.Id)" }

        # Loop through each partner relationship and remove it
        foreach ($partner in $partnerRelationships) {
            Write-Output "Removing partner relationship: $($partner.AppDisplayName) (ObjectId: $($partner.Id))"
            try {
                Remove-MgServicePrincipal -ServicePrincipalId $partner.Id
                Write-Output "Successfully removed: $($partner.AppDisplayName)"
            } catch {
                Write-Warning "Failed to remove partner relationship: $($partner.AppDisplayName). Error: $_"
            }
        }

        Write-Output "All unused partner relationships have been removed. You can now add the new one as needed."

    } catch {
        Write-Error "An error occurred: $_"
    }
}

# Example usage of the function
Remove-AllPartnerRelationships
##################################################################################################################################################################
#==============================End================================================================================================================================
##################################################################################################################################################################
