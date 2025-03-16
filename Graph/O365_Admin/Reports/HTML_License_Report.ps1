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
This PowerShell script automates the retrieval, organization, and reporting of Microsoft 365 license assignments.
 It utilizes the Microsoft Graph API to fetch detailed license data, including total available licenses (PrepaidUnits), assigned licenses (ConsumedUnits), and the number of licenses still available for assignment. 
The script also identifies users assigned to each license, grouping them by license type for clarity. 
A comprehensive and interactive HTML report is then generated, featuring a dropdown menu that allows users to select specific licenses and dynamically view associated user details. 
The report includes a summary of license utilization for each type, ensuring administrators can easily monitor and manage license allocations.
The script is highly modular, with separate functions for retrieving licenses, fetching user assignments, and generating the HTML report.
 It incorporates robust error handling and detailed logging to track progress and capture any issues during execution. 
To enhance user experience, the script employs real-time progress bars and automatically opens the HTML report upon completion in the system's default browser. 
Designed to streamline the license management process, this script reduces manual effort, improves visibility into license usage, and ensures IT administrators can make informed decisions about resource allocation within their Microsoft 365 environment.

#>
##################################################################################################################################################################
#==============================Beginning of script================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================Functions==========================================================================================================================
##################################################################################################################################################################
# Define the global log file path
$Global:LogFile = "C:\Rsanchezc169ScriptLogs\Log_$(Get-Date -Format 'MM_dd_yyyy_hh_mm_tt').log"

# Global logging function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter()]
        [string]$LogFile = $Global:LogFile
    )

    $LogDirectory = Split-Path -Path $LogFile
    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    $Timestamp = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss tt")
    Add-Content -Path $LogFile -Value "$Timestamp : $Message"
}
##################################################################################################################################################################
##################################################################################################################################################################
function Get-Licenses {
    try {
        Write-Progress -Activity "Retrieving Licenses" -Status "Fetching all available licenses" -PercentComplete 30
        Write-Log "Retrieving all available licenses in the tenant..."

        # Retrieve all available licenses
        $licenses = Get-MgSubscribedSku
        if (-not $licenses) {
            Write-Log "No licenses found in the tenant."
            throw "No licenses found in the tenant."
        }

        # Include license counts (available and assigned)
        $licensesWithDetails = $licenses | ForEach-Object {
            [PSCustomObject]@{
                SkuPartNumber  = $_.SkuPartNumber
                SkuId          = $_.SkuId
                PrepaidUnits   = $_.PrepaidUnits.Enabled # Total available licenses
                ConsumedUnits  = $_.ConsumedUnits        # Assigned licenses
            }
        }

        Write-Log "Retrieved $($licensesWithDetails.Count) licenses with usage details."
        return $licensesWithDetails
    } catch {
        Write-Log "Failed to retrieve licenses. Error: $_"
        throw
    }
}
##################################################################################################################################################################
##################################################################################################################################################################
function Get-UsersByLicense {
    param (
        [array]$Licenses
    )

    $allLicenseUsers = @()
    Write-Log "Retrieving all users and their assigned licenses..."

    try {
        # Fetch all users in the tenant
        $allUsers = Get-MgUser -All
        if (-not $allUsers) {
            Write-Log "No users found in the tenant."
            throw "No users found."
        }

        # Iterate through each user to retrieve license details
        foreach ($user in $allUsers) {
            Write-Log "Processing licenses for user: $($user.DisplayName)"
            try {
                # Fetch license details for the user
                $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
                foreach ($license in $licenseDetails) {
                    $skuPartNumber = ($Licenses | Where-Object { $_.SkuId -eq $license.SkuId }).SkuPartNumber
                    if ($skuPartNumber) {
                        $allLicenseUsers += [PSCustomObject]@{
                            UserName      = $user.DisplayName
                            EmailAddress  = $user.UserPrincipalName
                            License       = $skuPartNumber
                        }
                        Write-Log "User $($user.DisplayName) assigned license: $skuPartNumber"
                    }
                }
            } catch {
                Write-Log "Failed to retrieve license details for user: $($user.DisplayName). Error: $_"
            }
        }
    } catch {
        Write-Log "Failed to retrieve users. Error: $_"
        throw
    }

    if ($allLicenseUsers.Count -eq 0) {
        Write-Log "No license-user assignments found."
    }

    return $allLicenseUsers
}
##################################################################################################################################################################
##################################################################################################################################################################
function Generate-HTMLReport {
    param (
        [array]$AllLicenseUsers,
        [array]$Licenses,
        [string]$OutputHtmlPath
    )

    if ($AllLicenseUsers.Count -eq 0) {
        Write-Log "No license-user data found to generate an HTML report."
        throw "No license-user data available."
    }

    # Group users by license
    $groupedLicenses = $AllLicenseUsers | Group-Object -Property License
    Write-Log "Building HTML report with $($AllLicenseUsers.Count) user-license assignments."

    # Begin building the HTML content
    $htmlContent = @"
<html>
<head>
    <title>Microsoft 365 License Assignments</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f4f4f4; }
        .license { display: none; }
        h1 { text-align: center; }
    </style>
    <script>
        function showLicense(licenseId) {
            const licenses = document.querySelectorAll('.license');
            licenses.forEach(license => {
                if (licenseId === 'all') {
                    license.style.display = 'block';
                } else if (license.id === licenseId) {
                    license.style.display = 'block';
                } else {
                    license.style.display = 'none';
                }
            });
        }
    </script>
</head>
<body>
    <h1>Microsoft 365 License Assignments</h1>
    <label for="licenseSelect">Select a License:</label>
    <select id="licenseSelect" onchange="showLicense(this.value)">
        <option value="all">All Licenses</option>
"@

    # Add dropdown options for licenses and their usage details
    foreach ($license in $Licenses) {
        $licenseId = $license.SkuPartNumber -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $availableLicenses = $license.PrepaidUnits - $license.ConsumedUnits
        $htmlContent += "<option value='$licenseId'>$($license.SkuPartNumber) (Available: $availableLicenses)</option>"
    }

    $htmlContent += "</select>"

    # Add sections for each license and its users
    foreach ($group in $groupedLicenses) {
        $licenseId = $group.Name -replace '\s', '_'  # Replace spaces with underscores for HTML IDs
        $licenseDetails = $Licenses | Where-Object { $_.SkuPartNumber -eq $group.Name }
        $availableLicenses = $licenseDetails.PrepaidUnits - $licenseDetails.ConsumedUnits

        $htmlContent += "<div class='license' id='$licenseId'>"
        $htmlContent += @"
        <h2>$($group.Name)</h2>
        <p><strong>Total Licenses:</strong> $($licenseDetails.PrepaidUnits) | <strong>Assigned:</strong> $($licenseDetails.ConsumedUnits) | <strong>Available:</strong> $availableLicenses</p>
        <table>
            <thead>
                <tr>
                    <th>Name</th>
                    <th>Email Address</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($user in $group.Group) {
            $htmlContent += @"
                <tr>
                    <td>$($user.UserName)</td>
                    <td>$($user.EmailAddress)</td>
                </tr>
"@
        }

        $htmlContent += @"
            </tbody>
        </table>
        </div>
"@
    }

    # Finalize HTML content
    $htmlContent += @"
</body>
</html>
"@

    # Ensure the output directory exists
    $OutputDirectory = Split-Path -Path $OutputHtmlPath -Parent
    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created missing output directory: $OutputDirectory"
    }

    # Write HTML content to file
    try {
        $htmlContent | Set-Content -Path $OutputHtmlPath -Force
        Write-Log "HTML report successfully written to: $OutputHtmlPath"
    } catch {
        Write-Log "Failed to write HTML report. Error: $_"
        throw
    }

    # Validate if the file was created
    if (-not (Test-Path -Path $OutputHtmlPath)) {
        Write-Log "Error: HTML file does not exist at the specified path: $OutputHtmlPath"
        throw "HTML file was not created."
    } else {
        Write-Log "HTML file exists at: $OutputHtmlPath"
    }

    # Automatically open the file
    Write-Log "Opening the HTML file: $OutputHtmlPath"
    Start-Process -FilePath $OutputHtmlPath -Verb Open
}
##################################################################################################################################################################
#=============================End of Functions====================================================================================================================
##################################################################################################################################################################
#==============================Main===============================================================================================================================
##################################################################################################################################################################
try {
    Write-Progress -Activity "Execution" -Status "Starting script" -PercentComplete 0
    Write-Log "Script execution started."

    # Step 1: Ensure the Microsoft Graph SDK is installed
    Ensure-MicrosoftGraphSDK

    # Step 2: Connect to Microsoft Graph
    Connect-ToMicrosoftGraph

    # Step 3: Retrieve all licenses with their details (e.g., PrepaidUnits and ConsumedUnits)
    Write-Progress -Activity "Execution" -Status "Retrieving licenses" -PercentComplete 20
    $licenses = Get-Licenses

    # Step 4: Retrieve users assigned to licenses
    Write-Progress -Activity "Execution" -Status "Retrieving users for licenses" -PercentComplete 50
    $licenseUsers = Get-UsersByLicense -Licenses $licenses

    # Step 5: Generate the HTML report with license availability data
    $outputPath = "C:\Temp\LicenseAssignments.html"
    Write-Progress -Activity "Execution" -Status "Generating HTML report" -PercentComplete 80
    Generate-HTMLReport -AllLicenseUsers $licenseUsers -Licenses $licenses -OutputHtmlPath $outputPath

    # Mark script completion
    Write-Progress -Activity "Execution" -Status "Script completed successfully" -PercentComplete 100
    Write-Log "Script execution completed successfully. Report generated at: $outputPath"
} catch {
    Write-Log "An error occurred during execution: $_"
    Write-Error "An error occurred: $_"
}
##################################################################################################################################################################
#==============================End of Main========================================================================================================================
##################################################################################################################################################################
##################################################################################################################################################################
#==============================End of Script======================================================================================================================
##################################################################################################################################################################
