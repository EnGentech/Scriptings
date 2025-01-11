############################################################################################
# Script Title: resetProfile.ps1
# Script Function: Clear user corrupt profile
# Script Author: Iberedem Inyang [EnGentech]
# Date Created: 10/01/2024
############################################################################################

# Function to check if the script is run as an administrator
function Test-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Error "Warning! This script must be run as an administrator."
        exit
   }
}

# Check if the script is run as an administrator
Test-Admin

# Define the registry path
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"

# Function to find and delete the SID with the specified ProfileImagePath
function Delete-SID {
    param (
        [string]$LanID
    )

    # Define the registry path
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $SIDs = Get-ChildItem -Path $RegPath

    foreach ($SID in $SIDs) {
        $ProfileImagePath = (Get-ItemProperty -Path "$RegPath\$($SID.PSChildName)").ProfileImagePath
        if ($ProfileImagePath -eq "C:\Users\$LanID") {
            Remove-Item -Path "$RegPath\$($SID.PSChildName)" -Recurse
            Write-Host "Deleted SID: $($SID.PSChildName) with ProfileImagePath: $ProfileImagePath"
            return $true
        }
    }
    return $false
}


# Define the user's LAN ID
Write-Host "    Clear corrupt profile on SEPNU machine"
Write-Host "    ======================================"
Write-Host " "
$LanID = Read-Host "Enter LAN_ID...$ "

# Confirm status
Write-Host "Kindly understand that $LanID profile will be deleted"
$confirm = Read-Host "To proceed, enter yes or no to cancel...$ "

# Delete the SID with the specified ProfileImagePath
if ($confirm -eq "yes") {
    if (Delete-SID -LanID $LanID) {
        # Rename the user's folder
        Rename-Item -Path "C:\Users\$LanID" -NewName "${LanID}_OLD"
        Write-Host "Renamed folder: C:\Users\$LanID to C:\Users\${LanID}_OLD"

        # Restart the computer
        Write-Host "Restarting computer..."
        Start-Sleep -Seconds -3
        Restart-Computer -Force
    } else {
        Write-Host "No SID found with ProfileImagePath pointing to C:\Users\$LanID"
    }

} else {
    Write-Host "Operation cancelled"
    Start-Sleep -Seconds 3
    exit
}