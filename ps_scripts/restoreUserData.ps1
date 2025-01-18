############################################################################################
# Script Title: restoreUserData.ps1
# Script Function: This program is designed user files that was backup up to H-Drive,
# Script Author: Iberedem Inyang [EnGentech]
# Date Created: 16/01/2025
############################################################################################


# Define backup base
Function checkAndRestoreBackup {
    # Define the backup directory
    $backupBase = "H:\Backup_Directory"
    
    Write-Host "Looking for H drive..." -ForegroundColor cyan
    Start-Sleep -Seconds 1

    # Check for H drive existence
    if (Test-Path -Path "H:\") {
        Write-Host "H drive found" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Write-Host "Look up for Backup_Directory"
        Start-Sleep -Seconds 1

        # Check if the backup directory exists and is not empty
        if (Test-Path $backupBase) {
            # List all directories in the backup directory
            Write-Host "$backupBase found"
            Start-Sleep -Seconds 1
            $backups = Get-ChildItem -Path $backupBase -Directory

            if ($backups.Count -gt 1) {
                Write-Output "Backup_Directory contains several files, please select your restore point:"
                Write-Host " "
                $i = 1
                foreach ($backup in $backups) {
                    Write-Output "$i. $($backup.Name)"
                    $i++
                }

                # Get user input
                Write-Host " "
                $choice = Read-Host "Restore point option...$ "

                # Check if the choice is valid
                if ($choice -match '^\d+$' -and [int]$choice -gt 0 -and [int]$choice -le $backups.Count) {
                    # Create restore directory and save to a variable
                    $restoreRoot = "${backupBase}\$($backups[[int]$choice - 1].Name)"
                    restoreUserData -address $restoreRoot
                } else {
                    Write-Output "Invalid choice. Please run the script again and select a valid number."
                    Start-Sleep -Seconds 5
                    exit
                }
            } elseif ($backups.Count -eq 1){
                $restorRoot = "$backupBase\$backups"
                restoreUserData -address $restorRoot
            } else {
                Write-Output "No backups found in the directory."
                Start-Sleep -Seconds 5
                exit
            }
        } else {
            Write-Output "Backup_Directory is either empty or does not exist."
            Start-Sleep -Seconds 5
            exit
        }
    } else {
        Write-Host "H drive not found" -ForegroundColor Red
        Start-Sleep -Seconds 5
        exit
    }
}


# Define source paths
$user = $env:USERNAME
$restoreDesktop = "C:\Users\$user\Desktop"
$restoreDownload = "C:\Users\$user\Downloads"
$restorePictures = "C:\Users\$user\Pictures"
$basePath = "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data"


Function restoreUserData {
    param (
        [string]$address
    )

    # Define backup paths
    $desktop = "$address\Desktop_Files"
    $favorites = "$address\Favorites"
    $download = "$address\Downloads"
    $pictures = "$address\Pictures"

    Function Copy-WithConfirmation {
        param (
            [string]$SourcePath,
            [string]$DestinationPath
        )

        # Get all items in the source path
        $items = Get-ChildItem -Path $SourcePath -Recurse

        foreach ($item in $items) {
            # Construct the destination path for each item
            Write-Host "Restoring $item" -ForegroundColor Yellow
            $dest = Join-Path -Path $DestinationPath -ChildPath $item.FullName.Substring($SourcePath.Length)
            
            # Check if the destination file exists
            if (Test-Path -Path $dest) {
                $overwrite = ""
                
                # Loop until a valid input (y/n) is received
                while ($overwrite -ne 'y' -and $overwrite -ne 'n') {
                    $overwrite = Read-Host "The file $($item.Name) already exists. Overwrite? (y/n)"
                    
                    if ($overwrite -eq 'y') {
                        Copy-Item -Path $item.FullName -Destination $dest -Force
                        Write-Host "File overwritten." -ForegroundColor Cyan
                    } elseif ($overwrite -eq 'n') {
                        Write-Host "File skipped." -ForegroundColor DarkRed
                    } else {
                        Write-Host " "
                        Write-Host "Invalid input. Please enter 'y' or 'n'."
                    }
                }
            } else {
                Copy-Item -Path $item.FullName -Destination $dest
            }
        }

        # Restore files
        Write-Host "Restoration Completed" -ForegroundColor Green
    }

    # Copy files from each backup path to the respective restore locations with confirmation prompt
    Copy-WithConfirmation -SourcePath $desktop -DestinationPath $restoreDesktop
    Copy-WithConfirmation -SourcePath $download -DestinationPath $restoreDownload
    Copy-WithConfirmation -SourcePath $pictures -DestinationPath $restorePictures

    # Test path
    if (Test-Path -Path "$basePath\Default") {
        Copy-WithConfirmation -SourcePath $favorites -DestinationPath "$basePath\Default"
    } else {
        mkdir "$basePath\Default"
        Copy-WithConfirmation -SourcePath $favorites -DestinationPath "$basePath\Default"
    }
    Write-host " "
    Write-Host "All files have been restored successfully" -ForegroundColor Green
    Write-Host "Exiting..." 
    start-sleep -Seconds 5
}


function intro {
    Write-Host "Your back up files will be restored to its respective folders considering its latest backUp" -ForegroundColor cyan
    Write-Host "         This script will be terminated if the expected folder is not found" -ForegroundColor cyan
    for ($i = 0; $i -le 90; $i++) {
        Write-Host -NoNewLine "="
        Start-Sleep -MilliSeconds 15
    }
    Start-Sleep -seconds 2
    Write-Host " "
}

# Begin Process
intro

$networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
if ($networkAdapters) {
    Write-Output "Checking Connection..."
    Start-Sleep -Seconds 2
    try {
        $dnsResult = Resolve-DnsName -Name "www.google.com"
        if ($dnsResult) {
            Write-Output "Connection validated"
            Start-Sleep -Seconds 3
            # Call backup function
            checkAndRestoreBackup
        } else {
            Write-Output "Connect to VPN"
            Start-Sleep -Seconds 3
        }
    } catch {
        Write-Output "An error Occured: $_"
        Start-Sleep -Seconds 3
    }
} else {
    Write-Output "No Internet connection"
    Start-Sleep -Seconds 3
}