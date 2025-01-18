############################################################################################
# Script Title: backUpUserData.ps1
# Script Function: This program is designed to back up user data on desktop, favorites,
# pictures and downloads to Network drive [H drive]
# Script Author: Iberedem Inyang [EnGentech]
# Date Created: 16/01/2025
############################################################################################

# Define backup base
$timestamp = Get-Date -Format "dd-MM-yyyy"
$backupBase = "H:\Backup_Directory\Backup_$timestamp"

# Define backup paths
$desktop = "$backupBase\Desktop_Files"
$favorites = "$backupBase\Favorites"
$download = "$backupBase\Downloads"
$pictures = "$backupBase\Pictures"

# Define source paths
$user = $env:USERNAME
$copyDesktop = "C:\Users\$user\Desktop"
$copyDownload = "C:\Users\$user\Downloads"
$copyPictures = "C:\Users\$user\Pictures"
$basePath = "C:\Users\$user\AppData\Local\Microsoft\Edge\User Data"

# Get all files starting with 'bookmark'
$bookmarkFiles = Get-ChildItem -Path $basePath -Recurse -Filter "bookmark*" -ErrorAction SilentlyContinue

Function createBackupFolders {
    mkdir $backupBase
    mkdir $desktop
    mkdir $favorites
    mkdir $download
    mkdir $pictures
}

Function copyContent {
    param (
        [string]$sourcePath,
        [string]$destinationPath
    )

    if ((Get-ChildItem -Path $sourcePath -ErrorAction SilentlyContinue).Count -gt 0) { 
        $itemsCount = (Get-ChildItem -Path $sourcePath -ErrorAction SilentlyContinue).Count 
        Write-Host "The folder contains $itemsCount items."
        Write-Host "Copying from $sourcePath to $destinationPath..." -ForegroundColor DarkYellow
        Write-Host "Please be patient..." -ForegroundColor Yellow

        $files = Get-ChildItem -Path "$sourcePath\*" -Recurse
        $totalFiles = $files.Count
        $currentFile = 0

        foreach ($file in $files) {
            $currentFile++
            $percentComplete = [math]::Round(($currentFile / $totalFiles) * 100)
            Write-Progress -Activity "Copying files" -Status "Copying $($file.Name)" -PercentComplete $percentComplete

            $destFilePath = $file.FullName.Replace($sourcePath, $destinationPath)
            $destDir = Split-Path -Parent $destFilePath

            if (-not (Test-Path -Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir | Out-Null
            }

            Copy-Item -Path $file.FullName -Destination $destFilePath -Force
        }

        Write-Host "Copy operation completed." -ForegroundColor Green
        Start-Sleep -Seconds 2
    } else { 
        Write-Host "The folder is empty." -ForegroundColor Magenta
        Start-Sleep -Seconds 2
    }
}

Function copyBookmarkFiles {
    param (
        [System.IO.FileInfo[]]$files,
        [string]$destinationPath
    )

    foreach ($file in $files) {
        if (Test-Path -Path $file.FullName) {
            Write-Host "Copying $($file.FullName) to $destinationPath..."
            Copy-Item -Path $file.FullName -Destination $destinationPath -Force
            Write-Host "Bookmark file $($file.FullName) copied successfully."
        } else {
            Write-Host "Bookmark file $($file.FullName) not found."
        }
    }
}

Function callUp {
    Write-Host "Backup in progress..." -ForegroundColor Yellow
    copyBookmarkFiles -files $bookmarkFiles -destinationPath $favorites
    copyContent -sourcePath $copyDesktop -destinationPath $desktop
    copyContent -sourcePath $copyDownload -destinationPath $download
    copyContent -sourcePath $copyPictures -destinationPath $pictures
    Write-Host " "
    Write-Host ".................................................................................."
    Write-Host "Backup is now Completed" -ForegroundColor Green
    Write-Host "Exiting in 5 seconds..."
    Start-Sleep -Seconds 5
}

Function desktopBackup {
    Write-Host "Looking for H drive..." -ForegroundColor cyan
    Start-Sleep -Seconds 2

    # Check for H drive existence
    if (Test-Path -Path "H:\") {
        Write-Host "H drive found" -ForegroundColor Yellow
        Start-Sleep -Seconds 2

        if (-not (Test-Path -Path "H:\Backup_Directory")) {
            mkdir "H:\Backup_Directory"
        }

        if (Test-Path -Path $backupBase) {
            Write-Host "$backupBase already present in H drive"
            Write-Host "Renaming $backupBase to Backup_Old"
            Start-Sleep -Seconds 3

            # Check for backup_Old in H drive
            if (Test-Path -Path "H:\Backup_Directory\Backup_Old") {
                Write-Host "Backup_Old already present, please rename this folder and retry or visit SEIT" -ForegroundColor Magenta
                Start-Sleep -Seconds 10 
                exit
            }         
            Rename-Item -Path "$backupBase" -NewName "H:\Backup_Directory\Backup_Old"
            Start-Sleep -Seconds 2
        }

        Write-Host "Creating necessary folders on H drive" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        createBackupFolders            
        Write-Host "Folders created on H drive" -ForegroundColor Green
        Start-Sleep -Seconds 3
        callUp
    } else {
        Write-Host "Please map H drive and re-run the application" -ForegroundColor Magenta
        Start-Sleep -Seconds 3
        exit
    }
}

function intro {
    Write-Host "This script is designed to back up your data on Desktop, Downloads, Pictures, and Favorites to" -ForegroundColor cyan
    Write-Host "      your network drive [H drive]. If there's a previous backup, it will be renamed" -ForegroundColor cyan
    Write-Host "           to Backup_Old. If you encounter any issue, please visit SEPNU ITSC" -ForegroundColor cyan
    for ($i = 0; $i -le 94; $i++) {
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
            desktopBackup

        } else {
            Write-Output "Connect to VPN"
            Start-Sleep -Seconds 3
        }
    } catch {
        Write-Output "No Connection found"
        Start-Sleep -Seconds 3
    }
} else {
    Write-Output "No Internet connection"
    Start-Sleep -Seconds 3
}