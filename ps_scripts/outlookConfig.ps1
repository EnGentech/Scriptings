############################################################################################
# Script Title: outlookConfig.ps1
# Script Function: This script creates flexibility for outlook mail setup
# Script Author: Iberedem Inyang [EnGentech]
# Date Created: 08/01/2024
############################################################################################

Function Create-NewProfile {
    Write-Host "Creating a new Outlook profile..." -ForegroundColor Cyan

    # Define the new profile name
    $profileName = "Sepnu"

    # Create a new Outlook profile (dummy entry to ensure it's recognized)
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$profileName" -Force | Out-Null

    Write-Host "New profile created successfully." -ForegroundColor Green

    # Open Mail Setup in Control Panel to configure profile settings
    Start-Process "control.exe" -ArgumentList "mlcfg32.cpl"
    
    Write-Host "Configuring start-up mail configuration." -ForegroundColor Yellow

    # Pause to allow user to configure settings
    Start-Sleep -Seconds 5
}


Function add_secondary_profile {
    Create-NewProfile
    $SearchPattern = "mlcfg32.cpl"
    $SearchPaths = @("C:\Program Files", "C:\Program Files (x86)", "C:\Windows", "C:\Windows\System32")
    $mlcfgPath = $null

    foreach ($Path in $SearchPaths) {
        $File = Get-ChildItem -Path $Path -Filter $SearchPattern -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($File) {
            $mlcfgPath = $File.FullName
            break
        }
    }

    if ($mlcfgPath) {
        Write-Host "Path found"
        $process = Start-Process $mlcfgPath -PassThru
    
        # Wait for the Mail Setup to open
        Start-Sleep -Seconds 2

        # Activate Mail dialog box
        $mailWindowActivated = $false
        for ($i = 0; $i -lt 10; $i++) {
            if ($process.MainWindowHandle -ne 0) {
                $mailWindowActivated = $true
                break
            }
            Start-Sleep -Seconds 2
        }

        if ($mailWindowActivated) {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.SendKeys]::SendWait("%s")
            [System.Windows.Forms.SendKeys]::SendWait("%p")
            Start-Sleep -Seconds 1
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Write-Host "Keystrokes sent successfully."
            Start-Sleep -Seconds 1
            Write-Host "Starting Outlook"
            Start-Sleep -Seconds 1
            Write-Host "On Launching outlook, select Sepnu from the drop down"
            Start-Sleep -Seconds 5
            Start-Process "Outlook.exe"
            Start-Sleep -Seconds 2
            exit
        } else {
            Write-Host "Mail Setup window did not become active in time." -ForegroundColor Red
        }
    } else {
        Write-Host "Could not find mlcfg32.cpl in the specified directories."
    }

}


Function close_outlook_instance {
$OutlookProcesses = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue

if ($OutlookProcesses) {
    # Close each Outlook process
    foreach ($process in $OutlookProcesses) {
        try {
            $process.CloseMainWindow()
            Start-Sleep -Seconds 5
            # If it's still running, force kill the process
            if (!$process.HasExited) {
                $process.Kill()
            }
        } catch {
            Write-Host "An error occurred while trying to close Outlook: $_"
        }
    }
    Write-Host "All instances of Outlook have been closed."
} else {
    Write-Host "No instances of Outlook were found."
}

}


Function clearData {

    Write-Host "Closing Outlook if running..." -ForegroundColor Cyan
    Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
 
    # Remove Outlook profiles from registry
    Write-Host "Removing Outlook profiles from the registry..." -ForegroundColor Cyan
    Remove-Item -Path "HKCU:\Software\Microsoft\Office\*\Outlook\Profiles" -Recurse -ErrorAction SilentlyContinue
 
    # Remove Outlook cache and local app data
    Write-Host "Removing Outlook cache and local app data..." -ForegroundColor Cyan
    $localAppData = "$env:LocalAppData\Microsoft\Outlook"
    Remove-Item -Path $localAppData -Recurse -Force -ErrorAction SilentlyContinue
 
    # Resetting Office account credentials
    Write-Host "Removing Office cached credentials..." -ForegroundColor Cyan
    $genericCreds = Get-ChildItem -Path "HKCU:\Software\Microsoft\Office\*\Common\Identity\Identities" -ErrorAction SilentlyContinue
    if ($genericCreds) {
        Remove-Item -Path "HKCU:\Software\Microsoft\Office\*\Common\Identity\Identities" -Recurse -ErrorAction SilentlyContinue
    }
 
    Write-Host "Outlook profiles and cached data have been removed." -ForegroundColor Green

    # Start outlook
    Write-Host "Checking compatible outlook version..." -ForegroundColor Cyan
    $officeVersion = (Get-ItemProperty "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration").ProductReleaseIds
    Start-Sleep -Seconds 2
    if ($officeVersion -notlike "O365") {
        Write-Host "Outlook 365 is recommended. Please uninstall your current version and upgrade to Office 365 through SoftwareCenter." -ForegroundColor Red
        Start-Sleep -Seconds 2
        Start-Process "C:\Windows\CCM\ClientUX\scclient.exe"
        Start-Sleep -Seconds 2
        exit
    } else {
        Write-Host "Valid version found, Launching outlook in 3 seconds"
        Start-Sleep -Seconds 3
        Start-Process "outlook.exe"
        #EnGentech
}

}


Function clearProfile {
    # Display Warning dialog box
    Add-Type -AssemblyName "System.Windows.Forms"

    $response = [System.Windows.Forms.MessageBox]::Show(
    "Warning: This action will create a new profile for Outlook and open the Mail Setup window for you to configure profile selection.

    To proceed, select Yes or No to exit",
    "Warning",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    # Grant 3 attempt case of wrong input
    $countDown = 1
    if ($response -eq "Yes") {
        for ($i = 0; $i -le 10; $i++) {
           Write-Host "This script will create a new profile and open the Mail Setup window for you to configure profile selection." -ForegroundColor Yellow
           $confirm = Read-Host "Do you want to proceed? (yes/no)"
           if ($confirm -ne "yes") {
                if ($confirm -eq "no") {
                    Write-Host "Operation canceled." -ForegroundColor Red
                    Start-Sleep -Seconds 3
                    exit
                } else {
                     if ($countDown -eq 3) {
                        Write-Host "Attempts exceeded. Please visit ITSC for further assistance."
                        Start-Sleep -Seconds 3
                        exit
                    } else {
                        Write-Host "Invalid input, try again!" -ForegroundColor Red
                        Write-Host "Attempt = " $countDown 
                        Start-Sleep -Seconds 3     
                        cls
                        if ($countDown -eq 2) {
                            Write-Host "Last Attempt" -ForegroundColor Red
                        }
                        $countDown += 1
                    }
                }
           } else {
                clearData
                exit    
           }
        } 
    } else {
        Write-Host "Operation Cancelled." -ForegroundColor Red
        Start-Sleep -Seconds 3
        exit
    }
}


# Begin Process
Write-Host "Clear exxonmobil mails and reset outlook for a fresh setup or add secondary account" -ForegroundColor Cyan
Write-Host "===================================================================================" -ForegroundColor Cyan
Write-Host " "
$response = Read-Host "Enter yes to continue or no to quit...$ "


$countDown = 1
if ($response -eq "yes") {
    for ($i = 0; $i -le 10; $i++) {
        Write-Host "........................................."
        Write-Host " "
        Write-Host "  1 - Clear Exxon-Mail for a fresh Seplat-Mail setup"
        Write-Host "  2 - Add Secondary Mail to an existing mail"
        Write-Host "  0 - Cancel Operation" -ForegroundColor Red
        Write-Host " "
        $choice = Read-Host "Select an Option to continue...$ "

        if ($choice -ne "1" -and $choice -ne "2" -and $choice -ne "0") {
            if ($countDown -eq 3) {
                Write-Host "Attempts exceeded. Please visit ITSC for further assistance."
                Start-Sleep -Seconds 3
                exit
            } else {
                Write-Host "Invalid input, try again!" -ForegroundColor Red
                Write-Host "Attempt = " $countDown 
                Start-Sleep -Seconds 3     
                cls
                if ($countDown -eq 2) {
                    Write-Host "Last Attempt" -ForegroundColor Red
                }
                $countDown += 1
            }
        } else {
            if ($choice -eq "0") {
                Write-Host "Operation canceled." -ForegroundColor Red
                Start-Sleep -Seconds 3
                exit
            } elseif ($choice -eq "1") {
                # Call clearProfile function
                close_outlook_instance
                clearProfile
            } elseif ($choice -eq "2") {
                # Call Create-NewProfile function
                close_outlook_instance
                add_secondary_profile
            }
        }
    }
} elseif ($response -ne "no") {
    Write-Host "Invalid Selection, Operation Cancelled" -ForegroundColor DarkRed
    Start-Sleep -Seconds 2
} else {
    Write-Host "Operation Cancelled" -ForegroundColor DarkRed
    Start-Sleep -Seconds 2
    exit
}