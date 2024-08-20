# --- DriverCabCreation_1.0.9 --- 
# Creation date: 7/31/24 by Intern Jason Chin :)
# Recent addition: 
# 1) saves the logs into separate file in redacted folder to help in troubleshooting process

# Get current user's name for future pathing
$Current_user = (Get-WmiObject -Class win32_computersystem).UserName.split('\')[1]

# Start stopwatch
$timeTaken = [System.Diagnostics.Stopwatch]::StartNew()

# Add necessary .NET assemblies to use Windows forms for GUI input
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# Display input box and store the user input (make and model of the device)
$title = 'Driver Cabinet Creation'
$msg = 'What is the make and model of the device being certified?:'
$UserInput = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

# Define the base download path based on user input
$baseDownloadPath = "C:\redacted\$Current_user\redacted\$UserInput"

# Define the log file path
$logFilePath = "C:\redacted\$Current_user\redacted\ScriptLogs_$UserInput.log"

# Initialize log file
New-Item -Path $logFilePath -ItemType "file" -Force

# Logging function
function Write-Log {
    param (
        [string]$message
    )
    # Write message to console
    Write-Output $message
    # Append message to log file
    Add-Content -Path $logFilePath -Value $message
}

# Check if the specified folder exists in the Test directory
if (-Not (Test-Path -Path $baseDownloadPath)) {
    Write-Log "The specified folder '$UserInput' does not exist in the 'Test' directory. Please check the name and try again."
    exit
}

try {
    # Hashtable of destination categories and their corresponding (vendor side) folders that will be copied over
    # THIS is where to add NEW vendor driver folders
    $foldersToCopy = @{
        'redacted' = @('Audio', 'Communication', 'Communications')
        'redacted' = @('Chipset', 'Power Management', 'Power_Management')
        'redacted' = @('Application', 'Camera', 'Card Reader', 'Cardreader', 'Card_Reader', 'Display', 'Video', 'Graphics')
        'redacted' = @('Keyboard', 'Input', 'Other', 'Port', 'Sensor', 'Security', 'USB')
        'redacted' = @('HSAs')
        'redacted' = @('Bluetooth', 'Etherenet', 'Ethernet', 'Network', 'Thunderbolt', 'Wireless')
        'redacted' = @('Storage', 'Storeage', 'RAID')
        'redacted' = @('Docks', 'Docks_Stands', 'Dock Stands')
    }

    # Function to find 'Audio' folder
    function Find-AudioFolder {
        param ([string]$basePath)
        $audioFolder = Get-ChildItem -Path $basePath -Recurse -Directory | Where-Object { $_.Name -eq 'Audio' } | Select-Object -First 1
        return $audioFolder
    }

    # Find 'Audio' folder
    $audioFolder = Find-AudioFolder -basePath $baseDownloadPath

    # Check if 'Audio' folder exists
    if (-Not $audioFolder) {
        throw "The 'Audio' folder does not exist in the 'redacted' directory. The driver pack structure might be incorrect."
    }

    # Get parent folder of 'Audio' folder
    $parentFolder = $audioFolder.Parent.FullName

    # Get all sibling folders
    $siblingFolders = Get-ChildItem -Path $parentFolder -Directory | Select-Object -ExpandProperty Name

    # Flatten hashtable values to create list of all expected folders
    $expectedFolders = $foldersToCopy.Values | ForEach-Object { $_ } | Select-Object -Unique

    # Find folders not listed in expected folders
    $newFolders = $siblingFolders | Where-Object { $_ -notin $expectedFolders }

    # If there are new folders, stop the script and throw an error
    if ($newFolders.Count -gt 0) {
        $newFoldersList = $newFolders -join ", "
        throw "This driver pack contains the following new driver folders which are not yet accounted for: $newFoldersList"
    }

    # Construct path for the driver cabinet based on user input
    $DriverCab = "C:\redacted\$Current_user\redacted\$UserInput"

    # Check if driver cabinet folder already exists
    If (-Not (Test-Path -Path $DriverCab)) {
        # Create main folder and necessary subfolders
        New-Item -Path $DriverCab -ItemType "directory"
        $foldersToCopy.Keys + 'redacted' | ForEach-Object { New-Item -Path "$DriverCab\$_" -ItemType "directory" }
        New-Item -Path "$DriverCab\$UserInput.txt"
    } else {
        Write-Log "Path exists!"  # Print message if path exists
    }

    # Initialize an array to keep track of copied files
    $copiedFiles = @()

    # Function to find and copy files
    function Find-And-CopyFiles {
        param (
            [string]$basePath,         # Base path where to start searching
            [string]$targetFolderName, # Name of the folder to look for
            [string]$destinationFolder # Destination folder where files will be copied
        )

        # Create search pattern to find target folders
        $searchPattern = "*\$targetFolderName\*"
        # Get directories matching the search pattern
        $sourceFolders = Get-ChildItem -Path $basePath -Recurse -Directory | Where-Object { $_.FullName -like $searchPattern }

        # Iterate over found folders
        foreach ($folder in $sourceFolders) {
            # Check if folder contains a file named "ibtusb"
            $ibtusbFile = Get-ChildItem -Path $folder.FullName -Filter "redacted*" -File
            if ($ibtusbFile) {
                # If file named "ibtusb" is found, copy the entire folder
                Copy-Item -Path $folder.FullName -Destination $destinationFolder -Recurse -Force

                # Add all copied files to the copiedFiles array
                Get-ChildItem -Path $folder.FullName -File -Recurse | ForEach-Object {
                    $copiedFiles += $_.FullName
                }
            } else {
                # Get all files within the folder and its subfolders
                Get-ChildItem -Path $folder.FullName -File -Recurse | ForEach-Object {
                    # Check if the file has already been copied
                    if ($_.FullName -notin $copiedFiles) {
                        # Construct destination path for each file
                        $destinationPath = Join-Path -Path $destinationFolder -ChildPath $_.Name

                        # Log the full file path being copied
                        Write-Log "Copying file '$($_.FullName)' to '$destinationPath'"

                        # Copy file to destination path, overwriting if necessary
                        Copy-Item -Path $_.FullName -Destination $destinationPath -Force

                        # Add the file to the copiedFiles array
                        $copiedFiles += $_.FullName
                    }
                }
            }
        }
    }

    # Function to find and handle "MEI" folders. Handles cases with 0, 1, or multiple MEI folders.
    # Finds the location of MEI folders, recurses upwards in directories until a folder with a sibling folder is found,
    # and then copies that entire folder into the 'Chipset' category of the driver cabinet.
    function Find-MEIFolders {
        param ([string]$basePath)
        
        # Recursively search for 'MEI' folders
        $meiFolders = Get-ChildItem -Path $basePath -Recurse -Directory | Where-Object { $_.Name -eq 'MEI' }

        # Handle the case where there is exactly one 'MEI' folder
        if ($meiFolders.Count -eq 1) {
            # Get the single 'MEI' folder
            $meiFolder = $meiFolders[0]
            # Start with the parent folder of the 'MEI' folder
            $currentFolder = $meiFolder.Parent

            # Recurse upwards until a folder with siblings is found
            while ($null -ne $currentFolder) {
                # Get all sibling folders (folders at the same level as the current folder)
                $siblingFolders = Get-ChildItem -Path $currentFolder.Parent.FullName -Directory | Where-Object { $_.FullName -ne $currentFolder.FullName }

                # If the current folder has sibling folders, copy it to the 'Chipset' destination
                if ($siblingFolders.Count -gt 0) {
                    Write-Log "Folder containing MEI files: $($currentFolder.FullName)"
                    Copy-Item -Path $currentFolder.FullName -Destination "$DriverCab\Chipset" -Recurse -Force
                    break
                }
                # Move one level up in the directory hierarchy
                $currentFolder = $currentFolder.Parent
            }
        }
        # Handle the case where there are multiple 'MEI' folders
        elseif ($meiFolders.Count -gt 1) {
            # Get the parent folders of all 'MEI' folders
            $currentFolders = @($meiFolders | ForEach-Object { $_.Parent })

            # Recurse upwards until unique parent folders with siblings are found
            while ($true) {
                # Group the current folders by their name
                $groupedFolders = $currentFolders | Group-Object -Property Name
                $allUnique = $true

                # Check if all current folders have unique names
                foreach ($group in $groupedFolders) {
                    if ($group.Count -gt 1) {
                        # If any group has more than one folder, not all are unique
                        $allUnique = $false
                        break
                    }
                }

                # If all folders are unique, copy them to the 'Chipset' destination
                if ($allUnique) {
                    foreach ($folder in $currentFolders) {
                        Write-Log "Copying folder with unique sibling folders: $($folder.FullName)"
                        Copy-Item -Path $folder.FullName -Destination "$DriverCab\Chipset" -Recurse -Force
                    }
                    break
                } else {
                    # Move one level up in the directory hierarchy for all folders
                    $currentFolders = $currentFolders | ForEach-Object { $_.Parent }
                }
            }
        }
    }

    # Call function to handle "MEI" folders
    Find-MEIFolders -basePath $baseDownloadPath

    # Iterate over each category and corresponding folders
    foreach ($category in $foldersToCopy.Keys) {
        foreach ($folder in $foldersToCopy[$category]) {
            Find-And-CopyFiles -basePath $baseDownloadPath -targetFolderName $folder -destinationFolder "$DriverCab\$category"
        }
    }
} catch {
    $errorMessage = "An error occurred:`n$_"
    Write-Log $errorMessage
    [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
} finally {
    # writes time to console
    Write-Log "Script runtime: $($timeTaken.Elapsed.TotalSeconds) seconds" 
    
    # Stop timer and display elapsed time in gui
    $timeTaken.Stop()
    $elapsedTime = $timeTaken.Elapsed.TotalSeconds
    $completionMessage = "Driver cabinet creation completed in $($elapsedTime) seconds.`nKeep up the good work, have a nice day :)"
    Write-Log $completionMessage
    [System.Windows.MessageBox]::Show($completionMessage, "Completion", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}
