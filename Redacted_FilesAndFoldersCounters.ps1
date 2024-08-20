$Current_user = (Get-WmiObject -Class win32_computersystem).UserName.split('\')[1]

# Adds necessary .NET assemblies to use Windows forms for GUI input
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# Define the base test path
$baseTestPath = "C:\redacted\$Current_user\redacted"

# Check if the specified folder exists in the Test directory
if (-Not (Test-Path -Path $baseTestPath)) {
    Write-Error "The 'redacted' directory does not exist. Please check the path and try again."
    exit
}

# Function to count files and folders, and folders directly containing files
function Get-FilesAndFoldersCount {
    param (
        [string]$path
    )

    $items = Get-ChildItem -Path $path -Recurse
    $totalFiles = ($items | Where-Object { -not $_.PSIsContainer }).Count
    $totalFolders = ($items | Where-Object { $_.PSIsContainer }).Count

    # Count folders that directly contain files
    $foldersWithFiles = $items | Where-Object { $_.PSIsContainer } | ForEach-Object {
        $subItems = Get-ChildItem -Path $_.FullName
        if ($subItems | Where-Object { -not $_.PSIsContainer }) {
            $_
        }
    }
    $totalFoldersWithFiles = $foldersWithFiles.Count

    return [PSCustomObject]@{
        TotalFiles = $totalFiles
        TotalFolders = $totalFolders
        TotalFoldersWithFiles = $totalFoldersWithFiles
    }
}

# List all subfolders in the base test path
$subfolders = Get-ChildItem -Path $baseTestPath -Directory

$totalFilesList = @()
$totalFoldersList = @()
$totalFoldersWithFilesList = @()

foreach ($subfolder in $subfolders) {
    Write-Host "Processing folder: $($subfolder.Name)"
    $countResults = Get-FilesAndFoldersCount -path $subfolder.FullName
    Write-Host "Total number of files in $($subfolder.Name): $($countResults.TotalFiles)"
    Write-Host "Total number of folders in $($subfolder.Name): $($countResults.TotalFolders)"
    Write-Host "Total number of folders that directly contain files in $($subfolder.Name): $($countResults.TotalFoldersWithFiles)"

    $totalFilesList += $countResults.TotalFiles
    $totalFoldersList += $countResults.TotalFolders
    $totalFoldersWithFilesList += $countResults.TotalFoldersWithFiles
}

# Calculate averages
$averageFiles = ($totalFilesList | Measure-Object -Average).Average
$averageFolders = ($totalFoldersList | Measure-Object -Average).Average
$averageFoldersWithFiles = ($totalFoldersWithFilesList | Measure-Object -Average).Average

Write-Host "Average total number of files: $averageFiles"
Write-Host "Average total number of folders: $averageFolders"
Write-Host "Average total number of folders that directly contain files: $averageFoldersWithFiles"

