<#
.SYNOPSIS
    Safely backup your default Minecraft .minecraft folders and install a Prism Launcher instance into them using a GUI file picker for the ZIP.
.DESCRIPTION
    This PowerShell script will:
      1. Prompt you to select the Prism Launcher instance ZIP via a GUI file dialog.
      2. Optionally accept an InstanceName parameter for subfolders inside the ZIP.
      3. Create a timestamped backup of key .minecraft folders.
      4. Safely extract and install folders from the Prism ZIP into your .minecraft directory.
      5. Skip missing source folders with warnings and create or clean destination folders.
      6. Copy the 'options.txt' file from the instance, if present.
      7. Handle errors gracefully and log all messages to a single GUI window on exit.
.PARAMETER InstanceName
    (Optional) Specific subfolder inside the ZIP to use. If omitted, assumes ZIP root contains the proper structure.
.EXAMPLE
    .\install_prism_instance.ps1
    or
    .\install_prism_instance.ps1 -InstanceName "MyInstanceFolder"
.NOTES
    Tested on Windows PowerShell 5.1+, no external modules required.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string] $InstanceName
)

# Initialize log storage
$logBuilder = New-Object System.Text.StringBuilder
function Log($msg)       { $null = $logBuilder.AppendLine("[INFO]  $msg") }
function LogWarn($msg)   { $null = $logBuilder.AppendLine("[WARN]  $msg") }
function LogError($msg)  { $null = $logBuilder.AppendLine("[ERROR] $msg") }

# Display log in a single GUI window
function Show-LogWindow {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text           = 'Prism Instance Installer Log'
    $form.Width          = 600
    $form.Height         = 400
    $form.StartPosition  = 'CenterScreen'

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline   = $true
    $textBox.ReadOnly    = $true
    $textBox.ScrollBars  = 'Vertical'
    $textBox.Dock        = 'Fill'
    $textBox.Text        = $logBuilder.ToString()

    $form.Controls.Add($textBox)
    $form.Add_Shown({ $form.Activate() })
    [void] $form.ShowDialog()
}

# Exit handler to show log before exit
function Exit-WithLog {
    param([int] $code)
    Show-LogWindow
    exit $code
}

# Start script
Log 'Script started.'

try {
    # Prompt for Prism ZIP via GUI
    Log 'Prompting for Prism ZIP file.'
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'ZIP Files (*.zip)|*.zip'
    $dlg.Title  = 'Select a Prism Launcher Instance ZIP'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        LogError 'No ZIP file selected. Exiting.'
        Exit-WithLog 1
    }
    $PrismZip = $dlg.FileName
    Log "Selected ZIP: $PrismZip"

    # Resolve and validate ZIP path
    Log 'Resolving ZIP path.'
    $PrismZipPath = Resolve-Path -LiteralPath $PrismZip -ErrorAction Stop
} catch {
    LogError "Error accessing ZIP: $_"
    Exit-WithLog 1
}

try {
    # Locate .minecraft
    $mcRoot = Join-Path $env:APPDATA '.minecraft'
    if (-not (Test-Path $mcRoot -PathType Container)) {
        throw ".minecraft folder not found at $mcRoot"
    }
    Log "Located .minecraft at $mcRoot"
} catch {
    LogError $_
    Exit-WithLog 1
}

# Define folders and backup path
$foldersToBackup = @('mods','config','resourcepacks','shaderpacks','saves','versions','jarmods','nativelibraries')
$timestamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupRoot     = "${mcRoot}_backup_$timestamp"

# Confirm action
Log 'Starting backup and install process.'

try {
    # Backup existing folders
    Log "Creating backup directory at $backupRoot"
    New-Item -ItemType Directory -Path $backupRoot -ErrorAction Stop | Out-Null
    foreach ($f in $foldersToBackup) {
        $src = Join-Path $mcRoot $f
        if (Test-Path $src) {
            Log "Backing up $f"
            Copy-Item -Path $src -Destination (Join-Path $backupRoot $f) -Recurse -ErrorAction Stop
        } else {
            LogWarn "Backup source missing: $f (skipped)"
        }
    }

    # Extract Prism ZIP to temporary folder
    $tempExtract = Join-Path $env:TEMP "PrismExtract_$timestamp"
    if (Test-Path $tempExtract) { Remove-Item -Recurse -Force -ErrorAction Stop }
    New-Item -ItemType Directory -Path $tempExtract -ErrorAction Stop | Out-Null
    Log "Extracting ZIP to $tempExtract"
    Expand-Archive -LiteralPath $PrismZipPath -DestinationPath $tempExtract -Force -ErrorAction Stop

    # Determine source root inside extracted ZIP
    $sourceRoot = if ($InstanceName) { Join-Path $tempExtract $InstanceName } else { $tempExtract }
    if (-not (Test-Path $sourceRoot -PathType Container)) {
        throw "Instance folder '$InstanceName' not found in ZIP"
    }
    Log "Initial source root: $sourceRoot"

    # Adjust for embedded 'minecraft' folder
    $minecraftSub = Join-Path $sourceRoot 'minecraft'
    if (Test-Path $minecraftSub -PathType Container) {
        Log "'minecraft' folder detected inside ZIP. Using it as source root."
        $sourceRoot = $minecraftSub
    } else {
        Log "No embedded 'minecraft' folder; using initial source root."
    }

    # Install folders: clear existing then copy
    foreach ($f in $foldersToBackup) {
        $src  = Join-Path $sourceRoot $f
        $dest = Join-Path $mcRoot     $f
        if (Test-Path $src) {
            if (Test-Path $dest) {
                Log "Clearing existing contents of: $f"
                Get-ChildItem -Path $dest -Force | Remove-Item -Recurse -Force -ErrorAction Stop
            } else {
                Log "Creating destination folder: $f"
                New-Item -ItemType Directory -Path $dest -ErrorAction Stop | Out-Null
            }
            Log "Copying $f"
            Copy-Item -Path (Join-Path $src '*') -Destination $dest -Recurse -ErrorAction Stop
        } else {
            LogWarn "Install source missing: $f (skipped)"
        }
    }

    # Copy options.txt from instance root
    $optionsSrc = Join-Path $sourceRoot 'options.txt'
    if (Test-Path $optionsSrc -PathType Leaf) {
        $optionsDest = Join-Path $mcRoot 'options.txt'
        Log "Copying options.txt"
        Copy-Item -Path $optionsSrc -Destination $optionsDest -Force -ErrorAction Stop
    } else {
        LogWarn "options.txt not found in instance, skipping."
    }

    # Cleanup temporary extraction
    Log "Removing temporary folder $tempExtract"
    Remove-Item -Recurse -Force $tempExtract -ErrorAction Stop

    Log 'Installation successful.'
    Exit-WithLog 0
} catch {
    LogError "Fatal error: $_"
    Exit-WithLog 1
}
