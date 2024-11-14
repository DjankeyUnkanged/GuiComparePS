# GUI compare PowerShell script
# 

# Load the necessary assembly
Add-Type -AssemblyName System.Windows.Forms # This is for Explorer open/save/browse prompts
Add-Type -AssemblyName PresentationFramework # This is for GUI alert/dialog boxes
[System.Windows.Forms.Application]::EnableVisualStyles()

# Set metadata/system folder exception variable
$ExceptionList = ".Trashes",".Spotlight-V100",".fseventsd","System Volume Information"

# Define MessageBox function so all MessageBox objects appear on top
function Show-MessageBox {
    param (
        [string]$Message,
        [string]$Title,
        [string]$Buttons,
        [string]$Icon
    )

    # Preset parameters for the MessageBox
    $MsgWindow = New-Object System.Windows.Window
    $MsgWindow.Topmost = $true
    $MsgWindow.WindowStyle = 'None'
    $MsgWindow.ShowInTaskbar = $false
    $MsgWindow.ShowActivated = $false
    $MsgWindow.Width = 0
    $MsgWindow.Height = 0
    $MsgWindow.Show()

    # Show the MessageBox with $MsgWindow as the parent object
    $result = [System.Windows.MessageBox]::Show($MsgWindow, $Message, $Title, $Buttons, $Icon)

    # Close the parent window
    $MsgWindow.Close()

    return $result
}

# Define Xaml for progress bar window
$xamlTemplate = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{0}" Height="100" Width="400" WindowStartupLocation="CenterScreen" Topmost="True">
    <Grid>
        <ProgressBar Name="progressBar" Width="350" Height="30" Minimum="0" Maximum="100" Value="{1}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Grid>
</Window>
"@

function Show-ProgressBar {
    param (
        [int]$Progress,
        [string]$Title,
        [switch]$Done
    )

    begin {
        if (-not $script:window) {
            # Replace placeholders in the XAML template
            $xaml = [string]::Format($xamltemplate, $Title, $Progress)

            # Load the XAML
            $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader $xaml))
            $script:window = [Windows.Markup.XamlReader]::Load($reader)

            # Find the progress bar
            $script:progressBar = $script:window.FindName("progressBar")

            # Show the window
            $script:window.Show()
        }
    }

    process {
        if ($Done) {
            # Close the window once switch is called
            $script:window.Close()
            Remove-Variable -Name window -Scope Script
            Remove-Variable -Name progressBar -Scope Script
        } else {
            # Update progress
            $script:progressBar.Value = $Progress
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([action] { }, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    }
}

# Function to select a folder
function Select-FolderDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Please select a folder."
    $folderBrowser.ShowNewFolderButton = $true
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    $folderBrowser.ShowDialog($form) | Out-Null
    return $folderBrowser.SelectedPath
}

# Function to select a file save location
function Select-SaveFileDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Save Comparison Results"
    $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments)
    $saveFileDialog.ShowDialog($form) | Out-Null
    return $saveFileDialog.FileName
}
   

# Pick first disk or directory to compare
Show-MessageBox -Message "In the following window, please choose the first file or folder to be compared." -Title 'Select file or folder' -Buttons 'OK' -Icon 'Information'
$DataA = Select-FolderDialog
    
# Throw an error and exit if the choice is invalid
if ($DataA -eq "") {
    Show-MessageBox -Message 'The selection does not appear to be valid. Exiting...' -Title 'Invalid selection' -Buttons 'OK' -Icon 'Exclamation'
    return
}

# Pick second disk or directory to compare
Show-MessageBox -Message "In the following window, please choose the second file or folder to be compared." -Title 'Select file or folder' -Buttons 'OK' -Icon 'Information'
$DataB = Select-FolderDialog

# Throw an error and exit if both selections are the same or if selection invalid
if (($DataB -eq "") -or ($DataB -ieq $DataA)) {
    Show-MessageBox -Message 'Both selections are the same, or selection is otherwise invalid. Exiting...' -Title 'Invalid selection' -Buttons 'OK' -Icon 'Exclamation'
    return
}

# Run a loop - For each file in the source, gather data and get the SHA256 of each file. Mark each entry as 'Source' in the array.
$NumA = 0
$TotalSelA = (Get-ChildItem $DataA -Recurse -File -Exclude $ExceptionList).Count
$FilesSelA = Get-ChildItem $DataA -Recurse -File -Exclude $ExceptionList | ForEach-Object {
    [PSCustomObject]@{
        Path = $_.FullName
        Hash = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
        Size = $_.Length
        Name = $_.Name
        Source = 'Source'
    }
    $NumA++
    $SelAProgress = [math]::Round(($NumA / $TotalSelA) * 100)
    Show-ProgressBar -Title "Source Hash Progress" -Progress $SelAProgress
}
Show-ProgressBar -Done

# Run a loop - For each file in the destination, gather data and get the SHA256 of each file. Mark each entry as 'Destination' in the array.
$NumB = 0
$TotalSelB = (Get-ChildItem $DataB -Recurse -File -Exclude $ExceptionList).Count
$FilesSelB = Get-ChildItem $DataB -Recurse -File -Exclude $ExceptionList | ForEach-Object {
    [PSCustomObject]@{
        Path = $_.FullName
        Hash = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
        Size = $_.Length
        Name = $_.Name
        Source = 'Destination'
    }
    $NumB++
    $SelBProgress = [math]::Round(($NumB / $TotalSelB) * 100)
    Show-ProgressBar -Title "Destination Hash Progress" -Progress $SelBProgress
}
Show-ProgressBar -Done

# Initialize comparison array
$Comparison = @()
    
# Create array by adding the source file and destination file arrays together into one, organizing by SHA256 hash and file name. If all goes well in the copy, there should be two of every file (1 in source, 1 in destination), no more, no less.
# Grouping by object name alone doesn't work because sources can have multiple files with the same name, but different hashes - this throws the array off, breaking the script. Using two parameters avoids this issue.
$FileGroups = ($FilesSelA + $FilesSelB) | Group-Object Hash,Name
    
# Initialize variable (think canary in a coal mine)
$FailCanary = 'True'

# For each grouping, compare the hash between source and destination. If there is a single mismatch, our FailCanary variable switches to false to set up an error message.
foreach ($group in $FileGroups) {
    $GroupA = $group.Group | Where-Object { $_.Source -eq 'Source' }
    $GroupB = $group.Group | Where-Object { $_.Source -eq 'Destination' }

    $Comparison += [PSCustomObject]@{
        Name = if ([string]::IsNullOrEmpty($GroupA.Name)) { $GroupB.Name } else { $GroupA.Name }
        PathA = $GroupA.Path
        PathB = $GroupB.Path
        HashA = $GroupA.Hash
        HashB = $GroupB.Hash
        FileASizeInBytes = $GroupA.Size
        FileBSizeInBytes = $GroupB.Size
        Match = ($GroupA.Hash -eq $GroupB.Hash)
    }
    if ($GroupA.Hash -ne $GroupB.Hash) {$FailCanary = 'False'}
}

# Here's where that FailCanary variable comes in. If the variable has been set to false, there was a mismatch or unhandled error somewhere along the way
if ($FailCanary -ieq 'False') {
    Show-MessageBox -Message 'At least one file has failed to verify. Please check your media and try again.' -Title 'Mismatch detected' -Buttons 'OK' -Icon 'Exclamation'
}
    
# If the script made it to this point, the copy and compare has worked up to this point, even if there was a file mismatch or two. Now to choose the destination for the CSV containing the comparison results.
Show-MessageBox -Message "Comparison complete! In the following window, please choose where you would like to save the results." -Title 'Compare done!' -Buttons 'OK' -Icon 'Information'
$CsvPath = Select-SaveFileDialog
if ($CsvPath -eq "") {
    Show-MessageBox -Message 'The destination does not appear to be valid. Exiting...' -Title 'Invalid destination' -Buttons 'OK' -Icon 'Exclamation'
    return
}
$Comparison | Export-Csv -Path $CsvPath -NoTypeInformation

Show-MessageBox -Message "Comparison results have been saved to $CsvPath." -Title 'Done!' -Buttons 'OK' -Icon 'Information'

# (Optional) Display the comparison table in console
# $Comparison | Format-Table -Property Name, PathA, PathB, HashA, HashB, FileASizeInBytes, FileBSizeInBytes, Match