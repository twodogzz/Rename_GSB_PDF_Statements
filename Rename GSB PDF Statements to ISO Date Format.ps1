# Prompt user to select the source folder
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing GSB PDF statements"
$folderBrowser.ShowNewFolderButton = $false

# Load WinForms if needed
Add-Type -AssemblyName System.Windows.Forms

$dialogResult = $folderBrowser.ShowDialog()

if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No folder selected. Exiting."
    return
}

$sourceFolder = $folderBrowser.SelectedPath
Write-Host "Selected folder: $sourceFolder"

# Select multiple PDF files inside the chosen folder
$files = Get-ChildItem -Path $sourceFolder -Filter "GSB-*.pdf" |
         Out-GridView -Title "Select GSB PDF Statements to Rename" -PassThru

if (-not $files) {
    Write-Host "No files selected. Exiting."
    return
}

# Month lookup table
$monthMap = @{
    Jan = "01"; Feb = "02"; Mar = "03"; Apr = "04";
    May = "05"; Jun = "06"; Jul = "07"; Aug = "08";
    Sep = "09"; Oct = "10"; Nov = "11"; Dec = "12"
}

foreach ($file in $files) {

    # Match pattern: GSB-21Jan2026.pdf
    if ($file.BaseName -match '^GSB-(\d{1,2})([A-Za-z]{3})(\d{4})$') {

        $day   = $matches[1].PadLeft(2, '0')
        $monAbbrev = $matches[2]
        $year  = $matches[3]

        if (-not $monthMap.ContainsKey($monAbbrev)) {
            Write-Warning "Skipping '$($file.Name)' — unknown month abbreviation '$monAbbrev'"
            continue
        }

        $month = $monthMap[$monAbbrev]

        # Build new filename
        $newName = "GSB-$year-$month-$day.pdf"

        try {
            Rename-Item -LiteralPath $file.FullName -NewName $newName -ErrorAction Stop
            Write-Host "Renamed: $($file.Name) → $newName"
        }
        catch {
            Write-Warning "Failed to rename '$($file.Name)': $_"
        }

    }
    else {
        Write-Warning "Skipping '$($file.Name)' — filename does not match expected pattern."
    }
}