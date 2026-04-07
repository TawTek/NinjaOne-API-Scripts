<#
.SYNOPSIS
Downloads all scripts from Ninja's Automation library to local folder

.DESCRIPTION
Scripts are filtered to exclude built-in or native scripts and are then downloaded to a specified directory.
Initializes arrays to store script attributes and failed scripts, and a hashtable to track processed names.
Iterates through each script, setting the appropriate file extension based on the script's language.
Invalid characters in the script names are replaced, and duplicate names are handled by appending a "-copy#" suffix.
Script content, which is Base64-encoded, is decoded into a UTF-8 string and written to a file.
Progress is tracked and displayed, and any scripts that fail to save are saved to $global:FailedScripts array.
Details of the downloaded scripts are exported to a CSV file for reference that includes description (if any).

.PARAMETER ScriptFolder
The directory where the scripts will be downloaded.

.NOTES
Author: Tawhid Chowdhury
Date:   2024-10-23
#>

<#--------------------------------------------------------------------------------------------------------------------------
SCRIPT:WEBSESSION
--------------------------------------------------------------------------------------------------------------------------#>

# Paste web session code here from dev console

<#--------------------------------------------------------------------------------------------------------------------------
SCRIPT:PARAM_VAR
--------------------------------------------------------------------------------------------------------------------------#>

$Scripts       = $ScriptsResponse.content | ConvertFrom-Json | Where-Object { $_.language -notin @("native", "binary_install") }
$ScriptBaseURL = $ScriptsResponse.BaseResponse.RequestMessage.RequestUri.AbsoluteUri # URL of scripts library instance
$ScriptFolder  = "C:\Temp\NinjaScripts\Test" # sets directory to download scripts
$TotalScripts  = $Scripts.Count # total number of scripts in automation library

$global:FailedScripts   = @() # array to store names of failed scripts
$global:MultipleScripts = @() # array to store scripts with multiple categories
$global:ProcessedScript = @{} # hashtable to track processed names
$global:ScriptArray     = @() # array to store all script properties
$global:SavedScripts    = @() # Initialize this as an empty array

<#--------------------------------------------------------------------------------------------------------------------------
SCRIPT:FUNCTIONS
--------------------------------------------------------------------------------------------------------------------------#>

function Get-Scripts {

    param(
    [switch]$Debug
    )
    $CurrentScript = 0

    foreach ($Script in $Scripts) { # loop through each script in library

        switch ($Script.language) { # set file extension based on language
            'powershell' { $FileExtension = '.ps1' }
            'batchfile'  { $FileExtension = '.bat' }
            'vbscript'   { $FileExtension = '.vbs' }
            'sh'         { $FileExtension = '.sh' }
            default      { throw "Error: Unknown language type $($Script.language)" }
        }

        # Replaces invalid characters, trims whitespace, combines script name + extension
        $BaseFileName   = (($Script.name -replace '[\\/:*?"<>&|]', '_').TrimStart())
        $ScriptFileName = $BaseFileName + $FileExtension

        # Check for scripts in library with same filename
        if ($global:ProcessedScript.ContainsKey($BaseFileName)) { # first iteration -eq $false and else block executes
            $global:ProcessedScript[$BaseFileName]++ # count incremented if $true
            $ScriptFileName = "$BaseFileName-copy$($global:ProcessedScript[$BaseFileName])$FileExtension" # append '-copy#'
        } else {
            $global:ProcessedScript[$BaseFileName] = 0 # first iteration initalizes count for subsequent loops
        }

        # Update script count + progress bar that outputs the $ScriptFileName being processed
        $CurrentScript++
        $PercentComplete = [math]::Round(($CurrentScript / $TotalScripts) * 100)
        Write-Progress -Activity "Downloading Scripts" `
            -Status "$CurrentScript/$TotalScripts | $PercentComplete% Complete | $ScriptFileName" `
            -PercentComplete $PercentComplete       

        # Store script content in variable for easier property sourcing
        $ScriptContent = (Invoke-WebRequest -UseBasicParsing -Uri "$ScriptBaseURL/$($Script.id)" `
                        -WebSession $session -ContentType "application/json").Content | ConvertFrom-Json

        # Get category name(s)
        $CategoryName = ((Get-ScriptCategory -CategoryIDs $ScriptContent.categoriesIds) -replace '[\\/:*?"<>|]', '_')
        Set-Directory -Path "$ScriptFolder\$CategoryName" -Create # create category folder

        # Checks and skips iteration if multiple categories exist for script
        if ($ScriptContent.categoriesIds.Count -gt 1) {
            $global:MultipleScripts    += [PSCustomObject]@{ # store scripts name and category in custom array
                ScriptName    = $Script.name
                CategoryNames = $CategoryName
            }
            continue
        }

        # Decode the Base64-encoded string into a UTF-8 string so that script is readable
        $ScriptCode = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptContent.code))

        # Create a PSCustomObject to store script details
        $global:ScriptArray += [PSCustomObject]@{
            Name        = $Script.name
            Language    = $Script.language
            Category    = $CategoryName
            Description = $Script.description
            FileName    = $ScriptFileName
            FilePath    = Join-Path "$ScriptFolder\$CategoryName" $ScriptFileName
            Code        = $ScriptCode
        } 
        if ($Debug) { if (++$counter -ge 3) { break } }
    }
    Write-Progress -Activity "Downloading Scripts" -Completed
}

function Write-Script {
    foreach ($Script in $global:ScriptArray) { # loop through each script in the array            
        try {
            # Write the script code to the specified file path
            $Script.Code | Set-Content -Path $Script.FilePath -NoNewline
        } catch {
            try {
                # Fallback method using .NET to write the file
                [System.IO.File]::WriteAllText($Script.FilePath, $Script.Code)
            } catch {
                # If both methods fail, log the script name
                $global:FailedScripts += $Script.FileName # add failed script to array
                $Failed = $true
            }
        }
        if (!$Failed) {
            # If the script was saved successfully, log its details
            $global:SavedScripts += [PSCustomObject]@{
                Name        = $Script.name
                Language    = $Script.language
                Category    = $Script.category
                Description = $Script.description
                FilePath    = $Script.FilePath
            }
        }
    }
    Write-Progress -Activity "Writing to file" -Completed
}

function Get-Results {
    # Warns about scripts with multiple categories not being saved
    if ($global:MultipleScripts) {
        Write-Warning "The following scripts with multiple categories were not saved."
        Write-Warning "Go to Ninja Automation Library amd set scripts to one category only."
        $global:MultipleScripts | Format-Table -AutoSize
    }
    if ($global:FailedScripts.Count -gt 0) {
        Write-Output "The following scripts failed to save:"
        $global:FailedScripts | ForEach-Object { Write-Output $_ }
    }
    # Export list of scripts saved to CSV file
    $global:SavedScripts | Select-Object Name,Language,Category,Description,FilePath | Sort-Object Category | Export-Csv -Path "$ScriptFolder\Scripts.csv" -NoTypeInformation
    Write-Host "Exported saved script details to $ScriptFolder\Scripts.csv." -ForegroundColor Cyan    
}

<#--------------------------------------------------------------------------------------------------------------------------
SCRIPT:ANCILLARY_FUNCTIONS
--------------------------------------------------------------------------------------------------------------------------#>

function Get-ScriptCategory {
    param (
        [int[]]$CategoryIDs
    )

    $CategoriesHash = @{
        1   = 'Uncategorized'
        8   = '$ WS'
        140 = '# App'
        147 = '# Troubleshoot'
        167 = '! Client'
        169 = '$ KF'
    }

    $CategoryNames = @()
    foreach ($CategoryID in $CategoryIDs) {
        if ($CategoriesHash.ContainsKey($CategoryID)) {
            $CategoryNames += $CategoriesHash[$CategoryID]
        }
    }
    return $CategoryNames -join ", "
}

function Set-Directory {
    [CmdletBinding()]
    param (
        [string]$Path,
        [switch]$Create,
        [switch]$Remove
    )

    switch ($true) {
        { $Create.IsPresent } {
            if (-not (Test-Path -Path $Path)) {
                try {
                    Write-Verbose "Creating directory at $Path"
                    New-Item -Path $Path -ItemType "Directory" | Out-Null
                    Write-Verbose "Created diretory at $Path."
                } catch {
                    Write-Host "ERROR: Failed to create directory. $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Verbose "Directory exists at $Path"
            }
        }
        { $Remove.IsPresent } {
            try {
                Write-Verbose "Deleting directory."
                Remove-Item -Path $Path -Recurse -Force @EA_Stop
                Write-Verbose "Directory deleted."
            } catch {
                Write-Host "ERROR: Failed to remove directory. $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

<#--------------------------------------------------------------------------------------------------------------------------
SCRIPT:EXECUTIONS
--------------------------------------------------------------------------------------------------------------------------#>

Set-Directory -Path $ScriptFolder -Create
Get-Scripts
Write-Script
Get-Results