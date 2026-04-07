<#
.SYNOPSIS
Downloads all scripts from Ninja's Automation library with SSO/MFA support

.DESCRIPTION
Uses Selenium browser automation to handle SSO/MFA authentication.
Opens browser for login, captures session cookies automatically, then downloads all scripts.
Supports all authentication methods including SSO and multi-factor authentication.

.PARAMETER ScriptFolder
The directory where scripts will be downloaded

.NOTES
Author: Tawhid Chowdhury
Date: 2025-10-14
Requires: Selenium PowerShell module (auto-installed if missing)
Supports: Edge or Chrome browser
#>

param(
    [string]$ScriptFolder = "C:\Temp\NinjaScripts",
    [switch]$ClearCache,
    [switch]$Sequential
)

function Get-NinjaSession {
    <#
    .SYNOPSIS
    Authenticates to NinjaRMM using browser-based SSO/MFA and returns web session.
    
    .DESCRIPTION
    Opens browser for SSO/MFA login, captures session cookies automatically.
    Supports all authentication methods including SSO and MFA.
    Caches session for reuse to avoid repeated authentication.
    
    .PARAMETER ForceClear
    Force clear the cached session
    #>
    
    param(
        [switch]$ForceClear
    )
    
    $SessionCachePath = "$env:USERPROFILE\.ninja_session.xml"
    $CacheExpiryHours = 2  # Session valid for 2 hours
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "NinjaRMM Browser Authentication" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    # Clear cache if requested
    if ($ForceClear -and (Test-Path $SessionCachePath)) {
        Remove-Item $SessionCachePath -Force
        Write-Host "Session cache cleared." -ForegroundColor Yellow
    }
    
    # Check cached session
    if (Test-Path $SessionCachePath) {
        try {
            $CachedSession = Import-Clixml -Path $SessionCachePath
            
            if ($CachedSession.ExpiresAt -gt (Get-Date)) {
                $HoursRemaining = [math]::Round(($CachedSession.ExpiresAt - (Get-Date)).TotalHours, 1)
                Write-Host "Using cached session (expires in $HoursRemaining hours)..." -ForegroundColor Green
                
                # Recreate WebSession from cached data
                $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                $session.UserAgent = $CachedSession.UserAgent
                
                # Decrypt session key from DPAPI-encrypted SecureString
                $SecureKey = $CachedSession.SessionKey | ConvertTo-SecureString
                $BStr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureKey)
                try {
                    $DecryptedSessionKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BStr)
                } finally {
                    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BStr)
                }
                
                $session.Cookies.Add((New-Object System.Net.Cookie("sessionKey", $DecryptedSessionKey, "/", "app.ninjarmm.com")))
                
                # Test if session is still valid
                try {
                    Invoke-WebRequest -Uri "https://app.ninjarmm.com/swb/s21/scripting/scripts" `
                        -Method Get `
                        -WebSession $session `
                        -Headers @{ "Accept" = "application/json" } `
                        -ErrorAction Stop | Out-Null
                    
                    Write-Host "Cached session is valid!`n" -ForegroundColor Green
                    return $session
                } catch {
                    Write-Host "Cached session expired, re-authenticating..." -ForegroundColor Yellow
                }
            } else {
                Write-Host "Cached session expired, re-authenticating..." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Could not load cached session, re-authenticating..." -ForegroundColor Yellow
        }
    }
    
    # Import Selenium module
    if (-not (Get-Module -ListAvailable -Name Selenium)) {
        throw "Selenium module not found. Please run: Install-Module -Name Selenium"
    }
    Import-Module Selenium -ErrorAction Stop
    
    Write-Host "`nOpening browser for SSO/MFA login..." -ForegroundColor Yellow
    Write-Host "Please complete authentication in the browser window." -ForegroundColor Cyan
    Write-Host "The browser will close automatically after login.`n" -ForegroundColor Cyan
    
    $Driver = $null
    
    try {
        Write-Host "Starting Microsoft Edge..." -ForegroundColor Cyan
        
        # Get primary monitor resolution dynamically
        Add-Type -AssemblyName System.Windows.Forms
        $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
        $screenWidth = $primaryScreen.Bounds.Width
        $screenHeight = $primaryScreen.Bounds.Height
        Write-Host "Detected primary monitor: ${screenWidth}x${screenHeight}" -ForegroundColor DarkGray
        
        # Create Edge driver service with explicit path
        $driverPath = "C:\Git\tawtek\ninjaone-api-scripts"
        $service = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($driverPath, "msedgedriver.exe")
        $service.HideCommandPromptWindow = $true
        
        # Create Edge options to start window off-screen
        $options = New-Object OpenQA.Selenium.Edge.EdgeOptions
        
        # Try different methods to set window position off-screen
        try {
            # Try AddArguments (plural) first
            $options.AddArguments("--window-position=-32000,-32000")
        } catch {
            try {
                # Fallback to AddAdditionalCapability with raw EdgeDriver format
                $options.AddAdditionalCapability("ms:edgeOptions", @{ args = @("--window-position=-32000,-32000") })
            } catch {
                Write-Warning "Could not set off-screen position, window may flash briefly"
            }
        }
        
        # Create driver with off-screen window
        $Driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($service, $options)
        
        # Calculate window size and center position
        $windowWidth = 1100
        $windowHeight = 700
        $centerX = [Math]::Max(50, ($screenWidth - $windowWidth) / 2)
        $centerY = [Math]::Max(50, ($screenHeight - $windowHeight) / 2)
        
        Write-Host "Positioning window at (${centerX}, ${centerY})" -ForegroundColor DarkGray
        
        $windowSize = New-Object System.Drawing.Size($windowWidth, $windowHeight)
        $windowPosition = New-Object System.Drawing.Point([int]$centerX, [int]$centerY)
        
        # Navigate to URL while off-screen
        $Driver.Navigate().GoToUrl("https://app.ninjarmm.com/")
        
        # Wait less time for page to start loading
        Start-Sleep -Milliseconds 300
        
        # Move window to center and set size - window appears here
        $Driver.Manage().Window.Size = $windowSize
        $Driver.Manage().Window.Position = $windowPosition
        
        Write-Host "Edge browser started successfully (centered on screen).`n" -ForegroundColor Green
        
        $StartTime = Get-Date
        $TimeoutSeconds = 300
        
        Write-Host "Waiting for authentication..." -ForegroundColor Cyan
        
        while (((Get-Date) - $StartTime).TotalSeconds -lt $TimeoutSeconds) {
            Start-Sleep -Seconds 2
            
            try {
                $Cookies = $Driver.Manage().Cookies.AllCookies
                $SessionCookie = $Cookies | Where-Object { $_.Name -eq 'sessionKey' }
                
                if ($SessionCookie -and $SessionCookie.Value) {
                    Write-Host "Authentication successful! Session captured.`n" -ForegroundColor Green
                    
                    # Get real UserAgent from browser
                    $realUserAgent = $Driver.ExecuteScript("return navigator.userAgent")
                    
                    # Create WebSession and add the cookie
                    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $session.UserAgent = $realUserAgent
                    $session.Cookies.Add((New-Object System.Net.Cookie("sessionKey", $SessionCookie.Value, "/", "app.ninjarmm.com")))
                    
                    # Cache the session with DPAPI encryption
                    $SessionData = @{
                        SessionKey = $SessionCookie.Value | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
                        UserAgent = $realUserAgent
                        ExpiresAt = (Get-Date).AddHours($CacheExpiryHours)
                    }
                    $SessionData | Export-Clixml -Path $SessionCachePath
                    Write-Host "Session cached at: $SessionCachePath" -ForegroundColor DarkGray
                    Write-Host "Cache expires: $($SessionData.ExpiresAt)" -ForegroundColor DarkGray
                    
                    Start-Sleep -Seconds 1
                    return $session
                }
            } catch {
                # Continue waiting
            }
        }
        
        throw "Timeout: Authentication not completed within $TimeoutSeconds seconds."
        
    } finally {
        if ($Driver) {
            $Driver.Quit()
        }
    }
}

try {
    $session = Get-NinjaSession -ForceClear:$ClearCache
} catch {
    Write-Host "`nERROR: Authentication failed - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "Fetching categories from NinjaRMM..." -ForegroundColor Cyan

try {
    $CategoriesResponse = Invoke-RestMethod -Uri "https://app.ninjarmm.com/swb/s21/scripting/categories" `
        -Method Get `
        -WebSession $session `
        -Headers @{
            "Accept" = "application/json"
        } `
        -ErrorAction Stop
    
    Write-Host "Successfully retrieved categories!" -ForegroundColor Green
    
} catch {
    Write-Host "WARNING: Failed to fetch categories - $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "Using fallback categories..." -ForegroundColor Yellow
    $CategoriesResponse = @()
}

Write-Host "Fetching scripts from NinjaRMM..." -ForegroundColor Cyan

try {
    $ScriptsResponse = Invoke-RestMethod -Uri "https://app.ninjarmm.com/swb/s21/scripting/scripts" `
        -Method Get `
        -WebSession $session `
        -Headers @{
            "Accept" = "application/json"
        } `
        -ErrorAction Stop
    
    Write-Host "Successfully retrieved scripts!`n" -ForegroundColor Green
    
} catch {
    Write-Host "`nERROR: Failed to fetch scripts - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Initialize variables
$Scripts = $ScriptsResponse | Where-Object { $_.language -notin @("native", "binary_install", "file_transfer") }
$ScriptBaseURL = "https://app.ninjarmm.com/swb/s21/scripting/scripts"
$TotalScripts = $Scripts.Count

# Build dynamic categories hash from API response
$ExcludedCategories = @('Hardware', 'Linux OS Patching', 'Mac OS Patching', 'Maintenance', 'Patching')
$CategoriesHash = @{}
if ($CategoriesResponse -and $CategoriesResponse.Count -gt 0) {
    Write-Host "Building categories from API response ($($CategoriesResponse.Count) categories)..." -ForegroundColor Cyan
    $ExcludedCount = 0
    $CategoryStatus = @()
    
    foreach ($Category in $CategoriesResponse) {
        if ($Category.id -and $Category.name) {
            # Skip excluded native categories
            if ($ExcludedCategories -contains $Category.name) {
                $ExcludedCount++
                $CategoryStatus += [PSCustomObject]@{
                    Category = $Category.name
                    Path = 'Excluded'
                }
                continue
            }
            $CategoriesHash[[int]$Category.id] = $Category.name
            $CategoryStatus += [PSCustomObject]@{
                Category = $Category.name
                Path = "$ScriptFolder\$($Category.name)"
            }
        }
    }
    
    Write-Host "Successfully mapped $($CategoriesHash.Count) categories! (Excluded $ExcludedCount native categories)" -ForegroundColor Green
    Write-Host "`nCategory Mapping Status:" -ForegroundColor Cyan
    $CategoryStatus | Sort-Object Path, Category | Format-Table -AutoSize
    Write-Host ""
} else {
    Write-Host "Using fallback hardcoded categories..." -ForegroundColor Yellow
    $CategoriesHash = @{
        1   = 'Uncategorized'
        8   = '$ WS'
        140 = '# App'
        147 = '# Troubleshoot'
        167 = '! Client'
        169 = '$ KF'
    }
}

Write-Host "Processing $TotalScripts scripts..." -ForegroundColor Cyan

# Helper Functions
function Get-ScriptCategory {
    param(
        [int[]]$CategoryIDs,
        [hashtable]$CategoriesHash
    )
    
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
    param(
        [string]$Path,
        [switch]$Create
    )
    
    if ($Create.IsPresent) {
        if (-not (Test-Path -Path $Path)) {
            try {
                New-Item -Path $Path -ItemType "Directory" -Force | Out-Null
            } catch {
                Write-Host "ERROR: Failed to create directory. $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

# Pre-create all category directories to avoid race conditions in parallel mode
foreach ($CategoryName in $CategoriesHash.Values) {
    Set-Directory -Path "$ScriptFolder\$CategoryName" -Create
}
Set-Directory -Path "$ScriptFolder\- Duplicate-Categories" -Create

# Check PowerShell version and use parallel processing if available
if ($PSVersionTable.PSVersion.Major -ge 7 -and -not $Sequential) {
    Write-Host "Using parallel processing (ThrottleLimit: 20) to download scripts." -ForegroundColor Green
    
    # Create thread-safe concurrent dictionaries
    $ProcessedScripts = [System.Collections.Concurrent.ConcurrentDictionary[string,int]]::new()
    $ProgressCounter = [System.Collections.Concurrent.ConcurrentDictionary[string,int]]::new()
    $ProgressCounter['completed'] = 0
    $ProgressCounter['failed'] = 0
    
    $ScriptArray = $Scripts | ForEach-Object -Parallel {
        $Script = $_
        $session = $using:session
        $ScriptBaseURL = $using:ScriptBaseURL
        $ScriptFolder = $using:ScriptFolder
        $CategoriesHash = $using:CategoriesHash
        $ProcessedScripts = $using:ProcessedScripts
        $ProgressCounter = $using:ProgressCounter
        $TotalScripts = $using:TotalScripts
        
        switch ($Script.language) {
            'powershell' { $FileExtension = '.ps1' }
            'batchfile'  { $FileExtension = '.bat' }
            'vbscript'   { $FileExtension = '.vbs' }
            'sh'         { $FileExtension = '.sh' }
            default      { return }
        }
        
        $BaseFileName = (($Script.name -replace '[\\/:*?"<>&|]', '_').TrimStart())
        
        # Thread-safe duplicate filename handling
        $copyNumber = $ProcessedScripts.AddOrUpdate(
            $BaseFileName,
            0,
            { param($key, $oldValue) $oldValue + 1 }
        )
        
        if ($copyNumber -eq 0) {
            $ScriptFileName = $BaseFileName + $FileExtension
        } else {
            $ScriptFileName = "$BaseFileName-copy$copyNumber$FileExtension"
        }
        
        try {
            $ScriptContent = Invoke-RestMethod -Uri "$ScriptBaseURL/$($Script.id)" `
                -Method Get `
                -WebSession $session `
                -Headers @{
                    "Accept" = "application/json"
                } `
                -ErrorAction Stop
            
            # Determine category
            if ($ScriptContent.categoriesIds -and $ScriptContent.categoriesIds.Count -gt 1) {
                $CategoryName = "- Duplicate-Categories"
                $OriginalCategoryNames = @()
                foreach ($CategoryID in $ScriptContent.categoriesIds) {
                    if ($CategoriesHash.ContainsKey([int]$CategoryID)) {
                        $OriginalCategoryNames += $CategoriesHash[[int]$CategoryID]
                    }
                }
                if ($OriginalCategoryNames.Count -eq 0) {
                    $CategoryNamesStr = "Unknown IDs: $($ScriptContent.categoriesIds -join ', ')"
                } else {
                    $CategoryNamesStr = ($OriginalCategoryNames -join ", ")
                }
            } elseif ($ScriptContent.categoriesIds -and $ScriptContent.categoriesIds.Count -eq 1) {
                $CategoryID = [int]$ScriptContent.categoriesIds[0]
                if ($CategoriesHash.ContainsKey($CategoryID)) {
                    $CategoryName = $CategoriesHash[$CategoryID]
                    $CategoryNamesStr = $CategoryName
                } else {
                    $CategoryName = 'Uncategorized'
                    $CategoryNamesStr = "Unknown ID: $CategoryID"
                }
            } else {
                $CategoryName = 'Uncategorized'
                $CategoryNamesStr = 'Uncategorized'
            }
            
            $ScriptCode = [System.Text.Encoding]::UTF8.GetString(
                [System.Convert]::FromBase64String($ScriptContent.code)
            )
            
            # Write file immediately - don't store code in memory
            $FilePath = Join-Path "$ScriptFolder\$CategoryName" $ScriptFileName
            $WriteSuccess = $false
            try {
                $ScriptCode | Set-Content -Path $FilePath -NoNewline -ErrorAction Stop
                $WriteSuccess = $true
            } catch {
                try {
                    [System.IO.File]::WriteAllText($FilePath, $ScriptCode)
                    $WriteSuccess = $true
                } catch {
                    # Write failed, will be marked as failed
                }
            }
            
            # Return metadata only (no code)
            [PSCustomObject]@{
                Name              = $Script.name
                Language          = $Script.language
                Category          = $CategoryName
                OriginalCategories = $CategoryNamesStr
                Description       = $Script.description
                FileName          = $ScriptFileName
                FilePath          = $FilePath
                Failed            = -not $WriteSuccess
                MultiCat          = ($ScriptContent.categoriesIds -and $ScriptContent.categoriesIds.Count -gt 1)
            }
        } catch {
            [PSCustomObject]@{
                Name     = $Script.name
                FileName = $ScriptFileName
                Failed   = $true
                MultiCat = $false
            }
        }
    } -ThrottleLimit 50 | ForEach-Object {
        $result = $_
        if (-not $result.Failed) {
            $ProgressCounter.AddOrUpdate('completed', 1, { param($k, $v) $v + 1 }) | Out-Null
        } else {
            $ProgressCounter.AddOrUpdate('failed', 1, { param($k, $v) $v + 1 }) | Out-Null
        }
        $total = $ProgressCounter['completed'] + $ProgressCounter['failed']
        $percent = [math]::Round(($total / $TotalScripts) * 100)
        Write-Progress -Activity "Downloading Scripts" `
            -Status "$total/$TotalScripts | $percent% Complete" `
            -PercentComplete $percent
        $result  # Output result to be collected into $ScriptArray
    }
    
    Write-Progress -Activity "Downloading Scripts" -Completed
        
} else {
    Write-Host "Using sequential processing to download scripts." -ForegroundColor Yellow
    
    $ScriptArray = @()
    $ProcessedScript = @{}
    $CurrentScript = 0
        
    foreach ($Script in $Scripts) {
        switch ($Script.language) {
            'powershell' { $FileExtension = '.ps1' }
            'batchfile'  { $FileExtension = '.bat' }
            'vbscript'   { $FileExtension = '.vbs' }
            'sh'         { $FileExtension = '.sh' }
            default      { Write-Host "Unknown language: $($Script.language)" -ForegroundColor Yellow; continue }
        }
        
        $BaseFileName = (($Script.name -replace '[\\/:*?"<>&|]', '_').TrimStart())
        $ScriptFileName = $BaseFileName + $FileExtension
        
        if ($ProcessedScript.ContainsKey($BaseFileName)) {
            $ProcessedScript[$BaseFileName]++
            $ScriptFileName = "$BaseFileName-copy$($ProcessedScript[$BaseFileName])$FileExtension"
        } else {
            $ProcessedScript[$BaseFileName] = 0
        }
        
        $CurrentScript++
        $PercentComplete = [math]::Round(($CurrentScript / $TotalScripts) * 100)
        Write-Progress -Activity "Downloading Scripts" `
            -Status "$CurrentScript/$TotalScripts | $PercentComplete% Complete" `
            -PercentComplete $PercentComplete
        
        try {
            $ScriptContent = Invoke-RestMethod -Uri "$ScriptBaseURL/$($Script.id)" `
                -Method Get `
                -WebSession $session `
                -Headers @{
                    "Accept" = "application/json"
                } `
                -ErrorAction Stop
            
            $OriginalCategoryNames = ((Get-ScriptCategory -CategoryIDs $ScriptContent.categoriesIds -CategoriesHash $CategoriesHash) -replace '[\\/:*?"<>|]', '_')
            
            if ($ScriptContent.categoriesIds.Count -gt 1) {
                $CategoryName = "- Duplicate-Categories"
            } else {
                $CategoryName = $OriginalCategoryNames
            }
            
            $ScriptCode = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptContent.code))
            
            $ScriptArray += [PSCustomObject]@{
                Name               = $Script.name
                Language           = $Script.language
                Category           = $CategoryName
                OriginalCategories = $OriginalCategoryNames
                Description        = $Script.description
                FileName           = $ScriptFileName
                FilePath           = Join-Path "$ScriptFolder\$CategoryName" $ScriptFileName
                Code               = $ScriptCode
                Failed             = $false
                MultiCat           = ($ScriptContent.categoriesIds.Count -gt 1)
            }
        } catch {
            $ScriptArray += [PSCustomObject]@{
                Name     = $Script.name
                FileName = $ScriptFileName
                Failed   = $true
                MultiCat = $false
            }
        }
    }
    
    # Dismiss progress bar immediately after loop
    Write-Progress -Activity "Downloading Scripts" -Completed
    
    # For sequential mode, we need to write files since code is still in memory
    Write-Host "Script processing complete! Writing to disk..." -ForegroundColor Green

    Write-Host "`nWriting scripts to disk..." -ForegroundColor Cyan
    $WriteErrors = @()

    foreach ($Script in $ScriptArray) {
        try {
            # Ensure directory exists
            $directory = Split-Path -Path $Script.FilePath -Parent
            if (-not (Test-Path $directory)) {
                New-Item -Path $directory -ItemType Directory -Force | Out-Null
            }

            $Script.Code | Set-Content -Path $Script.FilePath -NoNewline -ErrorAction Stop
        } catch {
            try {
                [System.IO.File]::WriteAllText($Script.FilePath, $Script.Code)
            } catch {
                $WriteErrors += [PSCustomObject]@{
                    Name = $Script.Name
                    FilePath = $Script.FilePath
                    Error = $_.Exception.Message
                }
            }
        }
    }

    # Show write errors if any
    if ($WriteErrors.Count -gt 0) {
        Write-Host "`nFile write errors (first 5):" -ForegroundColor Red
        $WriteErrors | Select-Object -First 5 | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Error)" -ForegroundColor Yellow
        }
    }
}

# This runs regardless of which path was taken
$SavedScripts = $ScriptArray | Where-Object { -not $_.Failed } | Select-Object Name, Language, Category, Description, FilePath
$FailedScripts = $ScriptArray | Where-Object { $_.Failed }
$MultipleScripts = $ScriptArray | Where-Object { $_.MultiCat -and -not $_.Failed }

# Display results
if ($MultipleScripts) {
    Write-Warning "The following scripts with multiple categories were saved to '- Duplicate-Categories' folder:"
    $MultipleScripts | Select-Object Name, @{N='Categories';E={$_.OriginalCategories}} | 
        Sort-Object Categories, Name | Format-Table -AutoSize
}

if ($FailedScripts.Count -gt 0) {
    Write-Host "`nFailed scripts:" -ForegroundColor Red
    $FailedScripts | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Yellow }
}

$ScriptArray | Where-Object { -not $_.Failed } | Select-Object Name, Language, Category, Description, FilePath | 
    Sort-Object Category | 
    Export-Csv -Path "$ScriptFolder\Scripts.csv" -NoTypeInformation

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Download Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Total scripts: $TotalScripts" -ForegroundColor Cyan
Write-Host "Successfully saved: $($SavedScripts.Count)" -ForegroundColor Green
Write-Host "Failed: $($FailedScripts.Count)" -ForegroundColor $(if ($FailedScripts.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host "Multiple categories (saved to '- Duplicate-Categories'): $($MultipleScripts.Count)" -ForegroundColor Yellow
Write-Host "`nScripts saved to: $ScriptFolder" -ForegroundColor Cyan
Write-Host "Script details exported to: $ScriptFolder\Scripts.csv" -ForegroundColor Cyan
Write-Host "Session cached at: $env:USERPROFILE\.ninja_session.xml" -ForegroundColor Cyan