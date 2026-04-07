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
    [string]$ScriptFolder,
    [switch]$ClearCache
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
    
    .PARAMETER ScreenWidth
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
                    
                    # Create WebSession and add the cookie
                    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
                    $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"
                    $session.Cookies.Add((New-Object System.Net.Cookie("sessionKey", $SessionCookie.Value, "/", "app.ninjarmm.com")))
                    
                    # Cache the session with DPAPI encryption
                    $SessionData = @{
                        SessionKey = $SessionCookie.Value | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
                        UserAgent = $session.UserAgent
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
    # Alternative syntax options:
    # $session = Get-NinjaSession -ForceClear:$ClearCache           # Colon splatting (concise)
    # if ($ClearCache) { $session = Get-NinjaSession -ForceClear }   # Conditional (verbose)
    # $session = Get-NinjaSession @(@{ForceClear=$ClearCache})      # Hashtable splatting
    
    $session = Get-NinjaSession -ForceClear:$ClearCache
} catch {
    Write-Host "`nERROR: Authentication failed - $($_.Exception.Message)" -ForegroundColor Red
    exit 1
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

$global:FailedScripts = @()
$global:MultipleScripts = @()
$global:ProcessedScript = @{}
$global:ScriptArray = @()
$global:SavedScripts = @()

Write-Host "Processing $TotalScripts scripts..." -ForegroundColor Cyan

# Helper Functions
function Get-ScriptCategory {
    param([int[]]$CategoryIDs)
    
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
    param(
        [string]$Path,
        [switch]$Create
    )
    
    if ($Create.IsPresent) {
        if (-not (Test-Path -Path $Path)) {
            try {
                New-Item -Path $Path -ItemType "Directory" | Out-Null
            } catch {
                Write-Host "ERROR: Failed to create directory. $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

# Download scripts in batches for better performance
Write-Host "Downloading $TotalScripts scripts in batches of 50 (this will be much faster!)..." -ForegroundColor Cyan

# Reset global variables
$global:ScriptArray = @()
$global:FailedScripts = @()
$global:MultipleScripts = @()
$global:ProcessedScript = @{}
$global:SavedScripts = @()

# Process scripts in batches of 50 for better performance
$BatchSize = 50
$Batches = [Math]::Ceiling($TotalScripts / $BatchSize)
$CurrentScript = 0

for ($BatchIndex = 0; $BatchIndex -lt $Batches; $BatchIndex++) {
    $StartIndex = $BatchIndex * $BatchSize
    $EndIndex = [Math]::Min(($BatchIndex + 1) * $BatchSize - 1, $TotalScripts - 1)
    $BatchScripts = $Scripts[$StartIndex..$EndIndex]
    
    Write-Host "Processing batch $($BatchIndex + 1)/$Batches ($($BatchScripts.Count) scripts)..." -ForegroundColor Cyan
    
    # Process current batch
    foreach ($Script in $BatchScripts) {
        switch ($Script.language) {
            'powershell' { $FileExtension = '.ps1' }
            'batchfile'  { $FileExtension = '.bat' }
            'vbscript'   { $FileExtension = '.vbs' }
            'sh'         { $FileExtension = '.sh' }
            default      { Write-Host "Unknown language: $($Script.language)" -ForegroundColor Yellow; continue }
        }
        
        $BaseFileName = (($Script.name -replace '[\\/:*?"<>&|]', '_').TrimStart())
        $ScriptFileName = $BaseFileName + $FileExtension
        
        if ($global:ProcessedScript.ContainsKey($BaseFileName)) {
            $global:ProcessedScript[$BaseFileName]++
            $ScriptFileName = "$BaseFileName-copy$($global:ProcessedScript[$BaseFileName])$FileExtension"
        } else {
            $global:ProcessedScript[$BaseFileName] = 0
        }
        
        $CurrentScript++
        $PercentComplete = [math]::Round(($CurrentScript / $TotalScripts) * 100)
        Write-Progress -Activity "Downloading Scripts" `
            -Status "$CurrentScript/$TotalScripts | $PercentComplete% Complete | $ScriptFileName" `
            -PercentComplete $PercentComplete
        
        try {
            $ScriptContent = Invoke-RestMethod -Uri "$ScriptBaseURL/$($Script.id)" `
                -Method Get `
                -WebSession $session `
                -Headers @{
                    "Accept" = "application/json"
                } `
                -ErrorAction Stop
            
            # Get original category names for warning message
            $OriginalCategoryNames = ((Get-ScriptCategory -CategoryIDs $ScriptContent.categoriesIds) -replace '[\\/:*?"<>|]', '_')
            
            # Handle scripts with multiple categories by using '- Duplicates' folder
            if ($ScriptContent.categoriesIds.Count -gt 1) {
                $CategoryName = "- Duplicates"
                $global:MultipleScripts += [PSCustomObject]@{
                    ScriptName    = $Script.name
                    CategoryNames = $OriginalCategoryNames
                }
            } else {
                $CategoryName = $OriginalCategoryNames
            }
            
            Set-Directory -Path "$ScriptFolder\$CategoryName" -Create
            
            $ScriptCode = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptContent.code))
            
            $global:ScriptArray += [PSCustomObject]@{
                Name = $Script.name
                Language = $Script.language
                Category = $CategoryName
                Description = $Script.description
                FileName = $ScriptFileName
                FilePath = Join-Path "$ScriptFolder\$CategoryName" $ScriptFileName
                Code = $ScriptCode
            }
        } catch {
            Write-Host "Failed to download: $($Script.name)" -ForegroundColor Red
            $global:FailedScripts += $Script.name
        }
    }
    
    $CompletedScripts = ($BatchIndex + 1) * $BatchSize
    $CompletedScripts = [Math]::Min($CompletedScripts, $TotalScripts)
    $PercentComplete = [math]::Round(($CompletedScripts / $TotalScripts) * 100)
    Write-Host "Batch $($BatchIndex + 1) complete. Overall progress: $CompletedScripts/$TotalScripts ($PercentComplete%)" -ForegroundColor Green
}

Write-Progress -Activity "Downloading Scripts" -Completed
Write-Host "All batches complete! Processing results..." -ForegroundColor Green

# Write scripts to files
Write-Host "`nWriting scripts to disk..." -ForegroundColor Cyan
foreach ($Script in $global:ScriptArray) {
    $Failed = $false
    try {
        $Script.Code | Set-Content -Path $Script.FilePath -NoNewline
    } catch {
        try {
            [System.IO.File]::WriteAllText($Script.FilePath, $Script.Code)
        } catch {
            $global:FailedScripts += $Script.FileName
            $Failed = $true
        }
    }
    if (-not $Failed) {
        $global:SavedScripts += [PSCustomObject]@{
            Name = $Script.Name
            Language = $Script.Language
            Category = $Script.Category
            Description = $Script.Description
            FilePath = $Script.FilePath
        }
    }
}

# Display results
if ($global:MultipleScripts) {
    Write-Warning "The following scripts with multiple categories were saved to '- Duplicates' folder:"
    $global:MultipleScripts | Format-Table -AutoSize
}

if ($global:FailedScripts.Count -gt 0) {
    Write-Host "`nFailed scripts:" -ForegroundColor Red
    $global:FailedScripts | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

$global:SavedScripts | Select-Object Name, Language, Category, Description, FilePath | 
    Sort-Object Category | 
    Export-Csv -Path "$ScriptFolder\Scripts.csv" -NoTypeInformation

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Download Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Total scripts: $TotalScripts" -ForegroundColor Cyan
Write-Host "Successfully saved: $($global:SavedScripts.Count)" -ForegroundColor Green
Write-Host "Failed: $($global:FailedScripts.Count)" -ForegroundColor $(if ($global:FailedScripts.Count -gt 0) { 'Red' } else { 'Green' })
Write-Host "Multiple categories (saved to '- Duplicates'): $($global:MultipleScripts.Count)" -ForegroundColor Yellow
Write-Host "`nScripts saved to: $ScriptFolder" -ForegroundColor Cyan
Write-Host "Script details exported to: $ScriptFolder\Scripts.csv" -ForegroundColor Cyan
Write-Host "Session cached at: $env:USERPROFILE\.ninja_session.xml" -ForegroundColor Cyan