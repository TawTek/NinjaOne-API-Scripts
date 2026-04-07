BeforeAll {
    # Import script under test
    $script:ScriptPath = "$PSScriptRoot\..\Get-NinjaScripts.ps1"
    $script:ScriptContent = Get-Content $script:ScriptPath -Raw
    
    # Extract specific sections for focused testing
    $script:ParamBlock = ($script:ScriptContent -split "param\(")[1] -split "\)[^{]*{" | Select-Object -First 1
    $script:MainLogic = ($script:ScriptContent -split "if.*PSVersionTable.*Major.*ge.*7.*and.*not.*Sequential")[1] -split "} else {" | Select-Object -First 1
    $script:SequentialLogic = ($script:ScriptContent -split "} else {")[1] -split "Write-Host.*Processing.*scripts" | Select-Object -First 1
    $script:ProgressLogic = $script:ScriptContent | Select-String -Pattern "Write-Progress" -Context 2,2
    $script:ErrorHandling = $script:ScriptContent | Select-String -Pattern "try.*{" -Context 0,5
    
    # Mock session for testing without browser automation
    $script:MockSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $script:MockSession.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    $mockCookie = New-Object System.Net.Cookie("sessionKey", "mock-session-key-12345", "/", "app.ninjarmm.com")
    $script:MockSession.Cookies.Add($mockCookie)
    
    # Test data directory
    $script:TestDataDir = "$PSScriptRoot\.test-data"
    if (-not (Test-Path $script:TestDataDir)) {
        New-Item -Path $script:TestDataDir -ItemType Directory -Force | Out-Null
    }
    
    # Cache directory
    $script:CacheDir = "$PSScriptRoot\.tests-cache"
    if (-not (Test-Path $script:CacheDir)) {
        New-Item -Path $script:CacheDir -ItemType Directory -Force | Out-Null
    }
    
    # Cache TTL
    $script:CacheTTLHours = 24
    
    # Helper function to get cached results
    function Get-CachedResult {
        param(
            [string]$CacheKey
        )
        
        $cachePath = Join-Path $script:CacheDir "$CacheKey.json"
        
        if (Test-Path $cachePath) {
            $cache = Get-Content $cachePath -Raw | ConvertFrom-Json
            $cacheAge = (Get-Date) - [datetime]$cache.Timestamp
            
            if ($cacheAge.TotalHours -lt $script:CacheTTLHours) {
                return $cache.Results
            }
        }
        
        return $null
    }
    
    # Helper function to set cached results
    function Set-CachedResult {
        param(
            [string]$CacheKey,
            [object]$Results
        )
        
        $cachePath = Join-Path $script:CacheDir "$CacheKey.json"
        $cacheData = @{
            Timestamp = (Get-Date).ToString('o')
            Results = $Results
        }
        
        $cacheData | ConvertTo-Json -Depth 10 | Set-Content $cachePath
    }
}

Describe 'Get-NinjaScripts.ps1 - Code Structure' {
    
    Context 'Script metadata' {
        It 'has proper synopsis' {
            $script:ScriptContent | Should -Match '\.SYNOPSIS'
            $script:ScriptContent | Should -Match 'Downloads all scripts from Ninja'
        }
        
        It 'has proper description' {
            $script:ScriptContent | Should -Match '\.DESCRIPTION'
            $script:ScriptContent | Should -Match 'Selenium browser automation'
        }
        
        It 'documents SSO/MFA support' {
            $script:ScriptContent | Should -Match 'SSO/MFA'
        }
        
        It 'has author and date' {
            $script:ScriptContent | Should -Match 'Author:'
            $script:ScriptContent | Should -Match 'Date:'
        }
    }
    
    Context 'Parameters' {
        It 'has ScriptFolder parameter' {
            $script:ScriptContent | Should -Match '\[string\]\$ScriptFolder'
        }
        
        It 'has ClearCache switch parameter' {
            $script:ScriptContent | Should -Match '\[switch\]\$ClearCache'
        }
        
        It 'has default ScriptFolder value' {
            $script:ScriptContent | Should -Match 'ScriptFolder\s*=\s*"[^"]+"'
        }
    }
    
    Context 'Security implementation' {
        It 'uses DPAPI encryption for session cache' {
            $script:ScriptContent | Should -Match 'ConvertTo-SecureString'
            $script:ScriptContent | Should -Match 'ConvertFrom-SecureString'
        }
        
        It 'uses ZeroFreeBSTR for memory cleanup' {
            $script:ScriptContent | Should -Match 'ZeroFreeBSTR'
        }
        
        It 'has try/finally for SecureString cleanup' {
            $script:ScriptContent | Should -Match 'try\s*\{[^}]*PtrToStringAuto[^}]*\}\s*finally\s*\{[^}]*ZeroFreeBSTR'
        }
        
        It 'validates session before use' {
            $script:ScriptContent | Should -Match 'Test if session is still valid'
            $script:ScriptContent | Should -Match 'Invoke-WebRequest'
            $script:ScriptContent | Should -Match '-WebSession'
        }
    }
    
    Context 'Session management' {
        It 'has session cache path variable' {
            $script:ScriptContent | Should -Match '\$SessionCachePath\s*=.*\.ninja_session\.xml'
        }
        
        It 'has cache expiry configuration' {
            $script:ScriptContent | Should -Match '\$CacheExpiryHours\s*=\s*2'
        }
        
        It 'checks cached session before authentication' {
            $script:ScriptContent | Should -Match 'if \(Test-Path \$SessionCachePath\)'
        }
        
        It 'supports forced cache clearing' {
            $script:ScriptContent | Should -Match '\$ForceClear'
            $script:ScriptContent | Should -Match 'Remove-Item.*SessionCachePath'
        }
    }
    
    Context 'Selenium integration' {
        It 'checks for Selenium module' {
            $script:ScriptContent | Should -Match 'Get-Module.*Selenium'
        }
        
        It 'uses EdgeDriverService with explicit path' {
            $script:ScriptContent | Should -Match 'EdgeDriverService.*CreateDefaultService'
            $script:ScriptContent | Should -Match 'msedgedriver\.exe'
        }
        
        It 'creates EdgeDriver instance' {
            $script:ScriptContent | Should -Match 'New-Object.*EdgeDriver'
        }
        
        It 'navigates to NinjaRMM URL' {
            $script:ScriptContent | Should -Match 'Navigate.*GoToUrl.*app\.ninjarmm\.com'
        }
        
        It 'captures sessionKey cookie' {
            $script:ScriptContent | Should -Match 'sessionKey'
        }
        
        It 'has authentication timeout' {
            $script:ScriptContent | Should -Match '\$TimeoutSeconds\s*=\s*300'
        }
        
        It 'detects primary screen resolution dynamically' {
            $script:ScriptContent | Should -Match 'System.Windows.Forms.Screen'
            $script:ScriptContent | Should -Match 'PrimaryScreen'
        }
        
        It 'uses off-screen positioning for browser startup' {
            $script:ScriptContent | Should -Match '--window-position=-32000,-32000'
        }
        
        It 'hides EdgeDriver command prompt window' {
            $script:ScriptContent | Should -Match 'HideCommandPromptWindow.*true'
        }
    }
    
    Context 'Script processing' {
        It 'filters excluded languages' {
            $script:ScriptContent | Should -Match 'native.*binary_install.*file_transfer'
        }
        
        It 'has progress bar' {
            $script:ScriptContent | Should -Match 'Write-Progress'
        }
        
        It 'handles duplicate filenames' {
            $script:ScriptContent | Should -Match 'ProcessedScript.*ContainsKey'
            $script:ScriptContent | Should -Match '-copy'
        }
        
        It 'handles duplicate categories' {
            $script:ScriptContent | Should -Match '- Duplicate-Categories'
        }
        
        It 'decodes base64 script content' {
            $script:ScriptContent | Should -Match 'FromBase64String'
            $script:ScriptContent | Should -Match 'UTF8\.GetString'
        }
        
        It 'detects PowerShell version for processing mode' {
            $script:ScriptContent | Should -Match 'PSVersionTable.PSVersion.Major'
        }
    }
    
    Context 'Error handling' {
        It 'has try/catch for authentication' {
            $script:ScriptContent | Should -Match 'try\s*\{.*Get-NinjaSession.*\}\s*catch'
        }
        
        It 'has try/catch for script fetching' {
            $script:ScriptContent | Should -Match 'try\s*\{.*Invoke-RestMethod.*\}\s*catch'
        }
        
        It 'tracks failed scripts in ScriptArray' {
            $script:ScriptContent | Should -Match 'Failed.*=.*\$true'
        }
        
        It 'uses ErrorAction Stop for web requests' {
            $script:ScriptContent | Should -Match '-ErrorAction Stop'
        }
    }
    
    Context 'Output and reporting' {
        It 'exports to CSV' {
            $script:ScriptContent | Should -Match 'Export-Csv'
            $script:ScriptContent | Should -Match 'Scripts\.csv'
        }
        
        It 'reports multiple category scripts' {
            $script:ScriptContent | Should -Match 'multiple categories.*saved to.*Duplicate-Categories'
        }
        
        It 'reports failed scripts' {
            $script:ScriptContent | Should -Match 'Failed scripts'
        }
        
        It 'shows summary statistics' {
            $script:ScriptContent | Should -Match 'Total scripts'
            $script:ScriptContent | Should -Match 'Successfully saved'
        }
        
        It 'sorts duplicate scripts by category then name' {
            $script:ScriptContent | Should -Match 'Sort-Object Categories, Name'
        }
    }
}


Describe 'Get-NinjaScripts.ps1 - Security Tests' {
    
    Context 'Session cache encryption' {
        BeforeAll {
            $script:TestCachePath = Join-Path $script:TestDataDir "test_session.xml"
        }
        
        AfterAll {
            if (Test-Path $script:TestCachePath) {
                Remove-Item $script:TestCachePath -Force
            }
        }
        
        It 'encrypts session key with DPAPI' {
            $testSessionKey = "test-session-key-12345"
            $encrypted = $testSessionKey | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
            
            $sessionData = @{
                SessionKey = $encrypted
                UserAgent = "Test Agent"
                ExpiresAt = (Get-Date).AddHours(2)
            }
            
            $sessionData | Export-Clixml -Path $script:TestCachePath
            
            $fileContent = Get-Content $script:TestCachePath -Raw
            $fileContent | Should -Not -Match $testSessionKey
            $fileContent | Should -Match '01000000d08c9ddf'
        }
        
        It 'can decrypt encrypted session key' {
            $testSessionKey = "test-session-key-67890"
            $encrypted = $testSessionKey | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
            
            $sessionData = @{
                SessionKey = $encrypted
                UserAgent = "Test Agent"
                ExpiresAt = (Get-Date).AddHours(2)
            }
            
            $sessionData | Export-Clixml -Path $script:TestCachePath
            
            $loaded = Import-Clixml -Path $script:TestCachePath
            $secureKey = $loaded.SessionKey | ConvertTo-SecureString
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
            try {
                $decrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                $decrypted | Should -Be $testSessionKey
            } finally {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
    }
    
    Context 'Session validation' {
        It 'checks session expiry before use' {
            $script:ScriptContent | Should -Match 'if \(\$CachedSession\.ExpiresAt -gt \(Get-Date\)\)'
        }
        
        It 'tests session validity with API call' {
            $script:ScriptContent | Should -Match 'Invoke-WebRequest'
            $script:ScriptContent | Should -Match 'scripting/scripts'
            $script:ScriptContent | Should -Match '-WebSession'
            $script:ScriptContent | Should -Match '-ErrorAction Stop'
        }
        
        It 're-authenticates on expired session' {
            $script:ScriptContent | Should -Match 'Cached session expired, re-authenticating'
        }
    }
}

Describe 'Get-NinjaScripts.ps1 - Integration Patterns' {
    
    Context 'Language filtering' {
        It 'excludes native language' {
            $script:ScriptContent | Should -Match 'language -notin.*native'
        }
        
        It 'excludes binary_install language' {
            $script:ScriptContent | Should -Match 'language -notin.*binary_install'
        }
        
        It 'excludes file_transfer language' {
            $script:ScriptContent | Should -Match 'language -notin.*file_transfer'
        }
    }
    
    Context 'File extension mapping' {
        It 'maps powershell to .ps1' {
            $script:ScriptContent | Should -Match "'powershell'.*\.ps1"
        }
        
        It 'maps batchfile to .bat' {
            $script:ScriptContent | Should -Match "'batchfile'.*\.bat"
        }
        
        It 'maps vbscript to .vbs' {
            $script:ScriptContent | Should -Match "'vbscript'.*\.vbs"
        }
        
        It 'maps sh to .sh' {
            $script:ScriptContent | Should -Match "'sh'.*\.sh"
        }
    }
    
    Context 'Filename sanitization' {
        It 'removes invalid characters from filenames' {
            $script:ScriptContent | Should -Match '-replace.*\[\\\\/:.*?\].*_'
        }
        
        It 'trims whitespace from filenames' {
            $script:ScriptContent | Should -Match 'TrimStart'
        }
    }
    
    Context 'Parallel processing' {
        It 'uses ForEach-Object -Parallel for PowerShell 7+' {
            $script:ScriptContent | Should -Match 'ForEach-Object -Parallel'
        }
        
        It 'uses ThrottleLimit for parallel processing' {
            $script:ScriptContent | Should -Match 'ThrottleLimit'
        }
        
        It 'uses streaming ForEach-Object for progress updates' {
            $script:ScriptContent | Should -Match '\$result = \$_'
            $script:ScriptContent | Should -Match '\$ProgressCounter'
            $script:ScriptContent | Should -Match 'AddOrUpdate'
        }
        
        It 'uses thread-safe collections for parallel processing' {
            $script:ScriptContent | Should -Match 'ConcurrentDictionary'
        }
        
        It 'writes files immediately in parallel mode' {
            $script:ScriptContent | Should -Match 'Set-Content.*-Path.*FilePath'
            $script:ScriptContent | Should -Match 'WriteAllText'
        }
        
        It 'dismisses progress bar after completion' {
            $script:ScriptContent | Should -Match 'Write-Progress.*-Completed'
        }
        
        It 'suppresses progress counter return values' {
            $script:ScriptContent | Should -Match 'AddOrUpdate.*Out-Null'
        }
        
        It 'has sequential fallback for PowerShell 5.1' {
            $script:ScriptContent | Should -Match 'Using sequential processing to download scripts'
            $script:ScriptContent | Should -Match 'PSVersionTable\.PSVersion\.Major'
        }
    }
}

Describe 'Get-NinjaScripts.ps1 - Best Practices' {
    
    Context 'Code quality' {
        It 'uses proper error handling' {
            $script:ScriptContent | Should -Match 'try\s*\{[^}]*\}\s*catch'
        }
        
        It 'uses Write-Host for user feedback' {
            $script:ScriptContent | Should -Match 'Write-Host'
        }
        
        It 'uses Write-Progress for long operations' {
            $script:ScriptContent | Should -Match 'Write-Progress'
        }
        
        It 'uses Write-Warning for non-fatal issues' {
            $script:ScriptContent | Should -Match 'Write-Warning'
        }
        
        It 'cleans up resources in finally blocks' {
            $script:ScriptContent | Should -Match 'finally\s*\{[^}]*Quit'
        }
    }
    
    Context 'PowerShell conventions' {
        It 'uses approved verbs for functions' {
            $script:ScriptContent | Should -Match 'function Get-'
            $script:ScriptContent | Should -Match 'function Set-'
        }
        
        It 'uses param blocks for functions' {
            $script:ScriptContent | Should -Match 'function.*\{\s*param\s*\('
        }
        
        It 'uses CmdletBinding where appropriate' {
            $script:ScriptContent | Should -Match '\[CmdletBinding\(\)\]'
        }
    }
}