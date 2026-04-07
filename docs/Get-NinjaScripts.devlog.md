<!-- 
  AI Context Header
  This file documents the development journey, architecture decisions, debugging process, 
  and key invariants for Get-NinjaScripts.ps1. It serves as both historical record and 
  onboarding material for future maintainers.
  
  When modifying this script, update the relevant sections:
  - Add new bugs/fixes to "Debugging Journey"
  - Document rejected approaches in "Rejected Approaches"
  - Update invariants if assumptions change
  - Add changelog entries for significant changes
-->

# Get-NinjaScripts.ps1 — Dev Log

## Overview

A PowerShell script that downloads all scripts from NinjaRMM's Automation Library using browser-based SSO/MFA authentication via Selenium WebDriver. Built to solve the problem of accessing script content from the legacy scripting API endpoint that doesn't support OAuth tokens, only session cookies.

Supports PowerShell 5.1+ with secure session caching, DPAPI encryption, duplicate category handling, and batch processing for performance.

---

## How to Use This Script

**Parameters:**
```powershell
-ScriptFolder "C:\Temp\NinjaScripts"  # optional (defaults to C:\Temp\NinjaScripts\Selenium)
-ClearCache                            # optional — forces re-authentication
```

**Typical workflow:**
1. Run script — browser opens for SSO/MFA login
2. Complete authentication in browser
3. Session cached for 2 hours (DPAPI-encrypted)
4. Scripts downloaded in batches of 50
5. Results exported to CSV with metadata

**Session management:**
```powershell
# First run — opens browser
.\Get-NinjaScripts.ps1

# Subsequent runs within 2 hours — uses cached session
.\Get-NinjaScripts.ps1

# Force re-authentication
.\Get-NinjaScripts.ps1 -ClearCache

# Manual cache clear
Remove-Item "$env:USERPROFILE\.ninja_session.xml"
```

**Output:**
- Scripts organized by category in `$ScriptFolder`
- Duplicate category scripts in `- Duplicates` folder
- CSV inventory at `$ScriptFolder\Scripts.csv`

---

## Quick Reference

**Common Issues:**
- Browser doesn't open → Check EdgeDriver at `c:\git\msedgedriver.exe` (see [EdgeDriver path](#edgedriver-path-hardcoded))
- Session expired → Cache TTL is 2 hours, re-authenticate or use `-ClearCache` (see [Session cache expiry](#session-cache-expiry-time))
- Download failures → Session not passed to individual script requests (see [Session not passed to script downloads](#session-not-passed-to-individual-script-downloads))
- Plaintext session in cache → DPAPI encryption not applied (see [Session cache security](#session-cache-stored-as-plaintext))

**Key Functions:**
- `Get-NinjaSession` — Handles browser automation, SSO/MFA login, session capture and caching
- `Get-ScriptCategory` — Maps category IDs to friendly names
- `Set-Directory` — Creates category folders with error handling

---

## Architecture Decisions

### Why Selenium instead of OAuth

NinjaRMM provides an OAuth 2.0 API (`/api/v2/automation/scripts`) but it only returns script metadata, not the actual script content. The script content is only available via the legacy UI endpoint (`/swb/s21/scripting/scripts`) which requires session cookies from browser-based authentication, not OAuth bearer tokens.

Attempted OAuth `client_credentials` grant → 401 Unauthorized on legacy endpoint. The only way to access script content is through session cookies obtained via SSO/MFA login.

### Why not interactive browser popup like Connect-MgGraph

Microsoft Graph uses the OAuth Authorization Code Flow with PKCE, which is designed for interactive user consent. NinjaRMM's OAuth implementation only supports `client_credentials` (machine-to-machine), not authorization code flow. There's no official API endpoint for interactive user authentication that returns tokens usable for the scripting endpoint.

Selenium is the only reliable way to handle arbitrary SSO providers (Okta, Azure AD, Google, etc.) with MFA prompts.

### Session caching with DPAPI encryption

Session cookies are cached to avoid repeated SSO/MFA prompts during the 2-hour validity window. The session key is encrypted using Windows Data Protection API (DPAPI) via `ConvertTo-SecureString | ConvertFrom-SecureString`, making it user-bound and machine-bound. Only the Windows user account that created the cache can decrypt it.

Alternative approaches (plaintext, base64) would expose the session cookie to anyone with filesystem access. DPAPI provides enterprise-grade protection with zero configuration.

### PowerShell 7 parallel processing with streaming progress

PowerShell 7+ uses `ForEach-Object -Parallel` with thread-safe `ConcurrentDictionary` for duplicate filename tracking and progress counting. Files are written immediately inside parallel runspaces to optimize memory usage (no code stored in `$ScriptArray`).

**Critical invariant:** `$using:` scope only works inside the `-Parallel` scriptblock, NOT in the chained `ForEach-Object` on the main thread. The streaming `ForEach-Object` collects results directly into `$ScriptArray` via pipeline assignment without needing `$using:`.

**Progress counter pattern:**
- Progress updates happen ONLY in the streaming `ForEach-Object` (main thread)
- Each result passes through once, updating the counter once
- Avoids double-counting that occurs if counter is updated in both parallel block and streaming block

PowerShell 5.1 falls back to sequential processing with file writes in a separate loop after script metadata collection.

### Duplicate category handling

Scripts with multiple categories were originally skipped entirely. Changed to save once in a `- Duplicates` folder to prevent data loss. The warning table still shows original category names for reference, but the script is saved to a single predictable location.

This avoids creating multiple folders with similar names (e.g., "# App, ! Client" and "! Client, # App") and eliminates the need to decide which category takes precedence.

### EdgeDriver path hardcoded

The Selenium PowerShell module's `Start-SeEdge` helper function had inconsistent behavior finding the WebDriver. Direct instantiation of `EdgeDriverService` with an explicit path (`c:\git\msedgedriver.exe`) ensures the driver is always found.

This is a known limitation — the path should be parameterized for broader use, but for internal tooling the hardcoded path is acceptable.

---

## Key Invariants & Assumptions

These are the core architectural principles the script is built around. If any of these change, the corresponding logic needs revisiting.

| Invariant | Detail |
|---|---|
| Legacy scripting endpoint requires session cookies | OAuth tokens don't work on `/swb/s21/scripting/scripts` - only browser session cookies |
| Session cookie name is `sessionKey` | Selenium captures this specific cookie from `app.ninjarmm.com` domain |
| DPAPI encryption is user+machine bound | `ConvertFrom-SecureString` output can only be decrypted by the same user on the same machine |
| Session expires after unknown duration | NinjaRMM doesn't publish session TTL; 2-hour cache expiry is conservative |
| Script content is base64-encoded | All script code in API responses is base64; must decode with UTF-8 |
| `$using:` scope only works in `-Parallel` block | Chained `ForEach-Object` runs on main thread - use direct pipeline assignment to `$ScriptArray` |
| Progress counter updates ONLY in streaming block | Updating in both parallel and streaming blocks causes double-counting |
| Files written immediately in parallel mode | Optimizes memory - no script code stored in `$ScriptArray`, only metadata |
| Progress bar dismissal requires same scope | `Write-Progress -Completed` must be called from same scope that created the progress bar |
| `-Sequential` parameter forces sequential mode | Allows PowerShell 7+ users to opt for sequential processing when needed |
| `-ClearCache` parameter works with splatting syntax | `-ForceClear:$ClearCache` properly passes switch parameter to function |
| Parallel processing requires result output | `$result` must be output to pipeline for collection into `$ScriptArray` |
| Category folder names preserve original naming | Windows supports spaces in folder names, no character replacement needed |

---

## Implementation Details & Limitations

| Detail | Current Implementation |
|---|---|
| EdgeDriver path | Hardcoded to `C:\Git\tawtek\ninjaone-api-scripts\msedgedriver.exe` |
| Selenium EdgeOptions API | Uses fallback chain: `AddArguments()` then `AddAdditionalCapability()` |
| Off-screen positioning | Uses `--window-position=-32000,-32000` to start browser hidden |
| Screen detection | Dynamic detection via `System.Windows.Forms.Screen.PrimaryScreen` |
| Parallel ThrottleLimit | Set to 50 for balance of speed vs resource usage |
| Category sorting | Multi-level sort: `Sort-Object Categories, Name` for readability |

---

## Debugging Journey

### OAuth 401 Unauthorized on scripting endpoint

Initial attempt used OAuth `client_credentials` flow with `Authorization: Bearer $AccessToken` header.

**Error:**
```
Response status code does not indicate success: 401 (Unauthorized)
```

**Investigation:**
- OAuth token worked on `/api/v2/automation/scripts` (metadata only)
- Same token failed on `/swb/s21/scripting/scripts` (script content)
- Tested with different scopes (`monitoring`, `management`, `offline_access`) — no change

**Root cause:** The legacy scripting endpoint predates the OAuth API and was never updated to support bearer tokens. It only accepts session cookies from browser-based login.

**Fix:** Pivoted to Selenium browser automation to capture session cookies.

---

### Progress bar persistence in both processing modes

Progress bar remained visible after script completion in both sequential and parallel modes, requiring Ctrl+C to dismiss.

**Sequential Mode Issue:**
**Root cause:** `Write-Progress -Completed` only dismisses progress bar when called from the same scope that owns it. In sequential mode, the progress bar was being created inside the foreach loop but the dismissal call was outside the proper scope context.

**Parallel Mode Issue:**
**Root cause:** Progress updates were happening inside parallel runspaces (child scopes), not the main thread. The main thread never owned the progress bar, so `-Completed` on the main thread didn't dismiss it.

**Fix:** 
- **Sequential:** Added `Write-Progress -Activity "Downloading Scripts" -Completed` immediately after the foreach loop
- **Parallel:** Removed all `Write-Progress` calls from inside the parallel block and streamed results back to the main thread for progress updates

**Key Insight:** Progress bar ownership is scope-bound - the thread/scope that creates the progress bar must also be the one to dismiss it.

---

### Parallel processing showing zero script counts

Script output showed "Successfully saved: 0" and "Failed: 0" despite processing 1101 scripts in parallel mode.

**Root cause:** The parallel processing block was missing the `$result` output to the pipeline. Results were being generated inside the parallel block but not passed through to the streaming `ForEach-Object` for collection into `$ScriptArray`.

**Fix:** Added `$result` output at the end of the streaming `ForEach-Object` block to ensure results are collected into `$ScriptArray`.

---

### Selenium Start-SeEdge failing with "Edge not available"

Used `Start-SeEdge` from Selenium PowerShell module.

**Error:**
```
Edge not available
```

**Investigation:**
- Microsoft Edge installed and working
- `msedgedriver.exe` present at `c:\git\msedgedriver.exe`
- `Start-SeEdge` looking for `MicrosoftWebDriver.exe` (obsolete driver name)

**Attempted:**
- Adding driver path to `$env:PATH` — no effect
- `-WebDriverDirectory` parameter — parameter doesn't exist on `Start-SeEdge`

**Fix:** Bypass `Start-SeEdge` entirely and directly instantiate `EdgeDriverService`:
```powershell
$driverPath = "c:\git"
$service = [OpenQA.Selenium.Edge.EdgeDriverService]::CreateDefaultService($driverPath, "msedgedriver.exe")
$Driver = New-Object OpenQA.Selenium.Edge.EdgeDriver($service, $options)
```

---

### Session not passed to individual script downloads

Initial script list retrieved successfully, but individual script content downloads failed with generic "Failed to download" errors.

**Investigation:**
```powershell
# Initial list request — works
$ScriptsResponse = Invoke-RestMethod -Uri ".../scripts" -WebSession $session

# Individual script request — fails
$ScriptContent = Invoke-RestMethod -Uri ".../scripts/$id" -Headers @{
    "Authorization" = "Bearer $AccessToken"  # ← Wrong auth method
}
```

**Root cause:** Individual script downloads were using OAuth bearer tokens instead of the session cookie.

**Fix:** Changed to use `-WebSession $session` for all script content requests:
```powershell
$ScriptContent = Invoke-RestMethod -Uri "$ScriptBaseURL/$($Script.id)" `
    -Method Get `
    -WebSession $session `
    -Headers @{ "Accept" = "application/json" }
```

---

### Session cache stored as plaintext

Session key was visible in plaintext when viewing `~\.ninja_session.xml`.

**Investigation:**
```powershell
$SessionData = @{
    SessionKey = $SessionCookie.Value  # Plain string
    UserAgent = $session.UserAgent
    ExpiresAt = (Get-Date).AddHours(8)
}
$SessionData | Export-Clixml -Path $SessionCachePath
```

**Verification:**
```powershell
Get-Content ~\.ninja_session.xml | Select-String 'SessionKey'
# <S N="Value">5f1f3634-49ed-48e1-b861-afbf0dfb63a4</S>  ← Plaintext!
```

**Root cause:** `Export-Clixml` only encrypts `SecureString` types with DPAPI. Plain strings are serialized as `<S>` elements (plaintext).

**Fix:** Convert to SecureString before export:
```powershell
$SessionData = @{
    SessionKey = $SessionCookie.Value | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    UserAgent = $session.UserAgent
    ExpiresAt = (Get-Date).AddHours(8)
}
```

Now serializes as encrypted DPAPI string:
```xml
<S N="Value">01000000d08c9ddf0115d1118c7a00c04fc297eb...</S>
```

---

### SecureString decryption memory leak

Decryption code allocated unmanaged memory that was never freed.

**Code:**
```powershell
$DecryptedSessionKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureKey)
)
# ← Unmanaged memory never freed!
```

**Root cause:** `SecureStringToBSTR` allocates unmanaged memory that must be explicitly freed with `ZeroFreeBSTR`. Without cleanup, session keys linger in memory.

**Fix:** Proper try/finally cleanup:
```powershell
$BStr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureKey)
try {
    $DecryptedSessionKey = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BStr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BStr)  # Zeroes memory before freeing
}
```

---

### Stale comment referencing wrong cache file

Final summary message referenced the old OAuth token cache path.

**Code:**
```powershell
Write-Host "Token cached at: $env:USERPROFILE\.ninja_oauth_token.xml"
```

**Actual cache path:** `~\.ninja_session.xml`

**Impact:** Users trying to manually clear cache would delete the wrong file and be confused.

**Fix:** Updated message to match actual cache path:
```powershell
Write-Host "Session cached at: $env:USERPROFILE\.ninja_session.xml"
```

---

### Parallel processing $using: expression errors

Attempted to use `ForEach-Object -Parallel` with thread-safe collections.

**Errors:**
```
Expression is not allowed in a Using expression.
The assignment expression is not valid.
```

**Code:**
```powershell
$Added = $using:ProcessedScript.TryAdd($BaseFileName, 0)  # ← Error
$using:ProcessedScript[$BaseFileName] = $CurrentCount + 1  # ← Error
$using:ScriptArray.Add([PSCustomObject]@{ ... })  # ← Error
```

**Root cause:** PowerShell's `$using:` scope modifier doesn't support:
- Method calls on collections (`.TryAdd()`, `.Add()`)
- Index assignment (`[$key] = $value`)
- Complex expressions

**Attempted:** Simplifying expressions, using intermediate variables — still failed.

**Fix:** Abandoned parallel processing in favor of batch processing (50 scripts per batch) which provides good performance without threading complexity.

---

### Duplicate category scripts skipped entirely

Scripts with multiple categories were added to `$global:MultipleScripts` array and then `continue` skipped them.

**Impact:** Scripts with multiple categories were never saved to disk, causing data loss.

**Investigation:**
```powershell
if ($ScriptContent.categoriesIds.Count -gt 1) {
    $global:MultipleScripts += [PSCustomObject]@{ ... }
    continue  # ← Script never saved!
}
```

**Fix:** Changed to save scripts with multiple categories to `- Duplicates` folder:
```powershell
if ($ScriptContent.categoriesIds.Count -gt 1) {
    $CategoryName = "- Duplicates"
    $global:MultipleScripts += [PSCustomObject]@{
        ScriptName = $Script.name
        CategoryNames = $OriginalCategoryNames  # For warning table
    }
} else {
    $CategoryName = $OriginalCategoryNames
}
# Script continues to be saved
```

Updated warning message from "were not saved" to "were saved to '- Duplicates' folder".

---

### Session cache expiry time

Initial cache expiry was set to 8 hours.

**Concern:** Long expiry window increases exposure time if session is compromised.

**Fix:** Reduced to 2 hours as a balance between:
- **Security:** 75% reduction in exposure window
- **Convenience:** Still reasonable for daily work sessions

---

### file_transfer language not excluded

Scripts with `language = 'file_transfer'` were being processed.

**Request:** Add to exclusion filter alongside `native` and `binary_install`.

**Fix:**
```powershell
# Before
$Scripts = $ScriptsResponse | Where-Object { $_.language -notin @("native", "binary_install") }

# After
$Scripts = $ScriptsResponse | Where-Object { $_.language -notin @("native", "binary_install", "file_transfer") }
```

---

### Selenium EdgeOptions API compatibility issues

Multiple attempts to configure Edge browser window size and positioning failed due to PowerShell Selenium module API variations.

**Errors:**
```
Method invocation failed because [OpenQA.Selenium.Edge.EdgeOptions] does not contain a method named 'AddArgument'
Cannot find an overload for "AddAdditionalCapability" and the argument count: "3"
```

**Root cause:** PowerShell Selenium module has inconsistent EdgeOptions API across versions - some support `AddArgument()`, others support `AddArguments()`, and capability handling varies.

**Attempted approaches:**
1. `AddArgument("--headless")` - Method doesn't exist
2. `AddAdditionalCapability("headless", $true, $true)` - Wrong parameter count
3. `Arguments` property assignment - Property not settable
4. Chrome DevTools Protocol viewport emulation - Complex and unreliable

**Final solution:** Use fallback chain with proper error handling:
```powershell
try {
    $options.AddArguments("--window-position=-32000,-32000")
} catch {
    try {
        $options.AddAdditionalCapability("ms:edgeOptions", @{ args = @("--window-position=-32000,-32000") })
    } catch {
        Write-Warning "Could not set off-screen position, window may flash briefly"
    }
}
```

---

### Browser window positioning and startup optimization

Initial browser startup caused visual artifacts and slow performance.

**Issues:**
- Browser appeared large then resized (flash effect)
- 10+ second startup delays with off-screen positioning
- Multiple minimize/restore cycles
- Window appearing off-screen or partially hidden

**Evolution of solutions:**
1. **Direct positioning** - Browser appeared then resized (visible flash)
2. **Minimize approach** - Browser minimized, positioned, restored (still flashed)
3. **Off-screen positioning** - Started at (-32000,-32000), moved to center (slow but hidden)
4. **Headless-to-normal switching** - Complex and unreliable
5. **Optimized off-screen** - Reduced wait times from 1000ms to 300ms
6. **Dynamic screen detection** - Replaced hardcoded 5120x1440 with `System.Windows.Forms.Screen.PrimaryScreen`
7. **Parameter cleanup** - Removed ScreenWidth/ScreenHeight parameters since detection is automatic

**Current approach:** Off-screen start with minimal delay (300ms) for fastest hidden startup.

---

### Screen resolution detection failures

Hardcoded screen resolution failed on different monitor setups.

**Error:** Screen detection returned "x" (empty values) when using CIM/WMI queries.

**Attempted solutions:**
1. `Get-CimInstance Win32_DesktopMonitor` - Failed on some systems
2. `Get-WmiObject Win32_DesktopMonitor` - Also unreliable  
3. C# Win32 API calls - Complex and overkill

**Final solution:** `System.Windows.Forms.Screen.PrimaryScreen` - Most reliable:
```powershell
Add-Type -AssemblyName System.Windows.Forms
$primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
$screenWidth = $primaryScreen.Bounds.Width
$screenHeight = $primaryScreen.Bounds.Height
```

---

### Progress bar not dismissing after completion

Progress bar remained visible after script completion in parallel mode.

**Root cause:** `Write-Progress -Completed` only dismisses the progress bar when called from the same scope that owns it. In parallel mode, `Write-Progress` was being called inside runspaces (child scopes), not the main thread. The main thread never owned the progress bar, so `-Completed` on the main thread didn't dismiss it.

**Fix:** Remove all `Write-Progress` calls from inside the parallel block. Instead, stream results back to the main thread and update progress there:
```powershell
} -ThrottleLimit 50 | ForEach-Object {
    $result = $_
    # Update progress on main thread
    $total = $ProgressCounter['completed'] + $ProgressCounter['failed']
    $percent = [math]::Round(($total / $TotalScripts) * 100)
    Write-Progress -Activity "Downloading Scripts" `
        -Status "$total/$TotalScripts | $percent% Complete" `
        -PercentComplete $percent
    $result
}

Write-Progress -Activity "Downloading Scripts" -Completed  # Now works!
```

---

### Category folder names with invalid filesystem characters

Category names from API contained characters invalid for Windows folder names (`/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`).

**Error:** Failed to create directory with names like `"Hardware / Software"` or `"Client: Config"`.

**Fix:** Sanitize category names when building folder paths:
```powershell
$OriginalCategoryNames = ((Get-ScriptCategory -CategoryIDs $ScriptContent.categoriesIds -CategoriesHash $CategoriesHash) -replace '[\\/:*?"<>|]', '_')
```

This preserves the original category names in the CSV export and warning messages while using sanitized names for folder creation.

---

### Progress counter double-counting in parallel mode

Scripts were being counted twice, showing "Successfully saved: 2202" when total was 1101.

**Investigation:**
```powershell
# Progress counter updated in parallel block
if ($WriteSuccess) {
    $ProgressCounter.AddOrUpdate('completed', 1, { param($k, $v) $v + 1 })
}

# ALSO updated in streaming ForEach-Object
if (-not $result.Failed) {
    $ProgressCounter.AddOrUpdate('completed', 1, { param($k, $v) $v + 1 })
}
```

**Root cause:** Progress counter was being incremented twice - once inside the parallel runspace after file write, and again in the streaming `ForEach-Object` on the main thread. Each successful script was counted twice.

**Fix:** Remove progress counter updates from parallel block entirely. Only update in the streaming block:
```powershell
} -ThrottleLimit 50 | ForEach-Object {
    $result = $_
    if (-not $result.Failed) {
        $ProgressCounter.AddOrUpdate('completed', 1, { param($k, $v) $v + 1 }) | Out-Null
    } else {
        $ProgressCounter.AddOrUpdate('failed', 1, { param($k, $v) $v + 1 }) | Out-Null
    }
    # ... progress bar display ...
    $result  # pass through to $ScriptArray
}
```

Added `| Out-Null` to suppress the return value from `AddOrUpdate()` which was cluttering the pipeline.

---

## Rejected Approaches

| Approach | Why Rejected |
|---|---|
| OAuth Authorization Code Flow with PKCE | NinjaRMM only supports `client_credentials` grant, not authorization code flow |
| OAuth bearer tokens for script content | Legacy scripting endpoint doesn't support OAuth, only session cookies |
| Interactive browser popup like Connect-MgGraph | Requires OAuth authorization code flow which NinjaRMM doesn't provide |
| Selenium `Start-SeEdge` helper | Inconsistent driver detection, looking for obsolete `MicrosoftWebDriver.exe` |
| Plaintext session cache | Security vulnerability - anyone with filesystem access can steal session |
| Base64-encoded session cache | Not encryption, just encoding - trivial to decode |
| 8-hour session cache expiry | Too long - 2 hours provides better security/convenience balance |
| Skipping duplicate category scripts | Data loss - scripts with multiple categories were never saved |
| Sequential script downloads | Too slow for large libraries (1000+ scripts) |
| PowerShell 7 parallel processing | `$using:` scope limitations with complex expressions and method calls |
| Batch size of 20 | Too many batch iterations; 50 provides better performance |
| Progress bar removed | User feedback important; restored with batch processing |
| Hardcoded screen resolution | Failed on different monitor setups; dynamic detection more reliable |
| Selenium `AddArgument()` method | Doesn't exist in some PowerShell Selenium module versions |
| Browser zoom scaling via CSS | Unreliable; better to use appropriate window size |
| Complex minimize/restore cycles | Added unnecessary complexity and visual artifacts |
| Headless-to-normal browser switching | Too complex and unreliable for simple positioning needs |

---

## Known Limitations

- **EdgeDriver path hardcoded** - Script fails if driver not at `C:\Git\tawtek\ninjaone-api-scripts\msedgedriver.exe`
- **Category ID mapping hardcoded** - New categories require script update
- **Session TTL unknown** - NinjaRMM doesn't publish session expiry; 2-hour cache may be too conservative or too aggressive
- **Windows-only** - DPAPI encryption and EdgeDriver are Windows-specific
- **No retry logic** - Failed script downloads are logged but not retried
- **Selenium dependency** - Requires Selenium PowerShell module and EdgeDriver
- **Browser automation fragility** - Cookie capture relies on specific cookie name and domain
- **No parallel processing** - Batch processing is faster than sequential but slower than true parallelism

---

## TODO

- [ ] Parameterize EdgeDriver path instead of hardcoded location
- [ ] Add retry logic with exponential backoff for failed downloads
- [ ] Investigate NinjaRMM session TTL to optimize cache expiry
- [ ] Add `-ExportOnly` parameter to skip file writes and only generate CSV
- [ ] Consider `Start-ThreadJob` for PowerShell 5.1 parallelism if batch processing becomes bottleneck
- [ ] Add category ID auto-discovery instead of hardcoded mapping
- [ ] Add `-Verbose` support for detailed logging
- [ ] Add progress percentage to batch status messages

---

## Changelog

| Date | Summary |
|---|---|
| 2024-10-23 | Initial build - OAuth authentication, script metadata retrieval |
| 2024-10-23 | OAuth 401 on scripting endpoint - discovered OAuth doesn't work for script content |
| 2024-10-23 | Pivoted to Selenium browser automation for SSO/MFA support |
| 2024-10-23 | Fixed Selenium EdgeDriver detection - direct instantiation with explicit path |
| 2024-10-23 | Fixed script download failures - changed from OAuth tokens to session cookies |
| 2024-10-23 | Added session caching with DPAPI encryption |
| 2024-10-23 | Fixed plaintext session cache - convert to SecureString before export |
| 2024-10-23 | Fixed SecureString memory leak - added ZeroFreeBSTR cleanup |
| 2024-10-23 | Fixed stale comment - updated cache path reference |
| 2024-10-23 | Added `-ClearCache` parameter for manual session invalidation |
| 2024-10-23 | Reduced session cache expiry from 8 hours to 2 hours |
| 2024-10-23 | Added `file_transfer` to language exclusion filter |
| 2024-10-23 | Fixed duplicate category handling - save to `- Duplicates` folder instead of skipping |
| 2024-10-23 | Added batch processing (50 scripts per batch) for performance |
| 2024-10-23 | Restored progress bar with batch processing |
| 2026-04-06 | **Security review** - confirmed DPAPI encryption, memory cleanup, proper error handling |
| 2026-04-06 | **Refactor complete** - enterprise-grade security, session caching, batch processing, duplicate handling |
| 2026-04-07 | **Browser configuration** - Added dynamic screen detection, off-screen positioning, optimized startup |
| 2026-04-07 | **EdgeOptions compatibility** - Implemented fallback chain for different PowerShell Selenium module versions |
| 2026-04-07 | **Window positioning** - Replaced hardcoded screen resolution with dynamic detection via `System.Windows.Forms.Screen` |
| 2026-04-07 | **Parameter cleanup** - Removed ScreenWidth/ScreenHeight parameters since detection is automatic |
| 2026-04-07 | **Startup optimization** - Reduced browser startup delays from 10+ seconds to 2-3 seconds |
| 2026-04-07 | **Parallel processing** - Implemented PowerShell 7+ parallel mode with streaming progress and immediate file writes |
| 2026-04-07 | **Progress bar scope fix** - Moved progress updates to main thread to ensure proper dismissal with `-Completed` |
| 2026-04-07 | **Progress counter fix** - Resolved double-counting issue by updating counter only in streaming block |
| 2026-04-07 | **Category name sanitization** - Replace invalid filesystem characters in folder names while preserving originals in CSV |
| 2026-04-07 | **Memory optimization** - Files written immediately in parallel mode, no code stored in `$ScriptArray` |
| 2026-04-07 | **Duplicate scripts sorting** - Added multi-level sort (category A-Z, then name A-Z) for better readability |
| 2026-04-07 | **Sequential processing parameter** - Added `-Sequential` switch to force PowerShell 5.1-style sequential processing on PowerShell 7+ |
| 2026-04-07 | **Progress bar dismissal fix** - Fixed persistent progress bar in sequential mode by adding `Write-Progress -Completed` after foreach loop |
| 2026-04-07 | **Cache clearing parameter fix** - Fixed `-ClearCache` parameter passing to ensure proper session cache clearing |
| 2026-04-07 | **Parallel result collection fix** - Fixed missing `$result` output in parallel processing pipeline, causing zero script counts |
| 2026-04-07 | **Category mapping status table** - Added detailed category mapping table showing file paths for mapped categories and excluded status for native categories |
| 2026-04-07 | **Category folder name preservation** - Removed unnecessary character replacement, preserving original category names with spaces in folder paths |