# Configuration
$idleThresholdSeconds = 3  # 5 minutes (adjust as needed)
$presentationPath = Join-Path $PSScriptRoot "presentation.html"
$checkIntervalSeconds = 10   # How often to check idle time
$preferredBrowser = "chrome"  # Options: "chrome", "edge", "firefox", or "default"

Write-Host "Monitoring system idle time..." -ForegroundColor Green
Write-Host "Will open presentation after $($idleThresholdSeconds/60) minutes of inactivity" -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop monitoring`n" -ForegroundColor Cyan

$presentationOpened = $false

# Function to open browser in fullscreen/kiosk mode
function Open-PresentationFullscreen {
    param([string]$htmlPath)

    $fullPath = Resolve-Path $htmlPath
    $url = "file:///$($fullPath.Path.Replace('\', '/'))"

    $browserPaths = @{
        "chrome" = @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
            "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
        )
        "edge" = @(
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        )
        "firefox" = @(
            "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
            "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
        )
    }

    $opened = $false

    if ($preferredBrowser -ne "default") {
        # Try preferred browser
        foreach ($path in $browserPaths[$preferredBrowser]) {
            if (Test-Path $path) {
                Write-Host "Opening in $preferredBrowser (fullscreen mode)..." -ForegroundColor Cyan
                if ($preferredBrowser -eq "firefox") {
                    Start-Process $path -ArgumentList "--kiosk", $url
                } else {
                    Start-Process $path -ArgumentList "--start-fullscreen", $url
                }
                $opened = $true
                break
            }
        }
    }

    # Fallback: try all browsers in order
    if (-not $opened) {
        foreach ($browser in @("chrome", "edge", "firefox")) {
            foreach ($path in $browserPaths[$browser]) {
                if (Test-Path $path) {
                    Write-Host "Opening in $browser (fullscreen mode)..." -ForegroundColor Cyan
                    if ($browser -eq "firefox") {
                        Start-Process $path -ArgumentList "--kiosk", $url
                    } else {
                        Start-Process $path -ArgumentList "--start-fullscreen", $url
                    }
                    $opened = $true
                    break
                }
            }
            if ($opened) { break }
        }
    }

    # Last resort: use default browser (non-fullscreen)
    if (-not $opened) {
        Write-Host "No supported browser found. Opening with default browser..." -ForegroundColor Yellow
        Start-Process $fullPath
    }
}

# Add type for getting last input info
Add-Type @'
using System;
using System.Runtime.InteropServices;

public class IdleTimeChecker {
    [DllImport("user32.dll")]
    static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO {
        public uint cbSize;
        public uint dwTime;
    }

    public static uint GetIdleTime() {
        LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
        lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);

        if (!GetLastInputInfo(ref lastInputInfo)) {
            return 0;
        }

        return ((uint)Environment.TickCount - lastInputInfo.dwTime) / 1000;
    }
}
'@

while ($true) {
    $idleSeconds = [IdleTimeChecker]::GetIdleTime()

    if ($idleSeconds -ge $idleThresholdSeconds -and -not $presentationOpened) {
        Write-Host "System idle for $idleSeconds seconds. Opening presentation..." -ForegroundColor Green
        Open-PresentationFullscreen -htmlPath $presentationPath
        $presentationOpened = $true
    }
    elseif ($idleSeconds -lt $idleThresholdSeconds -and $presentationOpened) {
        Write-Host "Activity detected. Resetting monitor..." -ForegroundColor Yellow
        $presentationOpened = $false
    }
    elseif ($idleSeconds -ge $idleThresholdSeconds) {
        # Already opened, still idle
    }
    else {
        Write-Host "Idle: $idleSeconds s / $idleThresholdSeconds s" -NoNewline
        Write-Host "`r" -NoNewline
    }

    Start-Sleep -Seconds $checkIntervalSeconds
}
