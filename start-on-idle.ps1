# Configuration
$idleThresholdSeconds = 3  # 5 minutes (adjust as needed)
$presentationPath = Join-Path $PSScriptRoot "presentation.html"
$checkIntervalSeconds = 10   # How often to check idle time

Write-Host "Monitoring system idle time..." -ForegroundColor Green
Write-Host "Will open presentation after $($idleThresholdSeconds/60) minutes of inactivity" -ForegroundColor Yellow
Write-Host "Press Ctrl+C to stop monitoring`n" -ForegroundColor Cyan

$presentationOpened = $false

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
        Start-Process $presentationPath
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
