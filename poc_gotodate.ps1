# POC: verify that Outlook's View.GoToDate works on this machine.
# Usage: open Outlook and switch to the Calendar view first, then run:
#   powershell -ExecutionPolicy Bypass -File .\poc_gotodate.ps1
# It makes the calendar jump to 2026-12-25.
# NOTE: keep this file ASCII-only. Windows PowerShell 5.1 reads non-BOM
#       UTF-8 files as ANSI(GBK) and Chinese strings would break parsing.

$ErrorActionPreference = "Stop"
$target = [datetime]"2026-12-25"

try {
    # Attach to the running Outlook instance (fails with 0x800401E3 if not running)
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $explorer = $ol.ActiveExplorer()

    if ($null -eq $explorer) {
        Write-Host "ERROR: no active Outlook window. Open Outlook and go to Calendar view." -ForegroundColor Yellow
        exit 1
    }

    $view = $explorer.CurrentView
    if ($null -eq $view) {
        Write-Host "ERROR: no current view on the active explorer." -ForegroundColor Yellow
        exit 1
    }

    # The one line that matters: jump the calendar view to the target date
    $view.GoToDate($target)
    Write-Host "OK: calendar jumped to $($target.ToShortDateString()). Core logic works!" -ForegroundColor Green
}
catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Hints: Outlook must already be running and in Calendar view." -ForegroundColor Yellow
    Write-Host "       'New Outlook' does not support COM - this approach needs classic Outlook." -ForegroundColor Yellow
}
