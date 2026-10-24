# Verify the "date jump" add-in: one click on the Jump button must
#   (a) close the dialog, and
#   (b) actually move the Outlook calendar view.
# ASCII-only source; Chinese text is built from [char] codes.

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class Win {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    public static List<IntPtr> ByPid(uint pid) {
        var used = new List<IntPtr>();
        EnumWindows((h,l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid) used.Add(h); return true; }, IntPtr.Zero);
        return used;
    }
    public static string Title(IntPtr h) {
        var sb = new StringBuilder(512);
        GetWindowText(h, sb, 512);
        return sb.ToString();
    }
    public static IntPtr ByTitleContains(uint pid, string part) {
        foreach (var h in ByPid(pid)) {
            var t = Title(h);
            if (t.Length > 0 && t.Contains(part)) return h;
        }
        return IntPtr.Zero;
    }
    public static IntPtr ByTitleAny(uint pid, string[] titles) {
        foreach (var h in ByPid(pid)) {
            var t = Title(h);
            foreach (var want in titles) { if (t == want) return h; }
        }
        return IntPtr.Zero;
    }
}
"@

# --- Chinese literals ---
$dlgTitle   = ([char]0x8DF3) + ([char]0x8F6C) + ([char]0x5230) + ([char]0x65E5) + ([char]0x671F)   # dialog title
$datePart   = ([char]0x65E5) + ([char]0x671F)                                                        # "ri qi" part, loose match
$jumpText   = ([char]0x8DF3) + ([char]0x8F6C)                                                        # dialog OK button
$ribbonText = ([char]0x8DF3) + ([char]0x8F6C) + ([char]0x65E5) + ([char]0x671F)                      # ribbon button: tiao zhuan ri qi
$yearCh     = [char]0x5E74
$monthCh    = [char]0x6708
$dayCh      = [char]0x65E5
$viewCh     = ([char]0x89C6) + ([char]0x56FE)                                                        # shi tu
$eventCh    = ([char]0x4E8B) + ([char]0x4EF6) + ([char]0x603B) + ([char]0x6570)                      # "event total" - only in the month caption

function Find-ByName($root, $name, $typeName) {
    if (-not $root) { return $null }
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($el in $all) {
        try {
            if ($el.Current.Name -eq $name) {
                if (-not $typeName -or $el.Current.ControlType.ProgrammaticName -eq $typeName) { return $el }
            }
        } catch {}
    }
    return $null
}

function Invoke-El($el) {
    if (-not $el) { return $false }
    try {
        $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        return $true
    } catch {
        try {
            $el.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern).DoDefaultAction()
            return $true
        } catch {}
    }
    return $false
}

# Read the calendar view caption: element whose Name looks like
#   "back, month view, 2026 NIAN 10 YUE , 21 events"
function Get-ViewCaption($win) {
    if (-not $win) { return "NO_WIN" }
    try {
        $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($el in $all) {
            try {
                $n = $el.Current.Name
                # the month caption is the only element carrying "<n> events total";
                # other elements such as the "today" button also contain a year-month,
                # so this anchor is what makes the probe unambiguous.
                if ($n -and $n.Contains($eventCh)) {
                    return $n
                }
            } catch {}
        }
    } catch {}
    return "NOT_FOUND"
}

function Get-OutlookWin($pid2) {
    $rt = [System.Windows.Automation.AutomationElement]::RootElement
    $kids = $rt.FindAll([System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)
    $best = $null; $bestArea = 0
    foreach ($x in $kids) {
        try {
            if ($x.Current.ProcessId -eq $pid2 -and $x.Current.ControlType.ProgrammaticName -eq 'ControlType.Window') {
                $r = $x.Current.BoundingRectangle
                $area = $r.Width * $r.Height
                if ($area -gt $bestArea) { $bestArea = $area; $best = $x }
            }
        } catch {}
    }
    return $best
}

function Get-AllOutlookWins($pid2) {
    $rt = [System.Windows.Automation.AutomationElement]::RootElement
    $kids = $rt.FindAll([System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)
    $out = @()
    foreach ($x in $kids) {
        try {
            if ($x.Current.ProcessId -eq $pid2 -and $x.Current.ControlType.ProgrammaticName -eq 'ControlType.Window') { $out += $x }
        } catch {}
    }
    return $out
}

# find a ribbon button whose name contains the given text, across all windows.
# Several controls can match (built-in "go to date" item in the overflow menu, etc.),
# so collect them all, print their geometry, and return the LARGEST one (our large button).
function Find-RibbonButton($pid2, $text) {
    $hits = @()
    foreach ($w in (Get-AllOutlookWins $pid2)) {
        try {
            $all = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.Condition]::TrueCondition)
            foreach ($el in $all) {
                try {
                    if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and
                        $el.Current.Name -like "*$text*") { $hits += $el }
                } catch {}
            }
        } catch {}
    }
    if ($hits.Count -gt 0) { Write-Host ("ribbon candidates = " + $hits.Count) }
    $best = $null; $bestArea = -1
    foreach ($h in $hits) {
        try {
            $r = $h.Current.BoundingRectangle
            $area = $r.Width * $r.Height
            Write-Host ("   [" + $h.Current.Name + "] W=" + [int]$r.Width + " H=" + [int]$r.Height + " X=" + [int]$r.X + " Y=" + [int]$r.Y)
            if ($area -gt $bestArea) { $bestArea = $area; $best = $h }
        } catch {}
    }
    return $best
}

# find the calendar view caption across all windows
function Find-ViewCaptionAll($pid2) {
    foreach ($w in (Get-AllOutlookWins $pid2)) {
        $c = Get-ViewCaption $w
        if ($c -ne "NOT_FOUND" -and $c -ne "NO_WIN") { return $c }
    }
    return "NOT_FOUND"
}

# ---------- 0. make sure Outlook is running (it dies when the parent shell exits) ----------
$proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) {
    Write-Output "outlook not running -> starting"
    $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
    if (-not (Test-Path $exe)) { $exe = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\OUTLOOK.EXE' }
    Start-Process $exe
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Seconds 1
        $proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) { break }
    }
    if (-not $proc) { Write-Output "OUTLOOK_START_FAILED"; exit 1 }
    # wait until the main window really exists
    for ($i = 0; $i -lt 40; $i++) {
        $proc.Refresh()
        if ($proc.MainWindowHandle -ne [IntPtr]::Zero -and $proc.MainWindowTitle) { break }
        Start-Sleep -Seconds 1
    }
    Start-Sleep -Seconds 6
    $proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
}
$proc.Refresh()
Write-Output ("main window title = [" + $proc.MainWindowTitle + "]")

# ---------- 1. activate the Calendar module via COM ----------
$ol = $null
try {
    $ol = New-Object -ComObject Outlook.Application
    Write-Output "COM: attached via New-Object"
} catch {
    try {
        $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Output "COM: attached via GetActiveObject"
    } catch {
        Write-Output ("COM ATTACH FAILED: " + $_.Exception.Message)
    }
}
if ($ol) {
    try {
        $ns = $ol.GetNamespace('MAPI')
        $calFolder = $ns.GetDefaultFolder(9)     # olFolderCalendar
        # switch the MAIN explorer to the calendar folder (keeps a single window)
        $exp = $ol.ActiveExplorer()
        if ($exp) {
            $exp.CurrentFolder = $calFolder
            $exp.Display()
            Write-Output ("main explorer folder = " + $exp.CurrentFolder.Name)
        } else {
            $calFolder.Display()
            Write-Output "no active explorer -> used Folder.Display()"
        }
    } catch {
        Write-Output ("CAL SWITCH FAILED: " + $_.Exception.Message)
    }
}
Start-Sleep -Seconds 5

$proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) { Write-Output "NO_OUTLOOK"; exit 1 }
Write-Output ("outlook pid=" + $proc.Id)

$outlookWin = Get-OutlookWin $proc.Id
$before = "NOT_FOUND"
for ($i = 0; $i -lt 20; $i++) {
    $before = Find-ViewCaptionAll $proc.Id
    if ($before -ne "NOT_FOUND") { break }
    Start-Sleep -Seconds 1
}
Write-Output ("view BEFORE = " + $before)
if ($before -eq "NOT_FOUND" -or $before -eq "NO_WIN") {
    Write-Output "=== caption diagnostics (elements containing the month char) ==="
    try {
        $all = $outlookWin.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        $n = 0
        foreach ($el in $all) {
            try {
                $nm = $el.Current.Name
                if ($nm -and $nm.Contains($monthCh) -and $nm.Length -lt 90) {
                    $r = $el.Current.BoundingRectangle
                    Write-Output ("   [" + $nm + "] " + $el.Current.ControlType.ProgrammaticName + " X=" + [int]$r.X + " Y=" + [int]$r.Y)
                    $n++
                    if ($n -ge 12) { break }
                }
            } catch {}
        }
        if ($n -eq 0) { Write-Output "   (none)" }
    } catch { Write-Output ("   diag err: " + $_.Exception.Message) }
}

# ---------- 2. invoke the ribbon button ----------
# VSTO add-ins load asynchronously: right after a cold Outlook start the custom
# ribbon group may not exist yet, so retry for a while before giving up.
# Office only realises the controls of the ACTIVE ribbon tab, so make sure the
# Calendar tab is selected first.
$calTabText = ([char]0x65E5) + ([char]0x5386)
function Select-CalendarTab($pid2, $text) {
    $rt = [System.Windows.Automation.AutomationElement]::RootElement
    foreach ($x in $rt.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
        try {
            if ($x.Current.ProcessId -ne $pid2) { continue }
            $tabs = $x.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::TabItem)))
            foreach ($t in $tabs) {
                try {
                    if ($t.Current.Name -eq $text) {
                        $t.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
                        return $true
                    }
                } catch {}
            }
        } catch {}
    }
    return $false
}

$btn = $null
for ($try = 0; $try -lt 30; $try++) {
    $btn = Find-RibbonButton $proc.Id $ribbonText
    if ($btn) { break }
    if ($try -eq 0) { Write-Host ("select calendar tab = " + (Select-CalendarTab $proc.Id $calTabText)) }
    if ($try -eq 10 -and $exp) {
        Write-Host "still nothing after 10s -> toggling ribbon visibility"
        try { $exp.CommandBars.ExecuteMso("MinimizeRibbon") } catch { Write-Host ("expand failed: " + $_.Exception.Message) }
    }
    Start-Sleep -Seconds 1
}

if (-not $btn) {
    Write-Output "RIBBON_BUTTON_NOT_FOUND; dumping candidate buttons:"
    try {
        $all = $outlookWin.FindAll([System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        $n = 0
        foreach ($el in $all) {
            try {
                if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $el.Current.Name -like "*$ribbonText*") {
                    Write-Output ("   [" + $el.Current.Name + "]")
                    $n++
                    if ($n -gt 12) { break }
                }
            } catch {}
        }
        if ($n -gt 0) { Write-Output "no exact match; $n partial matches above" }
        if ($n -eq 0) {
            Write-Output "   (no partial match) buttons in the ribbon band (Y 80..240):"
            $m = 0
            foreach ($el in $all) {
                try {
                    if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $el.Current.Name) {
                        $rr = $el.Current.BoundingRectangle
                        if ($rr.Y -gt 80 -and $rr.Y -lt 240) {
                            Write-Output ("     [" + $el.Current.Name + "] W=" + [int]$rr.Width + " H=" + [int]$rr.Height + " X=" + [int]$rr.X + " Y=" + [int]$rr.Y)
                            $m++
                            if ($m -ge 30) { break }
                        }
                    }
                } catch {}
            }
            if ($m -eq 0) { Write-Output "     (none - ribbon likely collapsed)" }
        }
    } catch {}
    exit 1
}
Write-Output ("ribbon button rect: W=" + [int]$btn.Current.BoundingRectangle.Width + " H=" + [int]$btn.Current.BoundingRectangle.Height)
$ok = Invoke-El $btn
Write-Output ("ribbon invoked = " + $ok)

# ---------- 3. wait for our dialog ----------
$dlgHwnd = [IntPtr]::Zero
$dlgRealTitle = ""
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 200
    $dlgHwnd = [Win]::ByTitleContains([uint32]$proc.Id, $datePart)
    if ($dlgHwnd -ne [IntPtr]::Zero) { $dlgRealTitle = [Win]::Title($dlgHwnd); break }
}
if ($dlgHwnd -eq [IntPtr]::Zero) {
    Write-Output "DIALOG_NOT_FOUND (titles seen):"
    foreach ($h in [Win]::ByPid([uint32]$proc.Id)) { Write-Output ("   [" + [Win]::Title($h) + "]") }
    exit 1
}
Write-Output ("dialog hwnd=" + $dlgHwnd + "  real title=[" + $dlgRealTitle + "] len=" + $dlgRealTitle.Length)
Start-Sleep -Milliseconds 900

$dlg = [System.Windows.Automation.AutomationElement]::FromHandle($dlgHwnd)
$dlgRect = $dlg.Current.BoundingRectangle
Write-Output ("dialog size = " + [int]$dlgRect.Width + " x " + [int]$dlgRect.Height)

# ---------- 4. verify title structure (only ONE big title text, no small duplicate) ----------
$titleHits = 0
$titleInfo = @()
try {
    $all = $dlg.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($el in $all) {
        try {
            $nm = $el.Current.Name
            if ($nm -and $nm.Contains($datePart) -and $el.Current.ControlType.ProgrammaticName -eq 'ControlType.Text') {
                $r = $el.Current.BoundingRectangle
                $titleInfo += ("[" + $nm + "] W=" + [int]$r.Width + " H=" + [int]$r.Height)
                $titleHits++
            }
        } catch {}
    }
} catch {}
Write-Output ("title text elements = " + $titleHits)
foreach ($ti in $titleInfo) { Write-Output ("   " + $ti) }

# ---------- 5. verify calendar fills the form ----------
$calEl = $null
try {
    $all = $dlg.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($el in $all) {
        try { if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Calendar') { $calEl = $el; break } } catch {}
    }
} catch {}
if ($calEl) {
    $cr = $calEl.Current.BoundingRectangle
    $pctW = [Math]::Round(100.0 * $cr.Width / $dlgRect.Width, 1)
    $pctH = [Math]::Round(100.0 * $cr.Height / $dlgRect.Height, 1)
    Write-Output ("calendar = " + [int]$cr.Width + " x " + [int]$cr.Height + "  (" + $pctW + "% W, " + $pctH + "% H of dialog)")
} else {
    Write-Output "CALENDAR_NOT_FOUND"
}

# ---------- 6. move the calendar to another month, then pick a day ----------
$navigated = 0
if ($calEl) {
    try {
        $next = $calEl.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::AutomationIdProperty, 'PART_NextButton')))
        if ($next) {
            for ($k = 0; $k -lt 2; $k++) {
                if (Invoke-El $next) { $navigated++ }
                Start-Sleep -Milliseconds 700
            }
        }
    } catch {}
}
Write-Output ("month navigation invocations = " + $navigated)
Start-Sleep -Milliseconds 700

$dayCells = @()
try {
    $all = $dlg.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($el in $all) {
        try {
            $n = $el.Current.Name
            if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and
                $n -match '^\d{4}' -and $n.Contains($monthCh) -and $n.Contains($dayCh)) {
                $dayCells += $el
            }
        } catch {}
    }
} catch {}
Write-Output ("day cells found = " + $dayCells.Count)

$target = $null
foreach ($c in $dayCells) {
    if ($c.Current.Name -match '15') { $target = $c; break }
}
if (-not $target -and $dayCells.Count -gt 0) { $target = $dayCells[$dayCells.Count - 1] }
if ($target) {
    Write-Output ("target day = " + $target.Current.Name)
    Invoke-El $target | Out-Null
} else {
    Write-Output "NO_DAY_CELL"
}
Start-Sleep -Milliseconds 800

# ---------- 7. ONE click on the Jump button ----------
$jumpBtn = Find-ByName $dlg $jumpText 'ControlType.Button'
if (-not $jumpBtn) { Write-Output "JUMP_BUTTON_NOT_FOUND"; exit 1 }
$jumpRect = $jumpBtn.Current.BoundingRectangle
Write-Output ("jump button = " + [int]$jumpRect.Width + " x " + [int]$jumpRect.Height + " at " + [int]$jumpRect.X + "," + [int]$jumpRect.Y)
Write-Output "invoking jump ONCE ..."
$clickOk = Invoke-El $jumpBtn
Write-Output ("jump invoked = " + $clickOk)

# ---------- 8. did ONE click close the dialog? ----------
$closed = $false
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep -Milliseconds 200
    if (-not [Win]::IsWindow($dlgHwnd)) { $closed = $true; break }
    if ([Win]::Title($dlgHwnd) -ne $dlgRealTitle) { $closed = $true; break }
}
Write-Output ("RESULT dialog_closed_after_ONE_click = " + $closed)

Start-Sleep -Seconds 2

# ---------- 9. did the view move? ----------
$after = "NOT_FOUND"
for ($i = 0; $i -lt 12; $i++) {
    $after = Find-ViewCaptionAll $proc.Id
    if ($after -ne "NOT_FOUND") { break }
    Start-Sleep -Milliseconds 500
}
Write-Output ("view AFTER  = " + $after)
if ($before -ne "NOT_FOUND" -and $after -ne "NOT_FOUND" -and $before -ne $after) {
    Write-Output "RESULT VIEW_MOVED"
} elseif ($before -eq "NOT_FOUND" -or $after -eq "NOT_FOUND") {
    Write-Output "RESULT CAPTION_UNREADABLE"
} else {
    Write-Output "RESULT VIEW_NOT_MOVED"
}

# ---------- 10. trace tail ----------
$log = Join-Path $env:TEMP 'odjdbg.log'
Write-Output "=== trace tail ==="
if (Test-Path $log) { Get-Content $log -Tail 20 } else { Write-Output "(no trace log)" }
