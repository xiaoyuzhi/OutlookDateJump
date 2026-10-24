# Verify the date picker dialog with REAL mouse messages (WM_LBUTTONDOWN/UP),
# not UIA Invoke - this is what reproduces the "click twice" problem.
# Also dumps a screenshot and the UIA geometry of the dialog.
# ASCII-only source; Chinese text is built from [char] codes.

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class W2 {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr h, ref POINT p);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
    public static List<IntPtr> ByPid(uint pid) {
        var r = new List<IntPtr>();
        EnumWindows((h,l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid) r.Add(h); return true; }, IntPtr.Zero);
        return r;
    }
    public static string Title(IntPtr h) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString(); }
    public static IntPtr ByTitleContains(uint pid, string part) {
        foreach (var h in ByPid(pid)) { var t = Title(h); if (t.Length > 0 && t.Contains(part)) return h; }
        return IntPtr.Zero;
    }
    public static void Click(IntPtr h, int sx, int sy) {
        var p = new POINT(); p.X = sx; p.Y = sy;
        ScreenToClient(h, ref p);
        IntPtr lp = (IntPtr)((p.Y << 16) | (p.X & 0xFFFF));
        PostMessage(h, 0x0200, IntPtr.Zero, lp);              // WM_MOUSEMOVE
        PostMessage(h, 0x0201, (IntPtr)1, lp);                // WM_LBUTTONDOWN
        PostMessage(h, 0x0202, IntPtr.Zero, lp);              // WM_LBUTTONUP
    }
    public static string Shot(IntPtr h, string path) {
        RECT r;
        if (!GetWindowRect(h, out r)) return "no-rect";
        int w = r.R - r.L, ht = r.B - r.T;
        using (var bmp = new System.Drawing.Bitmap(w, ht)) {
            using (var g = System.Drawing.Graphics.FromImage(bmp)) {
                IntPtr hdc = g.GetHdc();
                PrintWindow(h, hdc, 2);
                g.ReleaseHdc(hdc);
            }
            bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }
        return w + "x" + ht;
    }
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
}
"@ -ReferencedAssemblies System.Drawing

$dlgKey = ([char]0x65E5) + ([char]0x671F)          # "ri qi"
$jumpText = ([char]0x8DF3) + ([char]0x8F6C)        # "tiao zhuan"
$ribbonText = ([char]0x8DF3) + ([char]0x8F6C) + ([char]0x65E5) + ([char]0x671F)
$monthCh = [char]0x6708
$dayCh = [char]0x65E5
$eventCh = ([char]0x4E8B) + ([char]0x4EF6) + ([char]0x603B) + ([char]0x6570)

function All($root) {
    return $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
}
function Invoke-El($el) {
    if (-not $el) { return $false }
    try { $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); return $true } catch {}
    return $false
}
function Find-Text($root, $text) {
    foreach ($el in (All $root)) {
        try { if ($el.Current.Name -eq $text -and $el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button') { return $el } } catch {}
    }
    return $null
}
function Caption($pid2) {
    $rt = [System.Windows.Automation.AutomationElement]::RootElement
    foreach ($x in $rt.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
        try {
            if ($x.Current.ProcessId -ne $pid2) { continue }
            if ($x.Current.ControlType.ProgrammaticName -ne 'ControlType.Window') { continue }
            foreach ($el in (All $x)) {
                try { $n = $el.Current.Name; if ($n -and $n.Contains($eventCh)) { return $n } } catch {}
            }
        } catch {}
    }
    return "NOT_FOUND"
}

# ---------- start / attach Outlook ----------
$proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) {
    $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE'
    if (-not (Test-Path $exe)) { $exe = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\OUTLOOK.EXE' }
    Start-Process $exe
    for ($i = 0; $i -lt 40; $i++) { Start-Sleep -Seconds 1; $proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1; if ($proc) { break } }
    if (-not $proc) { Write-Output "OUTLOOK_START_FAILED"; exit 1 }
    for ($i = 0; $i -lt 40; $i++) { $proc.Refresh(); if ($proc.MainWindowHandle -ne [IntPtr]::Zero) { break }; Start-Sleep -Seconds 1 }
    Start-Sleep -Seconds 6
    $proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1
}
$proc.Refresh()
$pid2 = $proc.Id
Write-Output ("outlook pid=" + $pid2)

# close a leftover dialog from a previous run
$old = [W2]::ByTitleContains([uint32]$pid2, $dlgKey)
if ($old -ne [IntPtr]::Zero) {
    Write-Output "closing leftover dialog"
    [W2]::PostMessage($old, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
    Start-Sleep -Seconds 1
}

$ol = $null
try { $ol = New-Object -ComObject Outlook.Application } catch { }
if ($ol) {
    try {
        $ns = $ol.GetNamespace('MAPI')
        $exp = $ol.ActiveExplorer()
        if ($exp) { $exp.CurrentFolder = $ns.GetDefaultFolder(9); $exp.Display() } else { $ns.GetDefaultFolder(9).Display() }
    } catch { Write-Output ("cal switch failed: " + $_.Exception.Message) }
}
Start-Sleep -Seconds 4
$before = Caption $pid2
Write-Output ("view BEFORE = " + $before)

# ---------- open the dialog through the ribbon ----------
$rt = [System.Windows.Automation.AutomationElement]::RootElement
$ribbon = $null
for ($try = 0; $try -lt 25; $try++) {
    foreach ($x in $rt.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)) {
        try {
            if ($x.Current.ProcessId -eq $pid2 -and $x.Current.ControlType.ProgrammaticName -eq 'ControlType.Window') {
                foreach ($el in (All $x)) { try { if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $el.Current.Name -eq $ribbonText) { $ribbon = $el } } catch {} }
            }
        } catch {}
    }
    if ($ribbon) { break }
    if ($try -eq 8 -and $exp) { try { $exp.CommandBars.ExecuteMso("MinimizeRibbon") } catch {} }
    Start-Sleep -Seconds 1
}
if (-not $ribbon) { Write-Output "RIBBON_NOT_FOUND"; exit 1 }
Invoke-El $ribbon | Out-Null
Start-Sleep -Seconds 2

$dlgHwnd = [W2]::ByTitleContains([uint32]$pid2, $dlgKey)
if ($dlgHwnd -eq [IntPtr]::Zero) {
    Write-Output "DIALOG_NOT_FOUND; titles:"
    foreach ($h in [W2]::ByPid([uint32]$pid2)) { Write-Output ("   [" + [W2]::Title($h) + "]") }
    exit 1
}
Write-Output ("dialog title=[" + [W2]::Title($dlgHwnd) + "]")
[W2]::SetForegroundWindow($dlgHwnd) | Out-Null
Start-Sleep -Milliseconds 600

# ---------- screenshot ----------
$shot = Join-Path (Get-Location) 'tools\shot_dialog.png'
Write-Output ("screenshot = " + [W2]::Shot($dlgHwnd, $shot) + " -> " + $shot)

$dlg = [System.Windows.Automation.AutomationElement]::FromHandle($dlgHwnd)
$dr = $dlg.Current.BoundingRectangle
Write-Output ("dialog rect = " + [int]$dr.X + "," + [int]$dr.Y + " " + [int]$dr.Width + "x" + [int]$dr.Height)

# ---------- UIA geometry dump ----------
Write-Output "=== UIA geometry ==="
foreach ($el in (All $dlg)) {
    try {
        $ct = $el.Current.ControlType.ProgrammaticName
        if ($ct -match 'Calendar|Button|ComboBox|Text|Edit|Pane|Document') {
            $n = $el.Current.Name
            if ($null -eq $n) { $n = "" }
            $n = $n.Replace("`r", " ").Replace("`n", " ")
            if ($n.Length -gt 30) { $n = $n.Substring(0, 30) }
            $r = $el.Current.BoundingRectangle
            if ($r.Width -gt 1) {
                Write-Output ("  " + $ct.Replace('ControlType.', '') + " [" + $n + "] " + [int]$r.X + "," + [int]$r.Y + " " + [int]$r.Width + "x" + [int]$r.Height)
            }
        }
    } catch {}
}

# ---------- pick a day with a REAL mouse click ----------
$days = @()
foreach ($el in (All $dlg)) {
    try {
        $n = $el.Current.Name
        if ($el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $n -match '^\d{4}' -and $n.Contains($monthCh) -and $n.Contains($dayCh)) { $days += $el }
    } catch {}
}
Write-Output ("day cells = " + $days.Count)
$target = $null
foreach ($d in $days) { if ($d.Current.Name -match '1([0-9])' + '$') { } }
if ($days.Count -gt 0) { $target = $days[[int]($days.Count / 2)] }
if (-not $target) { Write-Output "NO_DAY_CELL"; exit 1 }
$tr = $target.Current.BoundingRectangle
Write-Output ("clicking day [" + $target.Current.Name + "] at " + [int]$tr.X + "," + [int]$tr.Y)
[W2]::Click($dlgHwnd, [int]($tr.X + $tr.Width / 2), [int]($tr.Y + $tr.Height / 2))
Start-Sleep -Milliseconds 500

# ---------- ONE real click on the Jump button ----------
$okBtn = Find-Text $dlg $jumpText
if (-not $okBtn) { Write-Output "JUMP_BTN_NOT_FOUND"; exit 1 }
$orr = $okBtn.Current.BoundingRectangle
Write-Output ("jump button at " + [int]$orr.X + "," + [int]$orr.Y + " " + [int]$orr.Width + "x" + [int]$orr.Height)
Write-Output "posting ONE real mouse click on Jump ..."
[W2]::Click($dlgHwnd, [int]($orr.X + $orr.Width / 2), [int]($orr.Y + $orr.Height / 2))

$closed = $false
for ($i = 0; $i -lt 20; $i++) { Start-Sleep -Milliseconds 200; if (-not [W2]::IsWindow($dlgHwnd)) { $closed = $true; break } }
Write-Output ("RESULT dialog_closed_after_ONE_real_click = " + $closed)

Start-Sleep -Seconds 2
$after = Caption $pid2
Write-Output ("view AFTER  = " + $after)

Write-Output "=== trace tail ==="
$log = Join-Path $env:TEMP 'odjdbg.log'
if (Test-Path $log) { Get-Content $log -Tail 25 } else { Write-Output "(no log)" }
