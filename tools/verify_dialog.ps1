# One-shot: start Outlook -> calendar -> open picker -> read dialog geometry -> report.
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Start-Process (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\OUTLOOK.EXE')

$proc = $null
for ($i = 0; $i -lt 45; $i++) {
    Start-Sleep -Seconds 2
    $proc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    if ($proc) { break }
}
if (-not $proc) { Write-Output "NO_OUTLOOK"; exit 1 }
Write-Output ("pid=" + $proc.Id)
Start-Sleep -Seconds 8

function Get-WinByPid($pid2) {
    $rt = [System.Windows.Automation.AutomationElement]::RootElement
    $ws = $rt.FindAll([System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($x in $ws) {
        try { if ($x.Current.ProcessId -eq $pid2 -and $x.Current.Name -ne "") { return $x } } catch {}
    }
    return $null
}

$cal = [char]0x65E5 + [char]0x5386
$win = Get-WinByPid $proc.Id
for ($try = 0; $try -lt 10 -and $win; $try++) {
    $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($el in $all) {
        try {
            if ($el.Current.Name -eq $cal -and $el.Current.ControlType.ProgrammaticName -match 'TreeItem|ListItem|Button') {
                $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                break
            }
        } catch {}
    }
    Start-Sleep -Seconds 3
}
Start-Sleep -Seconds 2

$btnText = [char]0x8F6C + [char]0x5230
$win = Get-WinByPid $proc.Id
$all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)
foreach ($el in $all) {
    try {
        if ($el.Current.Name -eq $btnText -and $el.Current.ControlType.ProgrammaticName -eq 'ControlType.Button') {
            $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Write-Output "clicked goto"
            break
        }
    } catch {}
}
Start-Sleep -Seconds 5

$dlg = Get-WinByPid $proc.Id
if (-not $dlg) { Write-Output "NO_DIALOG"; exit 1 }
$r = $dlg.Current.BoundingRectangle
Write-Output ("dialog=[" + $dlg.Current.Name + "] X=" + [int]$r.X + " Y=" + [int]$r.Y + " W=" + [int]$r.Width + " H=" + [int]$r.Height)

$items = $dlg.FindAll([System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition)

$titleTxt = [char]0x8DF3 + [char]0x8F6C + [char]0x5230 + [char]0x65E5 + [char]0x671F
$todayTxt = [char]0x4ECA + [char]0x5929
$okTxt = [char]0x8DF3 + [char]0x8F6C
$cancelTxt = [char]0x53D6 + [char]0x6D88

foreach ($el in $items) {
    try {
        $n = $el.Current.Name
        $ct = $el.Current.ControlType.ProgrammaticName
        if ($n -eq $titleTxt -or $n -eq $todayTxt -or $n -eq $okTxt -or $n -eq $cancelTxt -or $ct -match 'ComboBox|Calendar|Image') {
            $rr = $el.Current.BoundingRectangle
            if ($rr.Width -gt 0) {
                Write-Output ("[" + $n + "] " + $ct + " X=" + [int]$rr.X + " Y=" + [int]$rr.Y + " W=" + [int]$rr.Width + " H=" + [int]$rr.Height)
            }
        }
    } catch {}
}
