# Outlook Date Jump - standalone WPF tool (no compile, no install)
# Usage (run from the folder containing this file):
#   powershell -ExecutionPolicy Bypass -File .\OutlookDateJump.ps1
# Shows a window with Year/Month/Day dropdowns plus a calendar. The three
# dropdowns are linked (day count follows month/year, including leap years).
# On "Jump" it makes the Outlook calendar view jump to the selected date
# via View.GoToDate.
# NOTE: ASCII-only. Windows PowerShell 5.1 misreads non-BOM UTF-8 as GBK.

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Jump to Date" Height="430" Width="330" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" ShowInTaskbar="False" Topmost="True">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,8">
      <ComboBox x:Name="CmbYear" Width="92" Margin="0,0,8,0"/>
      <ComboBox x:Name="CmbMonth" Width="64" Margin="0,0,8,0"/>
      <ComboBox x:Name="CmbDay" Width="64"/>
    </StackPanel>

    <Calendar x:Name="Cal" Grid.Row="1" SelectionMode="SingleDate"
              HorizontalAlignment="Center" VerticalAlignment="Top"/>

    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="BtnToday" Content="Today" Width="72" Margin="0,0,8,0"/>
      <Button x:Name="BtnOk" Content="Jump" Width="72" Margin="0,0,8,0" IsDefault="True"/>
      <Button x:Name="BtnCancel" Content="Cancel" Width="72" IsCancel="True"/>
    </StackPanel>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Set the program logo as the window icon. Resolve LOGO.png relative to this
# script so the tool stays portable (copy the .ps1 + LOGO.png together anywhere).
try {
    $iconPath = Join-Path $PSScriptRoot 'LOGO.png'
    if (Test-Path $iconPath) {
        $window.Icon = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$iconPath)
    }
} catch { }

$cal = $window.FindName("Cal")
$cmbYear = $window.FindName("CmbYear")
$cmbMonth = $window.FindName("CmbMonth")
$cmbDay = $window.FindName("CmbDay")
$btnToday = $window.FindName("BtnToday")
$btnOk = $window.FindName("BtnOk")
$btnCancel = $window.FindName("BtnCancel")

$now = [datetime]::Today

# populate year and month lists (wide practical range for mailbox archives)
$startYear = 1980
$endYear = 2100
for ($y = $startYear; $y -le $endYear; $y++) { [void]$cmbYear.Items.Add($y) }
for ($m = 1; $m -le 12; $m++) { [void]$cmbMonth.Items.Add($m) }

# refill the day list for the given year/month; keep the previous day if valid
function Fill-Days {
    param($year, $month)
    $prev = $cmbDay.SelectedItem
    $cmbDay.Items.Clear()
    $days = [datetime]::DaysInMonth($year, $month)
    for ($d = 1; $d -le $days; $d++) { [void]$cmbDay.Items.Add($d) }
    if ($prev -and $prev -le $days) { $cmbDay.SelectedItem = $prev }
    elseif ($cmbDay.Items.Count -ge 1) { $cmbDay.SelectedIndex = 0 }
}

$script:busy = $false

# calendar -> dropdowns
function Sync-CombosFromCal {
    $dt = $cal.SelectedDate
    if (-not $dt) { return }
    $script:busy = $true
    $cmbYear.SelectedItem = $dt.Year
    $cmbMonth.SelectedItem = $dt.Month
    Fill-Days $dt.Year $dt.Month
    $cmbDay.SelectedItem = $dt.Day
    $script:busy = $false
}

# dropdowns -> calendar (and each other)
function Sync-CalFromCombos {
    if ($script:busy) { return }
    $y = $cmbYear.SelectedItem
    $m = $cmbMonth.SelectedItem
    $d = $cmbDay.SelectedItem
    if (-not ($y -and $m -and $d)) { return }
    $script:busy = $true
    try {
        $dt = [datetime]::new($y, $m, $d)
        $cal.SelectedDate = $dt
        $cal.DisplayDate = $dt
    } catch { }
    $script:busy = $false
}

$cal.Add_SelectedDatesChanged({ Sync-CombosFromCal })
$cmbYear.Add_SelectionChanged({ Sync-CalFromCombos })
$cmbMonth.Add_SelectionChanged({ Sync-CalFromCombos })
$cmbDay.Add_SelectionChanged({ Sync-CalFromCombos })

# initial selection = today
$cmbYear.SelectedItem = $now.Year
$cmbMonth.SelectedItem = $now.Month
Fill-Days $now.Year $now.Month
$cmbDay.SelectedItem = $now.Day
$cal.SelectedDate = $now
$cal.DisplayDate = $now

$script:picked = $null

$btnToday.Add_Click({
    $cmbYear.SelectedItem = $now.Year
    $cmbMonth.SelectedItem = $now.Month
    Fill-Days $now.Year $now.Month
    $cmbDay.SelectedItem = $now.Day
    $cal.SelectedDate = $now
    $cal.DisplayDate = $now
})

$btnOk.Add_Click({
    $d = $cal.SelectedDate
    if ($d -ne $null) {
        $script:picked = $d
        $window.DialogResult = $true
    }
})

$btnCancel.Add_Click({
    $window.DialogResult = $false
})

$window.ShowDialog() | Out-Null

if ($script:picked -eq $null) {
    exit 0
}

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $explorer = $ol.ActiveExplorer()
    if ($null -eq $explorer) {
        [System.Windows.MessageBox]::Show("Outlook is not open or has no active window. Open Outlook and go to Calendar first.", "Outlook Date Jump", "OK", "Warning") | Out-Null
        exit 1
    }
    $view = $explorer.CurrentView
    if ($null -eq $view) {
        [System.Windows.MessageBox]::Show("No active view. Switch to the Calendar view and try again.", "Outlook Date Jump", "OK", "Warning") | Out-Null
        exit 1
    }
    $view.GoToDate($script:picked)
}
catch {
    [System.Windows.MessageBox]::Show("Failed: $($_.Exception.Message)", "Outlook Date Jump", "OK", "Error") | Out-Null
}
