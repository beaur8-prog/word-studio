param(
  [Parameter(Mandatory=$true)][string]$Setup,
  [Parameter(Mandatory=$true)][string]$Stage,
  [switch]$ExpectLanguageSelector
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') {
  throw 'Run native installer UI checks only on an isolated Windows CI runner.'
}
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class StudioSetupUI {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr data);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hwnd, EnumProc proc, IntPtr data);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
  [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr hwnd, int id);
  [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam);
  public static string Text(IntPtr hwnd) { var s = new StringBuilder(2048); GetWindowText(hwnd, s, s.Capacity); return s.ToString(); }
  public static string Class(IntPtr hwnd) { var s = new StringBuilder(256); GetClassName(hwnd, s, s.Capacity); return s.ToString(); }
  public static IntPtr[] Windows(HashSet<int> ids) {
    var result = new List<IntPtr>();
    EnumWindows((hwnd, data) => { int pid; GetWindowThreadProcessId(hwnd, out pid); if(ids.Contains(pid) && IsWindowVisible(hwnd)) result.Add(hwnd); return true; }, IntPtr.Zero);
    return result.ToArray();
  }
  public static IntPtr[] Children(IntPtr hwnd) {
    var result = new List<IntPtr>();
    EnumChildWindows(hwnd, (child, data) => { if(IsWindowVisible(child)) result.Add(child); return true; }, IntPtr.Zero);
    return result.ToArray();
  }
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam, uint flags, uint timeout, out IntPtr result);
  public static bool Checked(IntPtr hwnd) { IntPtr result; if(SendMessageTimeout(hwnd, 0x00F0, IntPtr.Zero, IntPtr.Zero, 2, 2000, out result) == IntPtr.Zero) throw new Exception("Cannot read installer checkbox"); return result.ToInt64() == 1; }
  public static void Click(IntPtr hwnd) { if(hwnd == IntPtr.Zero || !PostMessage(hwnd, 0x00F5, IntPtr.Zero, IntPtr.Zero)) throw new Exception("Cannot click installer control"); }
}
'@
$root = Split-Path $PSScriptRoot -Parent
$evidence = Join-Path $root '.cache'
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
# This is also the updater's real /currentuser invocation, with visible progress.
$process = Start-Process -FilePath (Resolve-Path $Setup) -ArgumentList '/currentuser' -PassThru
$ids = [Collections.Generic.HashSet[int]]::new()
[void]$ids.Add($process.Id)
$deadline = [DateTime]::UtcNow.AddSeconds(180)
$installed = $false
$finished = $false
$selector = $false
$captured = $false
$passed = $false
try {
  while ([DateTime]::UtcNow -lt $deadline) {
    # NSIS can run from a temporary child process; include descendants, not other apps.
    $tree = @(Get-CimInstance Win32_Process | Select-Object ProcessId, ParentProcessId)
    do {
      $added = $false
      foreach ($item in $tree) {
        if ($ids.Contains([int]$item.ParentProcessId) -and $ids.Add([int]$item.ProcessId)) { $added = $true }
      }
    } while ($added)
    foreach ($window in [StudioSetupUI]::Windows($ids)) {
      $children = @([StudioSetupUI]::Children($window))
      if (@($children | Where-Object { [StudioSetupUI]::Class($_) -eq 'ComboBox' }).Count) {
        $selector = $true
        if (-not $ExpectLanguageSelector) { throw 'Unexpected language selector during native Setup.' }
      }
      if (-not $captured -and ($selector -or [StudioSetupUI]::Text($window) -like '*Word Studio*')) {
        $image = [Drawing.Bitmap]::new(1200, 900)
        $graphics = [Drawing.Graphics]::FromImage($image)
        try { $graphics.CopyFromScreen(0, 0, 0, 0, $image.Size); $image.Save((Join-Path $evidence "installer-language-$Stage.png")) }
        finally { $graphics.Dispose(); $image.Dispose() }
        $captured = $true
      }
      if ($ExpectLanguageSelector -and $selector) {
        [StudioSetupUI]::Click([StudioSetupUI]::GetDlgItem($window, 2))
        $passed = $true
        break
      }
      if ([StudioSetupUI]::Text($window) -notlike '*Word Studio*') { continue }
      $next = [StudioSetupUI]::GetDlgItem($window, 1)
      if (-not [StudioSetupUI]::IsWindowEnabled($next)) { continue }
      $label = [StudioSetupUI]::Text($next).Replace('&','')
      if (-not $installed -and $label -match '^(Install|설치)') {
        [StudioSetupUI]::Click($next)
        $installed = $true
      } elseif ($installed -and -not $finished -and $label -match '^(Finish|마침)') {
        # Keep this installer test independent of normal editor smoke tests.
        foreach ($control in $children) {
          if ([StudioSetupUI]::Class($control) -eq 'Button' -and
              [StudioSetupUI]::Text($control) -match '(Run|실행)' -and [StudioSetupUI]::Checked($control)) {
            [StudioSetupUI]::Click($control)
          }
        }
        [StudioSetupUI]::Click($next)
        $finished = $true
      }
    }
    if ($ExpectLanguageSelector -and $passed) { break }
    if ($finished -and @($tree | Where-Object { $ids.Contains([int]$_.ProcessId) }).Count -eq 0) {
      if ($process.ExitCode -ne 0) { throw "Setup exited with $($process.ExitCode)" }
      $passed = $true
      break
    }
    if (-not $finished -and @($tree | Where-Object { $ids.Contains([int]$_.ProcessId) }).Count -eq 0) { throw 'Setup exited before completing its wizard.' }
    Start-Sleep -Milliseconds 250
  }
  if (-not $passed) { throw 'Timed out waiting for native Setup UI.' }
  if ($ExpectLanguageSelector -and -not $selector) { throw 'Positive control did not reproduce the language selector.' }
} finally {
  foreach ($id in $ids) {
    $remaining = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($null -ne $remaining) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
  }
  [ordered]@{ stage=$Stage; passed=$passed; languageSelector=$selector; installed=$installed; finished=$finished } |
    ConvertTo-Json | Set-Content (Join-Path $evidence "installer-language-$Stage.json")
}
