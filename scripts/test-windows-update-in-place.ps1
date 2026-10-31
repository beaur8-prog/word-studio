$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') { throw 'Requires an isolated Windows CI runner.' }
$root = Split-Path $PSScriptRoot -Parent
$version = (Get-Content (Join-Path $root 'package.json') | ConvertFrom-Json).version
$setup = Join-Path $root "release/installer/Word-Studio-$version-windows-x64-setup.exe"
$legacySetup = Join-Path $root '.cache/legacy-update/Word-Studio-1.4.12-windows-x64-setup.exe'
$ps51 = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$env:STUDIO_DISABLE_UPDATES = '1'
$report = [ordered]@{ passed=$false; scenarios=@() }
$work = Join-Path $env:RUNNER_TEMP "Word Studio update O'Brien; & 한글"
New-Item -ItemType Directory -Path $work -Force | Out-Null
$recovery = Join-Path $env:APPDATA 'Word Studio/recovery/in-place-preservation.txt'
$document = Join-Path $env:USERPROFILE 'Documents/word-update-preservation.txt'
New-Item -ItemType Directory -Path (Split-Path $recovery -Parent), (Split-Path $document -Parent) -Force | Out-Null
[IO.File]::WriteAllText($recovery,'PRESERVE RECOVERY')
[IO.File]::WriteAllText($document,'PRESERVE DOCUMENT')

function Run-Setup([string]$File, [string]$Scope, [string]$Directory) {
  $process = Start-Process -FilePath $File -ArgumentList @('/S',$Scope,('/D=' + $Directory)) -PassThru
  $null = $process.Handle
  if (-not $process.WaitForExit(180000)) { throw 'Seed Setup timed out.' }
  $process.Refresh()
  if ($process.ExitCode -ne 0) { throw "Seed Setup failed: $($process.ExitCode)" }
}
function Entries {
  foreach ($hive in @('HKCU','HKLM')) {
    foreach ($entry in @(Get-ItemProperty "$($hive):/Software/Microsoft/Windows/CurrentVersion/Uninstall/*" -ErrorAction SilentlyContinue)) {
      if ($entry.PSObject.Properties['DisplayName'] -and $entry.DisplayName -eq 'Word Studio') { $entry }
    }
  }
}
function Desktop-Links {
  foreach ($folder in @([Environment]::GetFolderPath('DesktopDirectory'),[Environment]::GetFolderPath('CommonDesktopDirectory'))) {
    $file = Join-Path $folder 'Word Studio.lnk'
    if (Test-Path $file) { $file }
  }
}
function Assert-Preserved {
  if ([IO.File]::ReadAllText($recovery) -ne 'PRESERVE RECOVERY' -or [IO.File]::ReadAllText($document) -ne 'PRESERVE DOCUMENT') { throw 'Document or recovery changed.' }
}
function Stop-TestApp {
  # These are disposable CI installations, never a user's desktop.
  Get-Process -Name 'Word Studio' -ErrorAction SilentlyContinue | Stop-Process -Force
}
function Remove-TestInstall([string]$Directory,[string]$Scope) {
  Stop-TestApp
  $uninstall = Join-Path $Directory 'Uninstall Word Studio.exe'
  $process = Start-Process -FilePath $uninstall -ArgumentList @('/S',$Scope) -PassThru
  if (-not $process.WaitForExit(120000)) { throw 'CI uninstall timed out.' }
  $deadline = [DateTime]::UtcNow.AddSeconds(30)
  while (Test-Path (Join-Path $Directory 'Word Studio.exe')) {
    if ([DateTime]::UtcNow -gt $deadline) { throw 'CI uninstall left its executable.' }
    Start-Sleep -Milliseconds 250
  }
}
if (@(Entries).Count) { throw 'An earlier test left a Studio installation on this runner.' }
try {
  foreach ($scenario in @('current-user','all-users','legacy-helper-with-duplicates')) {
    $scope = if ($scenario -eq 'current-user') { '/currentuser' } else { '/allusers' }
    $directory = Join-Path $work "$scenario/Word Studio 사용자 지정"
    Run-Setup $legacySetup $scope $directory
    $target = Join-Path $directory 'Word Studio.exe'
    if (-not (Test-Path $target)) { throw 'Legacy setup ignored the custom directory.' }
    if ($scenario -eq 'legacy-helper-with-duplicates') {
      # Reproduce the old helper's forced /currentuser against an all-user install.
      Run-Setup $legacySetup '/currentuser' (Join-Path $env:LOCALAPPDATA 'Programs/Word Studio')
      if (@(Entries).Count -ne 2 -or @(Desktop-Links).Count -ne 2) { throw 'The old installation-scope defect was not reproduced.' }
    }
    $jobDirectory = Join-Path $work "$scenario/job"
    New-Item -ItemType Directory -Path $jobDirectory -Force | Out-Null
    $installer = Join-Path $jobDirectory 'update.exe'
    Copy-Item -LiteralPath $setup -Destination $installer
    $log = Join-Path $jobDirectory 'install.log'
    # A short-lived disposable parent exercises the real PowerShell 5.1 wait.
    $sleep = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes('Start-Sleep -Seconds 4'))
    $parent = Start-Process -FilePath $ps51 -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$sleep) -WindowStyle Hidden -PassThru
    $jobFile = Join-Path $jobDirectory 'install.json'
    @{ schema=1; id='word'; parentPid=$parent.Id; installer=$installer; bytes=(Get-Item $installer).Length;
       sha256=(Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant(); executable=$target; version=$version; log=$log } |
      ConvertTo-Json | Set-Content -LiteralPath $jobFile -Encoding UTF8
    $helper = if ($scenario -eq 'legacy-helper-with-duplicates') { Join-Path $root '.cache/legacy-update/helper-1.4.12.ps1' } else { Join-Path $root 'electron/studio-update-windows.ps1' }
    # Match the product loader's explicit UTF-8 decoding in PowerShell 5.1.
    $workerScript = Join-Path $jobDirectory 'helper.ps1'
    [IO.File]::WriteAllText($workerScript, (Get-Content -LiteralPath $helper -Raw -Encoding UTF8), [Text.UTF8Encoding]::new($true))
    $command = "& ([ScriptBlock]::Create([IO.File]::ReadAllText('" + $workerScript.Replace("'","''") + "', [Text.Encoding]::UTF8))) -JobFile '" + $jobFile.Replace("'","''") + "'"
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $worker = Start-Process -FilePath $ps51 -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -RedirectStandardOutput (Join-Path $jobDirectory 'stdout.log') -RedirectStandardError (Join-Path $jobDirectory 'stderr.log') -PassThru
    $null = $worker.Handle
    $worker.Refresh()
    if ($worker.HasExited) { throw ("Helper exited before Setup: " + (Get-Content (Join-Path $jobDirectory 'stderr.log') -Raw)) }
    & (Join-Path $PSScriptRoot 'test-windows-installer-language.ps1') -Setup $setup -Stage $scenario -ObserveProcessId $worker.Id
    if (-not $worker.WaitForExit(120000) -or $worker.ExitCode -ne 0) { throw "Actual helper failed: $(Get-Content $log -Raw)" }
    $text = Get-Content $log -Raw
    if ($text -match 'Update failed:' -or $text -notmatch "Installed word $([regex]::Escape($version))") { throw "Unexpected update result: $text" }
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
      $main = @(Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'Word Studio.exe' -and $_.ExecutablePath -ieq $target -and $_.CommandLine -notmatch '--type=' })
      if ($main.Count -eq 1) { break }
      Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($main.Count -ne 1) { throw "Expected one relaunched editor, found $($main.Count)." }
    if ((Get-Item -LiteralPath $target).VersionInfo.ProductVersion -ne $version) { throw 'The original installation was not updated.' }
    $entries = @(Entries)
    if ($entries.Count -ne 1 -or $entries[0].UninstallString -ine ('"' + (Join-Path $directory 'Uninstall Word Studio.exe') + '" ' + $scope)) { throw 'Update created another installation or moved the existing one.' }
    $links = @(Desktop-Links)
    if ($links.Count -ne 1) { throw "Expected one desktop shortcut, found $($links.Count)." }
    $shortcutTarget = (New-Object -ComObject WScript.Shell).CreateShortcut($links[0]).TargetPath
    $report.current = @{ scenario=$scenario; expectedTarget=$target; shortcutTarget=$shortcutTarget; entries=$entries.Count; mainProcesses=$main.Count }
    # Shell links can return an 8.3 path. Compare NTFS file identities, not path spelling.
    $identity = & node -e 'const fs=require("node:fs"); const a=fs.statSync(process.argv[1],{bigint:true}),b=fs.statSync(process.argv[2],{bigint:true}); const same=a.dev===b.dev&&a.ino===b.ino; console.log(JSON.stringify({same,shortcut:fs.realpathSync.native(process.argv[1]),expected:fs.realpathSync.native(process.argv[2])})); process.exit(same?0:1);' $shortcutTarget $target
    if ($LASTEXITCODE -ne 0) { throw "Desktop shortcut points elsewhere: $shortcutTarget (expected $target)" }
    $report.current.identity = $identity | ConvertFrom-Json
    Assert-Preserved
    $report.scenarios += [ordered]@{ scenario=$scenario; passed=$true; originalDirectoryPreserved=$true; originalScopePreserved=$true;
      installedVersion=$version; desktopShortcuts=$links.Count; mainProcesses=$main.Count; noFailurePopup=$true; documentAndRecoveryPreserved=$true }
    Copy-Item -LiteralPath $log -Destination (Join-Path $root ".cache/update-in-place-$scenario.log")
    Remove-TestInstall $directory $scope
    Assert-Preserved
    if (@(Entries).Count) { throw 'Uninstall left an extra installed copy.' }
  }
  $report.passed = $true
} finally {
  foreach ($file in @(Get-ChildItem -LiteralPath $work -Filter '*.log' -Recurse -File)) {
    $name = $file.FullName.Substring($work.Length + 1).Replace('\','-')
    Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $root ('.cache/update-in-place-' + $name))
  }
  Stop-TestApp
  $report | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $root '.cache/update-in-place.json')
  $report | ConvertTo-Json -Depth 8 | Write-Output
}
