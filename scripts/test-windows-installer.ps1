$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ($env:CI -ne 'true' -or $env:RUNNER_OS -ne 'Windows') {
  throw 'Run this install/reinstall/uninstall test only on an isolated Windows CI runner.'
}

$root = Split-Path $PSScriptRoot -Parent
$version = (Get-Content (Join-Path $root 'package.json') | ConvertFrom-Json).version
$setup = Join-Path $root "release/installer/Word-Studio-$version-windows-x64-setup.exe"
$installDir = Join-Path $env:LOCALAPPDATA 'Programs/Word Studio'
$executable = Join-Path $installDir 'Word Studio.exe'
$profile = Join-Path $env:APPDATA 'Word Studio'
$recovery = Join-Path $profile 'recovery/installer-preservation-test.wordstudio'
$document = Join-Path $env:USERPROFILE 'Documents/installer-preservation-test.wordstudio'
$desktopLink = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Word Studio.lnk'
$startLink = Join-Path ([Environment]::GetFolderPath('StartMenu')) 'Programs/Word Studio.lnk'
$reportPath = Join-Path $root '.cache/installer-test.json'
$report = [ordered]@{ version = $version; passed = $false }
$extensions = @((& node -e 'console.log(JSON.stringify([...require(process.argv[1]).OPEN_EXTENSIONS]))' (Join-Path $root 'electron/document-types.cjs')) | ConvertFrom-Json)
if ($LASTEXITCODE -ne 0 -or $extensions.Count -eq 0) { throw 'Could not read the supported document types.' }
$classes = 'HKCU:/Software/Classes'
$capabilities = 'HKCU:/Software/WordStudio/Capabilities'
$application = "$classes/Applications/Word Studio.exe"
$existingHandler = 'WordStudio.AssociationTest.Existing'
$userChoices = @{}

function Registry-Value([string] $key, [string] $name) {
  $item = Get-Item -LiteralPath $key -ErrorAction SilentlyContinue
  if ($null -ne $item) { return $item.GetValue($name, $null) }
  return $null
}

function Get-WordStudioUninstallEntries {
  $key = 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall'
  if (Test-Path $key) {
    Get-ChildItem $key -ErrorAction SilentlyContinue | Get-ItemProperty -ErrorAction SilentlyContinue |
      Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq 'Word Studio' }
  }
}

function Assert-OtherAssociations {
  foreach ($ext in $extensions) {
    if ((Registry-Value "$classes/.$ext" '') -ne $existingHandler) { throw "Installation changed the .$ext default." }
    $other = Get-Item -LiteralPath "$classes/.$ext/OpenWithProgids"
    if ($other.GetValueNames() -notcontains $existingHandler) { throw "Installation removed another .$ext handler." }
    $choice = "HKCU:/Software/Microsoft/Windows/CurrentVersion/Explorer/FileExts/.$ext/UserChoice"
    foreach ($name in @('ProgId', 'Hash')) {
      if ((Registry-Value $choice $name) -ne $userChoices[$ext][$name]) { throw "Installation changed the .$ext user choice." }
    }
  }
}

function Assert-DocumentHandlers {
  if ((Registry-Value 'HKCU:/Software/RegisteredApplications' 'Word Studio') -ne 'Software\WordStudio\Capabilities') { throw 'Missing Default Apps registration.' }
  if ((Registry-Value $capabilities 'ApplicationName') -ne 'Word Studio' -or -not (Registry-Value $capabilities 'ApplicationDescription')) { throw 'Missing application capabilities.' }
  $command = '"' + $executable + '" "%1"'
  if ((Registry-Value "$application/shell/open/command" '') -ne $command) { throw 'The executable or document path is not quoted.' }
  foreach ($ext in $extensions) {
    $progId = "WordStudio.Document.$ext"
    $openWith = Get-Item -LiteralPath "$classes/.$ext/OpenWithProgids"
    if ($openWith.GetValueNames() -notcontains $progId) { throw "Missing .$ext Open With candidate." }
    if ((Registry-Value "$capabilities/FileAssociations" ".$ext") -ne $progId) { throw "Missing .$ext Default Apps capability." }
    if ((Get-Item -LiteralPath "$application/SupportedTypes").GetValueNames() -notcontains ".$ext") { throw "Missing .$ext supported type." }
    if ((Registry-Value "$classes/$progId/shell/open/command" '') -ne $command) { throw "Incorrect .$ext document-open command." }
    if (-not (Registry-Value "$classes/$progId/DefaultIcon" '')) { throw "Missing .$ext icon." }
  }
  Assert-OtherAssociations
}

function Run-Installer([string] $file, [string[]] $arguments) {
  $process = Start-Process -FilePath $file -ArgumentList $arguments -PassThru
  if (-not $process.WaitForExit(180000)) {
    $process.Kill($true)
    throw "Installer timed out: $file"
  }
  if ($process.ExitCode -ne 0) { throw "Installer exit code: $($process.ExitCode)" }
}

function Assert-Installed {
  foreach ($file in @($executable, (Join-Path $installDir 'resources/app.asar'), $desktopLink, $startLink)) {
    if (-not (Test-Path $file)) { throw "Missing installed file: $file" }
  }
  $actual = (Get-Item $executable).VersionInfo.ProductVersion
  if ($actual -ne $version) { throw "Installed version $actual does not match $version" }
  $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($startLink)
  if ($shortcut.TargetPath -ne $executable) { throw 'Start menu shortcut points to another executable.' }
  $entry = @(Get-ItemProperty 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/*' |
    Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -eq 'Word Studio' })
  if ($entry.Count -ne 1 -or $entry[0].DisplayVersion -ne $version) {
    throw 'Expected one current-user installed-app entry with the release version.'
  }
  Assert-DocumentHandlers
}

function Assert-Preserved {
  foreach ($file in @($recovery, $document)) {
    if (-not (Test-Path $file) -or (Get-Content $file -Raw).Trim() -ne 'PRESERVE WORD STUDIO USER DATA') {
      throw "User data was removed or changed: $file"
    }
  }
}

if (Test-Path $installDir) { throw 'This runner already has Word Studio installed.' }
try {
  # Seed a different candidate and extension default on this disposable runner.
  # Install/reinstall/remove must leave both and Explorer's existing choice intact.
  foreach ($ext in $extensions) {
    $choice = "HKCU:/Software/Microsoft/Windows/CurrentVersion/Explorer/FileExts/.$ext/UserChoice"
    $userChoices[$ext] = @{ ProgId = (Registry-Value $choice 'ProgId'); Hash = (Registry-Value $choice 'Hash') }
    New-Item -Path "$classes/.$ext/OpenWithProgids" -Force | Out-Null
    Set-Item -LiteralPath "$classes/.$ext" -Value $existingHandler
    New-ItemProperty -LiteralPath "$classes/.$ext/OpenWithProgids" -Name $existingHandler -Value '' -PropertyType String -Force | Out-Null
  }
  New-Item -ItemType Directory -Path (Split-Path $recovery -Parent), (Split-Path $document -Parent) -Force | Out-Null
  Set-Content $recovery 'PRESERVE WORD STUDIO USER DATA'
  Set-Content $document 'PRESERVE WORD STUDIO USER DATA'
  & (Join-Path $PSScriptRoot 'test-windows-installer-language.ps1') -Setup $setup -Stage 'fresh'
  Assert-Installed
  Assert-Preserved
  # Fresh Setup owns its initial launch; verify it before isolated app tests.
  $deadline = [DateTime]::UtcNow.AddSeconds(20)
  do {
    $started = @(Get-Process -Name 'Word Studio' -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($started.Count -eq 1) { break }
    Start-Sleep -Milliseconds 250
  } while ([DateTime]::UtcNow -lt $deadline)
  if ($started.Count -ne 1) { throw 'Fresh Setup did not open exactly one editor window.' }
  $report.freshSetupRunAfterFinish = $true
  # Disposable blank CI window, never a user installation or document.
  $freshApps = @(Get-Process -Name 'Word Studio' -ErrorAction SilentlyContinue)
  $freshApps | Stop-Process -Force -ErrorAction SilentlyContinue
  Wait-Process -InputObject $freshApps -Timeout 15 -ErrorAction SilentlyContinue
  $report.freshInstall = $true
  $report.shortcuts = $true
  $report.installedAppsEntry = $true
  $report.registeredDocumentTypes = $extensions
  $report.openWithAndDefaultApps = $true
  $report.existingDefaultsPreserved = $true

  $env:WORD_STUDIO_EXECUTABLE = $executable
  & node (Join-Path $root 'scripts/desktop-smoke.cjs')
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows application smoke test failed.' }
  $report.installedApplication = $true
  & node (Join-Path $root 'scripts/hwp-desktop-smoke.cjs')
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows HWP reader test failed.' }
  $report.hwpReader = $true
  & node (Join-Path $root 'scripts/hwp-editing-desktop-smoke.cjs')
  if ($LASTEXITCODE -ne 0) { throw 'Installed Windows HWP editing test failed.' }
  $report.hwpEditing = $true

  $deadline = [DateTime]::UtcNow.AddSeconds(15)
  while (@(Get-Process -Name 'Word Studio' -ErrorAction SilentlyContinue).Count -ne 0) {
    if ([DateTime]::UtcNow -gt $deadline) { throw 'Native app tests left an editor process running before reinstall.' }
    Start-Sleep -Milliseconds 250
  }
  & (Join-Path $PSScriptRoot 'test-windows-installer-language.ps1') -Setup $setup -Stage 'update'
  Assert-Installed
  Assert-Preserved
  $report.reinstall = $true
  $report.nativeSetupWithoutLanguageSelector = $true

  $uninstaller = @(Get-ChildItem $installDir -Filter '*Uninstall*.exe')
  if ($uninstaller.Count -ne 1) { throw 'Expected exactly one uninstaller.' }
  Run-Installer $uninstaller[0].FullName @('/S', '/currentuser')
  # NSIS runs its uninstaller from a temporary child process. The original
  # process can exit before shortcut and registry cleanup has finished.
  $deadline = (Get-Date).AddSeconds(30)
  do {
    $leftoverFiles = @(@($executable, $desktopLink, $startLink) | Where-Object { Test-Path $_ })
    $remaining = @(Get-WordStudioUninstallEntries)
    if ($leftoverFiles.Count -eq 0 -and $remaining.Count -eq 0) { break }
    Start-Sleep -Milliseconds 250
  } while ((Get-Date) -lt $deadline)
  foreach ($file in @($executable, $desktopLink, $startLink)) {
    if (Test-Path $file) { throw "Uninstall left an application file or shortcut: $file" }
  }
  $remaining = @(Get-WordStudioUninstallEntries)
  if ($remaining.Count -ne 0) { throw 'Uninstall left the installed-app registration.' }
  foreach ($ext in $extensions) {
    if (Test-Path "$classes/WordStudio.Document.$ext") { throw "Uninstall left the .$ext ProgID." }
    if ((Get-Item -LiteralPath "$classes/.$ext/OpenWithProgids").GetValueNames() -contains "WordStudio.Document.$ext") { throw "Uninstall left the .$ext Open With registration." }
  }
  if ((Registry-Value 'HKCU:/Software/RegisteredApplications' 'Word Studio') -or (Test-Path $capabilities) -or (Test-Path $application)) { throw 'Uninstall left application file capabilities.' }
  Assert-OtherAssociations
  $report.handlerUnregistration = $true
  Assert-Preserved
  $report.uninstall = $true
  $report.userDataPreserved = $true
  $report.passed = $true
} finally {
  New-Item -ItemType Directory -Path (Split-Path $reportPath -Parent) -Force | Out-Null
  $report | ConvertTo-Json | Set-Content $reportPath
  $report | ConvertTo-Json | Write-Output
}
