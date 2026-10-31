param([Parameter(Mandatory=$true)][string]$JobFile)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$job = $null
$acknowledged = $false
$readyFile = "$JobFile.ready"

function Get-WordInstallTarget([string]$Executable) {
  $target = [IO.Path]::GetFullPath($Executable)
  $folder = [IO.Path]::GetDirectoryName($target).TrimEnd('\')
  if ([IO.Path]::GetFileName($target) -ne 'Word Studio.exe' -or $folder -match '["\r\n]') { throw 'Invalid Word installation path.' }
  $scopes = @()
  foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
      $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
      try {
        $uninstall = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall')
        if ($null -eq $uninstall) { continue }
        try {
          foreach ($name in $uninstall.GetSubKeyNames()) {
            $entry = $uninstall.OpenSubKey($name)
            if ($null -eq $entry) { continue }
            try {
              if ($entry.GetValue('DisplayName') -ne 'Word Studio') { continue }
              $location = [string]$entry.GetValue('InstallLocation')
              # electron-builder stores InstallLocation under its app GUID key;
              # the standard uninstall entry instead supplies a quoted command.
              if ([string]::IsNullOrWhiteSpace($location)) {
                $command = [string]$entry.GetValue('UninstallString')
                $match = [regex]::Match($command, '^"([^"\r\n]+)"(?:\s|$)')
                if (-not $match.Success -or [IO.Path]::GetFileName($match.Groups[1].Value) -ne 'Uninstall Word Studio.exe') { continue }
                $location = [IO.Path]::GetDirectoryName($match.Groups[1].Value)
              }
              if ([IO.Path]::GetFullPath($location).TrimEnd('\') -ieq $folder) {
                $scopes += if ($hive -eq [Microsoft.Win32.RegistryHive]::LocalMachine) { '/allusers' } else { '/currentuser' }
              }
            } finally { $entry.Dispose() }
          }
        } finally { $uninstall.Dispose() }
      } finally { $base.Dispose() }
    }
  }
  $scopes = @($scopes | Select-Object -Unique)
  if ($scopes.Count -ne 1) { throw 'Cannot identify the running Word installation. Run Setup manually to repair the installation.' }
  return @{ directory = $folder; scope = $scopes[0] }
}
try {
  Write-Output 'Studio update helper started.'
  if (-not (Test-Path -LiteralPath $JobFile)) { Write-Output 'Update request no longer exists.'; return }
  $job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($job.schema -ne 1 -or $job.id -notin @('sheet','slide','word') -or $job.sha256 -notmatch '^[a-f0-9]{64}$' -or $job.version -notmatch '^\d+\.\d+\.\d+$') { throw 'Invalid update job.' }
  $directory = [IO.Path]::GetFullPath((Split-Path -Parent $JobFile))
  $installer = [IO.Path]::GetFullPath($job.installer)
  if ((Split-Path -Parent $installer) -ne $directory -or [IO.Path]::GetExtension($installer) -ne '.exe') { throw 'Installer is outside the verified update directory.' }
  # Resolve the running copy before acknowledging: never switch user scope.
  $wordTarget = if ($job.id -eq 'word') { Get-WordInstallTarget $job.executable } else { $null }
  # Confirm startup before the editor requests a normal, cancellable quit.
  # No installer runs until the parent and its renderer processes have exited.
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value 'Helper ready; waiting for editor exit.'
  @{ schema = 1; id = $job.id; parentPid = $job.parentPid } | ConvertTo-Json -Compress | Set-Content -LiteralPath $readyFile -Encoding UTF8
  $acknowledged = $true
  $parent = Get-Process -Id $job.parentPid -ErrorAction SilentlyContinue
  $deadline = [DateTime]::UtcNow.AddSeconds(120)
  while ($null -ne $parent -and -not $parent.WaitForExit(250)) {
    if (-not (Test-Path -LiteralPath $JobFile)) { return }
    if ([DateTime]::UtcNow -ge $deadline) { throw 'The editor has not exited.' }
  }
  if (-not (Test-Path -LiteralPath $JobFile)) { return }
  $names = @{ sheet = 'Sheet Studio'; slide = 'Slide Studio'; word = 'Word Studio' }
  $deadline = [DateTime]::UtcNow.AddSeconds(10)
  while (@(Get-Process -Name ($names[$job.id]) -ErrorAction SilentlyContinue).Count -ne 0 -and [DateTime]::UtcNow -lt $deadline) {
    if (-not (Test-Path -LiteralPath $JobFile)) { return }
    Start-Sleep -Milliseconds 250
  }
  if (@(Get-Process -Name ($names[$job.id]) -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Another editor instance is still running. Close it and try the update again.' }
  if (-not (Test-Path -LiteralPath $JobFile)) { return }
  # Use .NET directly: GUI-launched Windows PowerShell may inherit a module
  # search path that does not expose the Get-FileHash script cmdlet.
  $stream = [IO.File]::OpenRead($installer)
  $hasher = [Security.Cryptography.SHA256]::Create()
  try { $digest = [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-','').ToLowerInvariant() }
  finally { $stream.Dispose(); $hasher.Dispose() }
  if ((Get-Item -LiteralPath $installer).Length -ne $job.bytes -or $digest -ne $job.sha256) { throw 'Installer checksum mismatch.' }
  # Keep Word's visible Setup progress while using the shared startup handshake.
  # NSIS /D must be last and unquoted, including paths with spaces. Start-Process
  # invokes the executable directly; these arguments do not pass through cmd.exe.
  $arguments = if ($job.id -eq 'word') { @($wordTarget.scope, '--updated', ('/D=' + $wordTarget.directory)) } elseif ($job.id -eq 'slide') { @('/S') } else { @('/S','/currentuser') }
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value ('Starting verified installer: ' + ($arguments -join ' '))
  $process = Start-Process -FilePath $installer -ArgumentList $arguments -PassThru
  # Retain the native process handle before waiting (Windows PowerShell 5.1).
  $null = $process.Handle
  if (-not $process.WaitForExit(600000)) { throw 'Update installer timed out.' }
  $process.Refresh()
  if ($null -eq $process.ExitCode -or $process.ExitCode -notin @(0,3010)) { throw "Update installer failed: $($process.ExitCode)" }
  # Refuse a success report if the installer did not replace the expected app.
  $installedVersion = (Get-Item -LiteralPath $job.executable).VersionInfo.ProductVersion
  if ($installedVersion -ne $job.version) { throw "Installed version mismatch: $installedVersion" }
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value "Installed $($job.id) $($job.version) at $([DateTime]::UtcNow.ToString('o'))"
  Start-Process -FilePath $job.executable
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value 'Relaunched verified app once.'
  Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
} catch {
  if ($null -ne $job -and $job.PSObject.Properties['log']) { Add-Content -LiteralPath $job.log -Encoding UTF8 -Value "Update failed: $($_.Exception.Message)" }
  if ($acknowledged -and $null -ne $job -and $job.id -eq 'word' -and (Test-Path -LiteralPath $JobFile)) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show('설치가 취소되었거나 업데이트를 완료하지 못했습니다. Word Studio를 다시 실행한 뒤 업데이트 메뉴에서 재시도하거나 내려받은 Setup을 직접 실행해 주세요.' + [Environment]::NewLine + '로그: ' + $job.log, 'Word Studio 업데이트') | Out-Null
  }
  Write-Error $_ -ErrorAction Continue
  exit 1
} finally {
  Remove-Item -LiteralPath $JobFile -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $readyFile -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath "$JobFile.ps1" -Force -ErrorAction SilentlyContinue
}
