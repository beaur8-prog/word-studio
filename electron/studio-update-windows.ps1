param([Parameter(Mandatory=$true)][string]$JobFile)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$job = $null
$readyFile = "$JobFile.ready"
try {
  Write-Output 'Studio update helper started.'
  if (-not (Test-Path -LiteralPath $JobFile)) { Write-Output 'Update request no longer exists.'; return }
  $job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($job.schema -ne 1 -or $job.id -notin @('sheet','slide','word') -or $job.sha256 -notmatch '^[a-f0-9]{64}$' -or $job.version -notmatch '^\d+\.\d+\.\d+$') { throw 'Invalid update job.' }
  $directory = [IO.Path]::GetFullPath((Split-Path -Parent $JobFile))
  $installer = [IO.Path]::GetFullPath($job.installer)
  if ((Split-Path -Parent $installer) -ne $directory -or [IO.Path]::GetExtension($installer) -ne '.exe') { throw 'Installer is outside the verified update directory.' }
  # Confirm startup before the editor requests a normal, cancellable quit.
  # No installer runs until the parent and its renderer processes have exited.
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value 'Helper ready; waiting for editor exit.'
  @{ schema = 1; id = $job.id; parentPid = $job.parentPid } | ConvertTo-Json -Compress | Set-Content -LiteralPath $readyFile -Encoding UTF8
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
  $arguments = if ($job.id -eq 'word') { @('/currentuser') } elseif ($job.id -eq 'slide') { @('/S') } else { @('/S','/currentuser') }
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value 'Starting verified installer.'
  $process = Start-Process -FilePath $installer -ArgumentList $arguments -PassThru
  if (-not $process.WaitForExit(600000)) { throw 'Update installer timed out.' }
  if ($process.ExitCode -notin @(0,3010)) { throw "Update installer failed: $($process.ExitCode)" }
  # Refuse a success report if the installer did not replace the expected app.
  $installedVersion = (Get-Item -LiteralPath $job.executable).VersionInfo.ProductVersion
  if ($installedVersion -ne $job.version) { throw "Installed version mismatch: $installedVersion" }
  Add-Content -LiteralPath $job.log -Encoding UTF8 -Value "Installed $($job.id) $($job.version) at $([DateTime]::UtcNow.ToString('o'))"
  Start-Process -FilePath $job.executable
  Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
} catch {
  if ($null -ne $job -and $job.PSObject.Properties['log']) { Add-Content -LiteralPath $job.log -Encoding UTF8 -Value "Update failed: $($_.Exception.Message)" }
  if ($null -ne $job -and $job.id -eq 'word' -and (Test-Path -LiteralPath $JobFile)) {
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
