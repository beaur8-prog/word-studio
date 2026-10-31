param([Parameter(Mandatory=$true)][string]$JobFile)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
try {
  if ($job.schema -ne 1 -or $job.id -notin @('sheet','slide','word') -or $job.sha256 -notmatch '^[a-f0-9]{64}$' -or $job.version -notmatch '^\d+\.\d+\.\d+$') { throw 'Invalid update job.' }
  $directory = [IO.Path]::GetFullPath((Split-Path -Parent $JobFile))
  $installer = [IO.Path]::GetFullPath($job.installer)
  if ((Split-Path -Parent $installer) -ne $directory -or [IO.Path]::GetExtension($installer) -ne '.exe') { throw 'Installer is outside the verified update directory.' }
  # Never terminate the editor. The helper is started only on Electron's quit
  # event, after all editor windows and their save/recovery guards have closed.
  $parent = Get-Process -Id $job.parentPid -ErrorAction SilentlyContinue
  if ($null -ne $parent -and -not $parent.WaitForExit(120000)) { throw 'The editor has not exited.' }
  $names = @{ sheet = 'Sheet Studio'; slide = 'Slide Studio'; word = 'Word Studio' }
  if (@(Get-Process -Name ($names[$job.id]) -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Another editor instance is still running. Close it and try the update again.' }
  if ((Get-Item -LiteralPath $installer).Length -ne $job.bytes -or (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant() -ne $job.sha256) { throw 'Installer checksum mismatch.' }
  # Word shows Setup's own installation progress instead of installing invisibly.
  $arguments = if ($job.id -eq 'word') { @('/currentuser') } elseif ($job.id -eq 'slide') { @('/S') } else { @('/S','/currentuser') }
  $process = Start-Process -FilePath $installer -ArgumentList $arguments -PassThru
  if (-not $process.WaitForExit(600000)) { throw 'Update installer timed out.' }
  if ($process.ExitCode -notin @(0,3010)) { throw "Update installer failed: $($process.ExitCode)" }
  # Refuse a success report if the installer did not replace the expected app.
  $installedVersion = (Get-Item -LiteralPath $job.executable).VersionInfo.ProductVersion
  if ($installedVersion -ne $job.version) { throw "Installed version mismatch: $installedVersion" }
  Set-Content -LiteralPath $job.log -Value "Installed $($job.id) $($job.version) at $([DateTime]::UtcNow.ToString('o'))"
  Start-Process -FilePath $job.executable
  Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
} catch {
  if ($job.PSObject.Properties['log']) { Set-Content -LiteralPath $job.log -Value "Update failed: $($_.Exception.Message)" }
  if ($job.id -eq 'word') {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show('설치가 취소되었거나 업데이트를 완료하지 못했습니다. Word Studio를 다시 실행한 뒤 업데이트 메뉴에서 재시도해 주세요.', 'Word Studio 업데이트') | Out-Null
  }
  exit 1
} finally {
  Remove-Item -LiteralPath $JobFile -Force -ErrorAction SilentlyContinue
}
