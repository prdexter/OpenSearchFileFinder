# One-time administrator setup. The task only repairs local backup read access.
$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this installer in Administrator PowerShell.'
}
if ($identity.Name -ne 'DESKTOP-51AIMCJ\Paul Dexter') {
    throw 'Run as Paul Dexter so the helper belongs to the same account as the backup app.'
}
$installDirectory = Join-Path $env:ProgramFiles 'OpenSearchBackupHelper'
$stateDirectory = Join-Path $env:LOCALAPPDATA 'OpenSearch\permission_repair'
New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $stateDirectory -Force | Out-Null
if ((Get-Item -LiteralPath $installDirectory -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'Helper install directory must not be a reparse point'
}
# The elevated task must never load executable code writable by ordinary apps.
& icacls.exe $installDirectory /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)(F)' '*S-1-5-32-544:(OI)(CI)(F)' '*S-1-5-32-545:(OI)(CI)(RX)' /Q
if ($LASTEXITCODE -ne 0) { throw 'Cannot secure helper install directory' }
& icacls.exe $installDirectory /setowner '*S-1-5-32-544' /Q
if ($LASTEXITCODE -ne 0) { throw 'Cannot set administrator ownership on helper directory' }
$worker = Join-Path $installDirectory 'backup_permission_worker.ps1'
if ((Test-Path -LiteralPath $worker) -and ((Get-Item -LiteralPath $worker -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw 'Helper script must not be a reparse point'
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'backup_permission_worker.ps1') -Destination $worker -Force
& icacls.exe $worker /reset /Q
if ($LASTEXITCODE -ne 0) { throw 'Cannot secure installed helper script' }
& icacls.exe $worker /setowner '*S-1-5-32-544' /Q
if ($LASTEXITCODE -ne 0) { throw 'Cannot set administrator ownership on helper script' }
$arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -StateDirectory "{1}" -BackupAccountSid "{2}"' -f $worker,$stateDirectory,$identity.User.Value
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$action = New-ScheduledTaskAction -Execute $powershell -Argument $arguments
$taskPrincipal = New-ScheduledTaskPrincipal -UserId $identity.Name -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances Queue -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName 'OpenSearch-RepairActiveResearchPermissions' -Action $action -Principal $taskPrincipal -Settings $settings -Description 'On-demand read-permission repair restricted to D:\Active research for NAS backups.' -Force | Out-Null
# Leave a tiny protected fixture so Codex can verify the real elevated task from
# the ordinary backup account after setup. It is inside the excluded scratch area.
$probeRoot = 'D:\Active research\OpenSearch\scratch'
$probe = Join-Path $probeRoot ('permission-helper-probe-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $probe -Force | Out-Null
Set-Content -LiteralPath (Join-Path $probe 'sample.txt') -Value 'Automatic backup permission repair verification' -Encoding UTF8
$probeAcl = New-Object Security.AccessControl.DirectorySecurity
$probeAcl.SetAccessRuleProtection($true, $false)
$adminSid = New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')
$systemSid = New-Object Security.Principal.SecurityIdentifier('S-1-5-18')
$probeAcl.SetOwner($adminSid)
foreach ($sid in @($adminSid, $systemSid)) {
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $probeAcl.AddAccessRule($rule)
}
Set-Acl -LiteralPath $probe -AclObject $probeAcl
@{task_name='OpenSearch-RepairActiveResearchPermissions'; account=$identity.Name; verification_probe=$probe; installed_at=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $stateDirectory 'installed.json') -Encoding UTF8
Write-Host 'Automatic backup permission repair installed. Return to Codex for verification.'
