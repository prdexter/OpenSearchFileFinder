# Run in an Administrator PowerShell window. Repairs only directories listed as
# access-denied in the recorded Active research backup; never deletes files.
param(
    [string]$LogPath = "$env:LOCALAPPDATA\OpenSearch\backup_logs\active_research_repair.log",
    [string]$BackupAccount = 'DESKTOP-51AIMCJ\Paul Dexter'
)
$ErrorActionPreference = 'Stop'
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Open PowerShell as Administrator, then run this script again.'
}
$sourceRoot = 'D:\Active research\'
$directories = @(Get-Content -LiteralPath $LogPath | ForEach-Object {
    if ($_ -match 'ERROR 5 .*Copying Directory (D:\\Active research\\.*)') {
        $matches[1].TrimEnd('\')
    }
} | Sort-Object -Unique)
if ($directories.Count -eq 0) { throw 'No access-denied source directories found in the log.' }
foreach ($directory in $directories) {
    $resolved = [IO.Path]::GetFullPath($directory)
    if (-not $resolved.StartsWith($sourceRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Directory outside backup source: $resolved"
    }
    # Explicitly visit descendants: some test directories disable inheritance.
    # /L and /SKIPSL keep permission repair from following symbolic links.
    & icacls.exe $resolved /grant "${BackupAccount}:(OI)(CI)(RX)" /T /C /L /Q
    if ($LASTEXITCODE -ne 0) {
        & takeown.exe /F $resolved /A /R /D Y /SKIPSL
        if ($LASTEXITCODE -ne 0) { throw "Cannot take ownership of $resolved" }
        & icacls.exe $resolved /grant "${BackupAccount}:(OI)(CI)(RX)" /T /C /L /Q
        if ($LASTEXITCODE -ne 0) { throw "Cannot grant backup read access to $resolved" }
    }
}
Write-Host "Repaired read access on $($directories.Count) directories. Rerun the backup to verify."
