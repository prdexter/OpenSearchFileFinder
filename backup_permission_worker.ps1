# Installed into an administrator-owned directory; invoked only by the fixed task.
param([Parameter(Mandatory)][string]$StateDirectory,
      [Parameter(Mandatory)][string]$BackupAccountSid)
$ErrorActionPreference = 'Stop'
$sourceRoot = 'D:\Active research\'
function Invoke-LoggedNative([string]$Executable, [string[]]$NativeArguments, [string]$LogPath) {
    # Windows PowerShell treats redirected native stderr as an ErrorRecord.
    # Inspect the exit code so an access denial can reach the ownership fallback.
    $ErrorActionPreference = 'Continue'
    & $Executable @NativeArguments 2>&1 | Out-File -LiteralPath $LogPath -Append
    return $LASTEXITCODE
}
function Assert-SourcePath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($sourceRoot, [StringComparison]::OrdinalIgnoreCase) -or
        $full.Substring(3).Contains(':')) { throw "Path outside backup source: $Path" }
    $current = $full
    while ($current -and $current.Length -ge $sourceRoot.TrimEnd('\').Length) {
        $item = Get-Item -LiteralPath $current -Force
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Refusing permission repair through reparse point: $current"
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
    return $full
}
foreach ($requestFile in Get-ChildItem -LiteralPath $StateDirectory -Filter '*.request.json' -File) {
    $requestId = $requestFile.Name -replace '\.request\.json$', ''
    if ($requestId -notmatch '^[a-f0-9]{32}$') { continue }
    $resultPath = Join-Path $StateDirectory "$requestId.result.json"
    if (Test-Path -LiteralPath $resultPath) { continue }
    $errorsFound = @()
    try {
        $request = Get-Content -LiteralPath $requestFile.FullName -Raw | ConvertFrom-Json
        if ($request.request_id -ne $requestId -or @($request.paths).Count -gt 10000) { throw 'Invalid repair request' }
        $logPath = Join-Path $StateDirectory "$requestId.helper.log"
        foreach ($path in $request.paths) {
            try {
                $full = Assert-SourcePath $path
                $grant = "*${BackupAccountSid}:(OI)(CI)(RX)"
                $exitCode = Invoke-LoggedNative 'icacls.exe' @($full,'/grant',$grant,'/T','/C','/L','/Q') $logPath
                if ($exitCode -ne 0) {
                    $ownershipArguments = @('/F',$full,'/A')
                    if ((Get-Item -LiteralPath $full -Force).PSIsContainer) {
                        $ownershipArguments += @('/R','/D','Y','/SKIPSL')
                    }
                    $exitCode = Invoke-LoggedNative 'takeown.exe' $ownershipArguments $logPath
                    if ($exitCode -ne 0) { throw "Cannot take ownership: $full" }
                    $exitCode = Invoke-LoggedNative 'icacls.exe' @($full,'/grant',$grant,'/T','/C','/L','/Q') $logPath
                    if ($exitCode -ne 0) { throw "Cannot grant read access: $full" }
                }
            } catch { $errorsFound += $_.Exception.Message }
        }
    } catch { $errorsFound += $_.Exception.Message }
    $result = @{request_id=$requestId; success=($errorsFound.Count -eq 0); errors=@($errorsFound)}
    $tempResult = "$resultPath.tmp"
    $result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $tempResult -Encoding UTF8
    Move-Item -LiteralPath $tempResult -Destination $resultPath -Force
}
