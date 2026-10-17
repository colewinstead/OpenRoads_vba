param([Parameter(Mandatory=$true)][string]$Request)
$ErrorActionPreference = 'Stop'
# Dedicated child environment: ORD's environment and configuration are never modified.
try {
    $requestPath = [IO.Path]::GetFullPath($Request)
    $cfg = Get-Content -LiteralPath $requestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not (Test-Path -LiteralPath $cfg.python_executable -PathType Leaf)) { throw 'Python executable not found. Configure a full python.exe path.' }
    if ([IO.Path]::GetExtension($cfg.python_executable) -ne '.exe') { throw 'Select python.exe, not a .bat/.cmd launcher.' }
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $cfg.python_executable
    $script = Join-Path $PSScriptRoot 'main.py'
    if ($script.Contains('"') -or $requestPath.Contains('"')) { throw 'Invalid quote in path.' }
    $start.Arguments = '"' + $script + '" "' + $requestPath + '"'
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.WorkingDirectory = [IO.Path]::GetDirectoryName($requestPath)
    if ($cfg.environment_file) {
        $envCfg = Get-Content -LiteralPath $cfg.environment_file -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($entry in $envCfg.PSObject.Properties) {
            if ($entry.Name -eq 'PATH_prepend') {
                $start.EnvironmentVariables['PATH'] = [string]$entry.Value + ';' + $start.EnvironmentVariables['PATH']
            } elseif ($entry.Name -in @('GDAL_DATA','PROJ_DATA','PYTHONHOME','PYTHONPATH','QT_PLUGIN_PATH','SSL_CERT_FILE')) {
                $start.EnvironmentVariables[$entry.Name] = [string]$entry.Value
            } else { throw ('Unsupported child environment key: ' + $entry.Name) }
        }
    }
    $start.EnvironmentVariables['PROJ_NETWORK'] = 'OFF'
    $process = [Diagnostics.Process]::Start($start)
    $stdout = $process.StandardOutput.ReadToEndAsync()
    $stderr = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    $folder = [IO.Path]::GetDirectoryName($requestPath)
    [IO.File]::WriteAllText((Join-Path $folder 'stdout.log'), $stdout.Result)
    [IO.File]::WriteAllText((Join-Path $folder 'stderr.log'), $stderr.Result)
    if ($process.ExitCode -ne 0 -and -not (Test-Path -LiteralPath (Join-Path $folder 'status.txt'))) {
        [IO.File]::WriteAllText((Join-Path $folder 'status.txt'), "failed`r`nPython did not initialize. Verify environment paths; see stderr.log.", [Text.Encoding]::Unicode)
        [IO.File]::WriteAllText((Join-Path $folder 'error.log'), $stderr.Result)
    }
    exit $process.ExitCode
} catch {
    $folder = Split-Path -Parent $Request
    $message = $_.Exception.Message
    [IO.File]::WriteAllText((Join-Path $folder 'status.txt'), "failed`r`n$message", [Text.Encoding]::Unicode)
    [IO.File]::WriteAllText((Join-Path $folder 'error.log'), ($_ | Out-String))
    exit 2
}
