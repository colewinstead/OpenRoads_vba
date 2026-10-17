param([Parameter(Mandatory=$true)][string]$Python)
$ErrorActionPreference = 'Stop'
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'python'
$root = Join-Path $env:TEMP ('terrain bridge Unicode ' + [char]0x0394 + ' & test-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $root > $null
try {
    foreach ($case in @('invalid-executable','invalid-request')) {
        $folder = Join-Path $root $case
        New-Item -ItemType Directory -Path $folder > $null
        $request = Join-Path $folder 'request.json'
        $exe = if ($case -eq 'invalid-executable') { Join-Path $root 'missing.exe' } else { $Python }
        @{ schema_version=0; action='search'; python_executable=$exe } | ConvertTo-Json | Set-Content -LiteralPath $request -Encoding UTF8
        & powershell.exe -NoProfile -NonInteractive -File (Join-Path $source 'launch_worker.ps1') -Request $request
        if ($LASTEXITCODE -eq 0) { throw "Invalid input unexpectedly succeeded: $case" }
        $status = [IO.File]::ReadAllText((Join-Path $folder 'status.txt'), [Text.Encoding]::Unicode)
        if (-not $status.StartsWith('failed')) { throw "Missing failure status: $case" }
        if (-not (Test-Path -LiteralPath (Join-Path $folder 'error.log'))) { throw "Missing diagnostic log: $case" }
    }
    Write-Output 'PASS: Windows launcher handles spaces, Unicode, ampersands, missing interpreter, and worker failure.'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force
}
