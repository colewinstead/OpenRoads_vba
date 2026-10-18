param([string]$Python = "python", [string]$DependencyPath = "")
$ErrorActionPreference = "Stop"
$buildScript = Join-Path $PSScriptRoot "build_windows.py"
$buildArguments = @($buildScript)
if ($DependencyPath) { $buildArguments += @("--dependency-path", $DependencyPath) }
& $Python @buildArguments
if ($LASTEXITCODE -ne 0) { throw "Contour Plotter build or validation failed." }
