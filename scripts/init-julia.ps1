param(
  [string]$Julia = "julia",
  [string]$Project = (Join-Path $PSScriptRoot "julia")
)

$projectPath = [System.IO.Path]::GetFullPath($Project)

Write-Host "Initializing Julia environment at $projectPath"
& $Julia --project=$projectPath -e "using Pkg; Pkg.instantiate(); Pkg.precompile()"

if ($LASTEXITCODE -ne 0) {
  throw "Julia environment initialization failed."
}

Write-Host "Julia environment is ready."

