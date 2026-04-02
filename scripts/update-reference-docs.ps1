param(
  [ValidateSet("espd", "report", "dissertation", "synopsis", "presentation", "study-guide", "all")]
  [string[]]$DocumentType = @("all")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$selectedTypes = if ($DocumentType -contains "all") {
  @("espd", "report", "dissertation", "synopsis", "presentation", "study-guide")
} else {
  $DocumentType | Select-Object -Unique
}

$gostdownTypes = @($selectedTypes | Where-Object { $_ -in @("espd", "report") })
$generatedTypes = @($selectedTypes | Where-Object { $_ -in @("dissertation", "synopsis", "presentation", "study-guide") })

if ($gostdownTypes.Count -gt 0) {
  & (Join-Path $PSScriptRoot "refresh-gostdown-reference-docs.ps1") -DocumentType $gostdownTypes
}

if ($generatedTypes.Count -gt 0) {
  $previousSelection = $env:QUARTOGOST_REFERENCE_TYPES
  try {
    $env:QUARTOGOST_REFERENCE_TYPES = ($generatedTypes -join ",")
    & (Join-Path $PSScriptRoot "generate-reference-docs.ps1")
  }
  finally {
    if ($null -eq $previousSelection) {
      Remove-Item Env:\QUARTOGOST_REFERENCE_TYPES -ErrorAction SilentlyContinue
    }
    else {
      $env:QUARTOGOST_REFERENCE_TYPES = $previousSelection
    }
  }
}
