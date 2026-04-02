param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("espd", "report", "dissertation", "synopsis", "presentation", "envelopes", "article", "study-guide")]
  [string]$DocumentType,

  [Parameter(Mandatory = $true)]
  [string]$Destination,

  [string]$Name
)

$templateMap = @{
  espd = "templates\espd\espd-template.qmd"
  report = "templates\report\report-template.qmd"
  dissertation = "templates\dissertation\dissertation-template.qmd"
  synopsis = "templates\synopsis\synopsis-template.qmd"
  presentation = "templates\presentation\presentation-template.qmd"
  envelopes = "templates\envelopes\envelopes-template.qmd"
  article = "templates\article\article-template.qmd"
  "study-guide" = "templates\study-guide\study-guide-template.qmd"
}

$workspace = (Get-Location).Path
$templateSource = if ([System.IO.Path]::IsPathRooted($templateMap[$DocumentType])) {
  $templateMap[$DocumentType]
} else {
  Join-Path -Path $workspace -ChildPath $templateMap[$DocumentType]
}
$destinationSource = if ([System.IO.Path]::IsPathRooted($Destination)) {
  $Destination
} else {
  Join-Path -Path $workspace -ChildPath $Destination
}

$template = [System.IO.Path]::GetFullPath($templateSource)
$targetDir = [System.IO.Path]::GetFullPath($destinationSource)
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

$targetName = if ([string]::IsNullOrWhiteSpace($Name)) { "$DocumentType.qmd" } else { "$Name.qmd" }
$targetFile = Join-Path $targetDir $targetName

Copy-Item -LiteralPath $template -Destination $targetFile -Force

Write-Host "Created $targetFile"
