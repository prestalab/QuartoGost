param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("espd", "report", "dissertation", "synopsis", "presentation", "envelopes", "article", "study-guide")]
  [string]$DocumentType,

  [string]$InputFile,
  [string]$OutputDir = "build",
  [ValidateSet("all", "docx", "pdf", "pptx")]
  [string]$Format = "all",
  [string]$Name,
  [string]$ReferenceDoc,
  [string]$AddressList,
  [string]$SenderName = "Организация-отправитель",
  [string]$SenderAddress = "Адрес отправителя",
  [string]$JuliaProject = (Join-Path $PSScriptRoot "julia"),
  [string]$Quarto = "quarto",
  [switch]$EmbedFonts,
  [switch]$Counters,
  [switch]$NoWordPostprocess,
  [switch]$NoPresentationHandout
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

try {
  Add-Type -AssemblyName Microsoft.Office.Interop.Word -ErrorAction Stop
}
catch {
  Add-Type -TypeDefinition @"
namespace Microsoft.Office.Interop.Word {
  public enum wdSaveFormat { wdFormatPDF = 17 }
}
"@
}

function Resolve-PathSafe {
  param([string]$PathValue)
  return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $PathValue))
}

function Test-CommandAvailable {
  param([string]$CommandName)
  return $null -ne (Get-Command $CommandName -ErrorAction SilentlyContinue)
}

function Convert-SimpleYamlScalar {
  param([string]$Value)

  $text = $Value.Trim()
  if ($text.Length -ge 2) {
    if (($text.StartsWith('"') -and $text.EndsWith('"')) -or ($text.StartsWith("'") -and $text.EndsWith("'"))) {
      return $text.Substring(1, $text.Length - 2)
    }
  }

  if ($text -eq "null" -or $text -eq "~") {
    return ""
  }

  $text = $text.Replace("`r`n", [Environment]::NewLine)
  $text = $text.Replace("`n", [Environment]::NewLine)
  $text = $text.Replace("`r", [Environment]::NewLine)

  return $text
}

function Get-PlaceholderMapFromQmd {
  param(
    [string]$QmdPath,
    [string[]]$BlockNames = @("gost")
  )

  $map = @{}
  if (-not (Test-Path -LiteralPath $QmdPath)) {
    return $map
  }

  $content = Get-Content -LiteralPath $QmdPath -Raw
  if ($content -notmatch "(?s)^---\r?\n(.*?)\r?\n---") {
    return $map
  }

  $frontMatter = $Matches[1] -split "\r?\n"
  $insideBlock = $false

  foreach ($line in $frontMatter) {
    if (-not $insideBlock) {
      foreach ($blockName in $BlockNames) {
        if ($line -match ("^\s*" + [regex]::Escape($blockName) + ":\s*$")) {
          $insideBlock = $true
          break
        }
      }

      if ($insideBlock) {
        continue
      }
      continue
    }

    if ($line -match "^\S") {
      $insideBlock = $false
      foreach ($blockName in $BlockNames) {
        if ($line -match ("^\s*" + [regex]::Escape($blockName) + ":\s*$")) {
          $insideBlock = $true
          break
        }
      }
      if (-not $insideBlock) {
        continue
      }
      continue
    }

    if ($line -match "^\s{2}([A-Za-z0-9_-]+):\s*(.*)$") {
      $key = $Matches[1]
      $value = Convert-SimpleYamlScalar -Value $Matches[2]
      $map[$key] = $value
    }
  }

  return $map
}

function Invoke-QuartoRender {
  param(
    [string]$Source,
    [string]$To,
    [string]$OutputDirectory,
    [string]$OutputName
  )

  $arguments = @(
    "render",
    $Source,
    "--to",
    $To,
    "--output-dir",
    $OutputDirectory
  )

  if (-not [string]::IsNullOrWhiteSpace($OutputName)) {
    $arguments += @("--output", $OutputName)
  }

  Write-Host ("Running: {0} {1}" -f $Quarto, ($arguments -join " "))
  & $Quarto @arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Quarto render failed for format '$To'."
  }
}

function Export-WordPdf {
  param(
    [string]$Docx,
    [string]$Pdf,
    [switch]$Embed
  )

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  try {
    $doc = $word.Documents.Open($Docx)
    if ($Embed) {
      $word.ActiveDocument.EmbedTrueTypeFonts = $true
      $word.ActiveDocument.DoNotEmbedSystemFonts = $true
      $word.ActiveDocument.SaveSubsetFonts = $true
    }
    $doc.SaveAs2([ref]$Pdf, [ref][Microsoft.Office.Interop.Word.wdSaveFormat]::wdFormatPDF)
  }
  finally {
    if ($doc) { $doc.Close() }
    $word.Quit()
  }
}

function New-EnvelopeSource {
  param(
    [string]$TsvPath,
    [string]$TempFile,
    [string]$OutgoingName,
    [string]$OutgoingAddress,
    [string]$ReferenceDocPath
  )

  $rows = Import-Csv -Delimiter "`t" -Path $TsvPath -Header "postcode", "city", "address", "organization", "recipient"
  $body = @()
  $body += @(
    "---",
    'title: "Конверты для рассылки"',
    "lang: ru",
    "format:",
    "  docx:",
    "    reference-doc: $ReferenceDocPath",
    "    toc: false",
    "---",
    "",
    "# Лист конвертов {.unnumbered}",
    ""
  )

  foreach ($row in $rows) {
    $recipient = if ([string]::IsNullOrWhiteSpace($row.recipient)) { "" } else { "$($row.recipient)`r`n" }
    $body += @(
      "::: {.callout-note appearance=`"minimal`"}",
      "## Отправитель {.unnumbered}",
      "$OutgoingName  ",
      "$OutgoingAddress",
      "",
      "## Получатель {.unnumbered}",
      "$recipient$($row.organization)  ",
      "$($row.address)  ",
      "$($row.city)  ",
      "$($row.postcode)",
      ":::",
      "",
      "{{< pagebreak >}}",
      ""
    )
  }

  Set-Content -LiteralPath $TempFile -Value $body -Encoding UTF8
}

if (-not (Test-CommandAvailable -CommandName $Quarto)) {
  throw "Quarto command '$Quarto' was not found. Install Quarto and ensure it is in PATH."
}

$outputRoot = Resolve-PathSafe -PathValue $OutputDir
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

$defaultInputMap = @{
  espd = "templates\espd\espd-template.qmd"
  report = "templates\report\report-template.qmd"
  dissertation = "templates\dissertation\dissertation-template.qmd"
  synopsis = "templates\synopsis\synopsis-template.qmd"
  presentation = "templates\presentation\presentation-template.qmd"
  article = "templates\article\article-template.qmd"
  "study-guide" = "templates\study-guide\study-guide-template.qmd"
}

if ([string]::IsNullOrWhiteSpace($InputFile) -and $DocumentType -ne "envelopes") {
  $InputFile = $defaultInputMap[$DocumentType]
}

$requestedFormats = switch ($Format) {
  "all" {
    if ($DocumentType -eq "presentation") { @("pptx") }
    else { @("docx", "pdf") }
  }
  default { @($Format) }
}

$nameBase = if ([string]::IsNullOrWhiteSpace($Name)) { $DocumentType } else { $Name }
$tempRoot = Join-Path $outputRoot ("_tmp_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

$env:JULIA_PROJECT = [System.IO.Path]::GetFullPath($JuliaProject)

try {
  if ($DocumentType -eq "presentation") {
    $source = Resolve-PathSafe -PathValue $InputFile
    $pptxName = "$nameBase.pptx"
    Invoke-QuartoRender -Source $source -To "pptx" -OutputDirectory $outputRoot -OutputName $pptxName
    if (-not $NoPresentationHandout) {
      $handoutScript = Join-Path $PSScriptRoot "export-presentation-handout.ps1"
      $handoutPdf = Join-Path $outputRoot "$nameBase-handout.pdf"
      & $handoutScript -InputPptx (Join-Path $outputRoot $pptxName) -HandoutPdf $handoutPdf
      if ($LASTEXITCODE -ne 0) {
        throw "Presentation handout export failed."
      }
    }
    return
  }

  if ($DocumentType -eq "envelopes") {
    if ([string]::IsNullOrWhiteSpace($AddressList)) {
      throw "For envelopes use -AddressList <path-to-tsv>."
    }
    $tempQmd = Join-Path $tempRoot "$nameBase.qmd"
    $envelopeReferenceDoc = Resolve-PathSafe -PathValue "resources\reference-docs\report\reference.docx"
    New-EnvelopeSource -TsvPath (Resolve-PathSafe -PathValue $AddressList) -TempFile $tempQmd -OutgoingName $SenderName -OutgoingAddress $SenderAddress -ReferenceDocPath $envelopeReferenceDoc
    $source = $tempQmd
  } else {
    $source = Resolve-PathSafe -PathValue $InputFile
  }

  $tempDocx = Join-Path $tempRoot "$nameBase.docx"
  $finalDocx = Join-Path $outputRoot "$nameBase.docx"
  $finalPdf = Join-Path $outputRoot "$nameBase.pdf"

  $needDocxIntermediate = $requestedFormats -contains "docx" -or $requestedFormats -contains "pdf"
  if ($needDocxIntermediate) {
    Invoke-QuartoRender -Source $source -To "docx" -OutputDirectory $tempRoot -OutputName "$nameBase.docx"
  }

  $mergeWithReference = ($DocumentType -eq "espd" -or $DocumentType -eq "report" -or $DocumentType -eq "dissertation" -or $DocumentType -eq "study-guide" -or $DocumentType -eq "synopsis") -and (-not $NoWordPostprocess)
  if ([string]::IsNullOrWhiteSpace($ReferenceDoc) -and $mergeWithReference) {
    $ReferenceDoc = switch ($DocumentType) {
      "espd" { "resources\reference-docs\espd\reference.docx" }
      "report" { "resources\reference-docs\report\reference.docx" }
      "dissertation" { "resources\reference-docs\dissertation\reference.docx" }
      "synopsis" { "resources\reference-docs\synopsis\reference.docx" }
      "study-guide" { "resources\reference-docs\study-guide\reference.docx" }
    }
  }

  if ($mergeWithReference) {
    $placeholderJsonPath = $null
    if ($DocumentType -ne "envelopes" -and $source.ToLowerInvariant().EndsWith(".qmd")) {
      $blockNames = @("gost")
      if ($DocumentType -eq "synopsis") {
        $blockNames = @("synopsis", "gost")
      } elseif ($DocumentType -eq "dissertation") {
        $blockNames = @("dissertation", "gost")
      }

      $placeholders = Get-PlaceholderMapFromQmd -QmdPath $source -BlockNames $blockNames
      if ($placeholders.Count -gt 0) {
        $placeholderJsonPath = Join-Path $tempRoot "gost-placeholders.json"
        $jsonText = $placeholders | ConvertTo-Json
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($placeholderJsonPath, $jsonText, $utf8Bom)
      }
    }

    $postProcessScript = Join-Path $PSScriptRoot "postprocess-word.ps1"
    $applyCounters = $Counters -or $DocumentType -eq "report"

    $postProcessParams = @{
      ReferenceDoc = (Resolve-PathSafe -PathValue $ReferenceDoc)
      InputDocx = $tempDocx
      EmbedFonts = $EmbedFonts
    }

    if ($requestedFormats -contains "docx") {
      $postProcessParams.OutputDocx = $finalDocx
    }
    if ($requestedFormats -contains "pdf") {
      $postProcessParams.Pdf = $finalPdf
    }
    if ($applyCounters) {
      $postProcessParams.Counters = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($placeholderJsonPath)) {
      $postProcessParams.PlaceholderJson = $placeholderJsonPath
    }

    & $postProcessScript @postProcessParams
    if ($LASTEXITCODE -ne 0) {
      throw "Word post-processing failed."
    }
  } else {
    if ($requestedFormats -contains "docx") {
      Copy-Item -LiteralPath $tempDocx -Destination $finalDocx -Force
    }
    if ($requestedFormats -contains "pdf") {
      Export-WordPdf -Docx $tempDocx -Pdf $finalPdf -Embed:$EmbedFonts
    }
  }
}
finally {
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
  }
}
