param(
  [ValidateSet("espd", "report", "all")]
  [string[]]$DocumentType = @("all")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Path $PSScriptRoot -Parent
$espdSource = Join-Path $root "ref\gostdown\demo-template-espd.docx"
$reportSource = Join-Path $root "ref\gostdown\demo-template-report.docx"
$espdTarget = Join-Path $root "resources\reference-docs\espd\reference.docx"
$reportTarget = Join-Path $root "resources\reference-docs\report\reference.docx"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.ScreenUpdating = $false

$selectedTypes = if ($DocumentType -contains "all") {
  @("espd", "report")
} else {
  $DocumentType
}

function Set-ParagraphText {
  param(
    $Document,
    [int]$ParagraphIndex,
    [string]$Text
  )

  $paragraph = $Document.Paragraphs.Item($ParagraphIndex)
  $range = $paragraph.Range.Duplicate
  if ($range.End -gt $range.Start) {
    $range.SetRange($range.Start, $range.End - 1)
  }
  $range.Text = $Text
}

function Apply-ParagraphMap {
  param(
    [string]$Path,
    [hashtable]$Map
  )

  $document = $word.Documents.Open($Path)
  try {
    foreach ($entry in ($Map.GetEnumerator() | Sort-Object Key)) {
      Set-ParagraphText -Document $document -ParagraphIndex ([int]$entry.Key) -Text ([string]$entry.Value)
    }
    $document.Save()
  }
  finally {
    $document.Close([ref]0)
  }
}

try {
  # Меняем только реально переменные реквизиты.
  # Служебный текст макета, включая "УТВЕРЖДАЮ", "СОГЛАСОВАНО",
  # названия служебных листов и подчеркивания для подписей, сохраняем как есть.
  $espdParagraphs = [ordered]@{
    9 = "%APPROVER_TITLE%"
    12 = "__________ %APPROVER_NAME%"
    15 = "%APPROVAL_DATE%"
    21 = "%DOC_TITLE%"
    22 = "%DOC_KIND%"
    24 = "%DOC_APPROVAL_CODE%"
    29 = "%AGREED_1_TITLE%"
    30 = "____________________ %AGREED_1_NAME%"
    33 = "%AGREED_2_TITLE%"
    34 = "____________________ %AGREED_2_NAME%"
    37 = "%AGREED_3_TITLE%"
    38 = "____________________ %AGREED_3_NAME%"
    41 = "%AGREED_4_TITLE%"
    42 = "____________________ %AGREED_4_NAME%"
    45 = "%DOC_YEAR%"
    47 = "%DOC_APPROVAL_CODE%"
    55 = "%DOC_TITLE%"
    56 = "%DOC_KIND%"
    57 = "%DOC_CODE%"
    72 = "%DOC_YEAR%"
  }

  $reportParagraphs = [ordered]@{
    2 = "%ORG_NAME%"
    6 = "%UDC_LINE%"
    7 = "%RESEARCH_REG_NUMBER%"
    8 = "%INVENTORY_NUMBER%"
    14 = "%APPROVER_TITLE%"
    15 = "__________ %APPROVER_NAME%"
    16 = "%APPROVAL_DATE%"
    24 = "%REPORT_TITLE%"
    26 = "%TOPIC_TITLE%"
    27 = "%REPORT_STAGE%"
    28 = "%RESEARCH_CODE%"
    30 = "%LEADER_TITLE%"
    34 = "%LEADER_NAME%"
    36 = "%EXECUTOR_TITLE%"
    40 = "%EXECUTOR_NAME%"
    45 = "%CITY_YEAR%"
  }

  if ($selectedTypes -contains "espd") {
    New-Item -ItemType Directory -Force -Path (Split-Path -Path $espdTarget -Parent) | Out-Null
    Copy-Item -LiteralPath $espdSource -Destination $espdTarget -Force
    Apply-ParagraphMap -Path $espdTarget -Map $espdParagraphs
  }

  if ($selectedTypes -contains "report") {
    New-Item -ItemType Directory -Force -Path (Split-Path -Path $reportTarget -Parent) | Out-Null
    Copy-Item -LiteralPath $reportSource -Destination $reportTarget -Force
    Apply-ParagraphMap -Path $reportTarget -Map $reportParagraphs
  }
}
finally {
  $word.Quit()
}
