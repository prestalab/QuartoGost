param(
  [Parameter(Mandatory = $true)]
  [string]$InputPptx,

  [Parameter(Mandatory = $true)]
  [string]$HandoutPdf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$inputPath = [System.IO.Path]::GetFullPath($InputPptx)
$outputPath = [System.IO.Path]::GetFullPath($HandoutPdf)
$outputDir = Split-Path -Path $outputPath -Parent
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$tempRoot = Join-Path $outputDir ("_handout_" + [System.Guid]::NewGuid().ToString("N"))
$slideExportDir = Join-Path $tempRoot "slides"
New-Item -ItemType Directory -Force -Path $slideExportDir | Out-Null

$wdOrientLandscape = 1
$wdPaperA4 = 7
$wdCollapseEnd = 0
$wdSaveFormatPDF = 17
$wdAlignParagraphCenter = 1
$wdRowAlignmentCenter = 1
$msoTrue = -1

function Get-SlideImagePaths {
  param([string]$DirectoryPath)

  return @(
    Get-ChildItem -LiteralPath $DirectoryPath -Filter "*.PNG" |
      Sort-Object {
        if ($_.BaseName -match "(\d+)$") { [int]$Matches[1] } else { 0 }
      } |
      Select-Object -ExpandProperty FullName
  )
}

function Add-HandoutPage {
  param(
    $Word,
    $Document,
    [string[]]$SlidePaths
  )

  $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
  $table = $Document.Tables.Add($range, 1, 2)
  $table.Borders.Enable = 0
  $table.Rows.Alignment = $wdRowAlignmentCenter
  $table.AllowAutoFit = $false
  $table.Columns.Item(1).Width = [single]$Word.CentimetersToPoints(13.95)
  $table.Columns.Item(2).Width = [single]$Word.CentimetersToPoints(13.95)

  for ($column = 1; $column -le 2; $column++) {
    $cell = $table.Cell(1, $column)
    $cell.VerticalAlignment = 0
    $cellRange = $cell.Range
    $cellRange.End = $cellRange.End - 1
    $cellRange.ParagraphFormat.Alignment = $wdAlignParagraphCenter

    if ($column -le $SlidePaths.Count) {
      $picture = $cellRange.InlineShapes.AddPicture($SlidePaths[$column - 1])
      $picture.LockAspectRatio = $msoTrue
      $picture.Width = [single]$Word.CentimetersToPoints(13.6)
    }
  }

  $afterRange = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
  $afterRange.InsertParagraphAfter() | Out-Null
}

$powerPoint = $null
$presentation = $null
$word = $null
$document = $null

try {
  $powerPoint = New-Object -ComObject PowerPoint.Application
  $powerPoint.Visible = 1
  $presentation = $powerPoint.Presentations.Open($inputPath, 0, 0, 0)
  $presentation.Export($slideExportDir, "PNG", 1920, 1080)
  $presentation.Close()
  $presentation = $null
  $powerPoint.Quit()
  $powerPoint = $null

  $slidePaths = Get-SlideImagePaths -DirectoryPath $slideExportDir
  if ($slidePaths.Count -eq 0) {
    throw "No slide images were exported from presentation."
  }

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.ScreenUpdating = $false
  $word.DisplayAlerts = 0
  $document = $word.Documents.Add()

  $setup = $document.Sections.Item(1).PageSetup
  $setup.Orientation = $wdOrientLandscape
  $setup.PaperSize = $wdPaperA4
  $setup.TopMargin = [single]$word.CentimetersToPoints(0.5)
  $setup.BottomMargin = [single]$word.CentimetersToPoints(1.5)
  $setup.LeftMargin = [single]$word.CentimetersToPoints(0.5)
  $setup.RightMargin = [single]$word.CentimetersToPoints(0.5)
  $setup.HeaderDistance = 0
  $setup.FooterDistance = [single]$word.CentimetersToPoints(0.7)

  $normal = $document.Styles.Item(-1)
  $normal.Font.Name = "Times New Roman"
  $normal.Font.Size = 10
  $normal.ParagraphFormat.SpaceBefore = 0
  $normal.ParagraphFormat.SpaceAfter = 0

  for ($index = 0; $index -lt $slidePaths.Count; $index += 2) {
    $pair = @($slidePaths[$index])
    if (($index + 1) -lt $slidePaths.Count) {
      $pair += $slidePaths[$index + 1]
    }

    Add-HandoutPage -Word $word -Document $document -SlidePaths $pair

    if (($index + 2) -lt $slidePaths.Count) {
      $selection = $word.Selection
      $selection.EndKey(6) | Out-Null
      $selection.Collapse($wdCollapseEnd) | Out-Null
      $selection.InsertBreak(7) | Out-Null
    }
  }

  $document.SaveAs2($outputPath, $wdSaveFormatPDF)
}
finally {
  if ($document) { $document.Close([ref]0) }
  if ($word) { $word.Quit() }
  if ($presentation) { $presentation.Close() }
  if ($powerPoint) { $powerPoint.Quit() }
  if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
