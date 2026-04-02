param(
  [Parameter(Mandatory = $true)]
  [string]$SourceQmd,

  [Parameter(Mandatory = $true)]
  [string]$InputPptx,

  [Parameter(Mandatory = $true)]
  [string]$OutputDocx
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourcePath = [System.IO.Path]::GetFullPath($SourceQmd)
$pptxPath = [System.IO.Path]::GetFullPath($InputPptx)
$outputPath = [System.IO.Path]::GetFullPath($OutputDocx)
$outputDir = Split-Path -Path $outputPath -Parent
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$tempRoot = Join-Path $outputDir ("_handout_" + [System.Guid]::NewGuid().ToString("N"))
$slideExportDir = Join-Path $tempRoot "slides"
New-Item -ItemType Directory -Force -Path $slideExportDir | Out-Null

$wdOrientLandscape = 1
$wdPaperA4 = 7
$wdPageBreak = 7
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdRowAlignmentCenter = 1
$wdLineSpaceSingle = 0
$msoTrue = -1

function Convert-MarkdownToPlainText {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return ""
  }

  $result = $Text
  $result = $result -replace "\r\n", "`n"
  $result = $result -replace "`r", "`n"
  $result = $result -replace '!\[[^\]]*\]\([^)]+\)', ""
  $result = $result -replace '\[([^\]]+)\]\([^)]+\)', '$1'
  $result = $result -replace '`([^`]+)`', '$1'
  $result = $result -replace '\*\*([^\*]+)\*\*', '$1'
  $result = $result -replace '\*([^\*]+)\*', '$1'
  $result = $result -replace '^#{1,6}\s*', ""
  $result = $result -replace '^\s*[-*+]\s+', "• "
  $result = $result -replace '^\s*\d+\.\s+', "• "
  $result = $result -replace '\{[^\}]+\}', ""
  $result = $result -replace '\s+$', ""
  return ($result -split "`n" | ForEach-Object { $_.TrimEnd() }) -join [Environment]::NewLine
}

function Get-TitleSlideNotes {
  param([string[]]$Lines)

  $insideFrontMatter = $false
  $capturing = $false
  $noteLines = New-Object System.Collections.Generic.List[string]

  for ($i = 0; $i -lt $Lines.Count; $i++) {
    $line = $Lines[$i]
    if ($i -eq 0 -and $line.Trim() -eq "---") {
      $insideFrontMatter = $true
      continue
    }

    if ($insideFrontMatter -and $line.Trim() -eq "---") {
      break
    }

    if (-not $insideFrontMatter) {
      break
    }

    if (-not $capturing -and $line -match "^\s*notes:\s*\|?\s*$") {
      $capturing = $true
      continue
    }

    if ($capturing) {
      if ($line -match "^\S") {
        break
      }
      $noteLines.Add(($line -replace "^\s{2}", ""))
    }
  }

  return Convert-MarkdownToPlainText -Text ($noteLines -join [Environment]::NewLine)
}

function Get-SlideNotesFromQmd {
  param([string]$QmdPath)

  $content = Get-Content -LiteralPath $QmdPath -Raw
  $content = $content -replace "\r\n", "`n"
  $content = $content -replace "`r", "`n"
  $lines = $content -split "`n"

  $frontMatterEnd = -1
  if ($lines.Count -gt 0 -and $lines[0].Trim() -eq "---") {
    for ($i = 1; $i -lt $lines.Count; $i++) {
      if ($lines[$i].Trim() -eq "---") {
        $frontMatterEnd = $i
        break
      }
    }
  }

  $bodyStart = if ($frontMatterEnd -ge 0) { $frontMatterEnd + 1 } else { 0 }
  $titleSlideNotes = Get-TitleSlideNotes -Lines $lines

  $slides = New-Object System.Collections.Generic.List[object]
  $slides.Add([PSCustomObject]@{
    Title = "Титульный слайд"
    Notes = $titleSlideNotes
  })

  $currentTitle = $null
  $currentLines = New-Object System.Collections.Generic.List[string]

  for ($i = $bodyStart; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    if ($line -match "^##\s+(.+?)\s*$") {
      if ($null -ne $currentTitle) {
        $slideText = ($currentLines -join "`n")
        $noteMatches = [regex]::Matches($slideText, "(?s)::: \{\.notes\}\s*(.*?)\s*:::")
        $noteText = ""
        if ($noteMatches.Count -gt 0) {
          $noteText = (($noteMatches | ForEach-Object { $_.Groups[1].Value.Trim() }) -join [Environment]::NewLine + [Environment]::NewLine)
        }
        $slides.Add([PSCustomObject]@{
          Title = Convert-MarkdownToPlainText -Text $currentTitle
          Notes = Convert-MarkdownToPlainText -Text $noteText
        })
      }

      $currentTitle = $Matches[1]
      $currentLines = New-Object System.Collections.Generic.List[string]
      continue
    }

    if ($null -ne $currentTitle) {
      $currentLines.Add($line)
    }
  }

  if ($null -ne $currentTitle) {
    $slideText = ($currentLines -join "`n")
    $noteMatches = [regex]::Matches($slideText, "(?s)::: \{\.notes\}\s*(.*?)\s*:::")
    $noteText = ""
    if ($noteMatches.Count -gt 0) {
      $noteText = (($noteMatches | ForEach-Object { $_.Groups[1].Value.Trim() }) -join [Environment]::NewLine + [Environment]::NewLine)
    }
    $slides.Add([PSCustomObject]@{
      Title = Convert-MarkdownToPlainText -Text $currentTitle
      Notes = Convert-MarkdownToPlainText -Text $noteText
    })
  }

  return $slides
}

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

function Set-CellText {
  param(
    $Cell,
    [string]$Text,
    [int]$Alignment,
    [string]$FontName,
    [double]$FontSize,
    [switch]$Bold
  )

  $range = $Cell.Range
  $range.End = $range.End - 1
  $range.Text = $Text
  $range.ParagraphFormat.Alignment = $Alignment
  $range.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle
  $range.Font.Name = $FontName
  $range.Font.Size = $FontSize
  $range.Font.Bold = $(if ($Bold.IsPresent) { 1 } else { 0 })
}

function Add-HandoutPage {
  param(
    $Word,
    $Document,
    [object[]]$Slides
  )

  $range = $Document.Range($Document.Content.End - 1, $Document.Content.End - 1)
  $table = $Document.Tables.Add($range, 3, 2)
  $table.Rows.Alignment = $wdRowAlignmentCenter
  $table.AllowAutoFit = $false
  $table.Columns.Item(1).Width = [single]$Word.CentimetersToPoints(14.1)
  $table.Columns.Item(2).Width = [single]$Word.CentimetersToPoints(14.1)
  $table.Rows.Item(1).Height = [single]$Word.CentimetersToPoints(7.9)
  $table.Rows.Item(2).Height = [single]$Word.CentimetersToPoints(1.1)
  $table.Rows.Item(3).Height = [single]$Word.CentimetersToPoints(8.7)

  for ($column = 1; $column -le 2; $column++) {
    $cell = $table.Cell(1, $column)
    $cell.VerticalAlignment = 1
    $cellRange = $cell.Range
    $cellRange.End = $cellRange.End - 1
    $cellRange.ParagraphFormat.Alignment = $wdAlignParagraphCenter

    if ($column -le $Slides.Count) {
      $picture = $cellRange.InlineShapes.AddPicture($Slides[$column - 1].ImagePath)
      $picture.LockAspectRatio = $msoTrue
      $picture.Width = [single]$Word.CentimetersToPoints(13.4)
    }
  }

  for ($column = 1; $column -le 2; $column++) {
    if ($column -le $Slides.Count) {
      Set-CellText -Cell $table.Cell(2, $column) `
        -Text ("Слайд {0}. {1}" -f $Slides[$column - 1].Number, $Slides[$column - 1].Title) `
        -Alignment $wdAlignParagraphCenter `
        -FontName "Times New Roman" `
        -FontSize 11 `
        -Bold

      $notesText = $Slides[$column - 1].Notes
      if ([string]::IsNullOrWhiteSpace($notesText)) {
        $notesText = "Комментарий к слайду не задан."
      }

      Set-CellText -Cell $table.Cell(3, $column) `
        -Text $notesText `
        -Alignment $wdAlignParagraphLeft `
        -FontName "Times New Roman" `
        -FontSize 11
    } else {
      Set-CellText -Cell $table.Cell(2, $column) -Text "" -Alignment $wdAlignParagraphCenter -FontName "Times New Roman" -FontSize 11
      Set-CellText -Cell $table.Cell(3, $column) -Text "" -Alignment $wdAlignParagraphLeft -FontName "Times New Roman" -FontSize 11
    }
  }
}

$powerPoint = $null
$presentation = $null
$word = $null
$document = $null

try {
  $slideNotes = Get-SlideNotesFromQmd -QmdPath $sourcePath

  $powerPoint = New-Object -ComObject PowerPoint.Application
  $powerPoint.Visible = 1
  $presentation = $powerPoint.Presentations.Open($pptxPath, 0, 0, 0)
  $presentation.Export($slideExportDir, "PNG", 1920, 1080)
  $presentation.Close()
  $presentation = $null
  $powerPoint.Quit()
  $powerPoint = $null

  $slideImages = Get-SlideImagePaths -DirectoryPath $slideExportDir
  if ($slideImages.Count -eq 0) {
    throw "No slide images were exported from presentation."
  }

  $slides = New-Object System.Collections.Generic.List[object]
  for ($index = 0; $index -lt $slideImages.Count; $index++) {
    $note = if ($index -lt $slideNotes.Count) { $slideNotes[$index] } else { [PSCustomObject]@{ Title = "Слайд"; Notes = "" } }
    $slides.Add([PSCustomObject]@{
      Number = $index + 1
      Title = $note.Title
      Notes = $note.Notes
      ImagePath = $slideImages[$index]
    })
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
  $normal.Font.Size = 11
  $normal.ParagraphFormat.SpaceBefore = 0
  $normal.ParagraphFormat.SpaceAfter = 0
  $normal.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle

  for ($index = 0; $index -lt $slides.Count; $index += 2) {
    $pair = @($slides[$index])
    if (($index + 1) -lt $slides.Count) {
      $pair += $slides[$index + 1]
    }

    Add-HandoutPage -Word $word -Document $document -Slides $pair

    if (($index + 2) -lt $slides.Count) {
      $word.Selection.EndKey(6) | Out-Null
      $word.Selection.InsertBreak($wdPageBreak) | Out-Null
    }
  }

  $document.SaveAs2($outputPath)
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
