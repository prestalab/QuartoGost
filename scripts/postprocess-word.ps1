param(
  [Parameter(Mandatory = $true)][string]$ReferenceDoc,
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [string]$OutputDocx,
  [string]$Pdf,
  [string]$PlaceholderJson,
  [switch]$EmbedFonts,
  [switch]$Counters
)

try {
  Add-Type -AssemblyName Microsoft.Office.Interop.Word -ErrorAction Stop
}
catch {
  Add-Type -TypeDefinition @"
namespace Microsoft.Office.Interop.Word {
  public enum wdFindWrap { wdFindStop = 0, wdFindContinue = 1, wdFindAsk = 2 }
  public enum wdReplace { wdReplaceNone = 0, wdReplaceOne = 1, wdReplaceAll = 2 }
  public enum wdCollapseDirection { wdCollapseEnd = 0, wdCollapseStart = 1 }
  public enum wdUnits { wdCharacter = 1, wdStory = 6 }
  public enum wdMovementType { wdMove = 0, wdExtend = 1 }
  public enum wdBuiltinStyle { wdStyleHeading1 = -2, wdStyleHeading2 = -3, wdStyleHeading3 = -4, wdStyleBodyText = -67 }
  public enum wdListNumberStyle { wdListNumberStyleArabic = 0, wdListNumberStyleLowercaseRoman = 2, wdListNumberStyleBullet = 23 }
  public enum wdListLevelAlignment { wdListLevelAlignLeft = 0 }
  public enum wdTrailingCharacter { wdTrailingTab = 0 }
  public enum wdAutoFitBehavior { wdAutoFitContent = 1 }
  public enum wdLineSpacing { wdLineSpaceSingle = 0 }
  public enum wdInformation { wdActiveEndPageNumber = 3, wdNumberOfPagesInDocument = 4 }
  public enum wdSaveFormat { wdFormatPDF = 17 }
}
"@
}

if ([string]::IsNullOrWhiteSpace($OutputDocx) -and [string]::IsNullOrWhiteSpace($Pdf)) {
  throw "Specify -OutputDocx or -Pdf."
}

$referenceDoc = [System.IO.Path]::GetFullPath($ReferenceDoc)
$inputDocx = [System.IO.Path]::GetFullPath($InputDocx)
$isTemporaryDocx = $false

if ([string]::IsNullOrWhiteSpace($OutputDocx)) {
  $OutputDocx = [System.IO.Path]::GetTempFileName() + ".docx"
  $isTemporaryDocx = $true
} else {
  $OutputDocx = [System.IO.Path]::GetFullPath($OutputDocx)
}

if (-not [string]::IsNullOrWhiteSpace($Pdf)) {
  $Pdf = [System.IO.Path]::GetFullPath($Pdf)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.ScreenUpdating = $false

function Get-PlaceholderEntries {
  param([string]$JsonPath)

  $items = @()
  if ([string]::IsNullOrWhiteSpace($JsonPath) -or -not (Test-Path -LiteralPath $JsonPath)) {
    return $items
  }

  $jsonObject = Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json
  if ($null -eq $jsonObject) {
    return $items
  }

  foreach ($property in $jsonObject.PSObject.Properties) {
    $items += [PSCustomObject]@{
      Key = $property.Name
      Value = [string]$property.Value
    }
  }

  return $items
}

function Replace-TextInDocument {
  param(
    $Document,
    [string]$FindText,
    [string]$ReplaceText
  )

  if ($Document.StoryRanges.Count -lt 1) {
    return
  }

  $story = $Document.StoryRanges.Item(1)
  while ($story -ne $null) {
    $find = $story.Find
    $find.ClearFormatting()
    $find.Replacement.ClearFormatting()
    $find.Execute(
      $FindText,
      $false,
      $false,
      $false,
      $false,
      $false,
      $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue,
      $false,
      $ReplaceText,
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceAll
    ) | Out-Null

    $story = $story.NextStoryRange
  }
}

try {
  $doc = $word.Documents.Open($referenceDoc)
  $doc.Activate()
  $selection = $word.Selection

  Write-Host "Saving merged document..."
  $doc.SaveAs([ref]$OutputDocx)

  $placeholderEntries = Get-PlaceholderEntries -JsonPath $PlaceholderJson
  foreach ($entry in $placeholderEntries) {
    $token = "%" + $entry.Key.ToUpperInvariant() + "%"
    Replace-TextInDocument -Document $doc -FindText $token -ReplaceText $entry.Value
  }

  Write-Host "Inserting main text into template..."
  if ($selection.Find.Execute("%MAINTEXT%^13", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceNone)) {
    $start = $Selection.Range.Start
    $Selection.InsertFile($inputDocx)
    $end = $Selection.Range.End
    $insertedTables = $doc.Range([ref]$start, [ref]$end).Tables

    $selection.WholeStory()
    $totalEnd = $Selection.Range.End
    if ($end -ge ($totalEnd - 1)) {
      $selection.Collapse([Microsoft.Office.Interop.Word.wdCollapseDirection]::wdCollapseEnd) | Out-Null
      $selection.MoveLeft([Microsoft.Office.Interop.Word.wdUnits]::wdCharacter, 1,
        [Microsoft.Office.Interop.Word.wdMovementType]::wdExtend) | Out-Null
      $selection.Delete() | Out-Null
    }
  } else {
    throw "Reference template does not contain %MAINTEXT% marker."
  }

  foreach ($style in $doc.Styles) {
    switch ($style.NameLocal) {
      "TableStyleContributors" { $TableStyleContributors = $style; break }
      "TableStyleAbbreviations" { $TableStyleAbbreviations = $style; break }
      "TableStyleGost" { $TableStyleGost = $style; break }
      "TableStyleGostNoHeader" { $TableStyleGostNoHeader = $style; break }
      "UnnumberedHeading1" { $UnnumberedHeading1 = $style; break }
      "UnnumberedHeading1NoTOC" { $UnnumberedHeading1NoTOC = $style; break }
      "UnnumberedHeading2" { $UnnumberedHeading2 = $style; break }
    }
  }

  $bodyText = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleBodyText
  $heading1 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading1
  $heading2 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading2
  $heading3 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading3

  $bullets = [char]0x2014, [char]0xB0, [char]0x2014, [char]0xB0
  $numberPosition = 0, 0.75, 1.75, 3
  $textPosition = 0.85, 1.75, 3, 3.5
  $tabPosition = 1, 1.75, 3, 3.5
  $formatNested = "%1)", "%1.%2)", "%1.%2.%3)", "%1.%2.%3.%4)"
  $formatHeaders = "%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"
  $formatSingle = "%1)", "%2)", "%3)", "%4)"

  foreach ($template in $doc.ListTemplates) {
    for ($index = 1; $index -le $template.ListLevels.Count -and $index -le 4; $index++) {
      $level = $template.ListLevels.Item($index)
      $bullet = $level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleBullet
      $arabic = $level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleArabic
      $roman = $level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleLowercaseRoman

      if ($bullet) {
        if ($level.NumberFormat -ne " ") {
          $level.NumberFormat = $bullets[$index - 1] + ""
        }
        $level.NumberPosition = $word.CentimetersToPoints($numberPosition[$index - 1])
        $level.Alignment = [Microsoft.Office.Interop.Word.wdListLevelAlignment]::wdListLevelAlignLeft
        $level.TextPosition = $word.CentimetersToPoints($textPosition[$index - 1])
        $level.TabPosition = $word.CentimetersToPoints($tabPosition[$index - 1])
        $level.ResetOnHigher = $index - 1
        $level.StartAt = 1
        $level.Font.Size = 12
        $level.Font.Name = "PT Serif"
        if ($index % 2 -eq 0) {
          $level.Font.Position = -4
        }
        $level.LinkedStyle = ""
        $level.TrailingCharacter = [Microsoft.Office.Interop.Word.wdTrailingCharacter]::wdTrailingTab
      }

      if (($arabic -and ($level.NumberFormat -ne $formatHeaders[$index - 1])) -or $roman) {
        if ($level.NumberFormat -ne " ") {
          if ($arabic) {
            $level.NumberFormat = $formatNested[$index - 1]
          }
          if ($roman) {
            $level.NumberStyle = [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleArabic
            $level.NumberFormat = $formatSingle[$index - 1]
          }
        }
        $level.NumberPosition = $word.CentimetersToPoints($numberPosition[$index - 1])
        $level.Alignment = [Microsoft.Office.Interop.Word.wdListLevelAlignment]::wdListLevelAlignLeft
        $level.TextPosition = $word.CentimetersToPoints($textPosition[$index - 1])
        $level.TabPosition = $word.CentimetersToPoints($tabPosition[$index - 1])
        $level.ResetOnHigher = $index - 1
        $level.StartAt = 1
        $level.Font.Size = 12
        $level.Font.Name = "PT Serif"
        $level.LinkedStyle = ""
        $level.TrailingCharacter = [Microsoft.Office.Interop.Word.wdTrailingCharacter]::wdTrailingTab
      }
    }
  }

  $doc.GrammarChecked = $true
  $tableCount = 0

  for ($tableIndex = 1; $tableIndex -le $insertedTables.Count; $tableIndex++) {
    $table = $insertedTables.Item($tableIndex)

    if ($table.Cell(1, 1).Range.Style.NameLocal -eq "ContributorsTable" -and $TableStyleContributors) {
      $table.Select()
      $selection.ClearParagraphAllFormatting()
      $paragraph = $selection.ParagraphFormat
      $paragraph.LeftIndent = 0
      $paragraph.RightIndent = 0
      $paragraph.SpaceBefore = 0
      $paragraph.SpaceBeforeAuto = $false
      $paragraph.SpaceAfter = 0
      $paragraph.SpaceAfterAuto = $false
      $table.Style = $TableStyleContributors
      continue
    }

    if ($table.Cell(1, 1).Range.Style.NameLocal -eq "AbbreviationsTable" -and $TableStyleAbbreviations) {
      $table.AllowAutoFit = $true
      $table.AutoFitBehavior([Microsoft.Office.Interop.Word.wdAutoFitBehavior]::wdAutoFitContent)
      $table.Style = $TableStyleAbbreviations
      continue
    }

    $table.AllowAutoFit = $true
    if (-not [string]::IsNullOrWhiteSpace($table.Title)) {
      $tableCount++
    }

    $table.Select()
    $paragraph = $selection.ParagraphFormat
    $paragraph.LineSpacingRule = [Microsoft.Office.Interop.Word.wdLineSpacing]::wdLineSpaceSingle

    $table.Cell(1, 1).Select()
    if ($selection.Rows.HeadingFormat -eq -1 -and $TableStyleGost) {
      $table.Style = $TableStyleGost
    } elseif ($TableStyleGostNoHeader) {
      $table.Style = $TableStyleGostNoHeader
    }
  }

  $heading1Name = $doc.Styles.Item([Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading1).NameLocal
  $chapterCount = 0
  $figureCount = 0
  $referenceCount = 0
  $appendixCount = 0

  foreach ($paragraph in $doc.Paragraphs) {
    $characterStyleName = $paragraph.Range.CharacterStyle.NameLocal
    if ($characterStyleName -eq "UnnumberedHeadingOne" -and $UnnumberedHeading1) {
      $paragraph.Style = $UnnumberedHeading1
      continue
    }
    if ($characterStyleName -eq "AppendixHeadingOne" -and $UnnumberedHeading1) {
      $paragraph.Style = $UnnumberedHeading1
      $appendixCount++
      continue
    }
    if ($characterStyleName -eq "UnnumberedHeadingOneNoTOC" -and $UnnumberedHeading1NoTOC) {
      $paragraph.Style = $UnnumberedHeading1NoTOC
      continue
    }
    if ($characterStyleName -eq "UnnumberedHeadingTwo" -and $UnnumberedHeading2) {
      $paragraph.Style = $UnnumberedHeading2
      continue
    }

    $styleName = $paragraph.Style.NameLocal
    if ($styleName -eq "Source Code") {
      $paragraph.Range.Font.Size = 10.5
    } elseif ($styleName -eq "First Paragraph") {
      $paragraph.Style = $bodyText
    } elseif ($styleName -eq $heading1Name) {
      $chapterCount++
    } elseif ($styleName -eq "Captioned Figure") {
      $figureCount++
    } elseif ($styleName -eq "ReferenceItem") {
      $referenceCount++
    }
  }

  if ($Counters) {
    $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
    $selection.Find.Execute("%NCHAPTERS%", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $chapterCount + "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null
    $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
    $selection.Find.Execute("%NFIGURES%", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $figureCount + "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null
    $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
    $selection.Find.Execute("%NTABLES%", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $tableCount + "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null
    $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
    $selection.Find.Execute("%NREFERENCES%", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $referenceCount + "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null
    $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
    $selection.Find.Execute("%NAPPENDICES%", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $appendixCount + "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null
  }

  foreach ($math in $doc.OMaths) {
    $math.Range.Font.Size = 12.5
  }

  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
  if ($selection.Find.Execute("%TOC%^13", $true, $true, $false, $false, $false, $true,
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, "",
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceNone)) {
    $doc.TablesOfContents.Add($selection.Range, $false, 9, 9, $false, "", $true, $true, "", $true) | Out-Null
    $toc = $doc.TablesOfContents.Item(1)
    $toc.UseHeadingStyles = $true
    if ($UnnumberedHeading1) { $toc.HeadingStyles.Add($UnnumberedHeading1, 1) | Out-Null }
    if ($UnnumberedHeading2) { $toc.HeadingStyles.Add($UnnumberedHeading2, 2) | Out-Null }
    $toc.HeadingStyles.Add($heading1, 1) | Out-Null
    $toc.HeadingStyles.Add($heading2, 2) | Out-Null
    $toc.HeadingStyles.Add($heading3, 3) | Out-Null
    $toc.Update() | Out-Null
  }

  $doc.Repaginate()
  if ($doc.Sections.Count -gt 1) {
    $pageCount = $doc.Sections.Item(2).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdActiveEndPageNumber) -
      $doc.Sections.Item(1).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdActiveEndPageNumber)
  } else {
    $pageCount = $doc.Sections.Item(1).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdNumberOfPagesInDocument)
  }

  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | Out-Null
  $selection.Find.Execute("%NPAGES%", $true, $true, $false, $false, $false, $true,
    [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $false, $pageCount + "",
    [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | Out-Null

  if ($EmbedFonts) {
    $word.ActiveDocument.EmbedTrueTypeFonts = $true
    $word.ActiveDocument.DoNotEmbedSystemFonts = $true
    $word.ActiveDocument.SaveSubsetFonts = $true
  }

  if (-not $isTemporaryDocx) {
    $doc.Save()
  }

  if (-not [string]::IsNullOrWhiteSpace($Pdf)) {
    $doc.SaveAs2([ref]$Pdf, [ref][Microsoft.Office.Interop.Word.wdSaveFormat]::wdFormatPDF)
  }
}
finally {
  if ($doc) { $doc.Close() }
  $word.Quit()
  if ($isTemporaryDocx -and (Test-Path -LiteralPath $OutputDocx)) {
    Remove-Item -LiteralPath $OutputDocx -Force
  }
}
