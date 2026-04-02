param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$target = Join-Path (Split-Path -Path $PSScriptRoot -Parent) "resources\reference-docs\synopsis\reference.docx"
$tempTarget = Join-Path ([System.IO.Path]::GetDirectoryName($target)) ("reference-" + [System.Guid]::NewGuid().ToString("N") + ".docx")
New-Item -ItemType Directory -Force -Path (Split-Path -Path $target -Parent) | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.ScreenUpdating = $false

$wdStyleTypeParagraph = 1
$wdStyleTypeCharacter = 2
$wdStyleTypeTable = 3
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignParagraphJustify = 3
$wdLineSpaceSingle = 0
$wdOrientPortrait = 0
$wdPageBreak = 7
$wdSectionBreakNextPage = 2
$wdHeaderFooterPrimary = 1

function Get-OrAddStyle {
  param(
    $Document,
    [string]$Name,
    [int]$Type
  )

  try {
    return $Document.Styles.Item($Name)
  }
  catch {
    return $Document.Styles.Add($Name, $Type)
  }
}

function Set-ParagraphStyleBase {
  param(
    $Style,
    [string]$FontName,
    [double]$FontSize,
    [int]$Alignment,
    [double]$FirstLineIndentCm,
    [double]$SpaceBeforePt,
    [double]$SpaceAfterPt,
    [int]$LineSpacingRule
  )

  $Style.Font.Name = $FontName
  $Style.Font.Size = [single]$FontSize
  $Style.ParagraphFormat.Alignment = $Alignment
  $Style.ParagraphFormat.FirstLineIndent = [single]$word.CentimetersToPoints($FirstLineIndentCm)
  $Style.ParagraphFormat.SpaceBefore = [single]$SpaceBeforePt
  $Style.ParagraphFormat.SpaceAfter = [single]$SpaceAfterPt
  $Style.ParagraphFormat.LineSpacingRule = $LineSpacingRule
}

function Set-SynopsisSectionLayout {
  param(
    $Section,
    [switch]$ShowPageNumber,
    [switch]$RestartPageNumber
  )

  $setup = $Section.PageSetup
  $setup.Orientation = $wdOrientPortrait
  $setup.PageWidth = [single]$word.CentimetersToPoints(14.8)
  $setup.PageHeight = [single]$word.CentimetersToPoints(21.0)
  $setup.MirrorMargins = $true
  $setup.TopMargin = [single]$word.CentimetersToPoints(1.4)
  $setup.BottomMargin = [single]$word.CentimetersToPoints(1.4)
  $setup.LeftMargin = [single]$word.CentimetersToPoints(1.8)
  $setup.RightMargin = [single]$word.CentimetersToPoints(1.0)
  $setup.HeaderDistance = 0
  $setup.FooterDistance = [single]$word.CentimetersToPoints(0.5)

  $footer = $Section.Footers.Item($wdHeaderFooterPrimary)
  $footer.Range.Text = ""
  $footer.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
  $footer.PageNumbers.RestartNumberingAtSection = $RestartPageNumber.IsPresent
  if ($RestartPageNumber) {
    $footer.PageNumbers.StartingNumber = 1
  }

  if ($ShowPageNumber) {
    $null = $footer.PageNumbers.Add()
  }
}

function Ensure-CommonSynopsisStyles {
  param($Document)

  $normal = $Document.Styles.Item(-1)
  Set-ParagraphStyleBase -Style $normal `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0.9 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpaceSingle

  $bodyText = Get-OrAddStyle -Document $Document -Name "Body Text" -Type $wdStyleTypeParagraph
  $bodyText.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $bodyText `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0.9 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpaceSingle

  $firstParagraph = Get-OrAddStyle -Document $Document -Name "First Paragraph" -Type $wdStyleTypeParagraph
  $firstParagraph.BaseStyle = $bodyText
  Set-ParagraphStyleBase -Style $firstParagraph `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpaceSingle

  $un1 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1" -Type $wdStyleTypeParagraph
  $un1.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un1 `
    -FontName "Times New Roman" `
    -FontSize 12.0 `
    -Alignment $wdAlignParagraphCenter `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 18 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $wdLineSpaceSingle
  $un1.Font.Bold = $true

  $un1NoToc = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1NoTOC" -Type $wdStyleTypeParagraph
  $un1NoToc.BaseStyle = $un1

  $un2 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading2" -Type $wdStyleTypeParagraph
  $un2.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un2 `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphLeft `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 12 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $wdLineSpaceSingle
  $un2.Font.Bold = $true

  $figure = Get-OrAddStyle -Document $Document -Name "Figure" -Type $wdStyleTypeParagraph
  $figure.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $figure `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphCenter `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 6 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $wdLineSpaceSingle

  $referenceItem = Get-OrAddStyle -Document $Document -Name "ReferenceItem" -Type $wdStyleTypeParagraph
  $referenceItem.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $referenceItem `
    -FontName "Times New Roman" `
    -FontSize 10.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 4 `
    -LineSpacingRule $wdLineSpaceSingle

  $sourceCode = Get-OrAddStyle -Document $Document -Name "Source Code" -Type $wdStyleTypeParagraph
  $sourceCode.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $sourceCode `
    -FontName "Courier New" `
    -FontSize 9.0 `
    -Alignment $wdAlignParagraphLeft `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 4 `
    -LineSpacingRule $wdLineSpaceSingle

  foreach ($markerName in @(
    "UnnumberedHeadingOne",
    "UnnumberedHeadingOneNoTOC",
    "UnnumberedHeadingTwo",
    "AppendixHeadingOne",
    "ContributorsTable",
    "AbbreviationsTable",
    "GostKeywords",
    "Figure",
    "ReferenceItem",
    "MyCustomStyle"
  )) {
    $markerStyle = Get-OrAddStyle -Document $Document -Name $markerName -Type $wdStyleTypeCharacter
    $markerStyle.Font.Name = "Times New Roman"
    $markerStyle.Font.Size = 10
  }

  foreach ($styleName in @("TableStyleGost", "TableStyleGostNoHeader", "TableStyleContributors", "TableStyleAbbreviations")) {
    $style = Get-OrAddStyle -Document $Document -Name $styleName -Type $wdStyleTypeTable
    $style.Font.Name = "Times New Roman"
    $style.Font.Size = 10
  }
}

function Add-StyledParagraph {
  param(
    $Selection,
    [string]$Text,
    [string]$StyleName = "Body Text",
    [int]$Alignment = $wdAlignParagraphLeft,
    [switch]$Bold,
    [switch]$Italic,
    [switch]$AllCaps
  )

  $Selection.Style = $StyleName
  $Selection.ParagraphFormat.Alignment = $Alignment
  $Selection.TypeText($Text)
  $Selection.Font.Bold = $(if ($Bold.IsPresent) { 1 } else { 0 })
  $Selection.Font.Italic = $(if ($Italic.IsPresent) { 1 } else { 0 })
  $Selection.Font.AllCaps = $(if ($AllCaps.IsPresent) { 1 } else { 0 })
  $Selection.TypeParagraph()
  $Selection.Font.Bold = 0
  $Selection.Font.Italic = 0
  $Selection.Font.AllCaps = 0
}

function Add-BlankParagraphs {
  param($Selection, [int]$Count)
  for ($index = 1; $index -le $Count; $index++) {
    $Selection.TypeParagraph()
  }
}

function Set-CellText {
  param(
    $Cell,
    [string]$Text,
    [string]$StyleName = "Body Text",
    [int]$Alignment = $wdAlignParagraphLeft,
    [switch]$Bold
  )

  $range = $Cell.Range
  $range.End = $range.End - 1
  $range.Text = $Text
  $range.Style = $StyleName
  $range.ParagraphFormat.Alignment = $Alignment
  $range.Font.Bold = $(if ($Bold.IsPresent) { 1 } else { 0 })
}

function Build-SynopsisReference {
  param($Document)

  $Document.Content.Delete() | Out-Null
  $Document.Activate()
  $selection = $word.Selection

  Add-StyledParagraph -Selection $selection -Text "На правах рукописи" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 5
  Add-StyledParagraph -Selection $selection -Text "%THESIS_AUTHOR%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter -Bold
  Add-BlankParagraphs -Selection $selection -Count 4
  Add-StyledParagraph -Selection $selection -Text "%THESIS_TITLE%" -StyleName "UnnumberedHeading1" -Alignment $wdAlignParagraphCenter -Bold
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "%SPECIALTY_LINE_1%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "%SPECIALTY_LINE_2%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "Автореферат" -StyleName "UnnumberedHeading1NoTOC" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "диссертации на соискание ученой степени" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "%THESIS_DEGREE%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 5
  Add-StyledParagraph -Selection $selection -Text "%THESIS_CITY% --- %THESIS_YEAR%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter

  $selection.InsertBreak($wdSectionBreakNextPage) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "Работа выполнена в %THESIS_IN_ORGANIZATION%." -StyleName "First Paragraph" -Alignment $wdAlignParagraphLeft
  Add-BlankParagraphs -Selection $selection -Count 1

  $tableRange = $selection.Range
  $table = $Document.Tables.Add($tableRange, 3, 2)
  $table.Borders.Enable = 0
  $table.Rows.LeftIndent = 0
  $table.Columns.Item(1).PreferredWidth = [single]$word.CentimetersToPoints(4.0)
  $table.Columns.Item(2).PreferredWidth = [single]$word.CentimetersToPoints(8.5)

  Set-CellText -Cell $table.Cell(1, 1) -Text "Научный руководитель:"
  Set-CellText -Cell $table.Cell(1, 2) -Text "%SUPERVISOR_REGALIA%`r%SUPERVISOR_FIO%`r%SUPERVISOR_TWO_BLOCK%"

  Set-CellText -Cell $table.Cell(2, 1) -Text "Официальные оппоненты:"
  Set-CellText -Cell $table.Cell(2, 2) -Text "%OPPONENT_1_BLOCK%`r`r%OPPONENT_2_BLOCK%`r`r%OPPONENT_3_BLOCK%"

  Set-CellText -Cell $table.Cell(3, 1) -Text "Ведущая организация:"
  Set-CellText -Cell $table.Cell(3, 2) -Text "%LEADING_ORGANIZATION_TITLE%"

  $selection.MoveDown() | Out-Null
  $selection.EndKey(6) | Out-Null
  $selection.TypeParagraph()

  Add-StyledParagraph -Selection $selection -Text "Защита состоится %DEFENSE_DATE% на заседании диссертационного совета %DEFENSE_COUNCIL_NUMBER% при %DEFENSE_COUNCIL_TITLE% по адресу: %DEFENSE_COUNCIL_ADDRESS%." -StyleName "Body Text"
  Add-BlankParagraphs -Selection $selection -Count 1
  Add-StyledParagraph -Selection $selection -Text "С диссертацией можно ознакомиться в библиотеке %SYNOPSIS_LIBRARY%." -StyleName "Body Text"
  Add-BlankParagraphs -Selection $selection -Count 1
  Add-StyledParagraph -Selection $selection -Text "Отзывы на автореферат в двух экземплярах, заверенные печатью учреждения, просьба направлять по адресу: %DEFENSE_COUNCIL_ADDRESS%, ученому секретарю диссертационного совета %DEFENSE_COUNCIL_NUMBER%." -StyleName "Body Text"
  Add-BlankParagraphs -Selection $selection -Count 1
  Add-StyledParagraph -Selection $selection -Text "Автореферат разослан %SYNOPSIS_DATE%." -StyleName "Body Text"
  Add-StyledParagraph -Selection $selection -Text "Телефон для справок: %DEFENSE_COUNCIL_PHONE%." -StyleName "Body Text"
  Add-BlankParagraphs -Selection $selection -Count 1

  $secTableRange = $selection.Range
  $secTable = $Document.Tables.Add($secTableRange, 1, 2)
  $secTable.Borders.Enable = 0
  $secTable.Columns.Item(1).PreferredWidth = [single]$word.CentimetersToPoints(7.0)
  $secTable.Columns.Item(2).PreferredWidth = [single]$word.CentimetersToPoints(5.5)
  Set-CellText -Cell $secTable.Cell(1, 1) -Text "Ученый секретарь диссертационного совета %DEFENSE_COUNCIL_NUMBER%, %DEFENSE_SECRETARY_REGALIA%"
  Set-CellText -Cell $secTable.Cell(1, 2) -Text "%DEFENSE_SECRETARY_FIO%" -Alignment $wdAlignParagraphCenter

  $selection.MoveDown() | Out-Null
  $selection.EndKey(6) | Out-Null
  $selection.InsertBreak($wdSectionBreakNextPage) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "%MAINTEXT%" -StyleName "Body Text"

  $selection.InsertBreak($wdSectionBreakNextPage) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "%THESIS_AUTHOR_SHORT%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter -Italic
  Add-BlankParagraphs -Selection $selection -Count 1
  Add-StyledParagraph -Selection $selection -Text "%THESIS_TITLE%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 1
  Add-StyledParagraph -Selection $selection -Text "Автореф. дис. на соискание ученой степени %THESIS_DEGREE_SHORT%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "Подписано в печать %PRINT_SIGN_DATE%. Заказ № %PRINT_ORDER_NUMBER%." -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "Формат 60×90/16. Усл. печ. л. 1. Тираж 100 экз." -StyleName "Body Text" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "Типография %PRINT_SHOP%" -StyleName "Body Text" -Alignment $wdAlignParagraphCenter

  Set-SynopsisSectionLayout -Section $Document.Sections.Item(1)
  Set-SynopsisSectionLayout -Section $Document.Sections.Item(2) -ShowPageNumber -RestartPageNumber
  Set-SynopsisSectionLayout -Section $Document.Sections.Item(3)
  Set-SynopsisSectionLayout -Section $Document.Sections.Item(4)
}

try {
  $document = $word.Documents.Add()
  try {
    Ensure-CommonSynopsisStyles -Document $document
    Build-SynopsisReference -Document $document
    $document.SaveAs($tempTarget)
  }
  finally {
    $document.Close([ref]0)
  }

  $document = $word.Documents.Open($tempTarget)
  try {
    for ($index = 1; $index -le $document.Sections.Count; $index++) {
      if ($index -eq 2) {
        Set-SynopsisSectionLayout -Section $document.Sections.Item($index) -ShowPageNumber -RestartPageNumber
      }
      else {
        Set-SynopsisSectionLayout -Section $document.Sections.Item($index)
      }
    }
    $document.Save()
  }
  finally {
    $document.Close([ref]0)
  }

  Move-Item -LiteralPath $tempTarget -Destination $target -Force
}
finally {
  $word.Quit()
}
