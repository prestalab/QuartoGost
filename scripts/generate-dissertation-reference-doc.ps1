param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$target = Join-Path (Split-Path -Path $PSScriptRoot -Parent) "resources\reference-docs\dissertation\reference.docx"
$tempTarget = Join-Path ([System.IO.Path]::GetDirectoryName($target)) ("reference-" + [System.Guid]::NewGuid().ToString("N") + ".docx")
New-Item -ItemType Directory -Force -Path (Split-Path -Path $target -Parent) | Out-Null

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.ScreenUpdating = $false
$word.DisplayAlerts = 0

$wdStyleTypeParagraph = 1
$wdStyleTypeCharacter = 2
$wdStyleTypeTable = 3
$wdAlignParagraphLeft = 0
$wdAlignParagraphCenter = 1
$wdAlignParagraphRight = 2
$wdAlignParagraphJustify = 3
$wdLineSpace1pt5 = 1
$wdOrientPortrait = 0
$wdHeaderFooterPrimary = 1
$wdSectionBreakNextPage = 2
$wdPageBreak = 7

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

function Ensure-DissertationStyles {
  param($Document)

  $normal = $Document.Styles.Item(-1)
  Set-ParagraphStyleBase -Style $normal `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 1.25 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpace1pt5

  $bodyText = Get-OrAddStyle -Document $Document -Name "Body Text" -Type $wdStyleTypeParagraph
  $bodyText.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $bodyText `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 1.25 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpace1pt5

  $firstParagraph = Get-OrAddStyle -Document $Document -Name "First Paragraph" -Type $wdStyleTypeParagraph
  $firstParagraph.BaseStyle = $bodyText
  Set-ParagraphStyleBase -Style $firstParagraph `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $wdLineSpace1pt5

  $sourceCode = Get-OrAddStyle -Document $Document -Name "Source Code" -Type $wdStyleTypeParagraph
  $sourceCode.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $sourceCode `
    -FontName "Courier New" `
    -FontSize 12.0 `
    -Alignment $wdAlignParagraphLeft `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $wdLineSpace1pt5

  $figure = Get-OrAddStyle -Document $Document -Name "Figure" -Type $wdStyleTypeParagraph
  $figure.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $figure `
    -FontName "Times New Roman" `
    -FontSize 12.0 `
    -Alignment $wdAlignParagraphCenter `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 6 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $wdLineSpace1pt5

  $referenceItem = Get-OrAddStyle -Document $Document -Name "ReferenceItem" -Type $wdStyleTypeParagraph
  $referenceItem.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $referenceItem `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphJustify `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $wdLineSpace1pt5

  $un1 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1" -Type $wdStyleTypeParagraph
  $un1.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un1 `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphCenter `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 18 `
    -SpaceAfterPt 18 `
    -LineSpacingRule $wdLineSpace1pt5
  $un1.Font.Bold = $true

  $un1NoToc = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1NoTOC" -Type $wdStyleTypeParagraph
  $un1NoToc.BaseStyle = $un1

  $un2 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading2" -Type $wdStyleTypeParagraph
  $un2.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un2 `
    -FontName "Times New Roman" `
    -FontSize 14.0 `
    -Alignment $wdAlignParagraphLeft `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 18 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $wdLineSpace1pt5
  $un2.Font.Bold = $true

  foreach ($styleName in @("Heading 1", "Heading 2", "Heading 3")) {
    try {
      $style = $Document.Styles.Item($styleName)
      $style.Font.Name = "Times New Roman"
      $style.Font.Size = [single]14.0
      $style.Font.Bold = $true
      $style.ParagraphFormat.Alignment = $wdAlignParagraphCenter
      $style.ParagraphFormat.FirstLineIndent = 0
      $style.ParagraphFormat.SpaceBefore = 18
      $style.ParagraphFormat.SpaceAfter = 18
      $style.ParagraphFormat.LineSpacingRule = $wdLineSpace1pt5
    } catch { }
  }

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
    $markerStyle.Font.Size = 14
  }

  foreach ($styleName in @("TableStyleGost", "TableStyleGostNoHeader", "TableStyleContributors", "TableStyleAbbreviations")) {
    $style = Get-OrAddStyle -Document $Document -Name $styleName -Type $wdStyleTypeTable
    $style.Font.Name = "Times New Roman"
    $style.Font.Size = 12
  }
}

function Add-StyledParagraph {
  param(
    $Selection,
    [string]$Text,
    [string]$StyleName = "Body Text",
    [int]$Alignment = $wdAlignParagraphLeft,
    [switch]$Bold
  )

  $Selection.Style = $StyleName
  $Selection.ParagraphFormat.Alignment = $Alignment
  $Selection.Font.Bold = $(if ($Bold.IsPresent) { 1 } else { 0 })
  $Selection.TypeText($Text)
  $Selection.TypeParagraph()
  $Selection.Font.Bold = 0
}

function Add-BlankParagraphs {
  param($Selection, [int]$Count)

  for ($index = 1; $index -le $Count; $index++) {
    $Selection.TypeParagraph()
  }
}

function Set-DissertationSectionLayout {
  param(
    $Section,
    [switch]$ShowPageNumber,
    [int]$StartingPageNumber = 1
  )

  $setup = $Section.PageSetup
  $setup.Orientation = $wdOrientPortrait
  $setup.TopMargin = [single]$word.CentimetersToPoints(2.0)
  $setup.BottomMargin = [single]$word.CentimetersToPoints(2.0)
  $setup.LeftMargin = [single]$word.CentimetersToPoints(2.5)
  $setup.RightMargin = [single]$word.CentimetersToPoints(1.0)
  $setup.HeaderDistance = [single]$word.CentimetersToPoints(1.0)
  $setup.FooterDistance = [single]$word.CentimetersToPoints(1.0)

  $header = $Section.Headers.Item($wdHeaderFooterPrimary)
  $header.Range.Text = ""
  $header.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
  $header.PageNumbers.RestartNumberingAtSection = $true
  $header.PageNumbers.StartingNumber = $StartingPageNumber
  if ($ShowPageNumber) {
    $null = $header.PageNumbers.Add()
  }
}

function Build-DissertationReference {
  param($Document)

  $Document.Content.Delete() | Out-Null
  $Document.Activate()
  $selection = $word.Selection

  Add-StyledParagraph -Selection $selection -Text "%THESIS_ORGANIZATION%" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 4
  Add-StyledParagraph -Selection $selection -Text "На правах рукописи" -Alignment $wdAlignParagraphRight
  Add-BlankParagraphs -Selection $selection -Count 4
  Add-StyledParagraph -Selection $selection -Text "%THESIS_AUTHOR%" -Alignment $wdAlignParagraphCenter -Bold
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "%THESIS_TITLE%" -StyleName "UnnumberedHeading1" -Alignment $wdAlignParagraphCenter -Bold
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "%SPECIALTY_LINE_1%" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "%SPECIALTY_LINE_2%" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 2
  Add-StyledParagraph -Selection $selection -Text "Диссертация на соискание учёной степени" -Alignment $wdAlignParagraphCenter
  Add-StyledParagraph -Selection $selection -Text "%THESIS_DEGREE%" -Alignment $wdAlignParagraphCenter
  Add-BlankParagraphs -Selection $selection -Count 4
  Add-StyledParagraph -Selection $selection -Text "Научный руководитель:" -Alignment $wdAlignParagraphRight
  Add-StyledParagraph -Selection $selection -Text "%SUPERVISOR_REGALIA%" -Alignment $wdAlignParagraphRight
  Add-StyledParagraph -Selection $selection -Text "%SUPERVISOR_FIO%" -Alignment $wdAlignParagraphRight
  Add-StyledParagraph -Selection $selection -Text "%SUPERVISOR_TWO_BLOCK%" -Alignment $wdAlignParagraphRight
  Add-BlankParagraphs -Selection $selection -Count 3
  Add-StyledParagraph -Selection $selection -Text "%THESIS_CITY% --- %THESIS_YEAR%" -Alignment $wdAlignParagraphCenter

  $selection.InsertBreak($wdSectionBreakNextPage) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "ОГЛАВЛЕНИЕ" -StyleName "UnnumberedHeading1NoTOC" -Alignment $wdAlignParagraphCenter -Bold
  Add-StyledParagraph -Selection $selection -Text "%TOC%" -Alignment $wdAlignParagraphLeft
  $selection.InsertBreak($wdPageBreak) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "%MAINTEXT%" -Alignment $wdAlignParagraphLeft

  Set-DissertationSectionLayout -Section $Document.Sections.Item(1) -StartingPageNumber 1
  if ($Document.Sections.Count -ge 2) {
    Set-DissertationSectionLayout -Section $Document.Sections.Item(2) -ShowPageNumber -StartingPageNumber 2
  }
}

try {
  Write-Host "Creating dissertation reference document..."
  $document = $word.Documents.Add()
  try {
    Write-Host "Applying styles..."
    Ensure-DissertationStyles -Document $document
    Write-Host "Building content..."
    Build-DissertationReference -Document $document
    Write-Host "Saving temp file..."
    $document.SaveAs2($tempTarget)
  }
  finally {
    Write-Host "Closing temp document..."
    $document.Close([ref]0)
  }

  Write-Host "Moving final file..."
  if (Test-Path -LiteralPath $target) {
    Remove-Item -LiteralPath $target -Force
  }
  Move-Item -LiteralPath $tempTarget -Destination $target
}
finally {
  Write-Host "Closing Word..."
  $word.Quit()
}
