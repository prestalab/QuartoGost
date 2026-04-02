param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
$wdLineSpace1pt5 = 1
$wdOrientPortrait = 0
$wdPaperA4 = 7
$wdPageBreak = 7
$wdStyleNormal = -1
$wdStyleHeading1 = -2
$wdStyleHeading2 = -3
$wdStyleHeading3 = -4
$wdGoToPage = 1
$wdGoToAbsolute = 1

$selectedTypes = @("dissertation", "study-guide", "synopsis", "presentation")
if (-not [string]::IsNullOrWhiteSpace($env:QUARTOGOST_REFERENCE_TYPES)) {
  $selectedTypes = @(
    $env:QUARTOGOST_REFERENCE_TYPES.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) |
      ForEach-Object { $_.Trim().ToLowerInvariant() } |
      Where-Object { $_ -in @("dissertation", "study-guide", "synopsis", "presentation") } |
      Select-Object -Unique
  )

  if ($selectedTypes.Count -eq 0) {
    throw "QUARTOGOST_REFERENCE_TYPES must contain dissertation, study-guide, synopsis, presentation, or a comma-separated subset."
  }
}

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
  $Style.ParagraphFormat.FirstLineIndent = $word.CentimetersToPoints($FirstLineIndentCm)
  $Style.ParagraphFormat.SpaceBefore = [single]$SpaceBeforePt
  $Style.ParagraphFormat.SpaceAfter = [single]$SpaceAfterPt
  $Style.ParagraphFormat.LineSpacingRule = $LineSpacingRule
}

function Set-PageLayout {
  param(
    $Document,
    [double]$LeftCm,
    [double]$RightCm,
    [double]$TopCm,
    [double]$BottomCm,
    [double]$FooterCm
  )

  $section = $Document.Sections.Item(1)
  $setup = $section.PageSetup
  $setup.Orientation = $wdOrientPortrait
  $setup.PaperSize = $wdPaperA4
  $setup.LeftMargin = [single]$word.CentimetersToPoints($LeftCm)
  $setup.RightMargin = [single]$word.CentimetersToPoints($RightCm)
  $setup.TopMargin = [single]$word.CentimetersToPoints($TopCm)
  $setup.BottomMargin = [single]$word.CentimetersToPoints($BottomCm)
  $setup.FooterDistance = [single]$word.CentimetersToPoints($FooterCm)
}

function Get-NormalizedParagraphText {
  param($Paragraph)

  $text = $Paragraph.Range.Text
  $text = $text -replace "[`r`a]", " "
  $text = $text -replace "\s+", " "
  return $text.Trim()
}

function Copy-ParagraphFormatting {
  param(
    $SourceParagraph,
    $TargetParagraph
  )

  $sourceParagraphFormat = $SourceParagraph.Range.ParagraphFormat
  $targetParagraphFormat = $TargetParagraph.Range.ParagraphFormat
  $sourceFont = $SourceParagraph.Range.Font
  $targetFont = $TargetParagraph.Range.Font

  $targetParagraphFormat.Alignment = $sourceParagraphFormat.Alignment
  $targetParagraphFormat.LeftIndent = $sourceParagraphFormat.LeftIndent
  $targetParagraphFormat.RightIndent = $sourceParagraphFormat.RightIndent
  $targetParagraphFormat.FirstLineIndent = $sourceParagraphFormat.FirstLineIndent
  $targetParagraphFormat.SpaceBefore = $sourceParagraphFormat.SpaceBefore
  $targetParagraphFormat.SpaceAfter = $sourceParagraphFormat.SpaceAfter
  $targetParagraphFormat.LineSpacingRule = $sourceParagraphFormat.LineSpacingRule
  $targetParagraphFormat.LineSpacing = $sourceParagraphFormat.LineSpacing
  $targetParagraphFormat.KeepTogether = $sourceParagraphFormat.KeepTogether
  $targetParagraphFormat.KeepWithNext = $sourceParagraphFormat.KeepWithNext
  $targetParagraphFormat.WidowControl = $sourceParagraphFormat.WidowControl

  $targetFont.Name = $sourceFont.Name
  $targetFont.Size = $sourceFont.Size
  $targetFont.Bold = $sourceFont.Bold
  $targetFont.Italic = $sourceFont.Italic
  $targetFont.AllCaps = $sourceFont.AllCaps
  $targetFont.SmallCaps = $sourceFont.SmallCaps
  $targetFont.Position = $sourceFont.Position
  $targetFont.Color = $sourceFont.Color
}

function Copy-StyleFormatting {
  param(
    $SourceDocument,
    $TargetDocument,
    [string]$StyleName
  )

  try {
    $sourceStyle = $SourceDocument.Styles.Item($StyleName)
    $targetStyle = $TargetDocument.Styles.Item($StyleName)
  }
  catch {
    return
  }

  $sourceFormat = $sourceStyle.ParagraphFormat
  $targetFormat = $targetStyle.ParagraphFormat
  $sourceFont = $sourceStyle.Font
  $targetFont = $targetStyle.Font

  $targetFormat.Alignment = $sourceFormat.Alignment
  $targetFormat.LeftIndent = $sourceFormat.LeftIndent
  $targetFormat.RightIndent = $sourceFormat.RightIndent
  $targetFormat.FirstLineIndent = $sourceFormat.FirstLineIndent
  $targetFormat.SpaceBefore = $sourceFormat.SpaceBefore
  $targetFormat.SpaceAfter = $sourceFormat.SpaceAfter
  $targetFormat.LineSpacingRule = $sourceFormat.LineSpacingRule
  $targetFormat.LineSpacing = $sourceFormat.LineSpacing
  $targetFormat.KeepTogether = $sourceFormat.KeepTogether
  $targetFormat.KeepWithNext = $sourceFormat.KeepWithNext
  $targetFormat.WidowControl = $sourceFormat.WidowControl

  $targetFont.Name = $sourceFont.Name
  $targetFont.Size = $sourceFont.Size
  $targetFont.Bold = $sourceFont.Bold
  $targetFont.Italic = $sourceFont.Italic
  $targetFont.AllCaps = $sourceFont.AllCaps
  $targetFont.SmallCaps = $sourceFont.SmallCaps
  $targetFont.Position = $sourceFont.Position
  $targetFont.Color = $sourceFont.Color
}

function Sync-CoverParagraphFormattingFromSource {
  param(
    $TargetDocument,
    [string]$SourcePath,
    [int]$MaxSourceParagraphs = 80
  )

  if (-not (Test-Path -LiteralPath $SourcePath)) {
    return
  }

  $sourceDocument = $word.Documents.Open($SourcePath, $false, $true)
  try {
    foreach ($styleName in @(
      "Название организации",
      "Обычный-центр",
      "УДК",
      "Заголовок",
      "Подписи",
      "Фамилии-по правому краю"
    )) {
      Copy-StyleFormatting -SourceDocument $sourceDocument -TargetDocument $TargetDocument -StyleName $styleName
    }

    $targetParagraphs = $TargetDocument.Paragraphs
    $sourceParagraphs = $sourceDocument.Paragraphs
    $limit = [Math]::Min($sourceParagraphs.Count, $MaxSourceParagraphs)

    $targetIndex = 1
    for ($sourceIndex = 1; $sourceIndex -le $limit; $sourceIndex++) {
      $sourceParagraph = $sourceParagraphs.Item($sourceIndex)
      $sourceText = Get-NormalizedParagraphText -Paragraph $sourceParagraph

      if ([string]::IsNullOrWhiteSpace($sourceText)) {
        continue
      }

      for (; $targetIndex -le $targetParagraphs.Count; $targetIndex++) {
        $targetParagraph = $targetParagraphs.Item($targetIndex)
        $targetText = Get-NormalizedParagraphText -Paragraph $targetParagraph

        if ($targetText -eq "%MAINTEXT%") {
          return
        }

        if ([string]::IsNullOrWhiteSpace($targetText)) {
          continue
        }

        if ($targetText -eq $sourceText) {
          Copy-ParagraphFormatting -SourceParagraph $sourceParagraph -TargetParagraph $targetParagraph
          $targetIndex++
          break
        }
      }
    }
  }
  finally {
    $sourceDocument.Close([ref]0)
  }
}

function Ensure-CommonStyles {
  param(
    $Document,
    [hashtable]$Profile
  )

  $normal = $Document.Styles.Item($wdStyleNormal)
  Set-ParagraphStyleBase -Style $normal `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphJustify) `
    -FirstLineIndentCm $Profile.FirstLineIndentCm `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $Profile.LineSpacingRule

  try {
    $Document.AutoHyphenation = $true
  } catch { }

  $bodyText = Get-OrAddStyle -Document $Document -Name "Body Text" -Type ($wdStyleTypeParagraph)
  $bodyText.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $bodyText `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphJustify) `
    -FirstLineIndentCm $Profile.FirstLineIndentCm `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $Profile.LineSpacingRule

  $firstParagraph = Get-OrAddStyle -Document $Document -Name "First Paragraph" -Type ($wdStyleTypeParagraph)
  $firstParagraph.BaseStyle = $bodyText
  Set-ParagraphStyleBase -Style $firstParagraph `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphJustify) `
    -FirstLineIndentCm $Profile.FirstLineIndentCm `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 0 `
    -LineSpacingRule $Profile.LineSpacingRule

  $sourceCode = Get-OrAddStyle -Document $Document -Name "Source Code" -Type ($wdStyleTypeParagraph)
  $sourceCode.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $sourceCode `
    -FontName "Courier New" `
    -FontSize $Profile.CodeSize `
    -Alignment ($wdAlignParagraphLeft) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 6 `
    -LineSpacingRule ($wdLineSpaceSingle)

  $captionedFigure = Get-OrAddStyle -Document $Document -Name "Captioned Figure" -Type ($wdStyleTypeParagraph)
  $captionedFigure.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $captionedFigure `
    -FontName $Profile.FontName `
    -FontSize $Profile.CaptionSize `
    -Alignment ($wdAlignParagraphCenter) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 6 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $Profile.LineSpacingRule

  $referenceItem = Get-OrAddStyle -Document $Document -Name "ReferenceItem" -Type ($wdStyleTypeParagraph)
  $referenceItem.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $referenceItem `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphJustify) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $Profile.LineSpacingRule
  $referenceItem.ParagraphFormat.LeftIndent = 0
  $referenceItem.ParagraphFormat.FirstLineIndent = ([single](-1 * $word.CentimetersToPoints(0.75)))

  $figure = Get-OrAddStyle -Document $Document -Name "Figure" -Type ($wdStyleTypeParagraph)
  $figure.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $figure `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphCenter) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 6 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $Profile.LineSpacingRule

  $gostKeywords = Get-OrAddStyle -Document $Document -Name "GostKeywords" -Type ($wdStyleTypeParagraph)
  $gostKeywords.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $gostKeywords `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphJustify) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 6 `
    -SpaceAfterPt 12 `
    -LineSpacingRule $Profile.LineSpacingRule
  $gostKeywords.Font.Bold = $true

  $myCustomStyle = Get-OrAddStyle -Document $Document -Name "MyCustomStyle" -Type ($wdStyleTypeParagraph)
  $myCustomStyle.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $myCustomStyle `
    -FontName $Profile.FontName `
    -FontSize $Profile.BodySize `
    -Alignment ($wdAlignParagraphLeft) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 0 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $Profile.LineSpacingRule
  $myCustomStyle.Font.Italic = $true

  $un1 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1" -Type ($wdStyleTypeParagraph)
  $un1.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un1 `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading1Size `
    -Alignment ($wdAlignParagraphCenter) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt $Profile.Heading1BeforePt `
    -SpaceAfterPt $Profile.Heading1AfterPt `
    -LineSpacingRule $Profile.LineSpacingRule
  $un1.Font.Bold = $true
  if ($Profile.UppercaseHeading1) { $un1.Font.AllCaps = $true } else { $un1.Font.AllCaps = $false }

  $un1NoToc = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading1NoTOC" -Type ($wdStyleTypeParagraph)
  $un1NoToc.BaseStyle = $un1
  Set-ParagraphStyleBase -Style $un1NoToc `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading1Size `
    -Alignment ($wdAlignParagraphCenter) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt $Profile.Heading1BeforePt `
    -SpaceAfterPt $Profile.Heading1AfterPt `
    -LineSpacingRule $Profile.LineSpacingRule
  $un1NoToc.Font.Bold = $true
  if ($Profile.UppercaseHeading1) { $un1NoToc.Font.AllCaps = $true } else { $un1NoToc.Font.AllCaps = $false }

  $un2 = Get-OrAddStyle -Document $Document -Name "UnnumberedHeading2" -Type ($wdStyleTypeParagraph)
  $un2.BaseStyle = $normal
  Set-ParagraphStyleBase -Style $un2 `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading2Size `
    -Alignment ($wdAlignParagraphLeft) `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt $Profile.Heading2BeforePt `
    -SpaceAfterPt $Profile.Heading2AfterPt `
    -LineSpacingRule $Profile.LineSpacingRule
  $un2.Font.Bold = $true

  $heading1 = $Document.Styles.Item($wdStyleHeading1)
  Set-ParagraphStyleBase -Style $heading1 `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading1Size `
    -Alignment $Profile.NumberedHeadingAlignment `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt $Profile.Heading1BeforePt `
    -SpaceAfterPt $Profile.Heading1AfterPt `
    -LineSpacingRule $Profile.LineSpacingRule
  $heading1.Font.Bold = $true
  if ($Profile.UppercaseHeading1) { $heading1.Font.AllCaps = $true } else { $heading1.Font.AllCaps = $false }

  $heading2 = $Document.Styles.Item($wdStyleHeading2)
  Set-ParagraphStyleBase -Style $heading2 `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading2Size `
    -Alignment $Profile.NumberedHeadingAlignment `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt $Profile.Heading2BeforePt `
    -SpaceAfterPt $Profile.Heading2AfterPt `
    -LineSpacingRule $Profile.LineSpacingRule
  $heading2.Font.Bold = $true

  $heading3 = $Document.Styles.Item($wdStyleHeading3)
  Set-ParagraphStyleBase -Style $heading3 `
    -FontName $Profile.FontName `
    -FontSize $Profile.Heading3Size `
    -Alignment $Profile.NumberedHeadingAlignment `
    -FirstLineIndentCm 0 `
    -SpaceBeforePt 12 `
    -SpaceAfterPt 6 `
    -LineSpacingRule $Profile.LineSpacingRule
  $heading3.Font.Bold = $true
  $heading3.Font.Italic = $true

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
    $markerStyle = Get-OrAddStyle -Document $Document -Name $markerName -Type ($wdStyleTypeCharacter)
    $markerStyle.Font.Name = $Profile.FontName
    $markerStyle.Font.Size = [single]$Profile.BodySize
  }

  $tableGost = Get-OrAddStyle -Document $Document -Name "TableStyleGost" -Type ($wdStyleTypeTable)
  $tableGost.Font.Name = $Profile.FontName
  $tableGost.Font.Size = [single]$Profile.TableSize
  $tableGost.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle

  $tableGostNoHeader = Get-OrAddStyle -Document $Document -Name "TableStyleGostNoHeader" -Type ($wdStyleTypeTable)
  $tableGostNoHeader.Font.Name = $Profile.FontName
  $tableGostNoHeader.Font.Size = [single]$Profile.TableSize
  $tableGostNoHeader.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle

  $tableContributors = Get-OrAddStyle -Document $Document -Name "TableStyleContributors" -Type ($wdStyleTypeTable)
  $tableContributors.Font.Name = $Profile.FontName
  $tableContributors.Font.Size = [single]$Profile.TableSize
  $tableContributors.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle

  $tableAbbreviations = Get-OrAddStyle -Document $Document -Name "TableStyleAbbreviations" -Type ($wdStyleTypeTable)
  $tableAbbreviations.Font.Name = $Profile.FontName
  $tableAbbreviations.Font.Size = [single]$Profile.TableSize
  $tableAbbreviations.ParagraphFormat.LineSpacingRule = $wdLineSpaceSingle
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

function Clear-Document {
  param($Document)
  $Document.Content.Delete()
}

function Build-StudyGuideReference {
  param($Document)

  Clear-Document -Document $Document
  $Document.Activate()
  $selection = $word.Selection

  Add-StyledParagraph -Selection $selection -Text "Министерство науки и высшего образования Российской Федерации" -Alignment ($wdAlignParagraphCenter)
  Add-StyledParagraph -Selection $selection -Text "Наименование организации" -Alignment ($wdAlignParagraphCenter)
  $selection.TypeParagraph()
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "И. О. Фамилия" -Alignment ($wdAlignParagraphCenter)
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "НАЗВАНИЕ РАБОТЫ" -StyleName "UnnumberedHeading1" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
  Add-StyledParagraph -Selection $selection -Text "Учебное пособие" -Alignment ($wdAlignParagraphCenter)
  $selection.TypeParagraph()
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Казань" -Alignment ($wdAlignParagraphCenter)
  Add-StyledParagraph -Selection $selection -Text "Издательство организации" -Alignment ($wdAlignParagraphCenter)
  Add-StyledParagraph -Selection $selection -Text "2026" -Alignment ($wdAlignParagraphCenter)
  $selection.InsertBreak($wdPageBreak) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "УДК 000" -Alignment ($wdAlignParagraphLeft)
  Add-StyledParagraph -Selection $selection -Text "ББК 000" -Alignment ($wdAlignParagraphLeft)
  Add-StyledParagraph -Selection $selection -Text "Авторский знак" -Alignment ($wdAlignParagraphLeft)
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Печатается по решению редакционно-издательского совета организации" -Alignment ($wdAlignParagraphLeft) -Italic
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Рецензенты:" -Alignment ($wdAlignParagraphLeft) -Italic
  Add-StyledParagraph -Selection $selection -Text "д-р наук, проф. И. О. Фамилия" -Alignment ($wdAlignParagraphLeft)
  Add-StyledParagraph -Selection $selection -Text "канд. наук, доц. И. О. Фамилия" -Alignment ($wdAlignParagraphLeft)
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Фамилия И. О. Название работы : учебное пособие / И. О. Фамилия. — Казань : Издательство организации, 2026. — 000 с." -Alignment ($wdAlignParagraphLeft)
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Краткая аннотация: работа содержит..., предназначено для..., подготовлено на кафедре..." -Alignment ($wdAlignParagraphLeft)
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "ISBN присваивается издательством после редакционно-издательской обработки" -Alignment ($wdAlignParagraphLeft)
  $selection.InsertBreak($wdPageBreak) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "СОДЕРЖАНИЕ" -StyleName "UnnumberedHeading1NoTOC" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
  Add-StyledParagraph -Selection $selection -Text "%TOC%" -Alignment ($wdAlignParagraphLeft)
  $selection.InsertBreak($wdPageBreak) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "%MAINTEXT%" -Alignment ($wdAlignParagraphLeft)
  $selection.InsertBreak($wdPageBreak) | Out-Null

  Add-StyledParagraph -Selection $selection -Text "ВЫХОДНЫЕ СВЕДЕНИЯ" -StyleName "UnnumberedHeading1NoTOC" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
  Add-StyledParagraph -Selection $selection -Text "УЧЕБНОЕ ИЗДАНИЕ" -Alignment ($wdAlignParagraphCenter) -Italic
  Add-StyledParagraph -Selection $selection -Text "Имя Отчество Фамилия" -Alignment ($wdAlignParagraphCenter)
  Add-StyledParagraph -Selection $selection -Text "НАЗВАНИЕ РАБОТЫ" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
  $selection.TypeParagraph()
  Add-StyledParagraph -Selection $selection -Text "Редактор __________________" -Alignment ($wdAlignParagraphLeft)
  Add-StyledParagraph -Selection $selection -Text "Подписано в печать __________________" -Alignment ($wdAlignParagraphLeft)
  Add-StyledParagraph -Selection $selection -Text "Формат, бумага, печать, тираж, заказ заполняются после редакционно-издательской обработки" -Alignment ($wdAlignParagraphLeft)
}

function Build-StudyGuideCover {
  param($Path, [hashtable]$Profile)

  $document = $word.Documents.Add()
  try {
    Set-PageLayout -Document $document -LeftCm $Profile.LeftMarginCm -RightCm $Profile.RightMarginCm -TopCm $Profile.TopMarginCm -BottomCm $Profile.BottomMarginCm -FooterCm $Profile.FooterDistanceCm
    Ensure-CommonStyles -Document $document -Profile $Profile
    Clear-Document -Document $document
    $document.Activate()
    $selection = $word.Selection
    Add-StyledParagraph -Selection $selection -Text "ОБЛОЖКА ПОДАЕТСЯ ОТДЕЛЬНЫМ ФАЙЛОМ" -StyleName "UnnumberedHeading1NoTOC" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
    $selection.TypeParagraph()
    Add-StyledParagraph -Selection $selection -Text "Авторы" -Alignment ($wdAlignParagraphCenter)
    Add-StyledParagraph -Selection $selection -Text "НАЗВАНИЕ РАБОТЫ" -StyleName "UnnumberedHeading1" -Alignment ($wdAlignParagraphCenter) -Bold -AllCaps
    Add-StyledParagraph -Selection $selection -Text "Учебное пособие" -Alignment ($wdAlignParagraphCenter)
    $selection.TypeParagraph()
    Add-StyledParagraph -Selection $selection -Text "Место для иллюстрации 300 dpi и выше" -Alignment ($wdAlignParagraphCenter) -Italic
    $selection.TypeParagraph()
    Add-StyledParagraph -Selection $selection -Text "2026" -Alignment ($wdAlignParagraphCenter)
    $document.SaveAs([ref]$Path)
  }
  finally {
    $document.Close()
  }
}

function Open-OrCreateDoc {
  param([string]$Path)

  if (Test-Path -LiteralPath $Path) {
    return $word.Documents.Open($Path)
  }

  $doc = $word.Documents.Add()
  $doc.SaveAs([ref]$Path)
  return $doc
}

$profiles = @{
  espd = @{
    FontName = "PT Serif"
    BodySize = 12.0
    CodeSize = 10.5
    CaptionSize = 11.0
    TableSize = 12.0
    LeftMarginCm = 2.0
    RightMarginCm = 2.0
    TopMarginCm = 2.0
    BottomMarginCm = 2.0
    FooterDistanceCm = 1.25
    FirstLineIndentCm = 1.25
    LineSpacingRule = $wdLineSpace1pt5
    Heading1Size = 12.0
    Heading2Size = 12.0
    Heading3Size = 12.0
    Heading1BeforePt = 18.0
    Heading1AfterPt = 12.0
    Heading2BeforePt = 12.0
    Heading2AfterPt = 6.0
    UppercaseHeading1 = $false
    NumberedHeadingAlignment = $wdAlignParagraphLeft
  }
  report = @{
    FontName = "PT Serif"
    BodySize = 12.0
    CodeSize = 10.5
    CaptionSize = 11.0
    TableSize = 12.0
    LeftMarginCm = 3.0
    RightMarginCm = 1.5
    TopMarginCm = 2.0
    BottomMarginCm = 2.0
    FooterDistanceCm = 1.25
    FirstLineIndentCm = 1.25
    LineSpacingRule = $wdLineSpace1pt5
    Heading1Size = 12.0
    Heading2Size = 12.0
    Heading3Size = 12.0
    Heading1BeforePt = 18.0
    Heading1AfterPt = 12.0
    Heading2BeforePt = 12.0
    Heading2AfterPt = 6.0
    UppercaseHeading1 = $false
    NumberedHeadingAlignment = $wdAlignParagraphLeft
  }
  dissertation = @{
    FontName = "Times New Roman"
    BodySize = 14.0
    CodeSize = 12.0
    CaptionSize = 12.0
    TableSize = 12.0
    LeftMarginCm = 2.5
    RightMarginCm = 1.0
    TopMarginCm = 2.0
    BottomMarginCm = 2.0
    FooterDistanceCm = 1.25
    FirstLineIndentCm = 1.25
    LineSpacingRule = $wdLineSpace1pt5
    Heading1Size = 14.0
    Heading2Size = 14.0
    Heading3Size = 13.0
    Heading1BeforePt = 18.0
    Heading1AfterPt = 18.0
    Heading2BeforePt = 18.0
    Heading2AfterPt = 12.0
    UppercaseHeading1 = $false
    NumberedHeadingAlignment = $wdAlignParagraphCenter
  }
  synopsis = @{
    FontName = "Times New Roman"
    BodySize = 12.0
    CodeSize = 10.5
    CaptionSize = 11.0
    TableSize = 11.0
    LeftMarginCm = 2.5
    RightMarginCm = 1.0
    TopMarginCm = 2.0
    BottomMarginCm = 2.0
    FooterDistanceCm = 1.25
    FirstLineIndentCm = 1.25
    LineSpacingRule = $wdLineSpace1pt5
    Heading1Size = 12.0
    Heading2Size = 12.0
    Heading3Size = 11.0
    Heading1BeforePt = 18.0
    Heading1AfterPt = 18.0
    Heading2BeforePt = 12.0
    Heading2AfterPt = 12.0
    UppercaseHeading1 = $false
    NumberedHeadingAlignment = $wdAlignParagraphCenter
  }
  "study-guide" = @{
    FontName = "Times New Roman"
    BodySize = 16.0
    CodeSize = 14.0
    CaptionSize = 14.0
    TableSize = 13.0
    LeftMarginCm = 1.9
    RightMarginCm = 1.9
    TopMarginCm = 1.9
    BottomMarginCm = 2.4
    FooterDistanceCm = 1.5
    FirstLineIndentCm = 1.25
    LineSpacingRule = $wdLineSpaceSingle
    Heading1Size = 16.0
    Heading2Size = 16.0
    Heading3Size = 16.0
    Heading1BeforePt = 0.0
    Heading1AfterPt = 30.0
    Heading2BeforePt = 60.0
    Heading2AfterPt = 30.0
    UppercaseHeading1 = $true
    NumberedHeadingAlignment = $wdAlignParagraphLeft
  }
}

$sourceTemplateMap = @{
  espd = "C:\projects\QuartoGost\ref\gostdown\demo-template-espd.docx"
  report = "C:\projects\QuartoGost\ref\gostdown\demo-template-report.docx"
}

try {
  $docMap = @{
    "study-guide" = "C:\projects\QuartoGost\resources\reference-docs\study-guide\reference.docx"
  }

  $filteredDocMap = @{}
  foreach ($entry in $docMap.GetEnumerator()) {
    if ($selectedTypes -contains $entry.Key) {
      $filteredDocMap[$entry.Key] = $entry.Value
    }
  }

  if ($filteredDocMap.ContainsKey("study-guide")) {
    $studyGuideDir = Split-Path -Path $filteredDocMap["study-guide"] -Parent
    if (-not (Test-Path -LiteralPath $studyGuideDir)) {
      New-Item -ItemType Directory -Force -Path $studyGuideDir | Out-Null
    }
  }

  foreach ($entry in $filteredDocMap.GetEnumerator()) {
    $doc = Open-OrCreateDoc -Path $entry.Value
    try {
      Set-PageLayout -Document $doc `
        -LeftCm $profiles[$entry.Key].LeftMarginCm `
        -RightCm $profiles[$entry.Key].RightMarginCm `
        -TopCm $profiles[$entry.Key].TopMarginCm `
        -BottomCm $profiles[$entry.Key].BottomMarginCm `
        -FooterCm $profiles[$entry.Key].FooterDistanceCm
      Ensure-CommonStyles -Document $doc -Profile $profiles[$entry.Key]

      if ($sourceTemplateMap.ContainsKey($entry.Key)) {
        Sync-CoverParagraphFormattingFromSource -TargetDocument $doc -SourcePath $sourceTemplateMap[$entry.Key]
      }

      if ($entry.Key -eq "study-guide") {
        Build-StudyGuideReference -Document $doc
      }

      $doc.Save()
    }
    finally {
      $doc.Close()
    }
  }

  $dissertationGenerator = Join-Path $PSScriptRoot "generate-dissertation-reference-doc.ps1"
  if (($selectedTypes -contains "dissertation") -and (Test-Path -LiteralPath $dissertationGenerator)) {
    & $dissertationGenerator
  }

  $synopsisGenerator = Join-Path $PSScriptRoot "generate-synopsis-reference-doc.ps1"
  if (($selectedTypes -contains "synopsis") -and (Test-Path -LiteralPath $synopsisGenerator)) {
    & $synopsisGenerator
  }

  $presentationGenerator = Join-Path $PSScriptRoot "generate-presentation-reference-pptx.ps1"
  if (($selectedTypes -contains "presentation") -and (Test-Path -LiteralPath $presentationGenerator)) {
    & $presentationGenerator
  }

}
finally {
  $word.Quit()
}







