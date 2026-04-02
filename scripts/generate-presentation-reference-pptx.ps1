param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Path $PSScriptRoot -Parent
$target = Join-Path $root "resources\reference-pptx\reference.pptx"
$logoPath = Join-Path $root "resources\assets\images\logo.emf"
New-Item -ItemType Directory -Force -Path (Split-Path -Path $target -Parent) | Out-Null

function Get-OleColor {
  param(
    [int]$Red,
    [int]$Green,
    [int]$Blue
  )

  Add-Type -AssemblyName System.Drawing
  return [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb($Red, $Green, $Blue))
}

function Set-ShapeTextStyle {
  param(
    $Shape,
    [string]$FontName,
    [double]$FontSize,
    [int]$Color,
    [switch]$Bold
  )

  try {
    if (-not $Shape.HasTextFrame) { return }
    if (-not $Shape.TextFrame.HasText) { return }
    $textRange = $Shape.TextFrame.TextRange
    $textRange.Font.Name = $FontName
    $textRange.Font.Size = $FontSize
    $textRange.Font.Color.RGB = $Color
    $textRange.Font.Bold = $(if ($Bold.IsPresent) { -1 } else { 0 })
  } catch { }
}

function Set-ShapeGeometry {
  param(
    $Shape,
    [double]$Left,
    [double]$Top,
    [double]$Width,
    [double]$Height
  )

  $Shape.Left = $Left
  $Shape.Top = $Top
  $Shape.Width = $Width
  $Shape.Height = $Height
}

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1

$ppLayoutTitle = 1
$ppLayoutText = 2
$ppPlaceholderTitle = 1
$ppPlaceholderBody = 2
$msoTrue = -1

$blue = Get-OleColor -Red 25 -Green 74 -Blue 128
$lightBlue = Get-OleColor -Red 232 -Green 240 -Blue 250
$darkGray = Get-OleColor -Red 64 -Green 64 -Blue 64
$black = Get-OleColor -Red 0 -Green 0 -Blue 0
$white = Get-OleColor -Red 255 -Green 255 -Blue 255

try {
  $presentation = $pp.Presentations.Add()
  try {
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540

    $master = $presentation.SlideMaster
    $master.Background.Fill.ForeColor.RGB = $white

    try {
      $master.HeadersFooters.SlideNumber.Visible = $msoTrue
      $master.HeadersFooters.Footer.Visible = $msoTrue
      $master.HeadersFooters.Footer.Text = "Dissertation presentation"
    } catch { }

    $footerLine = $master.Shapes.AddLine(36, 504, 924, 504)
    $footerLine.Line.ForeColor.RGB = $blue
    $footerLine.Line.Weight = 1.25

    $footerText = $master.Shapes.AddTextbox(1, 36, 508, 500, 18)
    $footerText.TextFrame.TextRange.Text = "QuartoGost / dissertation presentation"
    $footerText.TextFrame.TextRange.Font.Name = "Calibri"
    $footerText.TextFrame.TextRange.Font.Size = 10
    $footerText.TextFrame.TextRange.Font.Color.RGB = $blue

    $titleLayout = $master.CustomLayouts.Item(1)
    $contentLayout = $master.CustomLayouts.Item(2)
    $sectionLayout = $master.CustomLayouts.Item(3)
    $titleOnlyLayout = $master.CustomLayouts.Item(6)

    if (Test-Path -LiteralPath $logoPath) {
      $null = $titleLayout.Shapes.AddPicture($logoPath, $msoTrue, $msoTrue, 36, 24, 76, 76)
      $null = $sectionLayout.Shapes.AddPicture($logoPath, $msoTrue, $msoTrue, 36, 24, 60, 60)
    }

    foreach ($layout in @($titleLayout, $contentLayout, $sectionLayout, $titleOnlyLayout)) {
      $layout.Background.Fill.ForeColor.RGB = $white
    }

    $sectionBand = $sectionLayout.Shapes.AddShape(1, 0, 0, 960, 540)
    $sectionBand.Fill.ForeColor.RGB = $lightBlue
    $sectionBand.Line.Visible = 0
    $sectionBand.ZOrder(1) | Out-Null

    foreach ($shape in $titleLayout.Shapes) {
      try {
        if ($shape.PlaceholderFormat.Type -eq $ppPlaceholderTitle) {
          Set-ShapeGeometry -Shape $shape -Left 120 -Top 156 -Width 720 -Height 96
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 26 -Color $blue -Bold
          $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        } else {
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 16 -Color $darkGray
          try { $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 } catch { }
        }
      } catch { }
    }

    foreach ($shape in $contentLayout.Shapes) {
      try {
        if ($shape.PlaceholderFormat.Type -eq $ppPlaceholderTitle) {
          Set-ShapeGeometry -Shape $shape -Left 42 -Top 20 -Width 876 -Height 36
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 22 -Color $blue -Bold
        } elseif ($shape.PlaceholderFormat.Type -eq $ppPlaceholderBody) {
          Set-ShapeGeometry -Shape $shape -Left 48 -Top 78 -Width 852 -Height 384
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 18 -Color $black
        } else {
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 16 -Color $black
        }
      } catch { }
    }

    foreach ($shape in $sectionLayout.Shapes) {
      try {
        if ($shape.PlaceholderFormat.Type -eq $ppPlaceholderTitle) {
          Set-ShapeGeometry -Shape $shape -Left 96 -Top 176 -Width 768 -Height 96
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 28 -Color $blue -Bold
          $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        } else {
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 16 -Color $darkGray
        }
      } catch { }
    }

    foreach ($shape in $titleOnlyLayout.Shapes) {
      try {
        if ($shape.PlaceholderFormat.Type -eq $ppPlaceholderTitle) {
          Set-ShapeGeometry -Shape $shape -Left 42 -Top 24 -Width 876 -Height 44
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 22 -Color $blue -Bold
        } else {
          Set-ShapeTextStyle -Shape $shape -FontName "Calibri" -FontSize 16 -Color $black
        }
      } catch { }
    }

    $presentation.SaveAs($target)
  }
  finally {
    $presentation.Close()
  }
}
finally {
  $pp.Quit()
}
