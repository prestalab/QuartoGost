param(
  [string]$Quarto = "quarto",
  [string]$Julia = "julia",
  [string]$Pandoc = "pandoc"
)

function Show-ToolStatus {
  param(
    [string]$Name,
    [string]$Command,
    [string]$VersionArgument = "--version"
  )

  $tool = Get-Command $Command -ErrorAction SilentlyContinue
  if ($null -eq $tool) {
    Write-Host ("[MISSING] {0} ({1})" -f $Name, $Command)
    return
  }

  Write-Host ("[FOUND] {0}: {1}" -f $Name, $tool.Source)
  & $Command $VersionArgument
  Write-Host ""
}

Show-ToolStatus -Name "Quarto" -Command $Quarto
Show-ToolStatus -Name "Julia" -Command $Julia
Show-ToolStatus -Name "Pandoc" -Command $Pandoc

try {
  $word = New-Object -ComObject Word.Application
  $version = $word.Version
  $word.Quit()
  Write-Host ("[FOUND] Microsoft Word COM: version {0}" -f $version)
} catch {
  Write-Host "[MISSING] Microsoft Word COM"
}
