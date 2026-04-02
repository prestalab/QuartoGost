@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType report -InputFile examples\report-demo\main.qmd -OutputDir build\report-demo -Name report-demo -EmbedFonts -Counters
