@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType espd -InputFile examples\espd-demo\main.qmd -OutputDir build\espd-demo -Name espd-demo -EmbedFonts
