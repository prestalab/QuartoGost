@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType study-guide -InputFile templates\study-guide\study-guide-template.qmd -OutputDir build\study-guide-demo -Name study-guide-demo
