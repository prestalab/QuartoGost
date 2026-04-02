@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType synopsis -InputFile examples\synopsis-demo\main.qmd -OutputDir build\synopsis-demo -Name synopsis-demo
