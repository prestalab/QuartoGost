@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType presentation -InputFile examples\presentation-demo\main.qmd -OutputDir build\presentation-demo -Name presentation-demo
