@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType dissertation -InputFile examples\dissertation-demo\main.qmd -OutputDir build\dissertation-demo -Name dissertation-demo
