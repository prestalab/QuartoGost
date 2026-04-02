@echo off
powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType envelopes -AddressList templates\common\data\sample-addresses.tsv -OutputDir build\envelopes-demo -Name envelopes-demo -SenderName "ФГБОУ ВО Пример" -SenderAddress "119991, г. Москва, Ленинские горы, д. 1"
