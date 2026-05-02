@echo off
chcp 65001 > nul
title Kevin Lust - Installateur

:: Telecharge toujours la derniere version du script depuis GitHub, puis l'execute.
:: Le .bat n'a jamais besoin d'etre remis a jour.

set SCRIPT_URL=https://raw.githubusercontent.com/KevinDSM/kevin-lust-musiques/main/Install-KevinLust.ps1
set TMP_SCRIPT=%TEMP%\Install-KevinLust.ps1

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { Invoke-WebRequest -Uri '%SCRIPT_URL%' -OutFile '%TMP_SCRIPT%' -UseBasicParsing -TimeoutSec 15 } catch { Write-Host '[WARN] Impossible de telecharger la mise a jour du script. Utilisation de la version locale.' -ForegroundColor Yellow; Copy-Item '%~dp0Install-KevinLust.ps1' '%TMP_SCRIPT%' -ErrorAction SilentlyContinue }"

powershell -NoProfile -ExecutionPolicy Bypass -File "%TMP_SCRIPT%"
