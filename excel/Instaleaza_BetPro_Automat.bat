@echo off
title BET PRO - instalare pornire automata
echo.
echo  BET PRO - creez pornirea automata zilnica la ora 09:00 ...
echo.
if not exist "%~dp0BetPro_Automat.vbs" (
  echo  EROARE: BetPro_Automat.vbs nu e in acest folder.
  echo  Pune acest fisier, BetPro_Automat.vbs si BET_PRO.xlsm in acelasi folder.
  echo.
  pause
  exit /b 1
)
schtasks /delete /tn "BET PRO" /f >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "$a=New-ScheduledTaskAction -Execute 'wscript.exe' -Argument '\"%~dp0BetPro_Automat.vbs\"'; $t=New-ScheduledTaskTrigger -Daily -At 9am; $s=New-ScheduledTaskSettingsSet -WakeToRun -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries; Register-ScheduledTask -TaskName 'BET PRO' -Action $a -Trigger $t -Settings $s -Force | Out-Null"
if errorlevel 1 (
  echo  Varianta 1 nu a mers, incerc varianta 2 ...
  schtasks /create /tn "BET PRO" /tr "wscript.exe \"%~dp0BetPro_Automat.vbs\"" /sc daily /st 09:00 /it /f
)
echo.
schtasks /query /tn "BET PRO" >nul 2>&1
if errorlevel 1 (
  echo  NU a mers. Fa o poza la aceasta fereastra si trimite-o.
) else (
  echo  GATA. Botul porneste singur in fiecare zi la 09:00 si iti trimite emailul.
  echo.
  echo  Vrei sa-l testezi acum? Inchide intai Excel-ul, apoi apasa o tasta.
  pause >nul
  schtasks /run /tn "BET PRO"
  echo.
  echo  Pornit. Emailul ar trebui sa vina in 2-3 minute.
)
echo.
pause
