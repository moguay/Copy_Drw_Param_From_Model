@echo off
IF EXIST %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam.html goto getConfirmation

IF NOT EXIST %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam.html goto deployCode

:getConfirmation
set /p confirmDeploy =Copy_param found, do you want to update [y/n] ?: 
if "%confirmDeploy%"=="y". goto backupCode
if "%confirmDeploy%"=="n". goto cancelDeploy

:backupCode
ECHO Create Backup...
IF NOT EXIST %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\backup mkdir IF NOT EXIST %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\backup
XCOPY /Q /Y %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam*.* %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\backup >nul 2>&1

:deployCode
ECHO Start Deploying...
XCOPY /Q /Y .\CopyParam\CopyParam.html %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\ >nul 2>&1
XCOPY /Q /Y .\CopyParam\CopyParam.css %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\ >nul 2>&1
XCOPY /Q /Y .\CopyParam\CopyParam.js %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\ >nul 2>&1
XCOPY /Q /Y .\CopyParam\CopyParam_default.ini %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\ >nul 2>&1

IF NOT EXIST %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam.ini COPY %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam_default.ini %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam.ini >nul 2>&1

REGEDIT.EXE /S ".\IE_Rondo-ACTIVEX.reg"

ECHO.
ECHO Installation finished.
ECHO Copy_Param in windows location : %HOMEDRIVE%%HOMEPATH%\Personal\WebLinks\CopyParam.html
ECHO Copy_Param in Creo    location : $home\WebLinks\CopyParam.html
PAUSE
explorer "%HOMEDRIVE%%HOMEPATH%\Personal\WebLinks"
goto end

:cancelDeploy
ECHO DEPLOY CANCELLED
PAUSE
goto end

:end