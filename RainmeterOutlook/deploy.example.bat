@ECHO OFF

SET RAINMETER-ROOT=C:\Program Files\Rainmeter\
SET RAINMETER-SKINS=C:\Users\YourName\Documents\Rainmeter\Skins\
SET BIT=64

SET RAINMETER=%RAINMETER-ROOT%Rainmeter.exe
SET RAINMETER-PLUGINS=%RAINMETER-ROOT%Plugins\

ECHO Stopping Rainmeter
START "" /WAIT "%RAINMETER%" !RainmeterQuit
: wait 2 seconds for rainmeter to close
@ping -n 2 localhost> nul

ECHO Copying files
: copy plugin file
COPY _compiled\x%BIT%\OutlookPlugin.dll "%RAINMETER-PLUGINS%"> nul
: copy skin files
XCOPY Skins\* "%RAINMETER-SKINS%" /S /V /Q /Y> nul

ECHO Restarting Rainmeter
START "" "%RAINMETER%"