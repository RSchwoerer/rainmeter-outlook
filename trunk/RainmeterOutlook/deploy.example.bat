@ECHO OFF

SET RAINMETER-ROOT=C:\Program Files\Rainmeter\
SET RAINMETER-SKINS=C:\Users\YourName\Documents\Rainmeter\Skins\
SET BIT=64

SET RAINMETER=%RAINMETER-ROOT%Rainmeter.exe
SET RAINMETER-PLUGINS=%RAINMETER-ROOT%Plugins\

ECHO Stopping Rainmeter
START "" /WAIT "%RAINMETER%" !RainmeterQuit
: wait 3 seconds for rainmeter to close
@ping -n 3 localhost> nul

: copy files to rmskin template
COPY _compiled\x64\OutlookPlugin.dll Template\Plugins\64bit> nul
COPY _compiled\x32\OutlookPlugin.dll Template\Plugins\32bit> nul
DEL Template\Skins /Q
XCOPY Skins\* Template\Skins /S /V /Q /Y> nul

: copy files to rainmeter
ECHO Copying files
COPY _compiled\x%BIT%\OutlookPlugin.dll "%RAINMETER-PLUGINS%"> nul
XCOPY Skins\* "%RAINMETER-SKINS%" /S /V /Q /Y> nul

ECHO Restarting Rainmeter
START "" "%RAINMETER%"