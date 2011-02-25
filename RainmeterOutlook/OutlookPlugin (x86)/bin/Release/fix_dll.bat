@ECHO OFF

:start
ECHO **
ECHO **  POST-BUILD
ECHO **
ECHO.
ECHO **
ECHO **  Bridging x86 plugin for C++ access...
ECHO **
CS_DLL_for_C.exe %2
GOTO ErrHandler%ERRORLEVEL%

:ErrHandler0
IF %1 == Release (
	ECHO.
	ECHO **
	ECHO **  Copying result to _compiled folder
	ECHO **
	IF NOT EXIST ..\..\..\_compiled MD ..\..\..\_compiled
	IF NOT EXIST "..\..\..\_compiled\x32" MD "..\..\..\_compiled\x32"
	COPY /Y %2 "..\..\..\_compiled\x32"
)

GOTO success

:ErrHandler1
SET _PostBuild=no filename given
GOTO error

:ErrHandler2
SET _PostBuild=.dll file missing (%2)
GOTO error

:ErrHandler3
SET _PostBuild=ildasm.exe is missing
GOTO error

:error
ECHO.
ECHO **
ECHO **  POST-BUILD: ERROR %_PostBuild% (%_PostBuildErr%)
ECHO **
GOTO end

:success
ECHO.
ECHO **
ECHO **  POST-BUILD: SUCCESS
ECHO **
:end
