cd c:\esther
goto boucle


:relance

start wscript.exe "%cd%\plugins\skype\Skype.vbs"



:boucle

tasklist /fi "imagename eq wscript.exe" | find "wscript.exe"
if %ERRORLEVEL% neq 0 goto relance

goto boucle