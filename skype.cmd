cd c:\esther
goto boucle


:relance
echo "Lancement de skype.vbs"
start wscript.exe "%cd%\plugins\skype\Skype.vbs"



:boucle
ping localhost -n 15 >NUL
tasklist /fi "imagename eq wscript.exe" | find "wscript.exe"
if %ERRORLEVEL% neq 0 goto relance

goto boucle