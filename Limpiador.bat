@echo off
rem "Elimina los archivos viejos termporales."

rem "del /S /F /Q C:\Users\%USERNAME%\AppData\Local\Temp\*.*"
rmdir /S /Q C:\Users\%USERNAME%\AppData\Local\Temp\


rem "del /S /F /Q C:\Windows\Temp\*.*"
rmdir /S /Q C:\Windows\Temp\

rem "del /q/f/s %TEMP%\"

echo Carpetas temporales eliminadas :-P