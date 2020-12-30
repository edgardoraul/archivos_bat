 @echo off
Title Copia de Seguridad
echo              =========================================
echo              =                                       =
echo              =         Copia de Seguridad            =
echo              =                                       =
echo              =========================================
echo.
echo Este comando copiara archivos y carpetas del disco D al disco de resguardo en la unidad F.
pause
@echo off

ROBOCOPY D:\ F:\EDGARD\ /S  /R:0 /w:0