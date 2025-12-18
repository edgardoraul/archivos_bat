@echo off
setlocal enabledelayedexpansion

echo [+] Solicitando seleccion de carpeta...

:: --- Parte 1: Usar PowerShell para obtener el dialogo de carpeta ---
for /f "delims=" %%i in ('powershell.exe -NoProfile -Command "& {Add-Type -AssemblyName System.Windows.Forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'Selecciona la carpeta con imagenes WEBP'; $f.ShowDialog() | Out-Null; $f.SelectedPath}"') do set "FOLDER_PATH=%%i"

:: Verificar si el usuario cancelo el dialogo
if not defined FOLDER_PATH (
    echo [!] Operacion cancelada por el usuario.
    pause
    exit /b
)

echo [+] Carpeta seleccionada: %FOLDER_PATH%

:: --- Parte 2: Verificar si ImageMagick esta instalado ---
where magick >nul 2>nul
if %errorlevel% neq 0 (
    echo [X] ERROR: ImageMagick no esta instalado o no se encuentra en el PATH.
    echo [!] Por favor, instalalo desde https://imagemagick.org/ y asegurate
    echo [!] de marcar la casilla "Add to system path" durante la instalacion.
    pause
    exit /b
)

:: Moverse a la carpeta seleccionada
cd /d "%FOLDER_PATH%"
if %errorlevel% neq 0 (
    echo [X] ERROR: No se pudo acceder a la carpeta: %FOLDER_PATH%
    pause
    exit /b
)

:: --- Parte 3: Conversion y renombrado ---
echo [+] Iniciando conversion...
set count=1

:: *** CAMBIO AQUI ***
:: Usar DIR /B para listar archivos, es mas robusto.
for /f "delims=" %%f in ('dir /b *.webp 2^>nul') do (
    
    set "outfile=!count!.jpg"
    echo [ ] Convirtiendo "%%f"   --->   "!outfile!"

    :: Ejecutar la conversion con ImageMagick
    magick "%%f" "!outfile!"
    
    if !errorlevel! equ 0 (
        echo [V] Conversion exitosa. Borrando original "%%f"...
        del "%%f"
        
        :: Incrementar el contador solo si la conversion fue exitosa
        set /a count+=1
    ) else (
        echo [X] ERROR: Fallo la conversion de "%%f". El original NO sera borrado.
        del "!outfile!" 2>nul
    )
    echo.
)

echo [+] Proceso finalizado.
echo [+] Se convirtieron un total de !count!-1 imagenes.
pause
endlocal