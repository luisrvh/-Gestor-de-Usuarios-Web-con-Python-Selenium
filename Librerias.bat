@echo off
echo Buscando Python instalado...

REM Buscar python en el PATH
for /f "delims=" %%P in ('where python 2^>nul') do (
    set PYTHON=%%P
    goto :found
)

:found
IF NOT DEFINED PYTHON (
    echo Python no está instalado o no está en el PATH.
    echo Instala Python y marca "Add Python to PATH".
    pause
    exit /b 1
)

echo Usando Python en:
echo %PYTHON%
echo.

REM Actualizar pip
"%PYTHON%" -m pip install --upgrade pip

REM Instalar librerías
"%PYTHON%" -m pip install --upgrade selenium pandas numpy openpyxl ttkbootstrap

IF ERRORLEVEL 1 (
    echo Error al instalar los paquetes.
) ELSE (
    echo Instalación completada exitosamente.
)

pause
