@echo off
setlocal enabledelayedexpansion

REM Directorio del proyecto y archivo Excel
set "SOURCE_FILE=F:\OneDrive\PROGRAMA\Ganado-pruebas\Financiero.xlsx"

REM Carpeta base para copias
set "BACKUP_BASE_DIR=F:\OneDrive\PROGRAMA\copias"

REM Obtener fecha y hora
for /f "tokens=1-4 delims=/ " %%a in ('date /t') do (
    set "DATE=%%d-%%b-%%c"
)
for /f "tokens=1-2 delims=: " %%a in ('time /t') do (
    set "TIME=%%a-%%b"
)

REM Limpiar caracteres no válidos
set "DATETIME=%DATE%_%TIME%"
set "DATETIME=%DATETIME:/=-%"
set "DATETIME=%DATETIME::=-%"
set "DATETIME=%DATETIME: =_%"

REM Crear carpeta de copia
set "BACKUP_DIR=%BACKUP_BASE_DIR%\copia_%DATETIME%"
mkdir "%BACKUP_DIR%"

REM Copiar solo el archivo Excel
copy "%SOURCE_FILE%" "%BACKUP_DIR%" >nul
echo Copia de Financiero.xlsx completada en: %BACKUP_DIR%

REM Cambiar al directorio del proyecto
cd /d F:\OneDrive\PROGRAMA\Ganado-pruebas

REM Obtener IP local
for /f "tokens=2 delims=:" %%f in ('ipconfig ^| findstr /C:"IPv4"') do (
    set IP=%%f
    goto :continue
)

:continue
set IP=%IP: =%
echo Tu IP local es: %IP%
start http://%IP%:8000

REM Ejecutar el servidor de Django
python manage.py runserver 0.0.0.0:8000
pause

