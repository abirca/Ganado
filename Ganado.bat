@echo off
REM ===== COPIA DEL EXCEL =====
set ORIGEN="C:\gestor_proveedores\Financiero.xlsx"
set DESTINO="C:\gestor_proveedores\destino\Financiero_%date:~-4%%date:~3,2%%date:~0,2%.xlsx"

REM Copiar con fecha en el nombre para no sobrescribir
copy %ORIGEN% %DESTINO%

REM ===== INICIO DEL SERVIDOR DJANGO =====
cd C:\gestor_proveedores

REM Obtener la IP local del equipo
for /f "tokens=2 delims=:" %%i in ('ipconfig ^| findstr /c:"IPv4"') do set IP=%%i
set IP=%IP: =%

REM Iniciar servidor Django en esa IP
start cmd /k "python manage.py runserver %IP%:8000"

REM Esperar unos segundos para que el servidor arranque
timeout /t 15 >nul

REM Abrir navegador
start http://%IP%:8000
