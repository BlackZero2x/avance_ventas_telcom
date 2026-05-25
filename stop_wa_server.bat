@echo off
:: Detiene el servidor wa_server.js al final del dia laboral.
:: Ejecutar por el Programador de Tareas (ej. lunes-viernes 18:00).

setlocal

set LOG_FILE=C:\proyectos\AVANCE_MOVISTAR\logs\wa_server_start.log

echo [%date% %time%] Deteniendo servidor WhatsApp (fin de dia)... >> "%LOG_FILE%"
taskkill /F /IM node.exe >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Proceso node.exe detenido. >> "%LOG_FILE%"
) else (
    echo [%date% %time%] No habia proceso node.exe activo. >> "%LOG_FILE%"
)

endlocal
exit /b 0
