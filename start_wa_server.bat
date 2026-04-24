@echo off
:: Verifica si wa_server.js esta corriendo y lo inicia si no.
:: Ejecutar antes de main_v2.py en el Programador de Tareas.

setlocal

set NODE_EXE=C:\Program Files\nodejs\node.exe
set SERVER_JS=C:\AVANCE_MOVISTAR\whatsapp_server\wa_server.js
set SERVER_DIR=C:\AVANCE_MOVISTAR\whatsapp_server
set PYTHON_EXE=C:\Users\developer2\AppData\Local\Programs\Python\Python310\python.exe
set WA_CLIENT=C:\AVANCE_MOVISTAR\whatsapp_server\wa_client.py
set LOG_FILE=C:\AVANCE_MOVISTAR\logs\wa_server_start.log
set MY_NUMBER=51975155264@c.us

echo [%date% %time%] Verificando servidor WhatsApp... >> "%LOG_FILE%"

:: Verificar si el servidor ya responde en /health
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Servidor WhatsApp ya esta activo. >> "%LOG_FILE%"
    call :enviar_confirmacion
    goto :fin
)

:: No responde — matar cualquier node residual e iniciar de nuevo
echo [%date% %time%] Servidor no disponible. Iniciando wa_server.js... >> "%LOG_FILE%"
taskkill /F /IM node.exe >NUL 2>&1
ping 127.0.0.1 -n 4 >NUL

cd /d "%SERVER_DIR%"
start /min "wa_server" "%NODE_EXE%" "%SERVER_JS%"

:: Esperar hasta 90 segundos a que el servidor este listo (18 x 5s)
set intentos=0
:esperar
set /a intentos+=1
if %intentos% GTR 18 (
    echo [%date% %time%] ERROR: Servidor no respondio tras 90s. >> "%LOG_FILE%"
    exit /b 1
)
ping 127.0.0.1 -n 6 >NUL
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :esperar

echo [%date% %time%] Servidor WhatsApp recuperado tras %intentos% intentos. >> "%LOG_FILE%"
call :enviar_confirmacion

:fin
endlocal
exit /b 0

:: -------------------------------------------------------
:enviar_confirmacion
set reintento=0
:reintentar
set /a reintento+=1
"%PYTHON_EXE%" "%WA_CLIENT%" --send-text "%MY_NUMBER%" "[OK] Servidor WhatsApp activo" >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Confirmacion enviada exitosamente (intento %reintento%). >> "%LOG_FILE%"
    exit /b 0
)
echo [%date% %time%] Intento %reintento% fallido. >> "%LOG_FILE%"
if %reintento% LSS 3 (
    ping 127.0.0.1 -n 6 >NUL
    goto :reintentar
)
echo [%date% %time%] ADVERTENCIA: No se pudo enviar mensaje de confirmacion tras 3 intentos. >> "%LOG_FILE%"
exit /b 1
