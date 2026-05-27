@echo off
:: Verifica si wa_server.js esta corriendo y lo inicia si no.
:: Ejecutar antes de main_v2.py en el Programador de Tareas.

setlocal

set NODE_EXE=C:\Program Files\nodejs\node.exe
set SERVER_JS=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js
set SERVER_DIR=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server
set PYTHON_EXE=C:\proyectos\.venv\Scripts\python.exe
set WA_CLIENT=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py
set LOG_FILE=C:\proyectos\AVANCE_MOVISTAR\logs\wa_server_start.log
set MY_NUMBER=51975155264@c.us

echo [%date% %time%] Verificando servidor WhatsApp... >> "%LOG_FILE%"

:: Si el servidor ya responde al health check, no hace falta reiniciar
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Servidor ya activo — no se reinicia. >> "%LOG_FILE%"
    call :enviar_confirmacion
    goto :fin
)

:: El servidor no responde: matar node y chromium y arrancar fresco
echo [%date% %time%] Servidor no responde — reiniciando... >> "%LOG_FILE%"
taskkill /F /IM node.exe >NUL 2>&1
taskkill /F /IM chrome.exe >NUL 2>&1
ping 127.0.0.1 -n 4 >NUL

:: Limpiar todos los lock files que Chrome puede dejar al cerrar mal
set PROFILE_DIR=%SERVER_DIR%\session_data\_IGNORE_automation_session
if exist "%PROFILE_DIR%\SingletonLock"    del /F /Q "%PROFILE_DIR%\SingletonLock"    >NUL 2>&1
if exist "%PROFILE_DIR%\SingletonSocket"  del /F /Q "%PROFILE_DIR%\SingletonSocket"  >NUL 2>&1
if exist "%PROFILE_DIR%\SingletonCookie"  del /F /Q "%PROFILE_DIR%\SingletonCookie"  >NUL 2>&1
if exist "%PROFILE_DIR%\Default\LOCK"     del /F /Q "%PROFILE_DIR%\Default\LOCK"     >NUL 2>&1
if exist "%PROFILE_DIR%\Default\lock"     del /F /Q "%PROFILE_DIR%\Default\lock"     >NUL 2>&1
if exist "%PROFILE_DIR%\Default\lockfile" del /F /Q "%PROFILE_DIR%\Default\lockfile" >NUL 2>&1
echo [%date% %time%] Lock files de Chromium limpiados. >> "%LOG_FILE%"
ping 127.0.0.1 -n 3 >NUL

cd /d "%SERVER_DIR%"
start /min "wa_server" "%NODE_EXE%" "%SERVER_JS%"

:: Esperar hasta 5 minutos a que el servidor este listo (60 x 5s)
set intentos=0
:esperar
set /a intentos+=1
if %intentos% GTR 60 (
    echo [%date% %time%] ERROR: Servidor no respondio tras 5 minutos. >> "%LOG_FILE%"
    exit /b 1
)
ping 127.0.0.1 -n 6 >NUL
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :esperar

echo [%date% %time%] Servidor iniciado tras %intentos% intentos. >> "%LOG_FILE%"
call :enviar_confirmacion

:fin
endlocal
exit /b 0

:: -------------------------------------------------------
:enviar_confirmacion
set reintento=0
:reintentar
set /a reintento+=1
"%PYTHON_EXE%" "%WA_CLIENT%" --send-to "%MY_NUMBER%" --message "[OK] Servidor WhatsApp activo" >NUL 2>&1
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
