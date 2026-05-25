@echo off
:: Watchdog del servidor WhatsApp.
:: El Programador de Tareas lo ejecuta cada 15 minutos (lun-vie, 08:00-18:00).
:: Solo actua si el health falla — no mata ni reinicia si el servidor esta OK.

setlocal

set NODE_EXE=C:\Program Files\nodejs\node.exe
set SERVER_JS=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_server.js
set SERVER_DIR=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server
set PYTHON_EXE=C:\proyectos\.venv\Scripts\python.exe
set WA_CLIENT=C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py
set LOG_FILE=C:\proyectos\AVANCE_MOVISTAR\logs\watchdog_wa.log
set MY_NUMBER=51975155264@c.us

echo [%date% %time%] Watchdog: verificando health... >> "%LOG_FILE%"

:: Verificar si el servidor responde
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Watchdog: servidor OK, sin accion. >> "%LOG_FILE%"
    goto :fin
)

echo [%date% %time%] Watchdog: health FALLO — iniciando recuperacion... >> "%LOG_FILE%"

:: Verificar si node esta corriendo
tasklist /FI "IMAGENAME eq node.exe" 2>NUL | find /I "node.exe" >NUL
if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] Watchdog: node.exe no encontrado, lanzando servidor... >> "%LOG_FILE%"
    goto :lanzar
)

:: Node esta corriendo pero no responde — puede estar bloqueado o inicializando
:: Esperar 60s mas antes de forzar reinicio (puede ser que Chrome aun este arrancando)
echo [%date% %time%] Watchdog: node corre pero no responde, esperando 60s... >> "%LOG_FILE%"
ping 127.0.0.1 -n 61 >NUL

"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] Watchdog: servidor respondio tras espera, OK. >> "%LOG_FILE%"
    goto :fin
)

:: Sigue sin responder — reinicio forzado (matar AMBOS: node + chrome)
echo [%date% %time%] Watchdog: sigue sin responder, reiniciando... >> "%LOG_FILE%"
taskkill /F /IM node.exe   >NUL 2>&1
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

:lanzar
cd /d "%SERVER_DIR%"
start /min "wa_server" "%NODE_EXE%" "%SERVER_JS%"
echo [%date% %time%] Watchdog: servidor relanzado, esperando hasta 5 min... >> "%LOG_FILE%"

:: Esperar hasta 5 minutos a que responda (60 x 5s)
set intentos=0
:esperar
set /a intentos+=1
if %intentos% GTR 60 (
    echo [%date% %time%] Watchdog: ERROR — servidor no respondio tras 5 min. >> "%LOG_FILE%"
    goto :fin
)
ping 127.0.0.1 -n 6 >NUL
"%PYTHON_EXE%" "%WA_CLIENT%" --health >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :esperar

echo [%date% %time%] Watchdog: servidor recuperado tras %intentos% intentos. >> "%LOG_FILE%"

:: Notificar recuperacion
"%PYTHON_EXE%" "%WA_CLIENT%" --send-to "%MY_NUMBER%" --message "[Watchdog] Servidor WhatsApp recuperado automaticamente" >NUL 2>&1

:fin
endlocal
exit /b 0
