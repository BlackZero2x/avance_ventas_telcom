@echo off
REM ============================================================
REM setup_tarea_asistencia.bat
REM Crea la tarea "AsistenciaPlanilla_930" que ejecuta el informe
REM de asistencia PLANILLA de lunes a sabado a las 9:30 AM.
REM Ejecutar como developer7 (o el usuario que corre los otros scripts).
REM ============================================================

set PYTHON=C:\proyectos\.venv\Scripts\python.exe
set SCRIPT=C:\proyectos\AVANCE_MOVISTAR\asistencia_planilla\asistencia_planilla.py
set USER=developer7
set TAREA=AsistenciaPlanilla_930

echo.
echo === Eliminando tarea existente (si existe) ===
schtasks /delete /tn "\%TAREA%" /f 2>nul && echo [OK] Eliminada %TAREA% || echo [--] No existia %TAREA%

echo.
echo === Creando tarea %TAREA% (lunes a sabado, 09:30) ===
schtasks /create /tn "\%TAREA%" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\"" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 09:30 ^
  /ru %USER% /it /f

if errorlevel 1 (
    echo [ERROR] No se pudo crear la tarea. Verifica que ejecutas como administrador.
) else (
    echo [OK] Tarea creada: %TAREA% — lun a sab a las 09:30
)

echo.
echo === Verificacion ===
schtasks /query /tn "\%TAREA%" /fo TABLE 2>nul | findstr /i "%TAREA%" && echo     ^^ OK || echo [ERROR] Tarea no encontrada

echo.
echo Listo. Los logs se generan en: C:\proyectos\AVANCE_MOVISTAR\logs\asistencia_YYYYMMDD.log
pause
