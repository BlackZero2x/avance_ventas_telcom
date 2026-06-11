@echo off
REM ============================================================
REM setup_tareas_cortes.bat
REM Elimina y recrea las tareas de alertas e informes de cortes.
REM Ejecutar como developer7.
REM ============================================================

set PYTHON=C:\proyectos\.venv\Scripts\python.exe
set SCRIPT=C:\proyectos\AVANCE_MOVISTAR\cortes_ventas\cortes_ventas.py
set USER=developer7

echo.
echo === Eliminando tareas existentes ===
for %%T in (CortesAlerta_12PM CortesAlerta_2PM CortesAlerta_4PM CortesAlerta_6PM CortesAlerta_CIERRE CortesInforme_12PM CortesInforme_2PM CortesInforme_4PM CortesInforme_6PM CortesInforme_CIERRE) do (
    schtasks /delete /tn "\%%T" /f 2>nul && echo [OK] Eliminada %%T || echo [--] No existia %%T
)

echo.
echo === Creando tareas de ALERTA (5 min antes del corte) ===

schtasks /create /tn "\CortesAlerta_12PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 12PM --alerta" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 11:55 ^
  /ru %USER% /it /f
echo [OK] CortesAlerta_12PM (11:55)

schtasks /create /tn "\CortesAlerta_2PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 2PM --alerta" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 13:55 ^
  /ru %USER% /it /f
echo [OK] CortesAlerta_2PM (13:55)

schtasks /create /tn "\CortesAlerta_4PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 4PM --alerta" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 15:55 ^
  /ru %USER% /it /f
echo [OK] CortesAlerta_4PM (15:55)

schtasks /create /tn "\CortesAlerta_6PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 6PM --alerta" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 17:55 ^
  /ru %USER% /it /f
echo [OK] CortesAlerta_6PM (17:55)

schtasks /create /tn "\CortesAlerta_CIERRE" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte CIERRE --alerta" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 08:55 ^
  /ru %USER% /it /f
echo [OK] CortesAlerta_CIERRE (08:55)

echo.
echo === Creando tareas de INFORME (+5 min despues del corte) ===

schtasks /create /tn "\CortesInforme_12PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 12PM" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 12:05 ^
  /ru %USER% /it /f
echo [OK] CortesInforme_12PM (12:05)

schtasks /create /tn "\CortesInforme_2PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 2PM" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 14:05 ^
  /ru %USER% /it /f
echo [OK] CortesInforme_2PM (14:05)

schtasks /create /tn "\CortesInforme_4PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 4PM" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 16:05 ^
  /ru %USER% /it /f
echo [OK] CortesInforme_4PM (16:05)

schtasks /create /tn "\CortesInforme_6PM" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte 6PM" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 18:05 ^
  /ru %USER% /it /f
echo [OK] CortesInforme_6PM (18:05)

schtasks /create /tn "\CortesInforme_CIERRE" ^
  /tr "\"%PYTHON%\" \"%SCRIPT%\" --corte CIERRE" ^
  /sc WEEKLY /d MON,TUE,WED,THU,FRI,SAT /st 09:05 ^
  /ru %USER% /it /f
echo [OK] CortesInforme_CIERRE (09:05)

echo.
echo === Verificacion final ===
for %%T in (CortesAlerta_12PM CortesAlerta_2PM CortesAlerta_4PM CortesAlerta_6PM CortesAlerta_CIERRE CortesInforme_12PM CortesInforme_2PM CortesInforme_4PM CortesInforme_6PM CortesInforme_CIERRE) do (
    schtasks /query /tn "\%%T" /fo TABLE 2>nul | findstr /i "%%T" && echo     ^^ OK || echo [ERROR] %%T no encontrada
)

echo.
echo Listo. Los logs se generan en: C:\proyectos\AVANCE_MOVISTAR\logs\cortes_YYYYMMDD.log
pause
