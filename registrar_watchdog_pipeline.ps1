# Registra la tarea WATCHDOG_AVANCE_PIPELINE en el Programador de Tareas de Windows.
# Ejecutar UNA sola vez como Administrador:
#   powershell -ExecutionPolicy Bypass -File registrar_watchdog_pipeline.ps1
#
# Horarios lun-vie: 10:00, 11:00, 12:00 (manana) | 14:00, 16:00, 18:00 (tarde)
# - Manana: deteccion temprana y aviso a jefes/Jose si no hay trigger a las 13:00+
# - Tarde:  fallback robusto si el pipeline fallo o el trigger llego tarde

$pythonExe = "C:\proyectos\.venv\Scripts\python.exe"
$script    = "C:\proyectos\AVANCE_MOVISTAR\watchdog_process.py"
$workdir   = "C:\proyectos\AVANCE_MOVISTAR"

function New-Trigger($hora) {
    return @"
    <CalendarTrigger>
      <StartBoundary>2026-05-22T${hora}:00</StartBoundary>
      <EndBoundary>2099-12-31T23:59:59</EndBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday/><Tuesday/><Wednesday/><Thursday/><Friday/>
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
    </CalendarTrigger>
"@
}

$triggers = (New-Trigger "10:00") + (New-Trigger "11:00") + (New-Trigger "12:00") +
            (New-Trigger "14:00") + (New-Trigger "16:00") + (New-Trigger "18:00")

$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Watchdog pipeline AVANCE: detecta fallos y avisa si no llega trigger (lun-vie 10-18h).</Description>
  </RegistrationInfo>
  <Triggers>
$triggers
  </Triggers>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
    <StartWhenAvailable>true</StartWhenAvailable>
  </Settings>
  <Actions>
    <Exec>
      <Command>$pythonExe</Command>
      <Arguments>$script</Arguments>
      <WorkingDirectory>$workdir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

$xmlPath = "$env:TEMP\watchdog_pipeline_task.xml"
$xml | Out-File -FilePath $xmlPath -Encoding Unicode

schtasks /Create /TN "WATCHDOG_AVANCE_PIPELINE" /XML $xmlPath /F

if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Tarea WATCHDOG_AVANCE_PIPELINE creada correctamente." -ForegroundColor Green
    Write-Host "     Ejecuta lun-vie a las 10:00, 11:00, 12:00, 14:00, 16:00 y 18:00." -ForegroundColor Cyan
    Write-Host "     Verifica en: Programador de Tareas > Biblioteca de Programador de Tareas" -ForegroundColor Cyan
} else {
    Write-Host "[ERROR] No se pudo crear la tarea. Ejecuta este script como Administrador." -ForegroundColor Red
}

Remove-Item $xmlPath -ErrorAction SilentlyContinue
