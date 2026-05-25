# Registra la tarea WATCHDOG_WA_SERVER en el Programador de Tareas de Windows.
# Ejecutar UNA sola vez como Administrador:
#   powershell -ExecutionPolicy Bypass -File registrar_watchdog_tarea.ps1

$xml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Verifica cada 15 min que wa_server.js este activo (lun-vie 08:45-18:00). Reinicia automaticamente si falla.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>2026-05-15T08:45:00</StartBoundary>
      <EndBoundary>2099-12-31T23:59:59</EndBoundary>
      <Enabled>true</Enabled>
      <ScheduleByWeek>
        <DaysOfWeek>
          <Monday/>
          <Tuesday/>
          <Wednesday/>
          <Thursday/>
          <Friday/>
        </DaysOfWeek>
        <WeeksInterval>1</WeeksInterval>
      </ScheduleByWeek>
      <Repetition>
        <Interval>PT15M</Interval>
        <Duration>PT9H15M</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </CalendarTrigger>
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
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Enabled>true</Enabled>
    <StartWhenAvailable>true</StartWhenAvailable>
  </Settings>
  <Actions>
    <Exec>
      <Command>C:\proyectos\AVANCE_MOVISTAR\watchdog_wa_server.bat</Command>
      <WorkingDirectory>C:\proyectos\AVANCE_MOVISTAR\</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
'@

$xmlPath = "$env:TEMP\watchdog_wa_task.xml"
$xml | Out-File -FilePath $xmlPath -Encoding Unicode

schtasks /Create /TN "WATCHDOG_WA_SERVER" /XML $xmlPath /F

if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Tarea WATCHDOG_WA_SERVER creada correctamente." -ForegroundColor Green
    Write-Host "     Ejecuta cada 15 min lun-vie 08:45-18:00." -ForegroundColor Cyan
    Write-Host "     Verifica en: Programador de Tareas > Biblioteca de Programador de Tareas" -ForegroundColor Cyan
} else {
    Write-Host "[ERROR] No se pudo crear la tarea. Ejecuta este script como Administrador." -ForegroundColor Red
}

Remove-Item $xmlPath -ErrorAction SilentlyContinue
