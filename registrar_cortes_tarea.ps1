# registrar_cortes_tarea.ps1
# Registra 10 tareas en el Programador de Tareas de Windows:
#   5 alertas de supervisores pendientes (10 min antes de cada corte)
#   5 informes de corte (en el horario de cada corte)
# Ejecución: Lunes a Sábado únicamente.
# Ejecutar como Administrador.

$python  = "C:\proyectos\.venv\Scripts\python.exe"
$script  = "C:\proyectos\AVANCE_MOVISTAR\cortes_ventas.py"
$logDir  = "C:\proyectos\AVANCE_MOVISTAR\logs"
$usuario = $env:USERNAME

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Force $logDir | Out-Null }

# ── Tareas de ALERTA (10 min antes) ──────────────────────────────────────────
$alertas = @(
    @{ Nombre = "CortesAlerta_12PM";   Hora = "11:50"; Corte = "12PM"   },
    @{ Nombre = "CortesAlerta_2PM";    Hora = "13:50"; Corte = "2PM"    },
    @{ Nombre = "CortesAlerta_4PM";    Hora = "15:50"; Corte = "4PM"    },
    @{ Nombre = "CortesAlerta_6PM";    Hora = "17:50"; Corte = "6PM"    },
    @{ Nombre = "CortesAlerta_CIERRE"; Hora = "08:50"; Corte = "CIERRE" }
)

# ── Tareas de INFORME (en el horario del corte) ───────────────────────────────
$informes = @(
    @{ Nombre = "CortesInforme_12PM";   Hora = "12:10"; Corte = "12PM"   },
    @{ Nombre = "CortesInforme_2PM";    Hora = "14:10"; Corte = "2PM"    },
    @{ Nombre = "CortesInforme_4PM";    Hora = "16:10"; Corte = "4PM"    },
    @{ Nombre = "CortesInforme_6PM";    Hora = "18:10"; Corte = "6PM"    },
    @{ Nombre = "CortesInforme_CIERRE"; Hora = "09:00"; Corte = "CIERRE" }
)

function Registrar-Tarea($nombre, $hora, $corte, $extraArgs, $descripcion) {
    $logFile = "$logDir\tarea_$corte.log"
    $accion = New-ScheduledTaskAction `
        -Execute $python `
        -Argument "`"$script`" --corte $corte $extraArgs >> `"$logFile`" 2>&1" `
        -WorkingDirectory "C:\proyectos\AVANCE_MOVISTAR"

    # Trigger semanal Lunes a Sábado (DaysOfWeek: Monday=0x02 Tuesday=0x04 ... Saturday=0x40)
    # New-ScheduledTaskTrigger -Weekly acepta -DaysOfWeek como array de strings
    $trigger = New-ScheduledTaskTrigger `
        -Weekly `
        -At $hora `
        -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday, Saturday `
        -WeeksInterval 1

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -StartWhenAvailable `
        -DontStopIfGoingOnBatteries `
        -RunOnlyIfNetworkAvailable

    $principal = New-ScheduledTaskPrincipal `
        -UserId $usuario `
        -LogonType Interactive `
        -RunLevel Highest

    Unregister-ScheduledTask -TaskName $nombre -Confirm:$false -ErrorAction SilentlyContinue

    Register-ScheduledTask `
        -TaskName    $nombre `
        -Action      $accion `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description $descripcion

    Write-Host "Registrada: $nombre a las $hora (Lun-Sab)" -ForegroundColor Green
}

Write-Host "`n== Registrando tareas de ALERTA (10 min antes) ==" -ForegroundColor Cyan
foreach ($t in $alertas) {
    Registrar-Tarea `
        -nombre      $t.Nombre `
        -hora        $t.Hora `
        -corte       $t.Corte `
        -extraArgs   "--alerta" `
        -descripcion "Alerta supervisores pendientes — corte $($t.Corte) — Lun a Sab"
}

Write-Host "`n== Registrando tareas de INFORME (en el corte) ==" -ForegroundColor Cyan
foreach ($t in $informes) {
    Registrar-Tarea `
        -nombre      $t.Nombre `
        -hora        $t.Hora `
        -corte       $t.Corte `
        -extraArgs   "" `
        -descripcion "Informe corte $($t.Corte) — ventas supervisores — Lun a Sab"
}

Write-Host "`nListo. 10 tareas registradas (Lunes a Sabado)." -ForegroundColor Cyan
Write-Host "Verifica en Programador de Tareas > Biblioteca de Programador de Tareas." -ForegroundColor Gray
Write-Host "`nPrueba manual:" -ForegroundColor Yellow
Write-Host "  Alerta:  python cortes_ventas.py --corte 12PM --alerta" -ForegroundColor Yellow
Write-Host "  Informe: python cortes_ventas.py --corte 12PM" -ForegroundColor Yellow
