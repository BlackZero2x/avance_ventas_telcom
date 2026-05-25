"""
execution_log.py
Registro diario de estado de ejecucion del orquestador.

Formato de execution_log.json:
{
  "2026-05-20": {
    "trigger_detected": true,
    "trigger_id": "19e4632aac313e8d",
    "trigger_hora": "11:30:17",
    "avance_ok": true,
    "modulos": {
      "backs": true,
      "jefes": true,
      ...
    },
    "all_sent": true,
    "finish_hora": "11:45:00",
    "failure_reason": null
  }
}
"""
import json
import logging
from datetime import date, datetime
from pathlib import Path

_LOG_PATH = Path(__file__).parent.parent.parent / "execution_log.json"


def _leer() -> dict:
    if not _LOG_PATH.exists():
        return {}
    try:
        return json.loads(_LOG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _escribir(data: dict):
    try:
        _LOG_PATH.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        logging.error(f"execution_log: no se pudo escribir: {e}")


def _hoy() -> str:
    return date.today().isoformat()


def _hora() -> str:
    return datetime.now().strftime("%H:%M:%S")


def registrar_trigger(trigger_id: str):
    data = _leer()
    hoy  = _hoy()
    data.setdefault(hoy, {})
    data[hoy].update({
        "trigger_detected": True,
        "trigger_id":       trigger_id,
        "trigger_hora":     _hora(),
        "avance_ok":        False,
        "all_sent":         False,
        "failure_reason":   None,
        "modulos":          {},
    })
    _escribir(data)


def registrar_avance_ok():
    data = _leer()
    hoy  = _hoy()
    data.setdefault(hoy, {})
    data[hoy]["avance_ok"]  = True
    data[hoy]["avance_hora"] = _hora()
    _escribir(data)


def registrar_avance_fallo(motivo: str):
    data = _leer()
    hoy  = _hoy()
    data.setdefault(hoy, {})
    data[hoy]["avance_ok"]      = False
    data[hoy]["failure_reason"] = motivo
    _escribir(data)


def registrar_modulo(nombre: str, ok: bool):
    data = _leer()
    hoy  = _hoy()
    data.setdefault(hoy, {}).setdefault("modulos", {})
    data[hoy]["modulos"][nombre] = ok
    _escribir(data)


def registrar_fin(ok: bool, motivo: str = None):
    data = _leer()
    hoy  = _hoy()
    data.setdefault(hoy, {})
    data[hoy]["all_sent"]      = ok
    data[hoy]["finish_hora"]   = _hora()
    if motivo:
        data[hoy]["failure_reason"] = motivo
    _escribir(data)


def estado_hoy() -> dict:
    return _leer().get(_hoy(), {})
