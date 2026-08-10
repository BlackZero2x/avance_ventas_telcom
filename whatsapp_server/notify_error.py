"""
Script auxiliar invocado por wa_server.js para enviar alertas por correo
cuando el servidor WhatsApp detecta un error crítico de sesión.

Uso (desde Node via child_process):
    python notify_error.py --tipo "Promise was collected" --detalle "..."
"""
import sys
import os
import argparse
import logging
from pathlib import Path

# Rutas relativas a la raíz del proyecto
BASE_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(BASE_DIR))

logging.basicConfig(level=logging.WARNING, format="%(levelname)s: %(message)s")


def _autenticar_gmail():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    token_path = BASE_DIR / "token.json"
    creds_path = BASE_DIR / "credentials.json"
    # El token.json compartido del proyecto se emitio con gmail.modify + drive
    # (lo usa main_v2.py para leer el correo trigger). gmail.modify ya incluye
    # permiso de envio, asi que pedimos EXACTAMENTE los scopes del token para
    # evitar "invalid_scope: Bad Request" en el refresh (bug del 18-20/07/2026,
    # cuando pediamos solo gmail.send que el token no tenia).
    scopes = [
        "https://www.googleapis.com/auth/gmail.modify",
        "https://www.googleapis.com/auth/drive",
    ]

    if not token_path.exists():
        print(f"[notify_error] token.json no encontrado en {token_path}", file=sys.stderr)
        sys.exit(1)

    creds = Credentials.from_authorized_user_file(str(token_path), scopes)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tipo", required=True, help="Tipo de error detectado")
    parser.add_argument("--detalle", default="", help="Descripción adicional del error")
    parser.add_argument(
        "--destinatario",
        default="augusto.moreno@auren.com.pe",
        help="Correo al que enviar la alerta",
    )
    args = parser.parse_args()

    asunto = f"[WA Server] Alerta: {args.tipo}"

    cuerpo = f"""Se detectó un error crítico en el servidor de WhatsApp (wa_server.js).

ERROR DETECTADO
===============
Tipo:    {args.tipo}
Detalle: {args.detalle or '(sin detalle adicional)'}

QUÉ PASÓ
=========
El servidor WhatsApp perdió la conexión con la página de WhatsApp Web.
Todos los envíos de mensajes (imágenes, texto, archivos) fallarán hasta
que la sesión sea restaurada.

El servidor intentará reiniciarse automáticamente. Si los errores continúan
después de 5 minutos, es necesario intervención manual.

QUÉ HACER SI EL PROBLEMA PERSISTE
===================================
1. Abrir PowerShell y ejecutar:
       Stop-Process -Name node -Force
2. Esperar 5 segundos y reiniciar:
       cd C:\\proyectos\\AVANCE_MOVISTAR\\whatsapp_server
       node wa_server.js
3. Si sigue fallando (aparece QR de nuevo):
       - Escanear el QR con el celular de la cuenta WhatsApp
       - Esperar hasta ver "Cliente WhatsApp listo" en la consola

LOGS
====
Ruta: C:\\proyectos\\AVANCE_MOVISTAR\\whatsapp_server\\logs\\

Este correo fue enviado automáticamente por wa_server.js.
"""

    try:
        service = _autenticar_gmail()
        from modules.shared.gmail_helper import GmailHelper
        helper = GmailHelper(service)
        ok = helper.send_simple_email(args.destinatario, asunto, cuerpo)
        if ok:
            print(f"[notify_error] Alerta enviada a {args.destinatario}")
        else:
            print("[notify_error] Fallo al enviar la alerta", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"[notify_error] Error inesperado: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
