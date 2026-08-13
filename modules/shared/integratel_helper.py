"""
Descarga el adjunto Excel de riesgo (PRE-PENALIDAD) enviado por Integratel
y lo acumula sobre el histórico existente.

Remitente fijo: eduardo.pinco@integratel.com.pe
Asunto fijo:    "Reporte de casos observados sujetos a PRE-PENALIDAD - Socio AUREN (MASIVO)"

Cada adjunto nuevo trae solo los casos recientes (no el histórico completo), así
que se combina con lo ya acumulado en destino_path y se deduplica por ORDER_KEY
(gana la fila del adjunto nuevo si un ORDER_KEY se repite). Así el archivo sigue
sirviendo para meses anteriores — AVANCE.py ya filtra por coincidencia de
`peticion` del período actual, así que las filas de meses viejos no afectan el
resultado, solo quedan disponibles por si se necesitan.

Si no llega correo nuevo, el archivo acumulado existente se conserva tal cual
(AVANCE.py siempre tiene datos disponibles).
"""
import base64
import logging
import os

import pandas as pd

INTEGRATEL_SENDER  = "eduardo.pinco@integratel.com.pe"
INTEGRATEL_SUBJECT = "Reporte de casos observados sujetos a PRE-PENALIDAD - Socio AUREN (MASIVO)"


def descargar_adjunto_integratel(gmail_service, destino_path):
    """
    Busca el correo más reciente de Integratel con el asunto fijo, descarga su
    primer adjunto .xlsx/.xls, y lo acumula sobre `destino_path` deduplicando
    por ORDER_KEY (la fila del adjunto nuevo gana sobre la ya acumulada).

    Devuelve True si se descargó y acumuló un adjunto nuevo, False si no se
    encontró correo/adjunto (en cuyo caso `destino_path` conserva lo que ya
    tenía, sin cambios).
    """
    try:
        query = f'from:{INTEGRATEL_SENDER} subject:"{INTEGRATEL_SUBJECT}" has:attachment'
        results = gmail_service.users().messages().list(
            userId="me", q=query, maxResults=5
        ).execute()
        mensajes = results.get("messages", [])
        if not mensajes:
            logging.warning(f"[Integratel] Sin correos de {INTEGRATEL_SENDER} con el asunto esperado")
            return False

        # El primer resultado es el más reciente (orden por defecto de Gmail)
        msg_id = mensajes[0]["id"]
        msg = gmail_service.users().messages().get(userId="me", id=msg_id).execute()

        parte_adjunto = None
        for part in msg.get("payload", {}).get("parts", []) or []:
            filename = part.get("filename", "")
            if filename.lower().endswith((".xlsx", ".xls")):
                parte_adjunto = part
                break

        if parte_adjunto is None:
            logging.warning("[Integratel] Correo encontrado pero sin adjunto .xlsx/.xls")
            return False

        attachment_id = parte_adjunto["body"]["attachmentId"]
        attachment = gmail_service.users().messages().attachments().get(
            userId="me", messageId=msg_id, id=attachment_id
        ).execute()

        data = base64.urlsafe_b64decode(attachment["data"])
        os.makedirs(os.path.dirname(destino_path), exist_ok=True)

        nuevo_path = destino_path + ".nuevo.tmp"
        with open(nuevo_path, "wb") as f:
            f.write(data)

        df_nuevo = pd.read_excel(nuevo_path)
        if os.path.exists(destino_path):
            df_previo = pd.read_excel(destino_path)
            # concat con el nuevo AL FINAL: drop_duplicates(keep="last") conserva
            # la fila del adjunto nuevo cuando el ORDER_KEY se repite.
            df_final = pd.concat([df_previo, df_nuevo], ignore_index=True)
            df_final = df_final.drop_duplicates(subset="ORDER_KEY", keep="last")
        else:
            df_final = df_nuevo

        df_final.to_excel(destino_path, index=False)
        os.remove(nuevo_path)

        logging.info(
            f"[OK] Integratel: adjunto acumulado en {destino_path} "
            f"({len(df_nuevo)} filas nuevas, {len(df_final)} filas totales tras deduplicar)"
        )
        return True

    except Exception as e:
        logging.error(f"[Integratel] Error descargando/acumulando adjunto: {e}")
        return False
