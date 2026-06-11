import sys
sys.argv = ["cortes_ventas.py", "--corte", "CIERRE"]

sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\whatsapp_server")
sys.path.insert(0, r"C:\proyectos\AVANCE_MOVISTAR\cortes_ventas")

import cortes_ventas as cv
import wa_client as wac

MI_NUMERO = "51975155264@c.us"
OrigClient = wac.WhatsAppClient

class TestClient(OrigClient):
    def send_image(self, to, ruta, caption=""):
        to_safe = to.encode("cp1252", errors="replace").decode("cp1252")
        print(f"[TEST] send_image -> {MI_NUMERO} (original: {to_safe})")
        return super().send_image(MI_NUMERO, ruta, caption=caption)
    def send_text(self, to, texto):
        to_safe = to.encode("cp1252", errors="replace").decode("cp1252")
        print(f"[TEST] send_text -> {MI_NUMERO} (original: {to_safe})")
        return super().send_text(MI_NUMERO, texto)

wac.WhatsAppClient = TestClient
cv.WhatsAppClient  = TestClient

cv.main()
