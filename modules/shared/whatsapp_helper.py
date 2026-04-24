"""
Helper para envío de mensajes por WhatsApp Web.

Fix 2026-03-26 (v2.2):
- Múltiples XPATHs para el cuadro de búsqueda (WhatsApp cambió su estructura HTML)
- Timeouts más largos (hasta 60s para búsqueda)
- _get_search_box_with_fallback(): intenta 4 selectores diferentes
- Mejor manejo de errores y logging descriptivo
- Reintentos automáticos si falla la primera búsqueda
"""
import logging
import time
import os

import win32gui
import win32con
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


class WhatsAppHelper:
    def __init__(self, driver=None):
        self.driver = driver

    # ── Métodos internos ───────────────────────────────────────────

    def _bring_edge_to_front(self):
        """
        Trae la ventana de Microsoft Edge (WhatsApp Web) al frente usando win32gui.
        Tras procesos largos de Excel el foco del SO queda en Excel y el driver
        de Edge pierde el DOM. Forzar el foco al navegador antes de cada envío
        soluciona el TimeoutException de forma consistente.
        """
        try:
            def _enum_callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    # WhatsApp Web en Edge tiene "WhatsApp" en el título
                    if "WhatsApp" in title and "Edge" in title:
                        results.append(hwnd)
                    # Fallback: cualquier ventana de Edge
                    elif "Microsoft Edge" in title or "msedge" in title.lower():
                        results.append(hwnd)

            edge_windows = []
            win32gui.EnumWindows(_enum_callback, edge_windows)

            if edge_windows:
                hwnd = edge_windows[0]
                # Restaurar si está minimizado
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1)  # Dar tiempo al SO para cambiar el foco
                logging.info("  Edge traído al frente correctamente.")
            else:
                logging.warning("  No se encontró ventana de Edge para traer al frente.")
        except Exception as e:
            logging.warning(f"  win32gui no pudo traer Edge al frente: {e}")

    def _ensure_whatsapp_tab(self):
        """
        1. Trae Edge al frente (win32gui).
        2. Verifica que el DOM de WhatsApp responda.
        3. Si no responde, navega a web.whatsapp.com y espera hasta 90s.
        """
        try:
            self._bring_edge_to_front()

            # Chequeo rápido: si el buscador ya responde, no recargar
            search_xpaths = [
                '//div[@contenteditable="true"][@data-tab="3"]',
                '//div[@aria-label="Buscador de chats"]',
                '//div[@aria-label="Search input textbox"]',
                '//input[@placeholder*="Buscar"]',
                '//div[@class*="search"][@contenteditable="true"]',
            ]
            
            for xpath in search_xpaths:
                try:
                    elements = self.driver.find_elements(By.XPATH, xpath)
                    if elements:
                        logging.info("  WhatsApp Web ya estaba listo (búsqueda encontrada)")
                        return
                except Exception:
                    pass

            # DOM no responde → recargar
            logging.info("  Reactivando conexión con WhatsApp Web...")
            self.driver.get("https://web.whatsapp.com")
            time.sleep(5)  # Dar tiempo a cargar

            # Intentar múltiples XPATHs en orden de especificidad
            xpaths = [
                '//div[@contenteditable="true"][@data-tab="3"]',
                '//div[@aria-label="Buscador de chats"]',
                '//div[@aria-label="Search input textbox"]',
                '//input[@placeholder*="Buscar"]',
                '//div[@id="side"]',
                '//div[@id="app"]',
            ]
            
            for xpath in xpaths:
                try:
                    WebDriverWait(self.driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    logging.info("  WhatsApp Web reactivado correctamente.")
                    time.sleep(3)
                    return
                except TimeoutException:
                    continue
                except Exception as e:
                    logging.warning(f"  Error con xpath {xpath}: {e}")
                    continue

            logging.warning("  WhatsApp Web no respondió tras reactivación, pero se continúa.")
        except Exception as e:
            logging.warning(f"  No se pudo reactivar WhatsApp Web: {e}")

    def _get_search_box_with_fallback(self, timeout=60):
        """
        Obtiene el cuadro de búsqueda de chats con múltiples intentos.
        Intenta 5 XPATHs diferentes (WhatsApp cambió su estructura).
        """
        self._ensure_whatsapp_tab()
        
        xpaths_to_try = [
            '//div[@contenteditable="true"][@data-tab="3"]',
            '//div[@aria-label="Buscador de chats"]',
            '//div[@aria-label="Search input textbox"]',
            '//input[@placeholder*="Buscar"]',
            '//div[@class*="search"][@contenteditable="true"]',
        ]

        for attempt, xpath in enumerate(xpaths_to_try, 1):
            try:
                logging.info(f"  Buscando cuadro de búsqueda (intento {attempt}/{len(xpaths_to_try)})...")
                element = WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                logging.info(f"  ✓ Cuadro de búsqueda encontrado (selector #{attempt})")
                return element
            except TimeoutException:
                logging.warning(f"  Timeout en selector #{attempt}, probando siguiente...")
                continue
            except NoSuchElementException:
                logging.warning(f"  Selector #{attempt} no encontrado, probando siguiente...")
                continue
            except Exception as e:
                logging.warning(f"  Error en selector #{attempt}: {e}")
                continue

        # Si nada funcionó, intentar un último refresco
        logging.warning("  Ningún selector funcionó, intentando refresco final...")
        self.driver.refresh()
        time.sleep(5)
        
        try:
            element = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
                )
            )
            return element
        except Exception as e:
            logging.error(f"  No se pudo obtener el cuadro de búsqueda tras refresco: {e}")
            raise TimeoutException(
                "No se pudo encontrar el cuadro de búsqueda de WhatsApp. "
                "Verifica que WhatsApp Web esté abierto en Edge y que hayas escaneado el QR."
            )

    def _clear_and_type(self, element, text):
        """Limpia el campo de forma confiable y escribe el texto."""
        try:
            element.click()
            time.sleep(0.5)
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(Keys.DELETE)
            time.sleep(0.3)
            element.send_keys(text)
        except Exception as e:
            logging.error(f"Error al escribir en el campo: {e}")
            raise

    def _open_chat(self, name, timeout=20):
        """
        Busca un contacto o grupo por nombre y abre el chat.
        Retorna True si lo encontró, False si no.
        
        Intenta múltiples selectores para encontrar el resultado.
        """
        try:
            search_box = self._get_search_box_with_fallback()
            self._clear_and_type(search_box, name)
            logging.info(f"  Buscando chat: '{name}'...")
            time.sleep(3)

            # Intentar múltiples selectores para encontrar el resultado
            selectors = [
                (By.XPATH, f'//span[@title="{name}"]'),
                (By.XPATH, f'//div[@title="{name}"]'),
                (By.XPATH, f'//*[contains(@title, "{name}")]'),
                (By.XPATH, f'//div[contains(., "{name}")]//ancestor::div[@role="button"]'),
                (By.XPATH, f'//span[contains(., "{name}")]//ancestor::div[@role="button"]'),
            ]

            for selector_type, selector in selectors:
                try:
                    result = WebDriverWait(self.driver, timeout).until(
                        EC.presence_of_element_located((selector_type, selector))
                    )
                    result.click()
                    logging.info(f"  ✓ Chat '{name}' abierto exitosamente")
                    time.sleep(2)
                    return True
                except TimeoutException:
                    continue
                except Exception:
                    continue

            logging.error(f"  ❌ No se encontró el chat '{name}' (probados 5 selectores)")
            return False

        except Exception as e:
            logging.error(f"  Error abriendo chat '{name}': {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def _get_message_box(self, timeout=15):
        """
        Obtiene el cuadro de escritura del chat abierto.
        Intenta múltiples XPATHs.
        """
        xpaths_to_try = [
            '//div[@contenteditable="true"][@data-tab="10"]',
            '//div[@aria-label="Cuadro de escritura del mensaje"]',
            '//div[@aria-label="Message input textbox"]',
            '//div[@class*="message"][@contenteditable="true"]',
            '//div[@contenteditable="true"][@role="textbox"]',
        ]

        for xpath in xpaths_to_try:
            try:
                element = WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                return element
            except TimeoutException:
                continue
            except Exception:
                continue

        logging.error("No se pudo encontrar el cuadro de mensajes")
        raise TimeoutException("No se pudo encontrar el cuadro de escritura de mensajes")

    # ── API pública ────────────────────────────────────────────────

    def send_message_to_contact(self, contact_name, message):
        """Envía un mensaje de texto a un contacto."""
        try:
            if not self.driver:
                logging.error("WhatsApp Web no está inicializado")
                return False

            logging.info(f"Enviando mensaje a contacto: {contact_name}")

            if not self._open_chat(contact_name):
                return False

            msg_box = self._get_message_box()
            msg_box.click()
            time.sleep(0.5)
            msg_box.send_keys(message)
            time.sleep(0.5)
            msg_box.send_keys(Keys.ENTER)
            time.sleep(2)

            logging.info(f"✓ Mensaje enviado a {contact_name} exitosamente")
            return True

        except Exception as e:
            logging.error(f"Error enviando mensaje a {contact_name}: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def send_message_to_group(self, group_name, message):
        """Envía un mensaje de texto a un grupo."""
        try:
            if not self.driver:
                logging.error("WhatsApp Web no está inicializado")
                return False

            logging.info(f"Enviando mensaje al grupo: {group_name}")

            if not self._open_chat(group_name):
                return False

            msg_box = self._get_message_box()
            msg_box.click()
            time.sleep(0.5)
            msg_box.send_keys(message)
            time.sleep(0.5)
            msg_box.send_keys(Keys.ENTER)
            time.sleep(2)

            logging.info(f"✓ Mensaje enviado al grupo {group_name} exitosamente")
            return True

        except Exception as e:
            logging.error(f"Error enviando mensaje al grupo {group_name}: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False

    def send_image_to_group(self, group_name, image_path, caption=""):
        """Envía una imagen (con caption opcional) a un grupo."""
        try:
            if not self.driver:
                logging.error("WhatsApp Web no está inicializado")
                return False

            if not os.path.exists(image_path):
                logging.error(f"Imagen no encontrada: {image_path}")
                return False

            logging.info(f"Enviando imagen al grupo: {group_name}")

            if not self._open_chat(group_name):
                return False

            image_input = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
                )
            )
            image_input.send_keys(image_path)
            time.sleep(3)

            if caption:
                try:
                    caption_box = self._get_message_box()
                    caption_box.send_keys(caption)
                    time.sleep(1)
                except Exception:
                    logging.warning("  No se pudo escribir el caption de la imagen")

            send_btn = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//span[@data-icon="send"]')
                )
            )
            send_btn.click()
            time.sleep(4)

            logging.info(f"✓ Imagen enviada al grupo {group_name} exitosamente")
            return True

        except Exception as e:
            logging.error(f"Error enviando imagen al grupo {group_name}: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return False