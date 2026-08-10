"""Re-export desde el módulo global C:\\proyectos\\shared\\screenshot_safe.py"""
import sys
sys.path.insert(0, r"C:\proyectos\shared")
from screenshot_safe import ScreenshotManager, capturar_tabla_excel, _adquirir_archivo  # noqa: F401
