# core.py
import subprocess
import pyautogui
import time
import logging
from pathlib import Path

pyautogui.FAILSAFE = True 
pyautogui.PAUSE = 0.5

def run_powershell(cmd):
    try:
        # Ejecuta el comando de PowerShell
        result = subprocess.run(["powershell", "-Command", cmd],
                                capture_output=True, text=True, timeout=10, check=True, encoding='utf-8')
        return 0, result.stdout.strip(), result.stderr.strip()
    except subprocess.CalledProcessError as e:
        logging.error(f"Error de PowerShell al ejecutar '{cmd}': {e.stderr.strip()}")
        return e.returncode, e.stdout.strip(), e.stderr.strip()
    except subprocess.TimeoutExpired:
        logging.error(f"Timeout ejecutando PowerShell: '{cmd}'")
        return 1, "", "Timeout (10s) alcanzado."
    except Exception as e:
        logging.error(f"Error inesperado de subprocess: {e}")
        return 1, "", str(e)


def take_screenshot(name):
    try:
        out_dir = Path("out")
        out_dir.mkdir(exist_ok=True)
        path = out_dir / f"{name}.png"
        img = pyautogui.screenshot()
        img.save(path)
        logging.info(f"Captura de pantalla guardada en: {path}")
        return str(path)
    except Exception as e:
        logging.error(f"No se pudo tomar la captura de pantalla '{name}': {e}")
        return None


def fill_form(data, start_coords):
    logging.info(f"Iniciando llenado de formulario en coordenadas: {start_coords}")
    try:
        # 1. Captura antes de hacer nada
        take_screenshot("before")
        # 2. Click en las coordenadas de inicio
        pyautogui.click(start_coords[0], start_coords[1])
        logging.info(f"Click en {start_coords}")
        # 3. Llenar "nombre" y tabular
        pyautogui.typewrite(data["nombre"])
        pyautogui.press("tab")
        logging.info(f"Nombre escrito: {data['nombre']}")
        # 4. Llenar "correo" y tabular
        pyautogui.typewrite(data["correo"])
        pyautogui.press("tab")
        logging.info(f"Correo escrito: {data['correo']}")
        # 5. Captura durante el llenado
        take_screenshot("during")
        # 6. Llenar "equipo" y enviar
        pyautogui.typewrite(data["equipo"])
        pyautogui.press("enter")
        logging.info(f"Equipo escrito: {data['equipo']} y 'Enter' presionado")
        # 7. Esperar 1 segundo para que la página reacciones
        time.sleep(1)
        # 8. Captura final
        take_screenshot("after")
        
        logging.info("Llenado de formulario completado exitosamente.")
    
    except Exception as e:
        logging.error(f"Error durante la automatización con pyautogui: {e}")
        # Tomar una captura de pantalla del error si es posible
        take_screenshot("error_state")