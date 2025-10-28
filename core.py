#Archivo Core.py el cual realiza la logica
import logging 
import pyautogui 
import time
import os
import subprocess
from datetime import datetime
from pathlib import Path

pyautogui.FAILSAFE = True 
pyautogui.PAUSE = 0.3 

#Funcion que ejecuta comandos de powershell desde python y devuelve el codigo de salida, la salida estandar y el error estandar
def run_powershell(cmd):
    try:
        result = subprocess.run(["powershell", "-Command", cmd],
        capture_output=True, text=True, timeout=10)
        return result.returncode, result.stdout.strip(), result.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

#Funcion que toma screenshot de la pantalla actual y la guarda en una carpeta ScreenShots
def take_screenshot(name): 
  try:
    out = Path("ScreenShots") 
    out.mkdir(exist_ok=True) 
    ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") 
    path = out / f"{name}_{ts}.png" 
    img = pyautogui.screenshot() 
    img.save(path)
    logging.info(f"Se guardo la captura {path}")
    return path 
  except Exception as e:
     logging.error(f"Error de captura de pantalla:{e}")

#Funcion que automatiza el llenado de un forms
def fill_forms(coordenadas):
  # Se utilizo una resolucion de 1920 x 1080 para la pruba de este script y las coordenadas (519,513)
  try: 
    logging.info("Se empezo a contestar el forms")
  #Abre una nueva pestana de Microsoft Edge
    os.system("start msedge")
    time.sleep(3)
    pyautogui.typewrite("https://forms.office.com/pages/responsepage.aspx?id=EZDKymp73kSGHwlaLKiDt4wXC_YfIWlGrUcWrbkA4-NURjFZQjdBMkJNSlkwQkVCM0c2V0cyWTVHNSQlQCNjPTEu&classId=31f16070-5361-4de8-9624-98f60a6f24ae&assignmentId=c865c317-1511-4faa-8a46-565ecf1dd392&submissionId=d9a59b82-b4c2-7d09-d320-fac3711bd2c4&route=shorturl")
    pyautogui.press("enter")
    #Contesta el apartado de la fecha en el forms
    time.sleep(9)
    take_screenshot("Antes")
    pyautogui.click(coordenadas[0],coordenadas[1], duration=0.1)
    time.sleep(2)
    pyautogui.press("enter")
  #Contesta el apartado de los nombres en el forms
    nombres = ["Diego Francisco Martinez Reyes", "Eduardo Tamez Olivo", "Daniel Martinez Viruega"]
    time.sleep(2)
    pyautogui.press("tab")
    for i in nombres:
        pyautogui.typewrite(i)
        pyautogui.press("enter")
  # Suma las matriculas y contesta el apartado de las matriculas en el forms
    Matriculas = [2110198,2118638,2077767]
    sum = 0
    time.sleep(1)
    pyautogui.press("tab")
    for i in Matriculas:
       sum +=i
    pyautogui.typewrite(str(sum))
  # Contesta el apartado de las opciones en el forms
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("down")
  # Envia el formulario
    time.sleep(1)
    pyautogui.press("tab")
    take_screenshot("Durante")
    pyautogui.press("enter")
    time.sleep(2)
    take_screenshot("Despues")
    logging.info("Se termino de contestar el forms")
  except Exception as e:
     logging.error(f"Error al llenar el forms: {e}")

def core_main(): 
    logging.basicConfig(filename="run.log", level=logging.INFO, 
    format="%(asctime)s %(levelname)s %(name)s -> %(message)s", encoding="utf-8") 
    logging.info("Inicio del examen") 
    start_coords = (519,513)
    comandos = ["Get-Date", "Get-Location"]
    for com in comandos:
        code, out, err = run_powershell(com)
        logging.info(f'Comando ejecutado: {com}')
        logging.info(f"Resultado del Comando: {out}")
        if err:
            logging.warning(f"PS code: {code}")
            logging.error(f"PS error: {err}")
    fill_forms(start_coords) 
    logging.info("Fin del examen")