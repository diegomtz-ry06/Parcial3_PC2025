import pyautogui 
import time
import os
import logging
from datetime import datetime
from pathlib import Path

pyautogui.FAILSAFE = True 
pyautogui.PAUSE = 0.3 

#Funcion que toma screenshot de la pantalla actal y la guarda en una carpeta ScreenShots
def take_screenshot(name): 
  logger = logging.getLogger(__name__)
  try:
    out = Path("ScreenShots") 
    out.mkdir(exist_ok=True) 
    ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ") 
    path = out / f"{name}_{ts}.png" 
    img = pyautogui.screenshot() 
    img.save(path)
    logger.info(f"Se guardo la captura {path}")
    return path 
  except Exception as e:
     logger.error(f"Error de captura de pantalla:{e}")

#Funcion que automatiza el contestado de un forms
def fill_forms(coordenadas):
  try:
    logger = logging.getLogger(__name__) 
    logger.info("Se empezo a contestar el forms")

  #Abre una nueva pestana de Microsoft Edge
    os.system("start msedge")
    time.sleep(3)
    pyautogui.typewrite("https://forms.office.com/pages/responsepage.aspx?id=EZDKymp73kSGHwlaLKiDt4wXC_YfIWlGrUcWrbkA4-NURjFZQjdBMkJNSlkwQkVCM0c2V0cyWTVHNSQlQCNjPTEu&classId=31f16070-5361-4de8-9624-98f60a6f24ae&assignmentId=c865c317-1511-4faa-8a46-565ecf1dd392&submissionId=d9a59b82-b4c2-7d09-d320-fac3711bd2c4&route=shorturl")
    pyautogui.press("enter")
    print("Empezando")
    time.sleep(8)
    take_screenshot("Antes")
    pyautogui.click(coordenadas[0],coordenadas[1], duration=0.1)

  #Contesta el apartado de la fecha en el forms
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
    Matriculas = [1,2,3]
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
  #pyautogui.press("enter") <<-- descomentar esto cuando no se necesite

    take_screenshot("Despues")
    logger.info("Se termino de contestar el forms")
  except Exception as e:
     logger.error(f"Error al llenar el forms: {e}")
  finally:
     os.system("taskkill /IM msedge.exe /F")