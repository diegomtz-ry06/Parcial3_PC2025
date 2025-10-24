#Este seria el archivo core.py

import logging 
from forms_fill import fill_forms
from cmdpwsh import run_powershell
# Todo mover funciones a un archivo core.py y dejar solo ejecución en runner.py 

def main(): 
    logging.basicConfig(filename="run.log", level=logging.INFO, 
    format="%(asctime)s %(levelname)s %(name)s -> %(message)s", encoding="utf-8") 

    logger = logging.getLogger(__name__)
    logger.info("Inicio del examen") 

    start_coords = (519,513)  # ← deben ajustar esto según su pantalla 

    comandos = ["Get-Date", "Get-Location"]
    for com in comandos:
        code, out, err = run_powershell(com)
        logger.info(f'Comando ejecutado: {com}')
        logger.info(f"Resultado del Comando: {out}")
        if err:
            logger.warning(f"PS code: {code}")
            logger.error(f"PS error: {err}")

    fill_forms(start_coords) 
    logger.info("Fin del examen")
