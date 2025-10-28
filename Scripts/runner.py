# runner.py
import logging
import sys
import core 

def main():

    # Configuración del Logging 
    try:
        logging.basicConfig(filename="run.log",
                            level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s",
                            encoding="utf-8",
                            filemode='w')
    except Exception as e:
        with open("CRITICAL_LOG_ERROR.txt", "w") as f:
            f.write(f"No se pudo iniciar el logging: {e}")
        sys.exit(1) # Salir si no podemos registrar logs

    logging.info("Inicio del examen")

    # Definición de datos 
    data = {
        "nombre": "Alumno Ejemplo",
        "correo": "ejemplo@correo.com",
        "equipo": "Equipo 3"
    }

    # Nota diego: ajustar coordenadas
    start_coords = (450, 320) # (X, Y) del primer campo del formulario

    # Vlidaciones entrada
    logging.info("Validando datos de entrada...")
    required_keys = ["nombre", "correo", "equipo"]
    
    # Verifica que todas las llaves existan y no estén vacías
    if not all(key in data and data[key] for key in required_keys):
        logging.error("Validación fallida: Faltan datos en el diccionario 'data'.")
        logging.error(f"Datos recibidos: {data}")
        logging.info("Fin del examen (por error de validación).")
        sys.exit(1) 
    
    logging.info("Validación de datos exitosa.")

    # Ejecución de PowerShell
    logging.info("Ejecutando comandos de PowerShell...")
    cmds_to_run = ["Get-Date", "whoami"]
    
    for cmd in cmds_to_run:
        logging.info(f"Ejecutando: '{cmd}'")
        code, out, err = core.run_powershell(cmd)
        logging.info(f"Comando '{cmd}' - Código: {code}")
        logging.info(f"Comando '{cmd}' - Salida: {out}")
        if err:
            logging.warning(f"Comando '{cmd}' - Error: {err}")

    # Llenado del formulario
    core.fill_form(data, start_coords)

    logging.info("Fin del examen.")

if __name__ == "__main__":
    main()