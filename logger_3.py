import logging
from pathlib import Path
import inspect
from datetime import datetime
import csv
import os

def ejecutar_con_log(funcion):
    # Obtener nombre del script que llama
    nombre_script = Path(inspect.stack()[1].filename).name

    # Configurar logger
    logger = logging.getLogger(nombre_script)
    if not logger.hasHandlers():
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter(
            '%Y-%m-%d %H:%M:%S - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        consola = logging.StreamHandler()
        consola.setFormatter(formatter)
        logger.addHandler(consola)

    # Registrar inicio
    mensaje_inicio = "🔄 Inicio de ejecución"
    logger.info(mensaje_inicio)

    try:
        funcion()  # Ejecutar lógica del usuario
        mensaje_ok = "✅ Fin de ejecución OK"
        guardar_en_archivo('log_ok.log', nombre_script, mensaje_ok)
        guardar_en_csv('log_general.csv', nombre_script, 'OK', mensaje_ok)
        logger.info(mensaje_ok)

    except Exception as e:
        mensaje_error = f"❌ Fin de ejecución KO: {e}"
        guardar_en_archivo('log_ko.log', nombre_script, mensaje_error)
        guardar_en_csv('log_general.csv', nombre_script, 'KO', mensaje_error)
        logger.error(mensaje_error)

def guardar_en_archivo(nombre_archivo, nombre_script, mensaje):
    fecha = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    linea = f"{fecha} - {nombre_script} - {mensaje}\n"
    with open(nombre_archivo, 'a', encoding='utf-8') as f:
        f.write(linea)

def guardar_en_csv(nombre_csv, nombre_script, estado, mensaje):
    fecha = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    encabezado = ['fecha', 'script', 'estado', 'mensaje']
    fila = [fecha, nombre_script, estado, mensaje]

    archivo_nuevo = not os.path.exists(nombre_csv)
    with open(nombre_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if archivo_nuevo:
            writer.writerow(encabezado)
        writer.writerow(fila)