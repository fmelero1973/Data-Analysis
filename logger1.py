import logging
import os

def configurar_logger(nombre_script):
    logger = logging.getLogger(nombre_script)

    if not logger.hasHandlers():
        logger.setLevel(logging.INFO)

        # Formato: fecha - nombre_script - nivel - mensaje
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

        # Handler para archivo de log
        handler = logging.FileHandler('registro_ejecuciones.log', encoding='utf-8')
        handler.setFormatter(formatter)

        logger.addHandler(handler)

    return logger

def log_inicio(logger):
    logger.info("🔄 Inicio de ejecución")

def log_ok(logger):
    logger.info("✅ Finalización correcta")

def log_error(logger, error):
    logger.error(f"❌ Error fatal: {error}")