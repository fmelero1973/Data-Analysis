from logger_config import configurar_logger, log_inicio, log_ok, log_error
import sys

logger = configurar_logger(__file__)  # Usa el nombre del script automáticamente
log_inicio(logger)

try:
    # Tu lógica aquí
    resultado = 10 / 2
    log_ok(logger)

except Exception as e:
    log_error(logger, e)
    sys.exit(1)