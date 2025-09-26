import pandas as pd
import win32com.client
from pathlib import Path
from melvive.funciones import imp_mensaje_inicial, fechas_seleccionar, mensaje_imprimir

# INICIO
nombre_script = 'GUARDADO ARCHIVO ADJUNTO CARPETA STOCKER'

# RUN------------------------------------------------------------------------------------
imp_mensaje_inicial(nombre_script)

# CAPTURO FECHAS DE INFORME
start_date, end_date = fechas_seleccionar()
fecha_inicio = pd.to_datetime(start_date)
fecha_fin = pd.to_datetime(end_date)
mensaje_imprimir(f'SE GUARDARAN ADJUNTOS DESDE {fecha_inicio:%d-%m-%Y} A {fecha_fin:%d-%m-%Y} INCLUSIVE')

def mail_adjuntos(mail_carpeta: str,
                  carpeta_salida: Path,
                  mail_cuenta: str = 'fernando.melero@arval.es',
                  asunto_mail: str = '',
                  adjunto_nombre: str = '',
                  remitente_valido: list[str] = None,
                  extensiones_permitidas: list[str] = None,
                  archivo_rutas: dict[str, dict[str, Path | str]] = None,
                  convertir_xls: bool = True,
                  eliminar_original: bool = True,
                  mostrar_detalle: bool = True):

    if remitente_valido is None:
        remitente_valido = []
    if extensiones_permitidas is None:
        extensiones_permitidas = ['.xlsx', '.xls']
    if archivo_rutas is None:
        archivo_rutas = {}

    carpeta_descarga_adjunto = Path(carpeta_salida)
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    main_folder = namespace.Folders.Item(mail_cuenta)

    try:
        carpeta = main_folder.Folders.Item(mail_carpeta)
    except Exception as e:
        raise Exception(f'No se encontró la carpeta {mail_carpeta} dentro del buzón principal.') from e

    mensajes = carpeta.Items
    mensajes_total = mensajes.Count
    mensajes_filtrados = 0
    archivos_guardados = 0
    mensajes.Sort("[ReceivedTime]", True)

    for msg in mensajes:
        fecha_mensaje = msg.ReceivedTime.replace(tzinfo=None)
        remitente = msg.SenderEmailAddress.lower()
        asunto = msg.Subject.lower()

        filtro_fecha = (fecha_inicio <= fecha_mensaje <= (fecha_fin + pd.Timedelta(days=1)))
        filtro_asunto = asunto_mail.lower() in asunto
        filtro_remitente = remitente in [r.lower() for r in remitente_valido]
        condicion_mail = filtro_fecha & filtro_asunto & filtro_remitente

        if condicion_mail:
            if mostrar_detalle:
                print(f'{asunto} \t{remitente}')
            mensajes_filtrados += 1

            for attachment in msg.Attachments:
                nombre_archivo = Path(attachment.Filename)
                if ((nombre_archivo.suffix.lower() in extensiones_permitidas) and
                    adjunto_nombre.lower() in nombre_archivo.stem.lower()):

                    ruta_temporal = carpeta_descarga_adjunto / nombre_archivo.name
                    attachment.SaveAsFile(str(ruta_temporal))

                    destino_especifico = None
                    nombre_final = None
                    for clave, info in archivo_rutas.items():
                        if clave.lower() in nombre_archivo.stem.lower():
                            destino_especifico = Path(info['ruta'])
                            nombre_final = info['nombre_final']
                            break

                    if destino_especifico and nombre_final:
                        destino_especifico.mkdir(parents=True, exist_ok=True)

                        if nombre_archivo.suffix.lower() == '.xls' and convertir_xls:
                            try:
                                df = pd.read_excel(ruta_temporal, engine='xlrd')
                                ruta_final = destino_especifico / f'{nombre_final}.xlsx'
                                df.to_excel(ruta_final, index=False)
                                if eliminar_original:
                                    ruta_temporal.unlink()
                                if mostrar_detalle:
                                    print(f'Convertido y guardado: {ruta_final}')
                            except Exception as e:
                                print(f'Error al convertir {ruta_temporal.name}: {e}')
                        else:
                            ruta_final = destino_especifico / f'{nombre_final}{nombre_archivo.suffix.lower()}'
                            ruta_temporal.replace(ruta_final)
                            if mostrar_detalle:
                                print(f'Guardado en ruta específica: {ruta_final}')
                    else:
                        if mostrar_detalle:
                            print(f'Guardado en ruta por defecto: {ruta_temporal}')

                    archivos_guardados += 1

    mensaje_imprimir(f'\n{mensajes_total} MENSAJES TOTALES EN {mail_carpeta} CON LAS CONDICIONES:'
                     f'\nCORREO: {mail_cuenta}\nCARPETA: {mail_carpeta}\nASUNTO MAIL {asunto_mail}\nNOMBRE ARCHIVO: {adjunto_nombre}'
                     f'\nTIPO ARCHIVO: {extensiones_permitidas}', 'red')
    mensaje_imprimir(f'\n{mensajes_filtrados} CUMPLEN CRITERIOS DE FILTROS DE CORREO')
    mensaje_imprimir(f'\n{archivos_guardados} ARCHIVOS ADJUNTOS GUARDADOS EN\n{carpeta_descarga_adjunto}')

    return mensajes_total, mensajes_filtrados, archivos_guardados

# EJEMPLO DE USO
archivo_rutas = {
    'stock_diario': {
        'ruta': Path(r'C:\Informes\Stock'),
        'nombre_final': 'stock_actualizado'
    },
    'ventas_mensuales': {
        'ruta': Path(r'C:\Informes\Ventas'),
        'nombre_final': 'reporte_ventas'
    }
}

mail_adjuntos(
    mail_carpeta='1_CAMPAS',
    carpeta_salida=Path(r'C:\Users\a33300\OneDrive - BNP Paribas\Bureau\PERITACIONES_DIARIAS'),
    asunto_mail='',
    adjunto_nombre='stock',
    remitente_valido=["mailaccount@fcm-eu.com", "ejemplo@ejemplo.com"],
    extensiones_permitidas=['.xlsx', '.xls'],
    archivo_rutas=archivo_rutas,
    convertir_xls=True,
    eliminar_original=True,
    mostrar_detalle=True
)

print("Proceso completado.")
