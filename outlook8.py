from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import win32com.client

def convertir_xls_a_xlsx(ruta_xls: Path, ruta_destino: Path):
    """
    Convierte un archivo .xls a formato .xlsx usando pandas.

    :param ruta_xls: Ruta del archivo .xls original.
    :param ruta_destino: Ruta donde se guardará el archivo convertido .xlsx.
    """
    df = pd.read_excel(ruta_xls, engine='xlrd')
    df.to_excel(ruta_destino, index=False)

def limpiar_temporales(carpeta: Path):
    """
    Elimina todos los archivos contenidos en una carpeta.

    :param carpeta: Ruta de la carpeta que se desea limpiar.
    """
    for archivo in carpeta.glob('*'):
        archivo.unlink()

def mail_adjuntos(
    mail_carpeta: str,
    carpeta_salida: Path = Path(__file__).parent,
    mail_cuenta: str = 'fernando.melero@arval.es',
    carpeta_procesados: str = "Procesados",
    filtro_asunto_global: str | None = None,
    adjunto_nombre: str | None = None,
    remitente_valido: list[str] = None,
    extensiones_permitidas: list[str] = None,
    archivo_rutas: dict[str, dict[str, dict[str, dict[str, Path | str]]]] = None,
    convertir_xls: bool = True,
    mostrar_detalle: bool = True
):
    """
    Procesa solo el último correo válido en Outlook, guarda su adjunto en la ruta correspondiente,
    convierte .xls a .xlsx si se solicita, y mueve el correo a la subcarpeta indicada.

    :param mail_carpeta: Carpeta de Outlook donde están los correos (ej. "Inbox").
    :param carpeta_salida: Carpeta temporal para guardar adjuntos antes de procesarlos.
    :param mail_cuenta: Dirección de correo configurada en Outlook.
    :param carpeta_procesados: Subcarpeta dentro de mail_carpeta donde se moverán los correos procesados.
    :param filtro_asunto_global: Texto para filtrar correos por asunto (opcional).
    :param adjunto_nombre: Texto para filtrar adjuntos por nombre (opcional).
    :param remitente_valido: Lista de remitentes válidos.
    :param extensiones_permitidas: Lista de extensiones válidas para los adjuntos.
    :param archivo_rutas: Diccionario que define cómo enrutar y nombrar los archivos.
    :param convertir_xls: Si True, convierte archivos .xls a .xlsx.
    :param mostrar_detalle: Si True, imprime detalles del proceso.
    """
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
    carpeta_origen = main_folder.Folders.Item(mail_carpeta)
    carpeta_destino = carpeta_origen.Folders.Item(carpeta_procesados)

    mensajes = carpeta_origen.Items
    mensajes.Sort("[ReceivedTime]", True)  # Orden descendente: primero el más reciente

    mensajes_total = mensajes.Count
    mensajes_filtrados = 0
    archivos_guardados = 0

    for msg in mensajes:
        fecha_mensaje = msg.ReceivedTime.replace(tzinfo=None)
        remitente = msg.SenderEmailAddress.lower()
        asunto = msg.Subject.lower()

        filtro_remitente = remitente in [r.lower() for r in remitente_valido]
        filtro_asunto = filtro_asunto_global.lower() in asunto if filtro_asunto_global else True
        filtro_fecha = True

        condicion_mail = filtro_remitente and filtro_asunto and filtro_fecha

        if condicion_mail:
            mensajes_filtrados += 1

            for attachment in msg.Attachments:
                nombre_archivo = Path(attachment.Filename)

                filtro_extension = nombre_archivo.suffix.lower() in extensiones_permitidas
                filtro_nombre = adjunto_nombre.lower() in nombre_archivo.stem.lower() if adjunto_nombre else True

                condicion_adjunto = filtro_extension and filtro_nombre

                if condicion_adjunto:
                    clave_archivo = next((k for k in archivo_rutas if nombre_archivo.stem.lower().startswith(k)), None)
                    remitente_info = archivo_rutas.get(clave_archivo, {}).get(remitente) if clave_archivo else None
                    coincidencias_asunto = [clave for clave in remitente_info if asunto.startswith(clave.lower())] if remitente_info else []

                    if len(coincidencias_asunto) == 1:
                        clave_asunto = coincidencias_asunto[0]
                        destino = remitente_info[clave_asunto]
                        destino_ruta = Path(destino['ruta'])
                        nombre_final = destino['nombre_final']

                        if nombre_archivo.suffix.lower() == '.xls' and convertir_xls:
                            ruta_temporal = carpeta_descarga_adjunto / nombre_archivo.name
                            attachment.SaveAsFile(str(ruta_temporal))
                            ruta_final = destino_ruta / f"{nombre_final}.xlsx"
                            convertir_xls_a_xlsx(ruta_temporal, ruta_final)
                            if mostrar_detalle:
                                print(f'CONVERTIDO Y GUARDADO: {ruta_final}')
                        else:
                            ruta_final = destino_ruta / f"{nombre_final}{nombre_archivo.suffix.lower()}"
                            attachment.SaveAsFile(str(ruta_final))
                            if mostrar_detalle:
                                print(f'GUARDADO DIRECTAMENTE: {ruta_final}')
                        archivos_guardados += 1
                        break  # ✅ Solo procesamos el primer adjunto válido

            msg.UnRead = False
            msg.Move(carpeta_destino)
            break  # ✅ Solo procesamos el primer mensaje válido

    try:
        limpiar_temporales(carpeta_descarga_adjunto)
        if mostrar_detalle:
            print(f'CARPETA TEMPORAL LIMPIADA: {carpeta_descarga_adjunto}')
    except Exception as e:
        raise RuntimeError(f'ERROR AL LIMPIAR CARPETA TEMPORAL: {e}')

    print(f'\nMENSAJES TOTALES: {mensajes_total}')
    print(f'MENSAJES FILTRADOS: {mensajes_filtrados}')
    print(f'ARCHIVOS GUARDADOS: {archivos_guardados}')

    borrar_mensajes_antiguos(carpeta_destino, dias=7)

def borrar_mensajes_antiguos(carpeta_objetivo, dias: int = 7):
    """
    Elimina los correos de una carpeta de Outlook que fueron recibidos hace más de 'dias' días.

    :param carpeta_objetivo: Objeto de carpeta Outlook (ej. carpeta_destino).
    :param dias: Número de días a conservar. Los correos más antiguos serán eliminados.
    """
    fecha_limite = datetime.now() - timedelta(days=dias)
    mensajes = carpeta_objetivo.Items

    eliminados = 0
    for msg in mensajes:
        if msg.ReceivedTime.replace(tzinfo=None) < fecha_limite:
            msg.Delete()
            eliminados += 1

    print(f"🗑️ Correos eliminados de '{carpeta_objetivo.Name}': {eliminados} (anteriores a {dias} días)")
