# test pegar imagenes
import win32com.client as win32
import time
import os

from rich import print

from melvive.tableau import Tableau
from melvive.funciones import mensaje_imprimir
from melvive.utilidades import MensajeAviso


def iniciar_powerpow(ruta_powerpo: str, Visible: bool = True):
    """
    Abre PowerPoint y genera una presentación
    :param ruta_powerpo: ruta absoluta de la presentación a abrir
    :param Visible: True para ver PowerPoint en pantalla
    :return: presentación abierta
    """
    ppt = win32.Dispatch("PowerPoint.Application")
    ppt.Visible = Visible
    pres = ppt.Presentations.Open(ruta_powerpo)
    return pres


def cerrar_powerpo(pres, cerrar: bool = True, matar: bool = True):
    time.sleep(5)
    pres.Save()
    if cerrar:
        pres.Close()
    if matar:
        win32.Dispatch("PowerPoint.Application").Quit()
        os.system("taskkill /f /im POWERPNT.exe")


def borrar_imagen(presentacion, num_diapo):
    try:
        slide = presentacion.Slides(num_diapo)
        for shape in list(slide.Shapes):
            if shape.Type == 13:  # Imagen
                shape.Delete()
    except Exception as e:
        raise RuntimeError(f'ERROR AL BORRAR IMAGENES EN DIAPOSITIVA {num_diapo}:\n{e}')


def pegar_imagen(presentacion, num_diapo, ruta_imagen, izquierda, arriba):
    time.sleep(2)
    try:
        slide = presentacion.Slides(num_diapo)
        slide.Shapes.AddPicture(
            FileName=ruta_imagen,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=izquierda,
            Top=arriba
        )
    except Exception as e:
        raise RuntimeError(f'ERROR AL PEGAR IMAGENES EN DIAPOSITIVA {num_diapo}:\n{e}')


# Configuración de dashboards (un solo diccionario con todo)
Nombre_libro = 'REPORTING_VO'

dashboards = {
    "HYU_1": {
        "libro": Nombre_libro,
        "vista": "HYU_1",
        "slide": 3,
        "izquierda": 0,
        "arriba": 0
    },
    "HYU_2": {
        "libro": Nombre_libro,
        "vista": "HYU_2",
        "slide": 4,
        "izquierda": 0,
        "arriba": 0
    },
    "HYU_3": {
        "libro": Nombre_libro,
        "vista": "HYU_3",
        "slide": 5,
        "izquierda": 0,
        "arriba": 10
    }
}

# Rutas
ruta = r'C:\Python\JupyterLab\Lab\PMELERO\SCRIPTS_py\REPORTE_PARTNERS'
ruta_ataque_ppt = (
    r'C:\Users\a33300\OneDrive - BNP Paribas\1 REMARKETING\BI'
    r'\6_REPORTES_PARTNERS\REPORTE_PARTNERS.pptx'
)

# 1. GENERAR IMÁGENES Y GUARDAR EN RUTA
mensaje_imprimir('SE INICIA GUARDADO DASHBOARDS COMO IMAGENES')

# Creamos diccionario para el exportador: {"clave": (libro, vista)}
dashboards_exportar = {
    clave: (datos["libro"], datos["vista"])
    for clave, datos in dashboards.items()
}

exportador = Tableau(dashboards_exportar)
imagenes = exportador.imagenes_to_archivo(ruta)
print(imagenes)

# 2. ABRIR POWERPOINT Y BORRAR IMÁGENES EXISTENTES
pres = iniciar_powerpow(ruta_ataque_ppt)
time.sleep(1)

for datos in dashboards.values():
    borrar_imagen(pres, datos["slide"])

# 3. PEGAR NUEVAS IMÁGENES
for clave, datos in dashboards.items():
    nombre_archivo = f"Z_{clave}.png"
    ruta_imagen = imagenes.get(nombre_archivo)

    if ruta_imagen is None:
        raise FileNotFoundError(f"No se encontró la imagen esperada: {nombre_archivo}")

    pegar_imagen(
        presentacion=pres,
        num_diapo=datos["slide"],
        ruta_imagen=ruta_imagen,
        izquierda=datos["izquierda"],
        arriba=datos["arriba"]
    )

# 4. CERRAR Y GUARDAR POWERPOINT
time.sleep(5)
cerrar_powerpo(pres)

mensaje_imprimir('FOTO DASHBOARD TABLEAU COPIADO Y FOTOS ANTERIORES BORRADAS', 'green')