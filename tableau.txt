import tableauserverclient as TSC
import win32com.client as win32
import os

# === CONFIGURACIÓN DE ACCESO A TABLEAU ===
server_url = "https://tableau.miempresa.com"
token_name = "Fernando"
token_secret = "001"
site_id = "compras"

# === DASHBOARDS A EXPORTAR ===
# Clave = marcador en el HTML, Valor = nombre de la vista en Tableau
dashboards = {
    "DASH1": "VentasMensuales",
    "DASH2": "ComprasAnuales",
    "DASH3": "ResumenGeneral"
}

# === RUTA A LA PLANTILLA OUTLOOK ===
plantilla_msg = os.path.abspath("plantilla.msg")

# === AUTENTICACIÓN EN TABLEAU ===
auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_id)
server = TSC.Server(server_url, use_server_version=True)

with server.auth.sign_in(auth):
    all_views, _ = server.views.get()

    # === EXPORTAR CADA DASHBOARD COMO IMAGEN ===
    imagenes_exportadas = {}

    for marcador, vista_nombre in dashboards.items():
        vista = next((v for v in all_views if v.name == vista_nombre), None)
        if vista:
            server.views.populate_image(vista)
            nombre_imagen = f"{marcador.lower()}.png"
            with open(nombre_imagen, "wb") as f:
                f.write(vista.image)
            imagenes_exportadas[marcador] = os.path.abspath(nombre_imagen)
            print(f"{vista_nombre} exportado como {nombre_imagen}")
        else:
            print(f"Vista no encontrada: {vista_nombre}")

# === ABRIR LA PLANTILLA DE OUTLOOK Y REEMPLAZAR MARCADORES ===
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.Session.OpenSharedItem(plantilla_msg)

# === REEMPLAZAR MARCADORES POR IMÁGENES INCRUSTADAS ===
for marcador, ruta_imagen in imagenes_exportadas.items():
    cid = f"cid_{marcador.lower()}"
    if f"#{marcador}#" in mail.HTMLBody:
        mail.HTMLBody = mail.HTMLBody.replace(
            f"#{marcador}#", f'<img src="cid:{cid}">'
        )
    else:
        print(f"Marcador #{marcador}# no encontrado en la plantilla.")

    # Adjuntar imagen como archivo y embebida
    adjunto = mail.Attachments.Add(ruta_imagen)
    adjunto.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
    )

# === GUARDAR O ENVIAR EL MENSAJE FINAL ===
ruta_final = os.path.abspath("mensaje_final.msg")
mail.SaveAs(ruta_final)
print(f"Mensaje final guardado como: {ruta_final}")
# mail.Send()  # Descomenta para enviarlo automáticamente