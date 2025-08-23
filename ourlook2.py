Perfecto, FERNANDO. En Outlook, cuando accedes a tus carpetas mediante `win32com.client`, la estructura suele ser algo así:

• `namespace.Folders` → lista de buzones (por ejemplo, si tienes varias cuentas configuradas)
• Dentro de cada buzón, hay carpetas como “Bandeja de entrada”, “Elementos enviados”, y también tus carpetas personalizadas como “facturas”


Si la carpeta “facturas” está dentro de tu buzón principal (que se llama como tu dirección de correo), puedes acceder a ella directamente así:

---

🧩 Versión ajustada del código

import win32com.client
import os

# Configuración
EMAIL_ACCOUNT = 'entrada@gmail.com'  # Tu dirección de correo
SUBJECT_KEYWORD = 'Factura'
ATTACHMENT_KEYWORD = 'factura'
SAVE_FOLDER = r'C:\carpeta'

# Crear carpeta si no existe
if not os.path.exists(SAVE_FOLDER):
    os.makedirs(SAVE_FOLDER)

# Conectar a Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Acceder al buzón principal
main_folder = namespace.Folders.Item(EMAIL_ACCOUNT)

# Acceder a la subcarpeta "facturas"
try:
    facturas_folder = main_folder.Folders.Item("facturas")
except Exception as e:
    raise Exception("No se encontró la carpeta 'facturas' dentro del buzón principal.") from e

# Procesar correos
messages = facturas_folder.Items
messages.Sort("[ReceivedTime]", True)  # Opcional: ordenar por fecha

for msg in messages:
    try:
        if hasattr(msg, 'SenderEmailAddress') and msg.SenderEmailAddress == EMAIL_ACCOUNT:
            if SUBJECT_KEYWORD.lower() in msg.Subject.lower():
                for attachment in msg.Attachments:
                    if attachment.Type == 1 and 'pdf' in attachment.FileName.lower() and ATTACHMENT_KEYWORD.lower() in attachment.FileName.lower():
                        filepath = os.path.join(SAVE_FOLDER, attachment.FileName)
                        attachment.SaveAsFile(filepath)
                        print(f'Archivo guardado: {filepath}')
    except Exception as e:
        print(f"Error procesando mensaje: {e}")
        continue

print("Proceso completado.")


---

🧠 Tips adicionales

• Si tienes más de una cuenta en Outlook, asegúrate de que `EMAIL_ACCOUNT` coincida exactamente con el nombre del buzón (puede que no sea idéntico al correo).
• Puedes imprimir los nombres de las carpetas para verificar la estructura con:


for folder in namespace.Folders:
    print(folder.Name)


¿Quieres que el script se ejecute automáticamente cada día o que te notifique cuando se descargue una factura nueva? Puedo ayudarte a automatizarlo.#