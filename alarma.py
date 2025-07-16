import tkinter as tk
from tkinter import messagebox
import winsound
import time
import sys
from pathlib import Path

def alerta_usuario(error):
    # 🔉 Patrón de beep
    for _ in range(3):
        winsound.Beep(1500, 300)
        time.sleep(0.2)

    # 📍 Ruta del script principal
    ruta_main = Path(sys.argv[0]).resolve()
    tipo_error = type(error).__name__
    mensaje_error = str(error)

    # 🧾 Texto del error completo
    texto_error = (
        f"⚠️ El proceso ha fallado.\n\n"
        f"🛑 Tipo de error: {tipo_error}\n"
        f"📄 Mensaje: {mensaje_error}\n"
        f"📁 Script ejecutado: {ruta_main}"
    )

    # 🪟 Crear ventana
    ventana = tk.Tk()
    ventana.title("❌ Error crítico")
    ventana.geometry("420x240")
    ventana.resizable(False, False)
    ventana.attributes("-topmost", True)

    etiqueta = tk.Label(ventana, text=texto_error, font=("Segoe UI", 10), justify="left", padx=20, pady=10)
    etiqueta.pack()

    def copiar_portapapeles():
        ventana.clipboard_clear()
        ventana.clipboard_append(texto_error)
        messagebox.showinfo("Copiado", "✅ Detalles del error copiados al portapapeles.")

    btn_copiar = tk.Button(ventana, text="📋 Copiar al portapapeles", font=("Segoe UI", 10), command=copiar_portapapeles)
    btn_copiar.pack(pady=4)

    btn_cerrar = tk.Button(ventana, text="Cerrar", font=("Segoe UI", 10), command=ventana.destroy)
    btn_cerrar.pack(pady=4)

    ventana.mainloop()