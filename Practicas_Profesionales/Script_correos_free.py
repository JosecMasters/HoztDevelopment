import smtplib
import pandas as pd
from email.message import EmailMessage

# Configuración del servidor SMTP (Gmail en este caso)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "jjacal36115@ucq.edu.mx"
EMAIL_PASSWORD = "MiSi@100porciento"

# Cargar destinatarios desde un archivo Excel
EXCEL_FILE = "Prueba.xlsx"  # Asegúrate de que el archivo esté en la misma carpeta o proporciona la ruta completa
COLUMN_NAME = "CORREO_ELECTRONICO"  # Nombre de la columna en el archivo Excel que contiene los correos

try:
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    print("Nombres de columnas en el archivo:", df.columns.tolist())
    recipients = df[COLUMN_NAME].dropna().tolist()  # Elimina valores nulos y convierte a lista
except Exception as e:
    print(f"Error al leer el archivo Excel: {e}")
    recipients = []

# Ruta de la imagen a adjuntar
IMAGE_PATH = "Imagen.jpg"  # Cambia esto a la imagen que deseas adjuntar

# Verificar si hay destinatarios
if not recipients:
    print("No se encontraron destinatarios en el archivo Excel.")
else:
    # Crear y enviar los correos
    for recipient in recipients:
        try:
            msg = EmailMessage()
            msg["Subject"] = "Prueba"
            msg["From"] = EMAIL_SENDER
            msg["To"] = recipient
            msg.set_content("CON fOTITO")

            # Adjuntar la imagen
            with open(IMAGE_PATH, "rb") as img:
                msg.add_attachment(img.read(), maintype="image", subtype="jpeg", filename="imagen.jpg")


            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(EMAIL_SENDER, EMAIL_PASSWORD)
                server.send_message(msg)

            print(f"Correo enviado exitosamente a: {recipient}")
        except Exception as e:
            print(f"Error al enviar correo a {recipient}: {e}")
