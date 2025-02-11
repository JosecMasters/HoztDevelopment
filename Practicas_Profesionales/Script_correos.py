import smtplib
import pandas as pd
import os
import time
from email.message import EmailMessage
from dotenv import load_dotenv

# Cargar credenciales desde el archivo .env
load_dotenv()
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT"))

# Configuración del envío
EXCEL_FILE = "Base_de_datos_jubilados_copy.xlsx"
COLUMN_NAME = "CORREO_ELECTRONICO"
IMAGE_PATH = "Imagen.jpg"
PROGRESS_FILE = "progreso.txt"  # Archivo para guardar el progreso
BATCH_SIZE = 150  # Máximo de correos por tanda

# Cargar destinatarios desde el archivo Excel
try:
    df = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    df.columns = df.columns.str.strip()
    recipients = df[COLUMN_NAME].dropna().tolist()
except Exception as e:
    print(f"❌ Error al leer el archivo Excel: {e}")
    recipients = []

# Leer el progreso guardado
def leer_progreso():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            return int(f.read().strip())  # Leer el último índice enviado
    return 0  # Si no hay registro, empezar desde el inicio

# Guardar el progreso actual
def guardar_progreso(index):
    with open(PROGRESS_FILE, "w") as f:
        f.write(str(index))

# Iniciar desde el último punto guardado
start_index = leer_progreso()
end_index = min(start_index + BATCH_SIZE, len(recipients))

if start_index >= len(recipients):
    print("✅ Todos los correos han sido enviados.")
else:
    print(f"📩 Enviando correos {start_index + 1} - {end_index} de {len(recipients)}...")

    for i in range(start_index, end_index):
        recipient = recipients[i]
        try:
            msg = EmailMessage()
            msg["Subject"] = "Convocatoria: PIOESIS Compañía Femenil de Teatro Universitario"
            msg["From"] = EMAIL_SENDER
            msg["To"] = recipient
            msg.set_content("Sindicato Único del Personal Académico de la Universidad Autónoma de Querétaro\nComité ejecutivo 2024-2027")

            # Adjuntar imagen
            with open(IMAGE_PATH, "rb") as img:
                msg.add_attachment(img.read(), maintype="image", subtype="jpeg", filename="imagen.jpg")

            # Enviar correo
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(EMAIL_SENDER, EMAIL_PASSWORD)
                server.send_message(msg)

            print(f"✅ Correo enviado a {recipient}")

        except Exception as e:
            print(f"❌ Error al enviar a {recipient}: {e}")

        # Guardar progreso después de cada envío
        guardar_progreso(i + 1)

    print(f"📌 Se enviaron {end_index - start_index} correos. Reanuda desde el número {end_index}.")

