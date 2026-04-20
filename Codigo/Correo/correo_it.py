#-- Imports --
import os
from Apis.Webhook.post import enviarCorreoGeneral

# --- Variables de Entorno ---
remitente = os.getenv("remitente")
client_id = os.getenv("client_id")
client_secret = os.getenv("client_secret")
tenant_id = os.getenv("TENANT_ID")
SCOPE = os.getenv("SCOPE")

def enviarCaptcha(para, copia, puerto, cia):

    url = f"http://jishucloud.redirectme.net:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"

    mensaje = f"Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece\n\n 👉 {url}\n\nFinaliza con clic en 'Ingresar'"

    if enviarCorreoGeneral(para, copia, asunto, mensaje):
        return True
    else:
        return False
