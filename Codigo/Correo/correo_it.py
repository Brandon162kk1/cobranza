#-- Imports --
import base64
import json
import requests
import os

# --- Variables de Entorno ---
remitente = os.getenv("remitente")
client_id = os.getenv("client_id")
client_secret = os.getenv("client_secret")
tenant_id = os.getenv("TENANT_ID")
SCOPE = os.getenv("SCOPE")

def enviarCaptcha(para, copia, puerto, cia, imagen):
    url = f"http://jishucloud.redirectme.net:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"
    mensaje = f"""
    <p>Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece.</p>
    <p>
        👉 <a href="{url}" target="_blank">{url}</a>
    </p>
    <p>Finaliza con clic en <b>Ingresar</b>.</p>
    """

    #enviarCorreoIT(para, copia, asunto, mensaje, imagen, None)
