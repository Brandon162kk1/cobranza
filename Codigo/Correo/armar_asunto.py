#-- Imports --
import os
from Codigo.Apis.Webhook.post import enviar_correo

url_n8n_base = os.getenv("url_n8n_base")

def enviarCaptcha(copia,puerto,cia):

    url = f"{url_n8n_base}:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"

    mensaje = f"Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece\n\n 👉 {url}\n\nFinaliza con clic en 'Ingresar'"

    if enviar_correo(copia,asunto,mensaje):
        return True
    else:
        return False

def enviarAviso(copia,cia):

    asunto = f"🔑 Cambio de contraseña en {cia}"

    mensaje = f"Cambiar la contraseña en {cia} de manera urgente para las proximas automatizaciones."

    if enviar_correo(copia,asunto,mensaje):
        return True
    else:
        return False
