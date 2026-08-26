#-- Imports --
import os
from Codigo.Apis.Webhook.post import enviar_correo

url_n8n_base = os.getenv("url_n8n_base")

para = os.getenv("para")
para_lista = para.split(",") if para else []

copia = os.getenv("copia")
copias_lista = copia.split(",") if copia else []

def enviarCaptcha(puerto,cia):

    url = f"{url_n8n_base}:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"

    mensaje = f"Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece\n\n 👉 {url}\n\nFinaliza con clic en 'Ingresar'"

    if enviar_correo(para_lista,copias_lista,asunto,mensaje):
        return True
    else:
        return False

def enviarAviso(compania):

    asunto = f"🔑 Cambio de contraseña en {compania}"

    mensaje = f"Cambiar la contraseña en {compania} de manera urgente para las proximas automatizaciones."

    if enviar_correo(para_lista,copias_lista,asunto,mensaje):
        return True
    else:
        return False
