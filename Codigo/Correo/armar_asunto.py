#-- Imports --
from Codigo.Apis.Webhook.post import enviarCorreoGeneral

def enviarCaptcha(para, copia, puerto, cia):

    url = f"http://jishucloud.redirectme.net:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"

    mensaje = f"Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece\n\n 👉 {url}\n\nFinaliza con clic en 'Ingresar'"

    enviarCorreoGeneral(para,copia,asunto,mensaje)

def enviarAviso(para,copia,cia):

    asunto = f"🔑 Cambio de contraseña en {cia}"

    mensaje = f"Cambiar la contraseña en {cia} de manera urgente para las proximas automatizaciones."

    if enviarCorreoGeneral(para, copia, asunto, mensaje):
        return True
    else:
        return False
