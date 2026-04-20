import requests
import logging
import os

# --- Variables de Entorno ---
url_n8n_enviar_correo_general = os.getenv("url_n8n_enviar_correo_general")

def enviarCorreoGeneral(para, copia, asunto, mensaje):
    
    payload = {
        "Para": para,
        "Copia": copia,
        "Asunto": asunto,
        "Mensaje": mensaje
    }

    print(f"📩 Enviando correo a {para} con asunto '{asunto}'")

    try:
        response = requests.post(url_n8n_enviar_correo_general,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            print(f"✅ Correo enviado")
            return True
        else:
            print(f"Problemas en el envio de correo - Status : {response.status_code} - Resp : {response.text}")
            return False

    except Exception as e:
        print(f"Error enviando el correo por el webhook, Motivo : {e}")
        return False