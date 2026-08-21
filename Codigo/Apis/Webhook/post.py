from logging import raiseExceptions
import requests
import os

# --- Variables de Entorno ---
url_n8n_base = os.getenv("url_n8n_base")
path_enviar_correo = os.getenv("path_enviar_correo")

para = os.getenv("para")
para_lista = para.split(",") if para else []

url_n8n_enviar_correo = f"{url_n8n_base}{path_enviar_correo}"

def enviar_correo(copia,asunto,mensaje):
    
    payload = {
        "Para": para,
        "Copia": copia,
        "Asunto": asunto,
        "Mensaje": mensaje
    }

    try:

        print(f"⌛ Enviando correo")

        response = requests.post(url_n8n_enviar_correo,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            print(f"📩 Correo enviado")
            return True
        else:
            print(f"⚠️ Problemas en el envio de correo - Status : {response.status_code} - Resp : {response.text}")
            return False

    except Exception as e:
        print(f"⚠️ Excepción en el envio de correo - Error : {e}")
        return False