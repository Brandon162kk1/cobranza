import requests
import os

# --- Variables de Entorno ---
url_n8n_enviar_correo_general = os.getenv("url_n8n_enviar_correo_general")

def enviarCorreoGeneral(para,copia,asunto,mensaje):
    
    payload = {
        "Para": para,
        "Copia": copia,
        "Asunto": asunto,
        "Mensaje": mensaje
    }

    print(f"⌛ Enviando correo")

    try:
        response = requests.post(url_n8n_enviar_correo_general,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            print(f"📩 Correo enviado")
        else:
            print(f"⚠️ Problemas en el envio de correo - Status : {response.status_code} - Resp : {response.text}")

    except Exception as e:
        raise Exception(str(e))