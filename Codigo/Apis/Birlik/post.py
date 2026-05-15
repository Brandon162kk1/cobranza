import requests
import os

#----- Variables de Entorno -------
API_KEY = os.getenv("API_KEY")
AFTER_API_KEY = os.getenv("AFTER_API_KEY")
API_KEY_HEADER = AFTER_API_KEY+" "+API_KEY

headers = {
        "Authorization": API_KEY_HEADER,  # si tu API necesita auth
        "Content-Type": "application/json"
    }

def guardar_observacion_cuota(id_cuota, observacion):

    url = "https://plataformabirlik.azurewebsites.net/api/cuota/guardar-observacion"

    payload = {
        "IdCuota": id_cuota,
        "Observacion": observacion
    }

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=5)
        if response.status_code == 200:
            data = response.json()
            print(f"[API] Observación guardada en cuota {id_cuota}: {data}")
            return True
        else:
            print(f"[API] Error al guardar observación en cuota {id_cuota}: {response.text}")
            return False
    except Exception as e:
        print(f"[API] Excepción al llamar API para cuota {id_cuota}: {e}")
        return False