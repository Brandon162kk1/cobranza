import requests
import os
from requests.exceptions import ReadTimeout, ConnectTimeout, RequestException

#----- Variables de Entorno -------
API_KEY = os.getenv("API_KEY")
AFTER_API_KEY = os.getenv("AFTER_API_KEY")
API_KEY_HEADER = AFTER_API_KEY+" "+API_KEY

headers = {
        "Authorization": API_KEY_HEADER
    }

# Obtener todos los datos de una cuota por su codigo
def obtener_datos_cuota(codigo_cuota):

    url = f"{os.getenv('url_obtener_datos_cuota')}{codigo_cuota}"

    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                "estado": data.get("estado"),
                "importe": data.get("importe"),
                "fkusuario" : data.get("fkusuario")
                }
        else:
            print(f"[API]: {response.text}")
            return None
    except Exception as e:
        print(f"[API] Error consultando la API para el codigo de  cuota {codigo_cuota}: {e}")
        return None

# Obtener solo estado de esa cuota por Id
def obtener_estado_cuota(id_cuota):

    url = f"{os.getenv('url_obtener_estado_cuota')}{id_cuota}/estado"

    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return data.get("estado", None)
        else:
            print(f"[API] No se pudo obtener el estado de la cuota {id_cuota}: {response.text}")
            return None
    except Exception as e:
        print(f"[API] Error consultando la API para cuota {id_cuota}: {e}")
        return None

# Metodo para obtener una lista de datos por fk_compania
def ObtenerListadeDatosporFk_Compania(url,fk_compania):

    urlfinal = f"{url}{fk_compania}"

    try:
        response = requests.get(urlfinal,headers=headers,verify=True,timeout=(10, 60))
        if response.status_code == 200:
            return response.json()  # lista de dicts
        print(f"[API] Error HTTP {response.status_code}: {response.text}")
        return None
    except ReadTimeout:
        print("[API] ⏱️ Timeout leyendo respuesta de la API")
        return None
    except ConnectTimeout:
        print("[API] 🔌 Timeout conectando a la API")
        return None
    except RequestException as e:
        print(f"[API] ❌ Error de requests: {e}")
        return None