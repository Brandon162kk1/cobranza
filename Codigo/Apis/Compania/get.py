import requests
import time

def codigo_compania(url,api_key):

    headers = {
        "x-api-key": f"{api_key}"
    }

    while True:
        resp = requests.get(f"{url}",headers=headers)

        if resp.status_code == 200:
            codigo = resp.json()["codigo"]
            print(f"✅ Código recibido: {codigo}")
            break

        time.sleep(2)

    return codigo