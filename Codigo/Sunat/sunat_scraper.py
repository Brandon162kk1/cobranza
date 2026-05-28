from playwright.sync_api import Page
import os,time

#----- Variables de Entorno -------
url_ruc = os.getenv("url_ruc")

def open_sunat_page(page: Page, ruc: str):

    for intento in range(3):

        try:
            page.goto(url_ruc,wait_until="networkidle",timeout=60000)
            break
        except Exception as e:
            print(f"⚠️ Reintentando SUNAT ({intento+1}/3): {e}")
            time.sleep(3)
    else:
        raise Exception("SUNAT no respondió")

    page.wait_for_selector("#txtRuc", timeout=15000)

    campo = page.locator("#txtRuc")
    campo.fill("")
    campo.fill(ruc)

    page.locator("#btnAceptar").click()

    page.wait_for_selector(
        "xpath=//h4[contains(.,'Número de RUC') or contains(.,'Actividad')]",
        timeout=30000
    )

    return page