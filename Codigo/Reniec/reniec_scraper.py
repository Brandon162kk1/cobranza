from playwright.sync_api import Page
import time

def open_reniec_page(page: Page, url: str):

    for intento in range(3):

        try:
            #page.goto(url,wait_until="networkidle",timeout=60000)
            page.goto(url,wait_until="domcontentloaded",timeout=60000)
            break
        except Exception as e:
            print(f"⚠️ Reintentando RENIEC ({intento+1}/3): {e}")
            time.sleep(3)
    else:
        raise Exception("RENIEC no respondió")

    return page

def buscar_dni(page: Page, dni: str):

    page.locator("#dni").wait_for(state="visible",timeout=10000)
    page.locator("#dni").fill(dni)

    page.get_by_role("button",name="Buscar datos").click()