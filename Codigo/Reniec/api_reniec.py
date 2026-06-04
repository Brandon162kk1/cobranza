
from fastapi import FastAPI, Header, HTTPException, Depends
from pydantic import BaseModel
from playwright.sync_api import sync_playwright, Page
import os
import time

API_KEY = os.getenv("API_KEY_RENIEC")
url_reniec = os.getenv("url_reniec")

if not API_KEY or not url_reniec:
    raise Exception("Variables de entorno no cargadas")

app = FastAPI(
    title="API RENIEC",
    version="1.0.0"
)

def auth(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(
            status_code=401,
            detail="API Key inválida"
        )

class RucRequest(BaseModel):
    dni: str

def get_page():

    playwright = sync_playwright().start()

    browser = playwright.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--disable-web-security"
        ]
    )

    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        viewport={"width": 1366, "height": 768},
        locale="es-PE"
    )

    page = context.new_page()

    return playwright, browser, page

def validar_dni(dni: str):

    dni = dni.strip()

    if not dni.isdigit():
        raise HTTPException(
            status_code=400,
            detail="El DNI debe contener solo números"
        )

    if len(dni) != 8:
        raise HTTPException(
            status_code=400,
            detail="El DNI debe tener exactamente 8 dígitos"
        )

    return dni

@app.post("/consultar-dni")
def consultar(data: RucRequest, auth=Depends(auth)):

    playwright = None
    browser = None

    try:

        dni = validar_dni(str(data.dni))

        if not dni:
            raise HTTPException(
                status_code=400,
                detail="DNI inválido"
            )

        playwright, browser, page = get_page()

        return consultar_dni_service(page=page,dni=dni)

    except Exception as e:
        print(f"Error al consultar: {str(e)}")
        raise HTTPException(status_code=500,detail=str(e))

    finally:

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

def consultar_dni_service(page, dni: str):

    open_reniec_page(page=page,dni=dni)
    buscar_dni(page, dni)
    return parse_reniec(page=page,dni=dni)

def open_reniec_page(page: Page, dni: str):

    for intento in range(3):

        try:

            page.goto(url_reniec,wait_until="networkidle",timeout=60000)
            
            page.wait_for_timeout(2000)

            page.reload(wait_until="networkidle")

            page.wait_for_timeout(2000)

            break

        except Exception as e:

            print(f"⚠️ Reintentando Reniec ({intento+1}/3): {e}")
            time.sleep(3)

    else:
        raise Exception("Reniec no respondió")

    return page

def buscar_dni(page: Page, dni: str):

    # Esperar input visible
    dni_input = page.locator("#dni")
    dni_input.wait_for(state="visible", timeout=15000)

    # Llenar DNI
    dni_input.fill(dni)

    # Click buscar
    page.get_by_role("button", name="Buscar datos").click()

    # Esperar que cargue la tabla
    page.wait_for_selector("tbody tr", timeout=15000)

def parse_reniec(page: Page, dni: str):

    page.wait_for_selector("tbody tr", timeout=15000)

    fila = page.locator("tbody tr").first

    tds = fila.locator("td")

    if tds.count() < 4:
        raise HTTPException(status_code=404,detail="DNI no encontrado")

    numero = tds.nth(0).inner_text().strip()
    nombres = tds.nth(1).inner_text().strip()
    ap_paterno = tds.nth(2).inner_text().strip()
    ap_materno = tds.nth(3).inner_text().strip()

    if not numero or not nombres:
        raise HTTPException(status_code=404,detail="DNI no encontrado")

    return {
        "numero_documento": numero,
        "nombres": nombres,
        "apellido_paterno": ap_paterno,
        "apellido_materno": ap_materno
    }