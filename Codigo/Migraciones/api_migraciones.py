
from fastapi import FastAPI, Header, HTTPException, Depends
from pydantic import BaseModel
from playwright.sync_api import sync_playwright, Page
from PIL import Image
from Codigo.GoogleChrome.fecha_y_hora import get_timestamp
import pytesseract
import os
import re
import time

API_KEY = os.getenv("API_KEY_MIGRACIONES")
url_migraciones = os.getenv("url_migraciones")

if not API_KEY or not url_migraciones:
    raise Exception("Variables de entorno no cargadas")

app = FastAPI(
    title="API MIGRACIONES",
    version="1.0.0"
)

def auth(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(
            status_code=401,
            detail="API Key inválida"
        )

class RucRequest(BaseModel):
    ce: str

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

@app.post("/consultar-ce")
def consultar(data: RucRequest, auth=Depends(auth)):

    playwright = None
    browser = None

    try:

        ce = str(data.ce).strip()

        if not ce:
            raise HTTPException(
                status_code=400,
                detail="Carné de extranjería inválido"
            )

        playwright, browser, page = get_page()

        return consultar_ce_service(
            page=page,
            ce=ce
        )

    except Exception as e:

        print(f"Error al consultar: {str(e)}")

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

    finally:

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

def consultar_ce_service(page, ce: str):

    open_migraciones_page(
        page=page,
        ce=ce
    )

    return parse_migraciones(
        page=page,
        ce=ce
    )

def leer_captcha(page):

    captcha = page.locator(
        "img[src*='CaptchaImage.axd']"
    )

    timestamp = get_timestamp()

    nombre_archivo = f"captcha_{timestamp}.png"
    nombre_procesado = f"captcha_procesado_{timestamp}.png"

    captcha.screenshot(
        path=nombre_archivo
    )

    img = Image.open(nombre_archivo)

    img = img.resize(
        (img.width * 3, img.height * 3)
    )

    img = img.convert("L")

    img = img.point(
        lambda x: 0 if x < 150 else 255,
        mode="1"
    )

    img.save(nombre_procesado)

    texto = pytesseract.image_to_string(
        img,
        config="--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    )

    texto = re.sub(
        r'[^A-Za-z0-9]',
        '',
        texto
    )

    texto = texto.upper().strip()

    print(f"Archivo original : {nombre_archivo}")
    print(f"Archivo procesado: {nombre_procesado}")
    print(f"Codigo OCR: [{texto}]")

    return texto

def open_migraciones_page(page: Page, ce: str):

    for intento in range(3):

        try:

            page.goto(
                url_migraciones,
                wait_until="networkidle",
                timeout=60000
            )
            
            # Dar tiempo a que termine de cargar todo
            page.wait_for_timeout(2000)

            # Refrescar
            page.reload(
                wait_until="networkidle"
            )

            page.wait_for_timeout(2000)

            break

        except Exception as e:

            print(
                f"⚠️ Reintentando Migraciones ({intento+1}/3): {e}"
            )

            time.sleep(3)

    else:
        raise Exception(
            "Migraciones no respondió"
        )

    page.wait_for_selector(
        "#ctl00_bodypage_txtnumerodoc",
        timeout=15000
    )

    page.locator(
        "#ctl00_bodypage_txtnumerodoc"
    ).fill(ce)

    page.locator(
        "#ctl00_bodypage_cbodia"
    ).select_option("14")

    page.locator(
        "#ctl00_bodypage_cbomes"
    ).select_option("9")

    page.locator(
        "#ctl00_bodypage_cboanio"
    ).select_option("1992")

    for intento in range(5):

        img_captcha = page.locator(
            "img[src*='CaptchaImage.axd']"
        )

        page.wait_for_timeout(1500)

        src_antes = img_captcha.get_attribute("src")

        print(
            f"SRC antes OCR: {src_antes}"
        )

        codigo = leer_captcha(page)

        captcha_input = page.locator(
            "#ctl00_bodypage_txtvalidator"
        )

        captcha_input.click()

        captcha_input.clear()

        captcha_input.type(
            codigo,
            delay=100
        )

        # print(
        #     f"Valor ingresado: "
        #     f"[{captcha_input.input_value()}]"
        # )

        # src_antes_click = img_captcha.get_attribute(
        #     "src"
        # )

        # print(
        #     f"SRC antes click: "
        #     f"{src_antes_click}"
        # )

        # print(
        #     "CE:",
        #     page.locator(
        #         "#ctl00_bodypage_txtnumerodoc"
        #     ).input_value()
        # )

        # print(
        #     "DIA:",
        #     page.locator(
        #         "#ctl00_bodypage_cbodia"
        #     ).input_value()
        # )

        # print(
        #     "MES:",
        #     page.locator(
        #         "#ctl00_bodypage_cbomes"
        #     ).input_value()
        # )

        # print(
        #     "AÑO:",
        #     page.locator(
        #         "#ctl00_bodypage_cboanio"
        #     ).input_value()
        # )

        # print(
        #     "CAPTCHA:",
        #     page.locator(
        #         "#ctl00_bodypage_txtvalidator"
        #     ).input_value()
        # )

        page.locator(
            "#ctl00_bodypage_btnverificar"
        ).click()

        page.wait_for_timeout(3000)

        print(
            f"Valor despues del click: "
            f"[{captcha_input.input_value()}]"
        )

        mensaje_locator = page.locator(
            "#ctl00_bodypage_lblmensaje"
        )

        if mensaje_locator.count() > 0:
            mensaje = mensaje_locator.text_content()
        else:
            mensaje = None

        page.screenshot(
            path="resultado_actual.png",
            full_page=True
        )

        print(f"Mensaje devuelto: [{mensaje}]")
        raise Exception(
            f"Migraciones devolvió mensaje: {mensaje}"
        )

    else:

        raise Exception(
            "No se pudo resolver el captcha"
        )

    page.wait_for_selector(
        "#ctl00_bodypage_lblnombre",
        state="attached",
        timeout=30000
    )

    return page

def parse_migraciones(page: Page, ce: str):

    resultado = {
        "numero_documento": ce,
        "nombres": "",
        "apellidos": "",
        "nacionalidad": "",
        "fecha_nacimiento": ""
    }

    nombre_completo = page.locator(
        "#ctl00_bodypage_lblnombre"
    ).text_content().strip()

    nacionalidad = page.locator(
        "#ctl00_bodypage_lblnacionalidad"
    ).text_content().strip()

    fecha_nacimiento = page.locator(
        "#ctl00_bodypage_lblfecnac"
    ).text_content().strip()

    resultado["nacionalidad"] = nacionalidad
    resultado["fecha_nacimiento"] = fecha_nacimiento

    if "," in nombre_completo:

        apellidos, nombres = nombre_completo.split(
            ",",
            1
        )

        resultado["apellidos"] = apellidos.strip()
        resultado["nombres"] = nombres.strip()

    else:

        resultado["nombres"] = nombre_completo

    return resultado