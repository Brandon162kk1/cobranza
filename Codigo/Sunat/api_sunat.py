from fastapi import FastAPI, Header, HTTPException, UploadFile, File, Depends
from fastapi.responses import FileResponse
from pydantic import BaseModel
from playwright.sync_api import sync_playwright
import tempfile
import pandas as pd
import os

from Codigo.Sunat.sunat_service import consultar_ruc_service
from Codigo.GoogleChrome.fecha_y_hora import get_timestamp

# ENV
API_KEY = os.getenv("API_KEY_SUNAT")
url_ruc = os.getenv("url_ruc")

if not API_KEY or not url_ruc:
    raise Exception("Variables de entorno no cargadas")

app = FastAPI(title="API SUNAT RUC", version="1.0.0")
# ---------------- AUTH ----------------
def auth(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API Key inválida")

class RucRequest(BaseModel):
    ruc: str

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

@app.post("/consultar-ruc")
def consultar(data: RucRequest, auth=Depends(auth)):

    playwright = None
    browser = None

    try:

        # limpiar RUC/documento
        ruc = str(data.ruc).strip()

        if not ruc:
            raise HTTPException(
                status_code=400,
                detail="RUC inválido"
            )

        playwright, browser, page = get_page()

        return consultar_ruc_service(page, data.ruc)

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

@app.post("/consultar-ruc-excel")
def consultar_excel(file: UploadFile = File(...),auth=Depends(auth)):

    temp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    temp.write(file.file.read())
    temp.close()

    df = pd.read_excel(
        temp.name,
        dtype=str
    )

    df.columns = [c.strip().lower() for c in df.columns]

    COLUMNAS_RUC = [
        "ruc",
        "numerodocumento",
        "numero_documento",
        "numero documento",
        "nrodocumento",
        "documento"
    ]

    columna_ruc = next(
        (col for col in COLUMNAS_RUC if col in df.columns),
        None
    )

    if not columna_ruc:
        raise HTTPException(
            status_code=400,
            detail="No se encontró una columna válida de RUC/documento"
        )

    columnas = [
        "tipo_documento",
        "ruc",
        "numero_documento",
        "nombres",
        "apellidos",
        "razon_social",
        "fecha_inicio",
        "estado",
        "domicilio_fiscal",
        "provincia",
        "ciudad",
        "distrito",
        "nombre_comercial",
        "cod_principal",
        "actividad_principal",
        "cod_secundario_1",
        "actividad_1",
        "cod_secundario_2",
        "actividad_2"
    ]

    for col in columnas:

        if col not in df.columns:
            df[col] = ""

    playwright = None
    browser = None

    try:

        playwright, browser, page = get_page()

        for i, row in df.iterrows():

            ruc = str(row[columna_ruc]).strip()

            if not ruc or ruc.lower() == "nan":
                continue

            try:

                resultado = consultar_ruc_service(
                    page,
                    ruc
                )

            except Exception as e:

                print(f"Error consultando {ruc}: {e}")
                resultado = {}

            for k, v in resultado.items():
                df.at[i, k] = v

    finally:

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

    output = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    df.to_excel(output.name, index=False)

    return FileResponse(
        output.name,
        filename=f"Resultado_{get_timestamp()}.xlsx"
    )