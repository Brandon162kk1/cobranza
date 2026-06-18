from fastapi import FastAPI, Header,UploadFile,File, HTTPException, Depends
from pydantic import BaseModel
from fastapi.responses import FileResponse
from playwright.sync_api import sync_playwright
from Codigo.Reniec.reniec_service import consultar_dni_service
from Codigo.GoogleChrome.fecha_y_hora import get_timestamp
from fastapi.middleware.cors import CORSMiddleware

import os,time
import tempfile
import pandas as pd

app = FastAPI(title="API RENIEC",version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

#----- Variables de Entorno -------
API_KEY = os.getenv("API_KEY_RENIEC")
url_reniec = os.getenv("url_reniec")

if not API_KEY or not url_reniec:
    raise Exception("Variables de entorno no cargadas")

# ---------------- AUTH ----------------
def auth(x_api_key: str = Header(...)):
    if x_api_key != API_KEY:
        raise HTTPException(
            status_code=401,
            detail="API Key inválida"
        )

class DniRequest(BaseModel):
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
        raise HTTPException(status_code=400,detail="El DNI debe contener solo números")

    if len(dni) != 8:
        raise HTTPException(status_code=400,detail="El DNI debe tener exactamente 8 dígitos")

    return dni

@app.post("/consultar-dni")
def consultar(data: DniRequest, auth=Depends(auth)):

    playwright = None
    browser = None

    try:
        dni = validar_dni(str(data.dni))
        if not dni:
            raise HTTPException(status_code=400,detail="DNI inválido")
        playwright, browser, page = get_page()
        return consultar_dni_service(page=page,dni=dni)
    except HTTPException:
        raise
    except Exception as e:
        print(f"Error al consultar: {str(e)}")
        raise HTTPException(status_code=500,detail=f"Error interno: {str(e)}")
    finally:
        if browser:
            browser.close()
        if playwright:
            playwright.stop()

@app.post("/consultar-dni-excel")
def consultar_excel(file: UploadFile = File(...),auth=Depends(auth)):

    playwright = None
    browser = None

    try:

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

        COLUMNAS_DNI = [
            "dni",
            "numerodocumento",
            "numero_documento",
            "numero documento",
            "nrodocumento",
            "documento"
        ]

        columna_dni = next(
            (col for col in COLUMNAS_DNI if col in df.columns),
            None
        )

        if not columna_dni:
            raise HTTPException(
                status_code=400,
                detail="No se encontró una columna válida de DNI/documento"
            )

        columnas = [
            "nombres",
            "apellido_paterno",
            "apellido_materno"
        ]

        for col in columnas:

            if col not in df.columns:
                df[col] = ""

        # Nueva columna para errores/observaciones
        if "observacion" not in df.columns:
            df["observacion"] = ""

        playwright = None
        browser = None

        playwright, browser, page = get_page()

        contador = 0

        for i, row in df.iterrows():

            dni = str(row[columna_dni]).strip()

            if not dni or dni.lower() == "nan":
                df.at[i, "observacion"] = "DNI vacío"
                continue

            dni = dni.zfill(8)
            df.at[i, columna_dni] = dni

            resultado = {}

            try:
                dni = validar_dni(dni)
                resultado = consultar_dni_service(page,dni)
                df.at[i, "observacion"] = "OK"
            except HTTPException as e:
                df.at[i, "observacion"] = e.detail
            except Exception as e:
                df.at[i, "observacion"] = str(e)
            finally:
                contador += 1
                time.sleep(1)

                if contador >= 10:

                    print("Reiniciando página...")
                    try:
                        page.close()
                    except:
                        pass
                    page = browser.new_page()
                    contador = 0

            for k, v in resultado.items():
                df.at[i, k] = v

        output = tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx")

        #-------------
        # Crear nombre completo concatenado
        df["nombre_completo"] = (
            df["nombres"].fillna("") + " " +
            df["apellido_paterno"].fillna("") + " " +
            df["apellido_materno"].fillna("")
        ).str.strip()

        # Comparar con cliente
        df["coincide"] = (
            (
                df["cliente"]
                .fillna("")
                .str.upper()
                .str.split()
                .str.join(" ")
                ==
                df["nombre_completo"]
                .fillna("")
                .str.upper()
                .str.split()
                .str.join(" ")
            )
            .map({True: "SI", False: "NO"})
        )

        df["Actualizar"] = df.apply(
            lambda row:
                f"UPDATE CLIENTE SET NombreTipoPersona = '{row['nombre_completo']}' "
                f"WHERE Id_Cliente = {row['id_cliente']}"
                if (
                    row["coincide"] == "NO"
                    and pd.notna(row["nombre_completo"])
                    and str(row["nombre_completo"]).strip() != ""
                )
                else "",
            axis=1
        )
        #-------------

        df.to_excel(output.name, index=False)
        return FileResponse(output.name,filename=f"Resultado_DNI_{get_timestamp()}.xlsx")

    except HTTPException:
        raise

    except Exception as e:

        import traceback

        error = traceback.format_exc()

        print(error)

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )

    finally:

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

    # temp = tempfile.NamedTemporaryFile(
    #     delete=False,
    #     suffix=".xlsx"
    # )

    # temp.write(file.file.read())
    # temp.close()

    # df = pd.read_excel(
    #     temp.name,
    #     dtype=str
    # )

    # df.columns = [c.strip().lower() for c in df.columns]

    # COLUMNAS_DNI = [
    #     "dni",
    #     "numerodocumento",
    #     "numero_documento",
    #     "numero documento",
    #     "nrodocumento",
    #     "documento"
    # ]

    # columna_dni = next(
    #     (col for col in COLUMNAS_DNI if col in df.columns),
    #     None
    # )

    # if not columna_dni:
    #     raise HTTPException(
    #         status_code=400,
    #         detail="No se encontró una columna válida de DNI/documento"
    #     )

    # columnas = [
    #     "nombres",
    #     "apellido_paterno",
    #     "apellido_materno"
    # ]

    # for col in columnas:

    #     if col not in df.columns:
    #         df[col] = ""

    # # Nueva columna para errores/observaciones
    # if "observacion" not in df.columns:
    #     df["observacion"] = ""

    # playwright = None
    # browser = None

    # try:

    #     playwright, browser, page = get_page()

    #     for i, row in df.iterrows():

    #         dni = str(row[columna_dni]).strip()

    #         if not dni or dni.lower() == "nan":
    #             df.at[i, "observacion"] = "DNI vacío"
    #             continue

    #         dni = dni.zfill(8)
    #         df.at[i, columna_dni] = dni

    #         resultado = {}

    #         try:
    #             dni = validar_dni(dni)
    #             resultado = consultar_dni_service(page,dni)
    #             df.at[i, "observacion"] = "OK"
    #         except HTTPException as e:
    #             df.at[i, "observacion"] = e.detail
    #         except Exception as e:
    #             df.at[i, "observacion"] = str(e)

    #         for k, v in resultado.items():
    #             df.at[i, k] = v

    # finally:

    #     if browser:
    #         browser.close()

    #     if playwright:
    #         playwright.stop()

    # output = tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx")

    # #-------------
    # # Crear nombre completo concatenado
    # df["nombre_completo"] = (
    #     df["nombres"].fillna("") + " " +
    #     df["apellido_paterno"].fillna("") + " " +
    #     df["apellido_materno"].fillna("")
    # ).str.strip()

    # # Comparar con cliente
    # df["coincide"] = (
    #     (
    #         df["cliente"]
    #         .fillna("")
    #         .str.upper()
    #         .str.split()
    #         .str.join(" ")
    #         ==
    #         df["nombre_completo"]
    #         .fillna("")
    #         .str.upper()
    #         .str.split()
    #         .str.join(" ")
    #     )
    #     .map({True: "SI", False: "NO"})
    # )

    # df["Actualizar"] = df.apply(
    #     lambda row:
    #         f"UPDATE CLIENTE SET NombreTipoPersona = '{row['nombre_completo']}' "
    #         f"WHERE Id_Cliente = {row['id_cliente']}"
    #         if row["coincide"] == "NO"
    #         else "",
    #     axis=1
    # )
    # #-------------

    # df.to_excel(output.name, index=False)
    # return FileResponse(output.name,filename=f"Resultado_DNI_{get_timestamp()}.xlsx")