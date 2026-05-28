import re
from playwright.sync_api import Page

def get_val(label,page):
    return page.locator(f"//h4[contains(text(),'{label}')]/ancestor::div[@class='col-sm-3']/following-sibling::div[@class='col-sm-3'][1]/p").inner_text().strip()

def obt_nom_com(label, page):
    try:
        locator = page.locator(
            f"xpath=//h4[contains(.,'{label}')]/ancestor::div[contains(@class,'row')]/following-sibling::div//p | //h4[contains(.,'{label}')]/following::p[1]"
        )

        if locator.count() == 0:
            return ""

        texto = locator.first.inner_text().strip()
        return re.sub(r"\s+", " ", texto)

    except:
        return ""

def parse_sunat(page: Page, ruc: str):

    resultado = {

        "tipo_documento": "" if str(ruc).startswith("10") else "RUC",
        "ruc": ruc,
        "numero_documento": "",
        "nombres": "",
        "apellidos" : "",
        "razon_social": "",
        "fecha_inicio": "",
        "estado": "",
        "domicilio_fiscal": "",
        "provincia": "",
        "ciudad": "",
        "distrito": "",
        "nombre_comercial": "",
        "cod_principal": "",
        "actividad_principal": "",
        "cod_secundario_1": "",
        "actividad_1": "",
        "cod_secundario_2": "",
        "actividad_2": ""
    }

    razon_social_locator = page.locator("//h4[contains(text(),'Número de RUC:')]/ancestor::div[@class='row']//div[@class='col-sm-7']/h4")
    texto_razon = razon_social_locator.text_content().strip()

    if " - " in texto_razon:
        razon_social = texto_razon.split(" - ", 1)[1].strip()
        resultado["razon_social"] = razon_social
    

    if str(ruc).startswith("10"):

        texto_documento = page.locator(
            "div.list-group-item:has(h4:text('Tipo de Documento:')) p.list-group-item-text"
        ).inner_text().strip()

        texto_documento = " ".join(texto_documento.split())

        match = re.search(
            r"(\w+)\s+(\d+)\s*-\s*(.+)",
            texto_documento
        )

        if match:

            resultado["tipo_documento"] = match.group(1).strip()
            resultado["numero_documento"] = match.group(2).strip()

            nombre_persona = match.group(3).strip()

            partes_nombre = nombre_persona.split(",")

            resultado["apellidos"] = partes_nombre[0].strip()

            resultado["nombres"] = (
                partes_nombre[1].strip()
                if len(partes_nombre) > 1
                else ""
            )

    #resultado["fecha_inscripcion"] = get_val("Fecha de Inscripción",page)
    resultado["fecha_inicio"] = get_val("Fecha de Inicio de Actividades",page)

    estado = page.locator(
        "h4:has-text('Estado del Contribuyente:')"
    ).locator("..").locator("..").locator(
        "p.list-group-item-text"
    ).inner_text().strip()
    resultado["estado"] = estado

    resultado["nombre_comercial"] = obt_nom_com("Nombre Comercial:", page)

    try:

        domicilio = page.locator("//h4[contains(.,'Domicilio Fiscal')]/ancestor::div[contains(@class,'col-sm-5')]/following-sibling::div[contains(@class,'col-sm-7')]//p").inner_text().strip()       
        domicilio = re.sub(r"\s+", " ", domicilio).replace("\xa0", " ")
        resultado["domicilio_fiscal"] = domicilio  # siempre guardar completo

        match = re.search(
            r"(.*)\s([A-ZÁÉÍÓÚÑ ]+)\s-\s([A-ZÁÉÍÓÚÑ ]+)\s-\s([A-ZÁÉÍÓÚÑ ]+)\s*$",
            domicilio
        )

        if match:
            resultado["domicilio_fiscal"] = match.group(1).strip()
            resultado["provincia"] = match.group(2).strip()
            resultado["ciudad"] = match.group(3).strip()
            resultado["distrito"] = match.group(4).strip()

        else:
            partes = domicilio.split(" - ")

            if len(partes) == 3:
                resultado["provincia"] = partes[0].strip()
                resultado["ciudad"] = partes[1].strip()
                resultado["distrito"] = partes[2].strip()

            elif len(partes) == 2:
                resultado["provincia"] = partes[0].strip()
                resultado["ciudad"] = partes[1].strip()

            # si no hay nada → se queda vacío

    except:
            pass

    try:
        filas = page.locator("//table[contains(@class,'tblResultado')]//tr")
        actividades = {}

        for i in range(filas.count()):
            texto = filas.nth(i).inner_text().strip()

            if "-" not in texto:
                continue

            partes = [p.strip() for p in texto.split("-")]

            if len(partes) >= 3:
                actividades[partes[0]] = {
                    "codigo": partes[1],
                    "actividad": partes[2]
                }

        if "Principal" in actividades:
            resultado["cod_principal"] = actividades["Principal"]["codigo"]
            resultado["actividad_principal"] = actividades["Principal"]["actividad"]

        if "Secundaria 1" in actividades:
            resultado["cod_secundario_1"] = actividades["Secundaria 1"]["codigo"]
            resultado["actividad_1"] = actividades["Secundaria 1"]["actividad"]

        if "Secundaria 2" in actividades:
            resultado["cod_secundario_2"] = actividades["Secundaria 2"]["codigo"]
            resultado["actividad_2"] = actividades["Secundaria 2"]["actividad"]

    except:
        pass

    return resultado