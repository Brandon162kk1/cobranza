from playwright.sync_api import Page

def parse_reniec(page: Page, dni: str):

    try:

        filas = page.locator("tbody tr")

        if filas.count() == 0:
            return None

        fila = filas.first
        tds = fila.locator("td")

        if tds.count() < 4:
            return None

        #numero = tds.nth(0).inner_text().strip()
        nombres = tds.nth(1).inner_text().strip()
        ap_paterno = tds.nth(2).inner_text().strip()
        ap_materno = tds.nth(3).inner_text().strip()

        return {
            "nombres": nombres,
            "apellido_paterno": ap_paterno,
            "apellido_materno": ap_materno
        }

    except Exception as e:
        print(f"Error parseando datos: {e}")
        return None