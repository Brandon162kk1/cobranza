from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support.ui import Select
from Codigo.GoogleChrome.chromeDriver import tomar_captura
import os

#----- Variables de Entorno -------
url_sunat = os.getenv("url_factura")

def es_pagina_bloqueada(html: str) -> bool:
    return (
        "The requested URL was rejected" in html
        and "Your support ID is" in html
    )

def consultarValidezSunat(driver,wait,ruc_compania,tipo_doc_birlik,ruc_cliente,comprobante,fecha_emision,monto,ruta_imagen_sunat,ruta_carpeta_errores):
    
    encontro = False

    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    
    driver.get(url_sunat)

    print(f"⌛ Ingresando a SUNAT para probar con la fecha '{fecha_emision}'")

    html = driver.page_source

    if es_pagina_bloqueada(html):
        print("🚫 IP / sesión bloqueada por firewall")
        tomar_captura(driver, ruta_carpeta_errores,f"ip_bloq_{comprobante}")
        return None

    campo_ruc_emi = wait.until(EC.element_to_be_clickable((By.NAME, "num_ruc")))
    campo_ruc_emi.clear()
    campo_ruc_emi.send_keys(ruc_compania)
    print(f"✅ RUC de la compania ingresado : {ruc_compania}")

    campo_ruc_rec = wait.until(EC.element_to_be_clickable((By.NAME, "num_docide")))
    campo_ruc_rec.clear()
    campo_ruc_rec.send_keys(ruc_cliente)
    print(f"✅ Documento Cliente : {ruc_cliente}")
    
    # Primeros 4 caracteres del "comprobante_valor"
    serie = comprobante[:4]
    campo_serie = wait.until(EC.element_to_be_clickable((By.NAME, "num_serie")))
    campo_serie.clear()
    campo_serie.send_keys(serie)
    print(f"✅ Serie: {serie}")

    # Desde el sexto carácter en adelante del "comprobante_valor"
    numero = comprobante[5:]
    campo_comprobante = wait.until(EC.element_to_be_clickable((By.NAME, "num_comprob")))
    campo_comprobante.clear()
    campo_comprobante.send_keys(numero)
    print(f"✅ Numero: {numero}")

    # Paso 9: Insertar el valor de la fecha en Sunat en formato dd/mm/yyyy
    campo_fecha = wait.until(EC.element_to_be_clickable((By.NAME, "fec_emision")))
    campo_fecha.clear()
    campo_fecha.send_keys(fecha_emision)
    print(f"✅ Fecha Emision: {fecha_emision}")

    # Suponiendo que ya tienes el driver y estás en la página, click afuera para que se vea el boton "Buscar"
    wait.until(EC.element_to_be_clickable((By.TAG_NAME, "body"))).click()

    campo_monto = wait.until(EC.element_to_be_clickable((By.NAME, "cantidad")))
    campo_monto.clear() 
    campo_monto.send_keys(monto)
    print(f"✅ Monto: {monto}")

    if tipo_doc_birlik == "RUC":
        tip_doc_rec = '6' #RUC
        tipo_compro = '03' #FE
        doc = 'Factura Electrónica'
        nom  = 'RUC'
    elif tipo_doc_birlik == "CEX":
        tip_doc_rec = '4' #CEX
        tipo_compro = '06' #FE
        doc = 'Boleta Electrónica'
        nom  = 'Carne Extranejeria'
    else:
        tip_doc_rec = '1'   #DNI
        tipo_compro = '06' #BE
        doc = 'Boleta Electrónica'
        nom  = 'DNI'

    select_element = wait.until(EC.element_to_be_clickable((By.NAME, "tipocomprobante")))
    select1 = Select(select_element)
    select1.select_by_value(tipo_compro)
    print(f"✅ Tipo Comprobante: {nom} - {tipo_compro}")

    select_element2 = wait.until(EC.element_to_be_clickable((By.NAME, "cod_docide")))
    select2 = Select(select_element2)
    select2.select_by_value(tip_doc_rec)
    print(f"✅ Tipo Doc Cliente: {tip_doc_rec}")

    boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Buscar']")))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_buscar)
    boton_buscar.click()

    print("🕐 Esperando resultado de SUNAT")

    try:

        xpath_resultado = (
            "//td[contains(., 'ha sido informada a SUNAT') "
            "or contains(., 'comprobante de pago válido') "
            "or contains(., 'no existe en los registros de SUNAT')]"
        )

        resultado = wait.until(
            EC.any_of(
                EC.visibility_of_element_located((By.XPATH, xpath_resultado)),
                lambda d: es_pagina_bloqueada(d.page_source)
            )
        )

        if resultado is True:
            print("🚫 IP / sesión bloqueada por firewall")
            tomar_captura(driver, ruta_carpeta_errores,f"ip_bloq_{comprobante}")
            return None

        mensaje = resultado.text.strip()

        if (
            "ha sido informada a SUNAT" in mensaje
            or "comprobante de pago válido" in mensaje
        ):

            print(f"✅ La {doc} es válida")
            driver.save_screenshot(ruta_imagen_sunat)
            encontro = True

        elif "no existe en los registros de SUNAT" in mensaje:
            print(f"❌ No se encontró la {doc} con la fecha: {fecha_emision}")

    except TimeoutException:
        print("⏳ SUNAT no respondió dentro del tiempo esperado")
    except WebDriverException as e:
        print(f"⚠️ Error de Selenium: {e}")
    finally:
        driver.close()
        print("✅ Cerrando la Pestaña SUNAT")
    
    return encontro