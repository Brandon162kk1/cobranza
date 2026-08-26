#-- Imports ---
import os
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException,TimeoutException
from datetime import datetime
from .urls import url_agregar_comprobante

#----- Datos -------
fecha_creacion_birlik = datetime(2022,12,5)
#----- Variables de Entorno -------
login_birlik = os.getenv("login_birlik")
usuarioBirlik = os.getenv("usuarioBirlik")
passwordBirlik = os.getenv("passwordBirlik")

def agregar_comprobante_pago(driver, wait, id_cuota_birlik, ruta_factura):

    ventana_principal_cia = driver.current_window_handle

    try:

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        url_final = url_agregar_comprobante + id_cuota_birlik
        driver.get(url_final)

        id_input_comprobante = "customFile_comprobante"

        iniciar_sesion_birlik(wait,id_input_comprobante)

        archivo_comprobante = wait.until(EC.presence_of_element_located((By.ID,id_input_comprobante)))
        archivo_comprobante.send_keys(ruta_factura)
        print("✅ Comprobante de Pago cargado")

        boton_adjuntar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(text(), "ADJUNTAR")]')))
        boton_adjuntar.click()
        print("🖱️ Clic en 'ADJUNTAR'")
        print("✅ Proceso completado")

        return True

    except TimeoutException as e:
        print(f"⏳ Timeout en Birlik: {e}")
        return False
    except StaleElementReferenceException as e:
        print(f"⚠ DOM modificado: {e}")
        return False
    except Exception as e:
        print(f"❌ Error en Birlik: {e}")
        return False
    finally:
        driver.close()
        driver.switch_to.window(ventana_principal_cia)
        print("🔄 Cerrando Birlik y volviendo a la CIA")

def cancelar_y_agregar_cuota(driver, wait, id_cuota,comprobante_valor,fecha_emision,ruta_factura,ruta_imagen_sunat,resultado_importe):
    
    ventana_principal_cia = driver.current_window_handle

    try:

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        url_final = login_birlik + id_cuota
        driver.get(url_final)
        
        id_input_factura = "factura"

        iniciar_sesion_birlik(wait,id_input_factura)

        factura_input = wait.until(EC.presence_of_element_located((By.ID,id_input_factura)))
        factura_input.clear()
        factura_input.send_keys(ruta_factura)
        print("✅ Factura subida")
        
        constancia_input = wait.until(EC.presence_of_element_located((By.ID, "constancia")))
        constancia_input.clear()
        constancia_input.send_keys(ruta_imagen_sunat)
        print("✅ Constancia de SUNAT subida")

        archivo_input_1 = wait.until(EC.presence_of_element_located((By.ID, "comprobante")))
        archivo_input_1.send_keys(ruta_factura)
        print("✅ Comprobante subido")

        observacion_input = wait.until(EC.presence_of_element_located((By.ID, "observacionCuota")))
        observacion_input.clear()
        if not resultado_importe:
            observacion_input.send_keys("Importes no coinciden")
        else:
            observacion_input.send_keys("Cancelado por robot")
        print("✅ Observación ingresada")

        input_factura = wait.until(EC.presence_of_element_located((By.ID, "numeroFactura")))
        input_factura.clear()
        input_factura.send_keys(comprobante_valor)
        print(f"🧾 Número de factura ingresado: {comprobante_valor}")

        # Ojo BIRLIK espera el formato MM/DD/YYYY
        fecha_dt = datetime.strptime(fecha_emision, "%d/%m/%Y")
        fecha_vista_factura = fecha_dt.strftime("%d/%m/%Y") #Asi se ve en la factura , sin hora ni minutos
        fecha_formateada_js  = fecha_dt.strftime("%Y-%m-%d")  # 👈 Formato compatible con <input type="date">

        input_fecha = wait.until(EC.presence_of_element_located((By.ID, "fechaPago")))
        input_fecha.clear()
        driver.execute_script("arguments[0].value = arguments[1];", input_fecha, fecha_formateada_js)
        print(f"📅 Fecha de Emisión de la Factura: {fecha_vista_factura}, pero se ingresa así (Y/m/d): {fecha_formateada_js}")

        btn_registrar = wait.until(EC.element_to_be_clickable((By.ID, "btnRegistrar")))
        btn_registrar.click()
        print("🖱️ Clic en registrar")

        wait.until(EC.invisibility_of_element_located((By.ID, "btnRegistrar")))
        print("✅ Registro completado")
        return True

    except TimeoutException as e:
        print(f"⏳ Timeout en Birlik: {e}")
        return False
    except StaleElementReferenceException as e:
        print(f"⚠ DOM modificado: {e}")
        return False
    except Exception as e:
        print(f"❌ Error en Birlik: {e}")
        return False
    finally:
        driver.close()
        driver.switch_to.window(ventana_principal_cia)
        print("🔄 Cerrando Birlik y volviendo a la CIA")

def iniciar_sesion_birlik(wait,id_elemento):

    input_email = (By.ID, "signinSrEmail")
    input_password = (By.ID, "signupSrPassword")

    # Un elemento que SOLO aparece cuando ya estás logueado
    elemento_logueado = (By.ID,id_elemento)

    resultado = wait.until(
        EC.any_of(
            EC.presence_of_element_located(input_email),
            EC.presence_of_element_located(elemento_logueado)
        )
    )

    element_id = resultado.get_attribute("id")

    if element_id in ("factura", "customFile_comprobante"):
        print("✅ Sesión ya activa")
        return

    print("🔐 Login requerido")

    email_input = resultado

    password_input = wait.until(EC.presence_of_element_located(input_password))
    email_input.clear()
    email_input.send_keys(usuarioBirlik)

    password_input.clear()
    password_input.send_keys(passwordBirlik)
    password_input.send_keys(Keys.RETURN)

    # Esperar confirmación de login
    wait.until(EC.presence_of_element_located(elemento_logueado))
    print("🔑 Sesión iniciada correctamente")
    return