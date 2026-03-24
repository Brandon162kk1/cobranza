#--- Imports --
import pandas as pd 
import os
import time
import shutil
#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from dateutil.relativedelta import relativedelta
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.Birlik.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver,crearCarpetas,guardarJson,esperar_archivos_nuevos
from GoogleChrome.fecha_y_hora import get_timestamp
from Correo.correo_it import enviarCaptcha

#------------CRECER----------------
ruc_crecer_vly = '20600098633'
ids_compania = [32]
#-------------CREDENCIALES---------------
login_url_crecer_vida_ley = os.getenv("login_url_crecer_vida_ley")
username_crecer = os.getenv("username_crecer")
password_crecer = os.getenv("password_crecer")
para_venv = os.getenv("para")
para_lista = para_venv.split(",") if para_venv else []
copia_venv = os.getenv("copia_cuotas")
copias_lista = copia_venv.split(",") if copia_venv else []
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Crecer_VidaLey_{get_timestamp()}"
# --- Configuración 2Captcha---
API_KEY = "8b61feec172173ef48060a723af1b6c7"

#------- Errores comunes de 2Captcha
# ERROR_WRONG_USER_KEY → API key incorrecta.
# ERROR_KEY_DOES_NOT_EXIST → API key no válida.
# ERROR_ZERO_BALANCE → No tienes saldo.
# ERROR_WRONG_GOOGLEKEY → El sitekey no corresponde o no es válido.
# ERROR_PAGEURL → La URL de la página no es válida.

def procesar_fila(driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores):

    # Extraer valores y quitar espacios en blanco
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    ruc_cliente_birlik = str(row["numeroDocumento"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    numero_proforma_birlik = str(row["codigoCuota"]).strip()
    numero_poliza_birlik = str(row["numeroPoliza"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estado_Cuota_birlik = str(row["estadoCuota"]).strip()
    fk_Cliente_birlik = str(row["fk_Cliente"]).strip() 
    vigencia_inicio_birlik = str(row["vigenciaInicio"]).strip()

    # Convertir a objeto datetime
    fecha_inicio_dt = datetime.strptime(vigencia_inicio_birlik, "%d/%m/%Y")
    fecha_fin_dt = datetime.strptime(vigencia_inicio_birlik,"%d/%m/%Y")
    # Restar 1 mes
    fecha_menos_un_mes = fecha_inicio_dt - relativedelta(months=1)
    fecha_mas_un_mes = fecha_fin_dt + relativedelta(months=1)
    # Convertir de nuevo a string si lo necesitas en ese formato
    fecha_resultado_menos = fecha_menos_un_mes.strftime("%d/%m/%Y")
    fecha_resultado_mas = fecha_mas_un_mes.strftime("%d/%m/%Y")
   
    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    try:
        
        input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "sfullname")))
        input_fullname.clear()
        input_fullname.send_keys(ruc_cliente_birlik)
        print(f"✅ Se ingresó el RUC '{ruc_cliente_birlik}'.")

        # Quitar el atributo readonly con Javascript
        input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "solicitud_fecha_inicio")))
        #driver.find_element(By.ID, "solicitud_fecha_inicio")
        driver.execute_script("arguments[0].removeAttribute('readonly')", input_fecha)
        input_fecha.clear()
        input_fecha.send_keys(fecha_resultado_menos)
        print(f"✅ Se ingresó un mes antes de la fecha de inicio de la vigencia - {fecha_resultado_menos}")

        time.sleep(3)

        # Hacer click en un lugar vacío de la página (por ejemplo el body)
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element(body).click().perform()
        print("🖱️ Clic en un lugar vacío de la página para asegurar la interacción.")

        # # Quitar el atributo readonly con Javascript
        # input_fecha_fin = wait.until(EC.visibility_of_element_located((By.ID, "solicitud_fecha_fin")))
        # #driver.find_element(By.ID, "solicitud_fecha_fin")
        # driver.execute_script("arguments[0].removeAttribute('readonly')", input_fecha_fin)
        # input_fecha_fin.clear()
        # input_fecha_fin.send_keys(fecha_resultado_mas)
        # print("✅ Se ingresó un mes despues de la fecha fin de la vigencia")

        input_poliza = wait.until(EC.visibility_of_element_located((By.XPATH,"//label[normalize-space()='Póliza']/following-sibling::input")))
        input_poliza.clear()
        input_poliza.send_keys(numero_poliza_birlik)
        print(f"✅ Se ingresó la Póliza '{numero_poliza_birlik}'.")

        # Esperar que el select esté visible
        select_element = wait.until(EC.visibility_of_element_located((By.ID, "Estados")))
        select = Select(select_element)
        select.select_by_value("0")
        print("🖱️ Se selecciono la opcion 'Todos' ")

        # Esperar que el botón esté visible y clickable
        boton_consultar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Consultar')]")))
        # Click en el botón usando JavaScript (más efectivo en Angular/React)
        driver.execute_script("arguments[0].click();", boton_consultar)
        print("🖱️ Clic en Consultar")

        time.sleep(5)

        # Esperar a que el loader desaparezca
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

        # Hacer click en un lugar vacío de la página (por ejemplo el body)
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element(body).click().perform()
        print("🖱️ Clic en un lugar vacío de la página para asegurar la interacción.")

        # Seleccionar la opción "20" en el dropdown de cantidad de registros
        select_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select.select2.form-control.input-sm")))
        Select(select_elem).select_by_value("20")
        print("✅ Se seleccionó '20' registros para mostrar más filas.")
        time.sleep(5)

        # Esperar a que el loader desaparezca
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered")))
            rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered tbody tr")))
            print(f"✅ ¡Tabla cargada con {len(rows)} filas!")

        except TimeoutException:
            driver.save_screenshot(os.path.join(ruta_carpeta_errores,f"{numero_poliza_birlik}_{get_timestamp()}.png"))
            raise Exception("❌ La tabla no tiene filas")

        fila_encontrada = False

        print("-------------------------------------------------")
        num_poliza_mas_reno = f"{numero_poliza_birlik}-R{numero_proforma_birlik}"
        print(f"🔍 Buscando el código de cuota '{numero_proforma_birlik}' o el numero de poliza '{num_poliza_mas_reno}'.")

        for row in rows:

            cells = row.find_elements(By.TAG_NAME, "td")

            if len(cells) < 15:
                continue

            cod_cotizacion = cells[1].text.strip()                                  # Código Cotización
            prima_total = cells[4].text.replace("S/.", "").replace(",", "").strip() # Monto Total
            estado_compania = cells[5].text.strip()                                 # "Pendiente" "Aprobado" "Rechazado"
            estado_emision = cells[11].text.strip()                                 # "Pendiente" "En Proceso" "Realizado" "Anulado"
            numero_poliza_mas_reno = cells[9].text.strip()                          # Numero Poliza 
            fecha_emision = cells[12].text.strip()                                  # Fecha emision PDF
            estado_pago = cells [16].text.strip()                                   # Estado de pago: "Pagado" 
            comprobante = cells[17].text.strip()                                    # Numero de Comprobante
            fecha_emision_comprobante =cells[18].text.strip()                       # Fecha de Emision Comprobante

            diferencia = abs(float(prima_total) - float(importe_total_birlik))

            if cod_cotizacion == numero_proforma_birlik or num_poliza_mas_reno == numero_poliza_mas_reno or diferencia > 0.05:

                fila_encontrada = True
                print("✅ Fila encontrada")

                resultado_estado = estado_pago

                print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(prima_total)}")
                if diferencia > 0.05:
                    print("❌ Los importes No coinciden")
                else:
                    print("✅ Los importes Coinciden")
                    resultado_importe = True

                if estado_pago.lower() == "pagado":

                    if len(cells) > 18:

                        comprobante_pdf = cells[17]

                        try:
                            # Hacer clic en el link de esa celda
                            link = comprobante_pdf.find_element(By.TAG_NAME, "a")

                            ventana_principal_crecer = driver.current_window_handle

                            # Guardar archivos antes del clic
                            archivos_antes = set(os.listdir(ruta_carpeta_facturas))

                            driver.execute_script("arguments[0].click();", link)
                            print("✅ Se hizo clic con JS en el botón de descarga")

                            try:

                                mensaje_generando = WebDriverWait(driver,7).until(EC.presence_of_element_located((By.ID, "swal2-content")))
                                texto = mensaje_generando.text.strip()

                                if "Estamos generando tu comprobante" in texto:

                                    # Esperar botón OK y hacer clic
                                    boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'swal2-confirm')]")))
                                    boton_ok.click()
                                    print("🖱️ Clic en Ok")

                                    raise Exception(f"Mensaje detectado -> '{texto}'")

                            except TimeoutException:
                                pass

                            archivo_nuevo = esperar_archivos_nuevos(ruta_carpeta_facturas,archivos_antes,".pdf",cantidad=1)

                            if archivo_nuevo:
                                print(f"✅ Archivo .pdf descargado exitosamente")
                                ruta_original = archivo_nuevo[0]
                                ruta_final = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{comprobante}.pdf")
                                os.rename(ruta_original, ruta_final)
                                print(f"🔄 Archivo renombrado a '{numero_poliza_birlik}_{comprobante}.pdf'")
                            else:
                                raise Exception("No se encontró archivo nuevo después de descargar")

                            # Esperar a que el loader desaparezca
                            wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

                            if os.path.exists(ruta_final):
                                #------------INGRESA A SUNAT-------  
                                nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                                ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                resultado = consultarValidezSunat(driver,wait,ruc_crecer_vly,tipo_doc_birlik,ruc_cliente_birlik,comprobante,fecha_emision_comprobante,prima_total,ruta_imagen_sunat)

                                driver.switch_to.window(ventana_principal_crecer)
                                print("🔄 Volviendo a la ventana de la CIA")

                                if resultado is None:
                                        resultado_accion = f'=HYPERLINK("{login_sunat}", "Sunat Bloqueado")'
                                        break
                                elif resultado:

                                    resultado_sunat = True

                                    if estado_Cuota_birlik == "Pendiente-comprobante":

                                        print("📤 Subiendo comprobante a Birlik...")
                                        agregar_comprobante_pago(driver,wait,id_cuota_birlik,ruta_final)
                                        resultado_accion = "Factura Enviada Anteriormente"

                                    else:

                                        print("📤 Subiendo todos los documentos a Birlik...")
                                        cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,comprobante,fecha_emision,ruta_final,ruta_imagen_sunat,resultado_importe)
                                        resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
            
                                    resultado_birlik = True
                                    break
                                else:
                                    resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                                                          
                            else:
                                raise Exception(f"No se descargo la Factuta '{comprobante}'.")

                        except Exception as e:
                            print(f"No se pudo descargar el PDF,Flujo Terminado con estado {estado_pago}, error: {e}")

                else:
                    print(f"La Poliza {numero_poliza_birlik} tiene estado de Pago '{estado_pago}',estado de emision '{estado_emision}', y estado de compañia '{estado_compania}'.")
                    resultado_accion  = "Esperar a que pague"
                    break

        if not fila_encontrada:
            resultado_accion = "°PROFORMA INCORRECTO"
            resultado_estado  = "No se sabe"
            print(f"❌ No se encontró ninguna fila que coincida con el código de cuota {numero_proforma_birlik} o {numero_poliza_birlik}-R{numero_proforma_birlik}.")
        
    except Exception as e:
        print(f"Detalles del error: {e}")
    finally:
        driver.refresh()
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Registrado" if resultado_birlik else "No registrado" , resultado_estado, resultado_accion
    
def main():
    
    while True:

        ruta_salida_API,ruta_salida,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

        try:
        
            display_num = os.getenv("DISPLAY_NUM", "0")
            os.environ["DISPLAY"] = f":{display_num}"

            driver,wait = abrirDriver(ruta_carpeta_facturas)

            driver.get(login_url_crecer_vida_ley)
            print("⌛ Ingresando a la URL")

            user_input = wait.until(EC.presence_of_element_located((By.ID, "suser")))
            user_input.clear()
            user_input.send_keys(username_crecer)
            print("⌨️ Digitando el Username")
        
            pass_input = wait.until(EC.presence_of_element_located((By.ID, "spassword")))
            pass_input.clear()
            pass_input.send_keys(password_crecer)
            print("⌨️ Digitando el Password")
   
            print("🧩 Resuelve el CAPTCHA manualmente y clic en 'Ingresar'.")

            desbloquear_interaccion()

            puerto = os.getenv("NOVNC_PORT")
            enviarCaptcha(para_lista,copias_lista,puerto,"Crecer Vida Ley",ruta_imagen=None)

            # Espera humana (hasta 5 minutos)
            wait_humano = WebDriverWait(driver,300)
            wait_humano.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]")))

            bloquear_interaccion()

            print("✅ Login exitoso detectado (Cerrar sesión visible)")
            print("🚀 Continuando flujo automáticamente")

            # Esperamos que el SVG esté cargado (opcional pero recomendado)
            span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
            actions = ActionChains(driver)
            actions.move_to_element(span_element).click().perform()
            print("🖱️ Clic en 'Gestion de Cotizacion'")
            time.sleep(3)

            # Esperar que el enlace esté disponible
            link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Bandeja de Evaluación")))
            link.click()
            print("🖱️ Clic en 'Bandeja de Evaluación'")

            time.sleep(3)

            while True:

                json_cuotas = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

                try:

                    if not json_cuotas:
                        raise Exception("No hay cuotas pendientes para esta compañia")

                    print("\n📁 Iniciando procesamiento para Crecer Vida Ley...")

                    # Guardar data del Json en un Excel para procesar fila por fila
                    guardarJson(json_cuotas,ruta_salida_API)

                    try:
                        df = pd.read_excel(ruta_salida_API, engine="openpyxl",dtype={"numeroDocumento": str})
                    except Exception as e:
                        raise Exception(f" Error al leer el archivo Excel: {e}")
            
                    # Nuevas columnas para registrar los resultados
                    df["Importe"] = ""
                    df["Sunat"] = ""
                    df["Birlik"] = ""
                    df["Estado"] = ""
                    df["Acción"] = ""

                    total_filas = len(df)

                    # 2. Iterar sobre cada fila
                    for index, row in df.iterrows():
                        print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                        try:
                            # Procesar la fila usando tu lógica
                            importe_estado, sunat_estado, birlik_estado,estado_estado,accion_estado = procesar_fila(
                                driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores)

                            df.at[index, "Importe"] = importe_estado
                            df.at[index, "Sunat"] = sunat_estado
                            df.at[index, "Birlik"] = birlik_estado
                            df.at[index, "Estado"] = estado_estado
                            df.at[index, "Acción"] = accion_estado

                            print(f"✅ Fila {index} guardada correctamente")
                        except Exception as e:
                            print(f"❌ Error procesando fila {index}: {e}")
                        finally:
                            df.to_excel(ruta_salida, index=False)
                            time.sleep(1)
     
                        time.sleep(3)
    
                except Exception as e:
                    print(f"\n🟡 Proceso Detenido, Motivo: {e}")
                finally:   
                    
                    if json_cuotas:
                        os.remove(ruta_salida_API)
                        print(f"\n✅ Flujo finalizado, Intentando de nuevo en 10 segundos.")
                        time.sleep(10)

        except Exception as e:
            print(f"\n🟡 Proceso Detenido por fuerza Mayor, Motivo: {e}")
        finally:
            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                print("🧹 Carpeta eliminada correctamente")
            else:
                print("⚠️ La carpeta no existe")
            print(f"\n✅ Flujo finalizado, Intentando de nuevo en 10 segundos.")
            time.sleep(10)

if __name__ == "__main__":
    main()