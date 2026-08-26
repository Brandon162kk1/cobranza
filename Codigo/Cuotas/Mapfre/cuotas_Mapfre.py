#-- Imports ---
import time
import os
import pandas as pd
import shutil
#-- Froms ----
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import ActionChains
from datetime import datetime
from Codigo.Sunat.validar_factura import consultarValidezSunat,url_sunat
from Codigo.Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota
from Codigo.Birlik.urls import url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Codigo.Apis.Birlik.metodo import consultarAPI
from Codigo.GoogleChrome.chromeDriver import abrirDriver, crearCarpetas,guardarJson,esperar_archivos_nuevos, tomar_captura
from Codigo.GoogleChrome.fecha_y_hora import get_timestamp
from Codigo.Apis.Compania.get import codigo_compania
from Codigo.Correo.armar_asunto import enviarAviso

#--------- Datos------
ruc_mapfre_salud = '20517182673' # Salud
ruc_mapfre_pension = '20418896915' #Seguros y Re aseguros
ids_compania = [16,17,18]
#----- Variables de Entorno -------
login_url_mapfre = os.getenv("url_mapfre")
username = os.getenv("usernameMapfre")
password = os.getenv("passwordMapfre")
url_api_cod_cot = os.getenv("url_api_cod_map")
API_KEY = os.getenv("API_KEY_MAPFRE")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Mapfre_{get_timestamp()}"

def limpiar(valor):
    if valor is None:
        return ""
    valor = valor.strip()
    return valor if valor else ""

def procesar_fila(driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,index):

    #--Extraer valores y quitar espacios en blanco
    numero_poliza_birlik = str(row["numeroPoliza"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    ruc_cliente_birlik = str(row["numeroDocumento"]).strip()
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    fk_Cliente_birlik = str(row["fk_Cliente"]).strip() 
    fk_compania_birlik = str(row["fK_Compania"]).strip()
    fk_Ramo_birlik = int(str(row["fk_Ramo"]).strip())
    numero_proforma_birlik = str(row["codigoCuota"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()
    primaneta_birlik = str(row["primaNeta"]).strip()
    id_Poliza_birlik = str(row["id_Poliza"]).strip()
    fecha_inicioVig_Birlik = str(row["vigenciaInicio"]).strip()
    fecha_finVig_Birlik = str(row["vigenciaFin"]).strip()
    #--------------

    if fk_compania_birlik == '17':
        ruc_compania = ruc_mapfre_pension
    else:
        ruc_compania = ruc_mapfre_salud

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    # if id_cuota_birlik not in ("10802"):
    #      return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    try:
        # ------------------ Inicio del Flujo de Automatización ------------------
        label = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-label[normalize-space()='Nro. Póliza *']")))
        driver.execute_script("arguments[0].click();", label)
        #time.sleep(2)
        print("🖱️ Clic en el # de Póliza")

        #poliza_input = wait.until(EC.element_to_be_clickable((By.ID, "mat-input-5")))

        poliza_input = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//mat-form-field[.//mat-label[contains(.,'Nro. Póliza')]]//input"
            ))
        )

        poliza_input.clear()
        poliza_input.send_keys(numero_poliza_birlik)
        print(f"⌨️ Digitando la póliza '{numero_poliza_birlik}'")

        time.sleep(3)

        buscar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Buscar']")))
        buscar_btn.click()
        print("🖱️ Clic en 'Buscar'")

        try:
            print("⌛ Esperando que cargue la tabla")
            wait.until(
                EC.presence_of_element_located((
                    By.XPATH, "(//ul[contains(@class,'g-tbl-row')])[2]"
                ))
            )
            print("✅ La tabla cargó correctamente")
        except:
            raise Exception("No cargó la tabla")

        encontrado = False

        while True:

            filas = driver.find_elements(By.XPATH, "//ul[contains(@class,'g-tbl-row')]")
            filas = filas[1:]

            for fila in filas:

                columnas = fila.find_elements(By.TAG_NAME, "li")

                # Asegurar que la fila tenga suficientes columnas
                if len(columnas) < 4:
                    continue

                # CodigoCuota
                codigo_fila = limpiar(columnas[6].text)
                importe_fila = limpiar(columnas[7].text)
                fecha_emision_fila = limpiar(columnas[9].text)
                numero_factura = limpiar(columnas[11].text)
                fecha_pago_fila = limpiar(columnas[14].text)

                if codigo_fila == numero_proforma_birlik:

                    encontrado = True
                    print(f"✅ Fila encontrada ")
                    placeholder = "-"
                    print(
                        f"⌛ Código Cuota: {codigo_fila or placeholder}, "
                        f"Importe: {importe_fila or placeholder}, "
                        f"Fecha Emisión: {fecha_emision_fila or placeholder}, "
                        f"Numero Factura: {numero_factura or placeholder}, "
                        f"Fecha de Pago: {fecha_pago_fila or placeholder}"
                    )
                    diferencia = abs(float(importe_fila) - float(importe_total_birlik))
                    #print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe_fila)}")
                    if diferencia > 0.05:
                        print("❌ Los importes No coinciden")
                    else:
                        print("✅ Los importes Coinciden")
                        resultado_importe = True
                    
                    if numero_factura == '-' or fecha_emision_fila == '-' or fecha_pago_fila == '-':
                        resultado_estado = "No dice estado"
                        resultado_accion = "Esperar"
                        break

                    checkbox = fila.find_element(By.XPATH, ".//input[@type='checkbox']")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                    driver.execute_script("arguments[0].click();", checkbox)
                    print("🖱️ Clic en el 'Checkbox'")

                    time.sleep(2)

                    ventana_principal_mapfre = driver.current_window_handle

                    archivos_antes = set(os.listdir(ruta_carpeta_facturas))

                    driver.execute_script("arguments[0].click();", columnas[11])
                    print("🖱️ Clic con JS en el botón de descarga")

                    archivo_nuevo = esperar_archivos_nuevos(ruta_carpeta_facturas,archivos_antes,".pdf",cantidad=1)

                    if archivo_nuevo:
                        print(f"✅ Factura descargada exitosamente")
                        ruta_original = archivo_nuevo[0]
                        ruta_final = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{numero_factura}.pdf")
                        os.rename(ruta_original, ruta_final)
                        print(f"🔄 Archivo renombrado a '{numero_poliza_birlik}_{numero_factura}.pdf'")
                    else:
                        raise Exception("No se encontró archivo nuevo después de descargar")
                
                    fecha_habiles_factura = []

                    fecha_emision_probar = datetime.strptime(fecha_pago_fila, "%d/%m/%Y")

                    fecha_habiles_factura.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                    print(f"📅 Fechas de Emisión a probar:{fecha_habiles_factura}")

                    for fecha in fecha_habiles_factura:
                        print("---------------------------------------")

                        #------------INGRESA A SUNAT-------  
                        nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                        ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                        resultado = consultarValidezSunat(driver,wait,ruc_compania,tipo_doc_birlik,ruc_cliente_birlik,numero_factura,fecha,importe_fila,ruta_imagen_sunat,ruta_carpeta_errores)

                        driver.switch_to.window(ventana_principal_mapfre)
                        print("🔄 Volviendo a la ventana de la CIA")

                        if resultado is None:
                            resultado_accion = f'=HYPERLINK("{url_sunat}", "Sunat Bloqueado")'
                            break
                        elif resultado:

                            resultado_sunat = True

                            if estadoCuota_birlik == "Pendiente-comprobante":
                                print("📤 Subiendo comprobante a Birlik...")
                                agregar_comprobante_pago(driver,wait,id_cuota_birlik,ruta_final)
                                resultado_accion = "Factura Enviada Anteriormente"
                            else:
                                print("📤 Subiendo todos los documentos a Birlik...")
                                cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,numero_factura,fecha,ruta_final,ruta_imagen_sunat,resultado_importe)
                                resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
            
                            resultado_birlik = True
                            break

                        else:
                            resultado_accion = f'=HYPERLINK("{url_sunat}", "Ver Sunat")'
                            continue

                    break              

            if encontrado:
                break

            if not encontrado:

                try:
                    #btn_siguiente_list = driver.find_elements(By.XPATH, "//button[@aria-label='Siguiente página']")
                    btn_siguiente_list = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Siguiente página']")))

                    if not btn_siguiente_list:
                        raise Exception(f"No existe el botón 'Siguiente página' porque no hay resultados para la póliza '{numero_poliza_birlik}'")

                    #btn_siguiente = btn_siguiente_list[0]

                    if btn_siguiente_list.get_attribute("disabled"):
                        raise Exception("No hay más páginas")

                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",btn_siguiente_list)
                    time.sleep(2)
                    btn_siguiente_list.click()

                    print("🔄 Cambiando a la siguiente página...")
                    time.sleep(2)

                except Exception as e:
                    print(f"⛔ {e}")
                    break

        if not encontrado:
            resultado_estado = "No se encuentro cuota"
            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}", "Revisar Cuota")'
            raise Exception(f"No se encontró ninguna fila con el código {numero_proforma_birlik}")

    except Exception as e:
        # import traceback
        # print(type(e).__name__)
        # print(e)
        # traceback.print_exc()
        print(f"❌ Error Procesando toda la fila, Motivo: {e}")
        tomar_captura(driver,ruta_carpeta_errores,f"Error_fila_{index+2}")

    finally:
        limpiar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Limpiar']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", limpiar_btn)
        #time.sleep(0.3)
        wait.until(lambda d: limpiar_btn.location['y'] > 0)
        limpiar_btn.click()
        print("🖱️ Clic en 'Limpiar'")
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" , "No indica" if resultado_estado is None else resultado_estado ,resultado_accion

def main():
    
    while True:

        ruta_salida_API,ruta_salida,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errore,carpeta_compañia,carpeta_principal = crearCarpetas(nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

        try:
        
            display_num = os.getenv("DISPLAY_NUM", "0")
            os.environ["DISPLAY"] = f":{display_num}"

            driver,wait = abrirDriver(ruta_carpeta_facturas)

            driver.get(login_url_mapfre) 
            print("⌛ Ingresando a la URL")

            user_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-1")))
            user_input.clear()
            user_input.send_keys(username)
            print(f"⌨️ Digitando el Username {username}")
        
            pass_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-0")))
            pass_input.clear()
            pass_input.send_keys(password)
            print(f"⌨️ Digitando el Password {password}")

            ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
            ingresar_btn.click()
            print("🖱️ Clic en 'Ingresar'")

            elemento = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".card-modality__item--left")))
            elemento.click()
            print("🖱️ Clic en enviar por Correo Electronico")

            codigo = codigo_compania(url_api_cod_cot,API_KEY)

            inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.g-input-codes__code")))

            if len(inputs) == len(codigo):
                for i, inp in enumerate(inputs):
                    inp.clear()
                    inp.send_keys(codigo[i])
            else:
                raise Exception("Los inputs no coinciden con la longitud del código")

            time.sleep(1)

            comprobar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Comprobar')]")))
            comprobar_btn.click()
            print("🖱️ Clic en 'Comprobar'")

            try:

                modal_mensaje = (By.CSS_SELECTOR, "div.c-modal p.txt")
                boton_ok_locator = (By.XPATH, "//button[.//span[contains(text(), 'Ok')]]")

                resultado = wait.until(
                    EC.any_of(
                        EC.visibility_of_element_located(modal_mensaje),
                        EC.element_to_be_clickable(boton_ok_locator)
                    )
                )

                # --- Caso 1: apareció modal con mensaje ---
                if resultado.tag_name.lower() == "p":

                    mensaje = resultado.text.strip()
                    print(f"⚠️ Modal detectado: {mensaje}")

                    # if not enviarAviso("Mapfre"):
                    #     raise Exception("No se pudo enviar el correo")

                    boton_cerrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'Cerrar')]]")))
                    driver.execute_script("arguments[0].click();", boton_cerrar)
                    print("✅ Modal cerrado correctamente")

                # --- Caso 2: apareció botón Ok ---
                else:
                    resultado.click()
                    print("🖱️ Clic en 'Ok'")

            except TimeoutException:
                pass

            # Esperar que no exista ningún diálogo visible
            wait.until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "mat-dialog-container")
                )
            )

            # Esperar que desaparezca el backdrop de Angular Material (si existe)
            wait.until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, ".cdk-overlay-backdrop")
                )
            )
            # consulta_gestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='CONSULTAS DE GESTION']")))
            # consulta_gestion.click()

            consulta_gestion = wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//div[contains(@class,'c-box-module')][.//span[normalize-space()='CONSULTAS DE GESTION']]"
                    )
                )
            )

            driver.execute_script("arguments[0].click();", consulta_gestion)

            print("🖱️ Clic en 'CONSULTAS DE GESTION'")
          
            action = ActionChains(driver)   
            cobranzas_link = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[normalize-space()='COBRANZAS']")))
            action.move_to_element(cobranzas_link).perform()
            print("🖱️ Mouse sobre 'COBRANZAS'")
  
            cronograma_pagos = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='CRONOGRAMA DE PAGOS']")))
            action.move_to_element(cronograma_pagos).click().perform()
            print("🖱️ Clic en 'CRONOGRAMA DE PAGOS'")

            while True:

                json_cuotas = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

                try:

                    if not json_cuotas:
                        raise Exception("No hay cuotas pendientes para esta compañia")

                    print("\n📁 Iniciando procesamiento para Mapfre")

                    # Guardar data del Json en un Excel para procesar fila por fila
                    guardarJson(json_cuotas,ruta_salida_API)

                    try:
                        df = pd.read_excel(ruta_salida_API, engine="openpyxl",dtype={"numeroDocumento": str})
                    except Exception as e:
                        raise Exception(f" Error al leer el archivo Excel | Motivo: {e}")

                    df["Importe"] = ""
                    df["Sunat"] = ""
                    df["Birlik"] = ""
                    df["Estado"] = ""
                    df["Acción"] = ""

                    total_filas = len(df)

                    for index, row in df.iterrows():
                        print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                        try:

                            importe_estado,sunat_estado,birlik_estado,estado_estado,accion_estado = procesar_fila(
                                driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errore,index)

                            df.at[index, "Importe"] = importe_estado
                            df.at[index, "Sunat"] = sunat_estado
                            df.at[index, "Birlik"] = birlik_estado
                            df.at[index, "Estado"] = estado_estado
                            df.at[index, "Acción"] = accion_estado

                            print(f"\n✅ Fila {index} guardada correctamente")
                        except Exception as e:
                            print(f"❌ Error en fila {index}, Motivo: {e}")
                        finally:
                            driver.refresh()
                            df.to_excel(ruta_salida, index=False)

                except Exception as e:
                    print(f"❌ Proceso Detenido, por : {e}")
                finally:
                    if json_cuotas:
                        os.remove(ruta_salida_API)
                        print(f"\n✅ Flujo finalizado, Intentando de nuevo en 10 segundos")
                        time.sleep(10)

        except Exception as e:
            print(f"❌ Proceso Detenido por fuerza Mayor, Motivo: {e}")

            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                #print("🧹 Carpeta eliminada correctamente")

            print("⌛ Esperando 30 minutos para intentar reiniciar el proceso\n")
            time.sleep(1800)
            
if __name__ == "__main__":
    main()
