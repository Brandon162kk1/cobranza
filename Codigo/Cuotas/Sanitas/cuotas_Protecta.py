#-- Imports ---
import re
import time
import os
import pandas as pd
import shutil
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.expected_conditions import invisibility_of_element_located
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from Codigo.Cuotas.Sanitas.cuotas_Crecer import buscaryRegistrarenCrecer,obtener_fecha_emision
from Codigo.Sunat.validar_factura import consultarValidezSunat,url_sunat
from Codigo.Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota
from Codigo.Birlik.urls import url_cuotas,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Codigo.Apis.Birlik.metodo import consultarAPI
from Codigo.GoogleChrome.chromeDriver import abrirDriver, crearCarpetas,esperar_archivos_nuevos,guardarJson
from Codigo.GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_hoy

#----- Datos -------
ids_compania = [5,29,31]
ruc_sanitas = '20523470761'
ruc_sanitas_protecta = '20517207331'
ruc_sanitas_crecer = '20600098633'
#----- Variables de Entorno -------
username_sanitas = os.getenv("usernameSanitas")
password_sanitas = os.getenv("passwordSanitas")
login_url_sanitas_crecer = os.getenv("login_url_sanitas_crecer")
login_url_sanitas_protecta = os.getenv("login_url_sanitas_protecta")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Sanitas_SCTR_{get_timestamp()}"

def procesar_fila(row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores):

    numero_poliza = str(row["numeroPoliza"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    numero_ruc = str(row["numeroDocumento"]).strip()
    id_cuota = str(row["id_Cuota"]).strip()
    fk_Cliente = str(row["fk_Cliente"]).strip() 
    fk_compania = str(row["fK_Compania"]).strip()
    numero_proforma = str(row["codigoCuota"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_ocr = False
    resultado_estado = None
    resultado_accion = ""

    # if id_cuota not in ("7833"):
    #     return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_ocr else "Pagina Web en Mantenimiento", "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    if fk_compania == '29':
        nombre_Compania = "Sanitas"
        url_general = login_url_sanitas_protecta
        ruc_compania = ruc_sanitas
    elif fk_compania == '31':
        nombre_Compania = "Sanitas Protecta"
        url_general = login_url_sanitas_protecta
        ruc_compania = ruc_sanitas_protecta
    else:
        nombre_Compania = "Sanitas Crecer"
        url_general = login_url_sanitas_crecer
        ruc_compania = ruc_sanitas_crecer

    print(f"Compañía {nombre_Compania} - RUC: {ruc_compania}")

    display_num = os.getenv("DISPLAY_NUM", "0")  # fallback = 0
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver(ruta_carpeta_facturas)
    
    try:
        driver.get(url_general) 
        print("✅ Ingresando a la URL")
        
        user_input = wait.until(EC.presence_of_element_located((By.ID, "Login")))
        user_input.clear()
        user_input.send_keys(username_sanitas)
        print("✅ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "Password")))
        pass_input.clear()
        pass_input.send_keys(password_sanitas)
        print("✅ Digitando el Password")
        
        ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
        ingresar_btn.click()
        print("🖱️ Clic en 'Ingresar'")
        
        wait.until(EC.url_contains("Quotation/Index"))
        
        autogestion_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(text(),'Autogestión')]")))
        autogestion_link.click()
        print("🖱️ Clic en 'Autogestión'")
        
        estado_cuenta_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Estado de cuenta SCTR')]")))
        estado_cuenta_link.click()
        print("🖱️ Clic en 'Estado de cuenta SCTR'")
        
        # # 6. Seleccionar en el dropdown el tipo de documento: RUC (value="2")
        # select_identity = wait.until(EC.presence_of_element_located((By.ID, "IdentityTypeId")))
        # select_obj = Select(select_identity)
        # select_obj.select_by_value("2")
        # print("Se seleccionó 'RUC' en el dropdown.")

        ruc_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' and contains(@class, 'AccountStatusList')]")))
        ruc_input.clear()
        ruc_input.send_keys(numero_ruc)
        print(f"✅ Se ingresó el RUC: {numero_ruc}")
        
        status_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'btn-group') and contains(@class,'AccountStatusList')]//button")))
        status_btn.click()
        print("🖱️ Clic en el dropdown de estado")

        abonada_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[.//span[contains(text(),'Abonada')]]")))
        abonada_option.click()
        print("🖱️ Se seleccionó la opción 'Abonada'")
    
        anulada_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[.//span[contains(text(),'Anulada')]]")))
        anulada_option.click()
        print("🖱️ Se seleccionó la opción 'Anulada'")

        migracion_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[.//span[contains(text(),'Migración')]]")))
        migracion_option.click()
        print("🖱️ Se seleccionó la opción 'Migración'")
        
        driver.find_element(By.TAG_NAME, "body").click()
        print("🖱️ Clic en body")
        
        filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
        filter_btn.click()
        print("🖱️ Clic en 'Filtrar'")

        wait.until(invisibility_of_element_located((By.ID, "AccountStatusListTable_processing")))
        print("⌛ Cargando")

        select_elem = wait.until(EC.presence_of_element_located((By.NAME, "AccountStatusListTable_length")))
        Select(select_elem).select_by_value("1000")
        print("🖱️ Se seleccionó '1000' registros para mostrar más filas")

        # Esperar que desaparezca el mensaje "Un momento por favor..."
        wait.until(invisibility_of_element_located((By.ID, "AccountStatusListTable_processing")))

        try:
            wait.until(
                lambda d: len(
                    d.find_elements(By.XPATH, "//table[@id='AccountStatusListTable']//tr")
                ) > 1
            )
            print("✅ La tabla tiene al menos 1 fila.")
        except TimeoutException:
            raise Exception("No se cargaron las filas en el tiempo esperado")

        table = driver.find_element(By.ID, "AccountStatusListTable")
        rows = table.find_elements(By.TAG_NAME, "tr")
        
        fila_encontrada = False
        comprobante_valor = None 
        estado_valor = None         
        documento_valor = None      
        fecha_emision_valor = None
        importe_valor = 0
   
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) == 11:                        
                fecha_emision = cells[3].text.strip()   
                documento = cells[4].text.strip()       
                comprobante = cells[6].text.strip()     
                estado = cells[7].text.strip()
                importe = cells[8].text.strip()
                
                solo_numeros_doc = re.sub(r'\D', '', documento)

                if numero_proforma in solo_numeros_doc:
                    fila_encontrada = True
                    estado_valor = estado
                    resultado_estado = estado_valor     #Le asignamos el valor del estado al Campo Estado
                    comprobante_valor = comprobante
                    documento_valor = documento
                    fecha_emision_valor = fecha_emision
                    importe_valor = importe

                    print(f"✅ Fila encontrada: Documento='{documento}', Proforma= '{numero_proforma}', Estado='{estado}', Factura='{comprobante_valor}', Importe ='{importe_valor}', Fecha Emisión ='{fecha_emision_valor}'")
                    
                    if estado.lower() == "abonada":
                        #print(f"✅ La cuota con proforma {numero_proforma} está abonada. Comprobante: {comprobante_valor}")
                        pass
                    elif estado.lower() == 'anulada':
                        #print(f"❌ La cuota con proforma {numero_proforma} está anulada. Comprobante: {comprobante_valor}")
                        resultado_accion = f'=HYPERLINK("{url_cuotas}{fk_Cliente}", "Anular Cuota")'
                    else:
                        #print(f"⚠️ La cuota con proforma {numero_proforma} está pendiente (Estado='{estado}')")
                        resultado_accion = f'Sin Observación'
                    
                    break
                else:
                    resultado_importe = "Codigo Cuota Incorrecto"
        
        #print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe_valor)}")

        # Validación de importes con tolerancia
        diferencia = abs(float(importe_valor) - float(importe_total_birlik))

        if diferencia > 0.05:
            print("❌ Los importes No coinciden")
        else:
            resultado_importe = True
            print("✅ Los importes Coinciden")

        if not fila_encontrada:
            print(f"❌ No se encontró ninguna fila con Documento conteniendo '{numero_proforma}'")
        
        # 13. Si la cuota está abonada, navegar a "Consulta de Comprobantes de pago"
        if fila_encontrada and estado_valor.lower() == "abonada":
            #print(f"✅ La cuota con proforma {numero_proforma} está abonada. Comprobante: {comprobante_valor} | Documento: {documento_valor} | Fecha Emisión: {fecha_emision_valor} ")

            autogestion_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(text(),'Autogestión')]")))
            autogestion_link.click()
            time.sleep(2)
            
            consulta_comprobantes_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/PaymentVouchers/LegalDocumentsSelfManagementIndex' and contains(text(),'Comprobantes')]")))
            consulta_comprobantes_link.click()
            print("🌐 Navegando a 'Consulta de Comprobantes de pago'")
            time.sleep(3)

            contract_input = wait.until(EC.presence_of_element_located((By.ID, "ContractNumber")))
            contract_input.clear() 
            contract_input.send_keys(numero_poliza)
            print(f"✅ Se ingresó el número de póliza: {numero_poliza}")

            fecha = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")
            fecha_inicio = fecha - relativedelta(months=1)
            fecha_fin = fecha + relativedelta(months=1)

            # Formatear el rango de fechas en el formato requerido: "dd/mm/yyyy - dd/mm/yyyy"
            rango_fechas = f"{fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}"
            print(f"📅 Rango de fechas: {rango_fechas}")

            input_rango = wait.until(EC.presence_of_element_located((By.ID, "DatesRangeCreation")))
            input_rango.clear()
            input_rango.send_keys(rango_fechas)

            filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
            filter_btn.click()

            select_elem = wait.until(EC.presence_of_element_located((By.NAME, "LegalDocumentSelfManagementListTable_length")))
            Select(select_elem).select_by_value("1000")
            print("🖱️ Clic en '1000' registros para ver más filas")
            time.sleep(2)
            tabla = wait.until(EC.presence_of_element_located((By.ID, "LegalDocumentSelfManagementListTable")))
            filas_tabla = tabla.find_elements(By.TAG_NAME, "tr")
    
            fila_encontrada_descarga = False
            for fila in filas_tabla:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                # Supongamos que la columna "Documento" es la 3ª (índice 2) en esta tabla.
                # Ajusta según el orden real en tu tabla.
                if len(celdas) >= 4:
                    doc_valor = celdas[2].text.strip()  # Ajusta índice si es necesario
                    if comprobante_valor in doc_valor:
                        fila_encontrada_descarga = True
                        print(f"✅ Fila encontrada en 'Consulta de Comprobantes', Documento = {doc_valor}")

                        celda_accion = celdas[-1]
                        
                        try:

                            icono = celda_accion.find_element(By.XPATH, ".//a[contains(@class, 'dropdown-toggle')]")
                            driver.execute_script("arguments[0].scrollIntoView(true);", icono)
                            ActionChains(driver).move_to_element(icono).click().perform()
                            print("🖱️ Clic en el icono")
                            time.sleep(2) 

                            menu = celda_accion.find_element(By.XPATH, ".//ul[contains(@class, 'dropdown-menu')]")
                            driver.execute_script("arguments[0].style.display = 'block';", menu)
                            print("✅ Menú desplegable forzado a visible")
                            time.sleep(1)

                            #---Haciendo dinamico el encontrar el ID
                            # Encuentra TODOS los enlaces <a> en el menú desplegable
                            links = menu.find_elements(By.TAG_NAME, "a")
                            boton_descarga = None

                            for link in links:
                                texto = link.text.strip()
                                titulo = link.get_attribute("title")
                                #id_link = link.get_attribute("id")
                                #print(f"Probando link: texto='{texto}', title='{titulo}', id='{id_link}'")
                                # Busca por texto, título o clase
                                if "Descarga" in texto or "Descarga" in (titulo or ""):
                                    boton_descarga = link
                                    break

                            # Si aún no lo encuentras, podrías tomar el primer <a>
                            if boton_descarga is None and links:
                                print("❌ No se encontró 'Descarga' en texto/título, se usará el primer <a> del menú.")
                                boton_descarga = links[0]

                            # ------------------------------- fin 
                            wait.until(EC.visibility_of(boton_descarga))

                            # Siempre guarda tu ventana principal (X)
                            ventana_cia = driver.current_window_handle

                            #------------------------------------------------------
                            # Guardar archivos antes del clic
                            archivos_antes = set(os.listdir(ruta_carpeta_facturas))

                            driver.execute_script("arguments[0].click();", boton_descarga)
                            print("🖱️ Clic con JS en el botón de descarga")

                            archivo_nuevo = esperar_archivos_nuevos(ruta_carpeta_facturas,archivos_antes,".pdf",cantidad=1)

                            if archivo_nuevo:
                                print(f"✅ Factura descargado exitosamente")
                                ruta_original = archivo_nuevo[0]
                                ruta_final = os.path.join(ruta_carpeta_facturas, f"{numero_poliza}_{comprobante_valor}.pdf")
                                os.rename(ruta_original, ruta_final)
                                print(f"🔄 Archivo renombrado a '{numero_poliza}_{comprobante_valor}.pdf'")
                            else:
                                raise Exception("No se encontró archivo nuevo después de descargar")
                            #---------------------------------------------------------
              
                            time.sleep(3)

                            fechas_habiles = []

                            #--------
                            fecha_emision_extraida = obtener_fecha_emision(ruta_final)
                            #--------

                            fecha_emision_probar = datetime.strptime(fecha_emision_extraida, "%d/%m/%Y")
                            fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                            # Hasta tener 15 fechas consecutivas (incluye sábados y domingos)
                            while len(fechas_habiles) < 15:
                                fecha_emision_probar += timedelta(days=1)

                                # Si la siguiente fecha es mayor que hoy, se detiene
                                if fecha_emision_probar.date() >= get_fecha_hoy().date():
                                    break

                                fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                            print(f"📅 Fechas de Emisión a probar: {fechas_habiles}")

                            for fecha in fechas_habiles:
                                print("---------------------------------------")
                                #print(f"⌛ Probando con la Fecha hábil: {fecha}")

                                nombre_imagen_sunat = f"{numero_proforma}_{numero_poliza}.png"
                                ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                resultado = consultarValidezSunat(driver,wait,ruc_compania,tipo_doc_birlik,numero_ruc,comprobante_valor,fecha,importe_valor,ruta_imagen_sunat,ruta_carpeta_errores)

                                driver.switch_to.window(ventana_cia)
                                print("🔄 Volviendo a la ventana de la CIA")

                                if resultado is None:
                                    resultado_accion = f'=HYPERLINK("{url_sunat}", "Sunat Bloqueado")'
                                    break
                                elif resultado:

                                    resultado_sunat = True

                                    if estadoCuota_birlik == "Pendiente-comprobante":
                                        print("📤 Subiendo comprobante a Birlik")
                                        agregar_comprobante_pago(driver,wait,id_cuota,ruta_final)
                                        resultado_accion = "Factura Enviada Anteriormente"
                                    else:
                                        print("📤 Subiendo todos los documentos a Birlik")
                                        cancelar_y_agregar_cuota(driver,wait,id_cuota,comprobante_valor,fecha,ruta_final,ruta_imagen_sunat,resultado_importe)
                                        resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente}", "Enviar Factura")'
            
                                    resultado_birlik = True
                                    break  # Salir del bucle porque ya funcionó con esa fecha

                                else:
                                    resultado_accion = f'=HYPERLINK("{url_sunat}", "Ver Sunat")'
                                    continue # Si no es True, salta al siguiente intento

                        except Exception as ex:
                            print(f"Error general al intentar interactuar con el menú desplegable, Detalles {ex}")
    
            if not fila_encontrada_descarga:
                print("❌ No se encontró la fila con el Documento esperado \n⌛ Consultando en la otra Compañia")
                resultado_sunat_cre, resultado_birlik_cre ,resultado_ocr_cre ,resultado_accion_cre = buscaryRegistrarenCrecer(driver,wait,fecha_emision_valor,comprobante_valor,importe_valor,id_cuota,ruc_compania,numero_ruc,numero_proforma,numero_poliza,estadoCuota_birlik,fk_Cliente,tipo_doc_birlik,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,resultado_importe)
                resultado_sunat = resultado_sunat_cre
                resultado_birlik = resultado_birlik_cre
                resultado_ocr = resultado_ocr_cre
                resultado_accion = resultado_accion_cre

    except Exception as e:
        print(f"⚠️ Detalles del error: {e}")
    finally:
        driver.quit()
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" ,"Si" if resultado_ocr else "No",resultado_estado,resultado_accion

def main():
    
    #------API---------

    json_cuotas = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

    if not json_cuotas:
        print("❌ No hay cuotas pendientes para esta compañia")
    else:

        ruta_salida_API,ruta_salida,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

        guardarJson(json_cuotas,ruta_salida_API)

        try:
            print("\n📁 Iniciando procesamiento para Sanitas Protecta y Crecer...")

            try:
                df = pd.read_excel(ruta_salida_API, engine="openpyxl",dtype={"numeroDocumento": str})
            except Exception as e:
                raise Exception(f" Error al leer el archivo Excel | Motivo: {e}")
            
            df["Importe"] = ""
            df["Sunat"] = ""
            df["Birlik"] = ""
            df["OCR"] = ""
            df["Estado"] = ""
            df["Acción"] = ""

            total_filas = len(df)

            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                try:
                    importe_estado,sunat_estado,birlik_estado,fecha_detectada,estado_estado,accion_estado = procesar_fila(
                        row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores)
                    df.at[index, "Importe"] = importe_estado
                    df.at[index, "Sunat"] = sunat_estado
                    df.at[index, "Birlik"] = birlik_estado
                    df.at[index, "OCR"] = fecha_detectada
                    df.at[index, "Estado"] = estado_estado
                    df.at[index, "Acción"] = accion_estado

                    print(f"\n✅ Fila {index} guardada correctamente")
                except Exception as e:
                    print(f"❌ Error procesando fila {index}: {e}")
                finally:
                    df.to_excel(ruta_salida, index=False)
                
                time.sleep(2)
                    
            print(f"\n✅ Flujo finalizado")

        finally:

            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                #print("🧹 Carpeta eliminada correctamente")

            time.sleep(5)

if __name__ == "__main__":
    main()
