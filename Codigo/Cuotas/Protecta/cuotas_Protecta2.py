#--- Imports --
import pandas as pd 
import os
import time
import subprocess
import shutil
import zipfile
import random
#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait, Select
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.Birlik.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas,guardarJson,esperar_archivos_nuevos
from Cuotas.Positiva.cuotas_Positiva import mover_y_hacer_click_simple, escribir_lento
from GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_actual
from Correo.correo_it import enviarCaptcha
from Cuotas.Crecer.cuotas_Crecer import bloquear_interaccion,desbloquear_interaccion
#------------ PROTECTA ----------------
ruc_protecta_vly = '20517207331'
ids_compania = [25]
# -- Credenciales Protecta Vida Ley ---
url_protecta = os.getenv("url_protecta")
username_protecta = os.getenv("username_protecta")
password_protecta = os.getenv("password_protecta")
para_venv = os.getenv("para")
para_lista = para_venv.split(",") if para_venv else []
copia_venv = os.getenv("copia_cuotas")
copias_lista = copia_venv.split(",") if copia_venv else []
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Protecta_VidaLey_{get_timestamp()}"
# --- Configuración de 2captcha ---
API_KEY = "8b61feec172173ef48060a723af1b6c7"

#------- Errores comunes de 2Captcha
# ERROR_WRONG_USER_KEY → API key incorrecta.
# ERROR_KEY_DOES_NOT_EXIST → API key no válida.
# ERROR_ZERO_BALANCE → No tienes saldo.
# ERROR_WRONG_GOOGLEKEY → El sitekey no corresponde o no es válido.
# ERROR_PAGEURL → La URL de la página no es válida.

#fecha_hoy = datetime.now().strftime("%d/%m/%Y")

def formatear_fechas(fecha_inicio_str, fecha_fin_str):
    try:
        fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d %H:%M:%S")
        fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d %H:%M:%S")

        return fecha_inicio.strftime("%d/%m/%Y"), fecha_fin.strftime("%d/%m/%Y")
    except Exception as e:
        print("❌ Error al formatear fechas:", e)
        return fecha_inicio_str, fecha_fin_str

def procesar_fila(driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores):

    # Extraer valores y quitar espacios en blanco
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    ruc_cliente_birlik = str(row["numeroDocumento"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    numero_proforma_birlik = str(row["codigoCuota"]).strip()
    numero_poliza_birlik = str(row["numeroPoliza"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()
    fk_Cliente_birlik = str(row["fk_Cliente"]).strip() 
    vigencia_inicio_birlik = str(row["vigenciaInicio"]).strip()
    vigencia_fin_birlik = str(row["vigenciaFin"]).strip()

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    print(f"CIA: Protecta Vida Ley, RUC: {ruc_protecta_vly}")

    try:
        
        # input_ruc = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='documento']")))
        # input_ruc.clear()
        # input_ruc.send_keys(ruc_cliente_birlik)
        # print(f"🖱️ Se ingreso el numero de RUC {ruc_cliente_birlik}")

        # Esperar que el select esté presente
        select_element = wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='filterType']")))
        select = Select(select_element)
        select.select_by_value("2")
        print("✅ Se seleccionó la opcion Póliza")

        input_poliza = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[formcontrolname='poliza']")))
        input_poliza.clear()
        input_poliza.send_keys(numero_poliza_birlik)
        print(f"🖱️ Se ingreso el numero de Póliza {numero_poliza_birlik}")

        # fecha_inicio_input = wait.until(EC.presence_of_element_located((By.XPATH, "//span[normalize-space()='FECHA INICIO:']/following-sibling::input[@bsdatepicker]")))

        # driver.execute_script("""
        #     arguments[0].value = arguments[1];
        #     arguments[0].dispatchEvent(new Event('input'));
        #     arguments[0].dispatchEvent(new Event('change'));
        # """, fecha_inicio_input, vigencia_inicio_birlik)

        # fecha_inicio_input.send_keys(Keys.TAB)

        # print(f"🖱️ Fecha Vigencia Inicio OK → {vigencia_inicio_birlik}")

        # fecha_fin_input = wait.until(EC.presence_of_element_located((By.XPATH, "//span[normalize-space()='FECHA FIN:']/following-sibling::input[@bsdatepicker]")))

        # driver.execute_script("""
        #     arguments[0].value = arguments[1];
        #     arguments[0].dispatchEvent(new Event('input'));
        #     arguments[0].dispatchEvent(new Event('change'));
        # """, fecha_fin_input, vigencia_fin_birlik)

        # fecha_fin_input.send_keys(Keys.TAB)

        # print(f"🖱️ Fecha Vigencia Fin OK → {vigencia_fin_birlik}")

        btn_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Buscar']")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_buscar)
        btn_buscar.click()
        print("🖱️ Clic en Buscar")

        try:
            wait.until(EC.invisibility_of_element_located((By.TAG_NAME, "ngx-spinner")))
        except:
            pass

        wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[contains(@class,'table')]//tbody//tr")))
        filas = driver.find_elements(By.XPATH, "//table[contains(@class,'table')]//tbody//tr")

        if len(filas) == 1 and "No se encontraron" in filas[0].text:
            resultado_estado = 'No se sabe'
            resultado_accion = "Revisar Póliza y Código de Cuota"
            raise Exception("No se encontraron registros en la tabla")

        print(f"✅ Total filas encontradas: {len(filas)}")

        fila_encontrada = False

        for i, fila in enumerate(filas, start=1):

            columnas = fila.find_elements(By.TAG_NAME, "td")

            if len(columnas) < 7:
                continue

            print("-------------------------------------------------")
            producto = columnas[1].text.strip()
            doc_referencia = columnas[2].text.strip()
            #tipo_comprobante = columnas[3].text.strip()
            serie_numero = columnas[4].text.strip()
            monto = columnas[5].text.replace("S/", "").strip()
            #ruc = columnas[6].text.strip()
            #contratante = columnas[7].text.strip()
            poliza = columnas[8].text.strip()
            fecha_emision = columnas[10].text.strip()
            estado = columnas[11].text.strip()

            print(f"Fila {i},Producto : {producto} ,Documento: {doc_referencia},Serie: {serie_numero},Monto: {monto},Póliza: {poliza},Estado: {estado},Fecha Emisión: {fecha_emision}")

            # Validación de importes con tolerancia
            diferencia = abs(float(monto) - float(importe_total_birlik))

            if diferencia < 0.05:

                resultado_estado = estado
                print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(monto)}")

                if diferencia > 0.05:
                    print("❌ Los importes No coinciden")
                else:
                    resultado_importe = True
                    print("✅ Los importes Coinciden")

                fila_encontrada = True

                if estado.lower() == "cancelado":

                    ventana_principal_protecta = driver.current_window_handle

                    try:

                        checkbox = columnas[0].find_element(By.TAG_NAME, "input")

                        if not checkbox.is_selected():
                            driver.execute_script("arguments[0].click();", checkbox)             
                            print(f"🖱️ Checkbox marcado")

                        try:
                            wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class,'la-ball-triangle-path')]")))
                        except:
                            pass

                        btn_descargar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(normalize-space(),'Descargar')]]")))

                        # Guardar archivos antes del clic
                        archivos_antes = set(os.listdir(ruta_carpeta_facturas))

                        driver.execute_script("arguments[0].click();", btn_descargar)
                        print("🖱️ Clic con JS en el botón 'Descargar'")

                        archivo_nuevo = esperar_archivos_nuevos(ruta_carpeta_facturas,archivos_antes,".zip",cantidad=1)

                        if archivo_nuevo:
                            print(f"✅ Archivo .zip descargado exitosamente")
                            ruta_original = archivo_nuevo[0]
                        else:
                            raise Exception("No se encontró archivo .zip nuevo después de la descarga")

                        try:
                            if not os.path.exists(ruta_original):
                                raise Exception(f"No existe el ZIP descargado: {ruta_original}")

                            with zipfile.ZipFile(ruta_original, 'r') as z:
                                pdfs = [f for f in z.namelist() if f.endswith(".pdf")]
                                if not pdfs:
                                    raise Exception("El ZIP no contiene PDF")

                                pdf_interno = pdfs[0]

                                nuevo_nombre_pdf = f"{numero_poliza_birlik}_{serie_numero}.pdf"

                                ruta_pdf_salida = os.path.join(ruta_carpeta_facturas, nuevo_nombre_pdf)

                                with open(ruta_pdf_salida, "wb") as f:
                                    f.write(z.read(pdf_interno))

                            print(f"✅ PDF extraído en: {ruta_pdf_salida}")

                        finally:
                            if os.path.exists(ruta_original):
                                os.remove(ruta_original)
                                print("🧹 Archivo ZIP eliminado correctamente")
                            else:
                                print("⚠️ El archivo ZIP no existe")

                        fecha_habiles_factura = []

                        fecha_habiles_factura.append(fecha_emision)

                        for fecha in fecha_habiles_factura: #(*)
                            print("---------------------------------------")
                            print(f"⌛ Probando con la Fecha hábil: {fecha}")
                            #------------INGRESA A SUNAT-------  
                            nombre_imagen_sunat = f"{serie_numero}_{numero_poliza_birlik}.png"
                            ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)

                            resultado = consultarValidezSunat(driver,wait,ruc_protecta_vly,tipo_doc_birlik,ruc_cliente_birlik,serie_numero,fecha,monto,ruta_imagen_sunat)

                            driver.switch_to.window(ventana_principal_protecta)
                            print("🔄 Volviendo a la ventana de la CIA")

                            if resultado is None:
                                resultado_accion = f'=HYPERLINK("{login_sunat}", "Sunat Bloqueado")'
                                break
                            elif resultado:

                                resultado_sunat = True

                                if estadoCuota_birlik == "Pendiente-comprobante":

                                    print("📤 Subiendo comprobante a Birlik...")
                                    agregar_comprobante_pago(driver,wait,id_cuota_birlik,ruta_pdf_salida)
                                    resultado_accion = "Factura Enviada Anteriormente"

                                else:

                                    print("📤 Subiendo todos los documentos a Birlik...")
                                    if cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,serie_numero,fecha,ruta_pdf_salida,ruta_imagen_sunat,resultado_importe):
                                        resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
                                    else:
                                        resultado_accion = "No se pudo registrar en Birlik"
                                        
                                resultado_birlik = True
                                break #(*)
                            else:
                                resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'

                    except Exception as e:
                         raise Exception(f"Error descargando ZIP, Motivo -> {e}")
                    finally:
                        break

                else:
                    print(f"La Cuota {numero_proforma_birlik} tiene estado {estado.lower}.")

        if not fila_encontrada:
            print(f"❌ No se encontró ninguna fila que coincida con el monto {importe_total_birlik} para la cuota '{numero_proforma_birlik}'.")
        
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

            driver.get(url_protecta)
            print("\n⌛ Ingresando a la URL")

            user_input = wait.until(EC.presence_of_element_located((By.ID, "username")))
            user_input.clear()

            mover_y_hacer_click_simple(driver, user_input)
            time.sleep(random.uniform(0.97, 0.99))

            escribir_lento(user_input, username_protecta, min_delay=0.97, max_delay=0.99)
            print("⌨️ Digitando el Username")

            time.sleep(1 + random.random() * 1.5)

            pass_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
            pass_input.clear()

            mover_y_hacer_click_simple(driver, pass_input)
            time.sleep(random.uniform(0.97, 0.99))

            escribir_lento(pass_input, password_protecta, min_delay=0.97, max_delay=0.99)
            print(f"⌨️ Digitando el Password '{password_protecta}'")

            desbloquear_interaccion()

            #-------------
            print("🧩 Resuelve el CAPTCHA manualmente y clic en 'Ingresar'.")

            puerto = os.getenv("NOVNC_PORT")
            enviarCaptcha(para_lista,copias_lista,puerto,"Protecta Vida Ley",ruta_imagen=None)

            wait_humano = WebDriverWait(driver,300)
            wait_humano.until(EC.presence_of_element_located((By.XPATH,"//a[contains(@class,'menu-item-father')][not(@hidden)]//span[normalize-space()='Mis Comprobantes']/ancestor::a")))
            bloquear_interaccion()
            print("✅ Login exitoso detectado (Cerrar sesión visible)")
            print("🚀 Continuando flujo automáticamente")
            #-------------

            menu_mis_comprobantes = wait.until(EC.element_to_be_clickable((By.XPATH,"//a[contains(@class,'menu-item-father')][not(@hidden)]//span[normalize-space()='Mis Comprobantes']/ancestor::a")))

            driver.execute_script("""
                arguments[0].scrollIntoView({block:'center'});
                arguments[0].click();
            """, menu_mis_comprobantes)
            print("🖱️ Click en 'Mis Comprobantes'")

            while True:

                json_cuotas = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

                try:

                    if not json_cuotas:
                        raise Exception("No hay cuotas pendientes para esta compañia")

                    print("\n📁 Iniciando procesamiento para Protecta Vida Ley...")

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
                                driver,wait,row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores)

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