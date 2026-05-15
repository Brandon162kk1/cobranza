import pandas as pd
from openpyxl import load_workbook
from .get import ObtenerListadeDatosporFk_Compania

def guardarDatosAPI_excel(datos_cobranza,ruta_salida_API):

    excel_json = pd.DataFrame(datos_cobranza)
    excel_json.to_excel(ruta_salida_API, index=False)

    wb = load_workbook(ruta_salida_API)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Ej: "A", "B", etc.
        for cell in col:
           try:
              if cell.value:
                max_length = max(max_length, len(str(cell.value)))
           except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(ruta_salida_API)

def consultarAPI(url,ids_compania):

    # Lista para acumular resultados
    todos_los_datos = []

    #print(f"📡 Consultando API de Birlik para todas las CIAs.")
    for fk_compania in ids_compania:
        #print(f"📡 Consultando API para compañía {fk_compania}...")

        #-- Esta API para enviar factura
        datos_cobranza = ObtenerListadeDatosporFk_Compania(url,fk_compania)
    
        if datos_cobranza is None:
            #print(f"❌ Error al obtener datos para compañía {fk_compania}.")
            continue

        if not datos_cobranza:  # Si es [] o vacío
            #print(f"⚠️ No hay cuotas para la compañía {fk_compania}.")
            continue

        # Si hay datos, convertir a DataFrame y acumular
        df = pd.DataFrame(datos_cobranza)
        if df.empty:
            print(f"⚠️ DataFrame vacío para la compañía {fk_compania}.")
            continue

        todos_los_datos.append(df)
        #print(f"✅ Datos de la Cuota extraído correctamente de la compañía {fk_compania}")

    return todos_los_datos