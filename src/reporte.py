import pandas as pd
import os
from datetime import datetime

def generar_reporte_final(lista_resultados):
    """
    Recibe una lista de diccionarios con el estado de cada fila
    y genera un Excel en la carpeta 'reportes/'.
    """
    if not lista_resultados:
        print("⚠️ No hay datos para generar reporte.")
        return

    # 1. Crear carpeta si no existe
    carpeta = "reportes"
    os.makedirs(carpeta, exist_ok=True)

    # 2. Generar nombre único con Timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"{carpeta}/resultado_carga_{timestamp}.xlsx"

    # 3. Crear DataFrame y Guardar
    try:
        df = pd.DataFrame(lista_resultados)
        
        # Reordenamos columnas para que quede legible (si existen)
        cols_ordenadas = [col for col in ['Fila', 'Codigo', 'Producto', 'Estado', 'Detalle', 'Hora'] if col in df.columns]
        # Agregamos el resto de columnas que pudieran venir
        cols_ordenadas += [c for c in df.columns if c not in cols_ordenadas]
        
        df = df[cols_ordenadas]
        
        df.to_excel(nombre_archivo, index=False)
        print(f"\n📊 REPORTE GENERADO EXITOSAMENTE: {nombre_archivo}")
        print(f"   (Contiene {len(df)} registros procesados)")
        
    except Exception as e:
        print(f"❌ Error al guardar el Excel de reporte: {e}")