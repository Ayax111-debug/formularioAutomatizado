import pandas as pd
import os

# Aseguramos que exista la carpeta data
os.makedirs("data", exist_ok=True)

# Datos de prueba simulando productos de supermercado/tienda
datos = {
    'Rubro': [21,21],
    'Clasificacion': ['001','001'],
    'Linea': ['004','004'],
    'CodigoBarras': [
        '7800143152603',
        '5437265423018' 
    ],
    'Descripcion': [
                        
        'CAJA ORG PLASTICA',
        'USB SLIM KEYBOARD'                            
    ],
    'Marca': [
         '1250',
         '1250'
    ],
    'Peso': [13,1],
    'Impuesto': [
        'ILA',
        'ILA'   # Con impuesto
    ],
    'Compra': ['KG','UN'],
    'Venta': ['UN','KG'],
    'Capacidad': [1000,1000],
    'Embalaje_1': [1,1],
    'Embalaje_2': [99,99]
}

# Crear DataFrame
df = pd.DataFrame(datos)

# Guardar en la carpeta data
ruta_salida = "data/ejemplo_carga.xlsx"
df.to_excel(ruta_salida, index=False)

print(f"✅ Archivo generado exitosamente en: {ruta_salida}")
print("   -> Contiene 1 registro de prueba.")
print("   -> Incluye caso de descripción larga (Fila 3).")
print("   -> Incluye casos de impuesto vacío.")