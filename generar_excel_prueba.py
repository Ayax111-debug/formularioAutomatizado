import pandas as pd
import os

# Aseguramos que exista la carpeta data
os.makedirs("data", exist_ok=True)

# Datos de prueba simulando productos de supermercado/tienda
datos = {
    'Rubro': [22],
    'Clasificacion': ['001'],
    'Linea': ['006'],
    'CodigoBarras': [
        '7800143152603' 
    ],
    'Descripcion': [
        'Producto correcto prueba',          
    
                                 
    ],
    'Marca': [
         '1250',
         
         
    ],
    'Peso': [11],
    'Impuesto': [
        'ILA',
        
          # Con impuesto
    ],
    'Compra': ['KG',],
    'Venta': ['UN',],
    'Capacidad': [111,],
    'Embalaje_1': [1,],
    'Embalaje_2': [99,]
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