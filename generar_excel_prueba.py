import pandas as pd
import os

# Aseguramos que exista la carpeta data
os.makedirs("data", exist_ok=True)

# Datos de prueba simulando productos de supermercado/tienda
datos = {
    'Rubro': [101, 101, 102, 103, 104, 105, 106, 107, 108, 109],
    'Clasificacion': [10, 10, 20, 30, 15, 25, 35, 40, 45, 50],
    'Linea': [1, 1, 2, 3, 1, 2, 3, 4, 5, 6],
    'CodigoBarras': [
        '780001001', '780001002', '780001003', '780001004', '780001005',
        '780001006', '780001007', '780001008', '780001009', '780001010'
    ],
    'Descripcion': [
        'ARROZ GRADO 2',                    # Normal
        'ACEITE MARAVILLA 1LT',             # Normal
        'ESTA DESCRIPCION TIENE MAS DE 25 CARACTERES PARA PROBAR EL CORTE', # Test largo
        'FIDEOS ESPIRALES',                 # Normal
        'BEBIDA COLA 3LT',                  # Normal
        'DETERGENTE LIQUIDO',               # Normal
        'GALLETAS DE VINO',                 # Normal
        'LECHE ENTERA 1L',                  # Normal
        'ATUN EN AGUA',                     # Normal
        'CONFORT DOBLE HOJA'                # Normal
    ],
    'Marca': [
        'TUCAPEL', 'CHEF', 'LUCCHETTI', 'CAROZZI', 'COCACOLA', 
        'OMO', 'MCKAY', 'COLUN', 'ROBINSON', 'ELITE'
    ],
    'Peso': [1.0, 1.0, 0.4, 0.4, 3.2, 3.0, 0.15, 1.0, 0.17, 0.5],
    'Impuesto': [
        'ILA',   # Con impuesto
        '',      # Vacío (Para probar que salte escritura pero haga los TABS)
        '',      
        'ILA', 
        'IABA',  # Impuesto bebidas analcohólicas
        '',
        'ILA',
        '',
        '',
        ''
    ],
    'Compra': ['KG', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN'],
    'Venta': ['UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN', 'UN'],
    'Capacidad': [1000, 1000, 400, 400, 3000, 3000, 150, 1000, 170, 500],
    'Embalaje': [10, 12, 20, 20, 6, 4, 30, 12, 24, 4]
}

# Crear DataFrame
df = pd.DataFrame(datos)

# Guardar en la carpeta data
ruta_salida = "data/ejemplo_carga.xlsx"
df.to_excel(ruta_salida, index=False)

print(f"✅ Archivo generado exitosamente en: {ruta_salida}")
print("   -> Contiene 10 registros de prueba.")
print("   -> Incluye caso de descripción larga (Fila 3).")
print("   -> Incluye casos de impuesto vacío.")