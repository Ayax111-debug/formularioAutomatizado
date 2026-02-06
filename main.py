import pandas as pd
import time
import sys
import os
from datetime import datetime

# Importamos tus módulos
from src.robot_engine import RobotFormulario
from src.reporte import generar_reporte_final  # <--- NUEVO IMPORT

def ejecutar_carga_completa():
    print("🤖 INICIANDO EJECUCIÓN MASIVA CON REPORTE FINAL")
    
    # 1. Verificar datos
    ruta_excel = "data/ejemplo_carga.xlsx"
    if not os.path.exists(ruta_excel):
        print(f"❌ Error: No encuentro el Excel en {ruta_excel}")
        return

    # 2. Instanciar el Robot (Modo Producción)
    bot = RobotFormulario(modo_prueba=False, velocidad=0.3, titulo_ventana="Progress")

    # 3. Cargar Excel
    try:
        df = pd.read_excel(ruta_excel, dtype=str).fillna('')
    except Exception as e:
        print(f"❌ Error al leer el Excel de entrada: {e}")
        return

    total = len(df)
    print(f"🚀 Procesando {total} registros...")
    print("-" * 50)

    # --- LISTA PARA ACUMULAR LOS RESULTADOS ---
    bitacora_resultados = [] 

    try:
        for index, row in df.iterrows():
            fila_n = index + 2
            sku = row.get('CodigoBarras', 'N/A')
            desc = row.get('Descripcion', 'Sin Nombre')
            
            # Estructura base del reporte para esta fila
            resultado_fila = {
                "Fila": fila_n,
                "Codigo": sku,
                "Producto": desc,
                "Estado": "PENDIENTE",
                "Detalle": "",
                "Hora": datetime.now().strftime("%H:%M:%S")
            }

            try:
                # EJECUCIÓN DEL ROBOT
                # El robot retorna True (Éxito) o False (Error controlado/Saltado)
                exito = bot.procesar_producto(row, fila_n)
                
                if exito:
                    resultado_fila["Estado"] = "EXITOSO"
                    resultado_fila["Detalle"] = "Carga OK"
                else:
                    resultado_fila["Estado"] = "FALLIDO"
                    resultado_fila["Detalle"] = "Validación Negocio (Duplicado/Error Form)"

                # Pausa técnica entre productos
                time.sleep(1)

            except Exception as e_tecnico:
                # Si explota el código (crash), lo capturamos aquí
                resultado_fila["Estado"] = "ERROR CRITICO"
                resultado_fila["Detalle"] = str(e_tecnico)
                print(f"❌ CRASH en fila {fila_n}: {e_tecnico}")

            # GUARDAMOS LA FICHA EN LA BITÁCORA
            bitacora_resultados.append(resultado_fila)

    except KeyboardInterrupt:
        print("\n🛑 Proceso detenido manualmente.")
    
    finally:
        print("-" * 50)
        print("🏁 PROCESO FINALIZADO.")
        
        # 4. GENERAR EL EXCEL FINAL (AQUÍ OCURRE LA MAGIA)
        if bitacora_resultados:
            print("💾 Guardando evidencia...")
            generar_reporte_final(bitacora_resultados)
        else:
            print("⚠️ No se procesaron registros.")

if __name__ == "__main__":
    import pyautogui
    pyautogui.FAILSAFE = True
    ejecutar_carga_completa()