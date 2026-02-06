import time
import pandas as pd
import pyautogui
import os
import sys
from datetime import datetime

# Importamos el robot y el generador de reportes
from src.robot_engine import RobotFormulario
from src.reporte import generar_reporte_final # <--- ESTO FALTABA PARA CREAR EL EXCEL

def ejecutar_prueba_integrada():
    print("🤖 INICIANDO PROTOCOLO DE PRUEBA AUTOMÁTICA")
    
    # 1. Verificar datos
    ruta_excel = "data/ejemplo_carga.xlsx"
    if not os.path.exists(ruta_excel):
        print(f"❌ Error: No encuentro el Excel en {ruta_excel}")
        return

    # 2. Instanciar el Robot en MODO REAL (False)
    # velocidad=0.3 para verlo trabajar, luego bájalo a 0.01 si quieres que vuele
    bot = RobotFormulario(modo_prueba=False, velocidad=0.3, titulo_ventana="(Com-Mae6)")

    # 3. Cargar datos
    try:
        df = pd.read_excel(ruta_excel, dtype=str).fillna('')
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
        return

    print(f"🚀 Ejecutando carga de {len(df)} productos...")
    
    # --- ACUMULADOR DE RESULTADOS ---
    bitacora_resultados = [] 

    try:
        for index, row in df.iterrows():
            numero_fila = index + 2
            
            # Preparamos la ficha del reporte
            resultado_fila = {
                "Fila Excel": numero_fila,
                "Producto": row.get('Descripcion', 'N/A'),
                "Hora": datetime.now().strftime("%H:%M:%S"),
                "Estado": "PENDIENTE",
                "Detalle": ""
            }

            try:
                # Ejecutamos el robot
                exito = bot.procesar_producto(row, numero_fila)
                
                if exito:
                    resultado_fila["Estado"] = "EXITOSO"
                    resultado_fila["Detalle"] = "Carga Correcta"
                else:
                    resultado_fila["Estado"] = "FALLIDO"
                    resultado_fila["Detalle"] = "Error Negocio (Duplicado/Validación)"

                # Pausa visual entre productos
                time.sleep(1) 
            
            except Exception as e:
                resultado_fila["Estado"] = "ERROR CRITICO"
                resultado_fila["Detalle"] = str(e)
                print(f"❌ Error técnico: {e}")

            # ¡AQUÍ GUARDAMOS EL RESULTADO!
            bitacora_resultados.append(resultado_fila)
            
    except pyautogui.FailSafeException:
        print("🛑 FAILSAFE: Mouse movido a la esquina. Abortando.")
    except KeyboardInterrupt:
        print("🛑 Interrupción manual.")
    finally:
        # 4. Generar el Reporte Final
        print("-" * 40)
        print("🏁 Prueba finalizada.")
        
        if bitacora_resultados:
            print("💾 Generando Excel de resultados...")
            generar_reporte_final(bitacora_resultados) # <--- AQUÍ SE CREA LA CARPETA Y EL ARCHIVO
        else:
            print("⚠️ No se generaron registros para el reporte.")

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    ejecutar_prueba_integrada()