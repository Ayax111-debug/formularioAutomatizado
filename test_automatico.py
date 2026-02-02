import subprocess
import time
import pandas as pd
import pyautogui
import sys
import os
from src.robot_engine import RobotFormulario

def ejecutar_prueba_integrada():
    print("🤖 INICIANDO PROTOCOLO DE PRUEBA AUTOMÁTICA")
    
    # 1. Verificar datos
    ruta_excel = "data/ejemplo_carga.xlsx"
    if not os.path.exists(ruta_excel):
        print("❌ Error: No encuentro el Excel en data/ejemplo_carga.xlsx")
        print("   -> Ejecuta primero 'python generar_excel_prueba.py'")
        return

    # 2. Lanzar el Visor QA en un subproceso (segundo plano)
    print("🖥️  Abriendo ventana de pruebas...")
    # Usamos python para abrir el otro script
    proceso_visor = subprocess.Popen([sys.executable, "visor_qa.py"])
    
    # 3. TIEMPO MUERTO: Esperar a que la ventana aparezca
    print("⏳ Esperando 3 segundos para carga de interfaz...")
    time.sleep(3)

    # 4. AUTO-FOCUS (El truco sucio pero efectivo)
    # Hacemos clic en el centro de la pantalla para asegurar que el Visor tenga el foco
    ancho, alto = pyautogui.size()
    pyautogui.click(ancho / 2, alto / 2)
    print("🎯 Foco asegurado en ventana de pruebas.")

    # 5. Instanciar el Robot en MODO REAL (False)
    # Le bajamos la velocidad un poco para ver mejor lo que hace (0.3s)
    bot = RobotFormulario(modo_prueba=False, velocidad=0.3)

    # 6. Cargar datos (Solo procesaremos los primeros 3 para no estar años esperando)
    df = pd.read_excel(ruta_excel).head(3) # <--- SOLO 3 REGISTROS
    df = df.fillna('')

    print(f"🚀 Ejecutando carga de {len(df)} productos de prueba...")
    
    try:
        for index, row in df.iterrows():
            bot.procesar_producto(row, index + 1)
            time.sleep(1) # Pausa visual entre productos
            
    except pyautogui.FailSafeException:
        print("🛑 FAILSAFE: Mouse movido a la esquina. Abortando.")
    except Exception as e:
        print(f"❌ Error durante la ejecución: {e}")
    finally:
        # 7. Limpieza
        print("🏁 Prueba finalizada.")
        print("   -> Cerrando visor en 5 segundos...")
        time.sleep(5)
        proceso_visor.terminate() # Matar el proceso del visor
        print("👋 Bye.")

if __name__ == "__main__":
    # Seguridad: FailSafe activado
    pyautogui.FAILSAFE = True
    ejecutar_prueba_integrada()