import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import time
import sys
from src.robot_engine import RobotFormulario

def validar_columnas(df):
    """
    Verifica que el Excel tenga las columnas críticas que el robot necesita.
    Retorna (True, "") si todo está bien, o (False, mensaje_error).
    """
    # Estas deben coincidir EXACTAMENTE con las que usamos en robot_engine.py
    columnas_requeridas = [
        'Rubro', 'Clasificacion', 'Linea', 'CodigoBarras', 
        'Descripcion', 'Marca', 'Peso', 'Impuesto', 
        'Compra', 'Venta', 'Capacidad', 'Embalaje'
    ]
    
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    
    if faltantes:
        return False, f"Faltan columnas en el Excel:\n{', '.join(faltantes)}"
    return True, ""

def main():
    # 1. Configuración de la ventana (oculta)
    root = tk.Tk()
    root.withdraw()

    # 2. Selección del archivo
    ruta_excel = filedialog.askopenfilename(
        title="Selecciona la Planilla de Carga (Excel)",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not ruta_excel:
        print("Operación cancelada por el usuario.")
        return

    # 3. Lectura del Excel
    try:
        df = pd.read_excel(ruta_excel)
        # Limpieza básica: convertir NaN a string vacío para evitar errores
        df = df.fillna('') 
    except Exception as e:
        messagebox.showerror("Error de Lectura", f"No se pudo leer el archivo:\n{e}")
        return

    # 4. Validación de Columnas
    es_valido, mensaje = validar_columnas(df)
    if not es_valido:
        messagebox.showerror("Excel Inválido", mensaje)
        return

    # 5. Decisión: ¿Prueba o Realidad?
    es_prueba = messagebox.askyesno(
        "Configuración de Ejecución",
        f"Se cargaron {len(df)} registros.\n\n"
        "¿Quieres ejecutar en MODO PRUEBA (Dry Run)?\n\n"
        "SÍ = Solo genera un archivo de texto (Seguro).\n"
        "NO = Toma el control del mouse y escribe (Cuidado)."
    )

    # 6. Inicializar el Robot
    # Velocidad 0.5 es un buen punto de partida para Progress
    bot = RobotFormulario(modo_prueba=es_prueba, velocidad=0.5)

    if not es_prueba:
        confirmacion = messagebox.askokcancel(
            "Última Advertencia",
            "⚠️ MODO REAL ACTIVADO ⚠️\n\n"
            "1. Abre el software de la empresa.\n"
            "2. Pon el cursor en el campo 'RUBRO'.\n"
            "3. Al dar OK, tendrás 5 segundos para cambiar de ventana.\n"
            "4. Mueve el mouse a la esquina superior izquierda para abortar."
        )
        if not confirmacion:
            return
        
        print("⏳ INICIANDO EN 5 SEGUNDOS... CAMBIA DE VENTANA AHORA.")
        time.sleep(5)

    # 7. Ejecución del Bucle
    print("🚀 Iniciando proceso automatizado...")
    
    registros_exitosos = 0
    errores = 0

    try:
        for index, row in df.iterrows():
            # Sumamos 2 al index porque Excel empieza en fila 1 y tiene cabecera
            numero_fila_excel = index + 2 
            
            try:
                bot.procesar_producto(row, numero_fila_excel)
                registros_exitosos += 1
                
                # Pequeña pausa entre productos para que el sistema respire
                if not es_prueba:
                    time.sleep(1.5) 
                    
            except Exception as e:
                errores += 1
                print(f"❌ Error en fila {numero_fila_excel}: {e}")
                # Aquí podrías decidir si parar o seguir. Por ahora seguimos.

    except KeyboardInterrupt:
        print("\n🛑 Ejecución detenida manualmente.")
        messagebox.showwarning("Interrupción", "El proceso fue detenido por el usuario.")
    except pyautogui.FailSafeException:
        print("\n🛑 FAILSAFE ACTIVADO: Mouse en esquina de seguridad.")
        messagebox.showerror("Emergencia", "Se activó el FailSafe. Proceso abortado.")

    # 8. Reporte Final
    mensaje_final = (
        f"Proceso finalizado.\n\n"
        f"✅ Procesados: {registros_exitosos}\n"
        f"❌ Errores: {errores}\n"
    )
    
    if es_prueba:
        mensaje_final += f"\nRevisa el log en: {bot.log_path}"
    
    messagebox.showinfo("Resumen", mensaje_final)

if __name__ == "__main__":
    main()