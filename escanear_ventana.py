from pywinauto import Application
import time

print("🕵️ INTENTO FINAL DE CONEXIÓN")
print("1. Ejecuta esto.")
print("2. Haz CLIC RÁPIDO en la ventana del software.")
print("3. Espera...")
time.sleep(5)

try:
    # Conectamos directamente a la aplicación activa, sea la que sea
    app = Application(backend="win32").connect(active_only=True)
    ventana = app.top_window()
    
    print(f"✅ Conectado a: {ventana.window_text()}")
    print("-----------------------------------------")
    ventana.print_control_identifiers()
    
except Exception as e:
    print(f"❌ No se pudo: {e}")