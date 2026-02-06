import pyautogui
import time
import os
import keyboard 
from datetime import datetime

# --- IMPORTACIÓN ROBUSTA DE PYWINAUTO ---
try:
    from pywinauto import Application, findwindows
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("⚠️ ADVERTENCIA: 'pywinauto' no está instalado. Auto-enfoque desactivado.")

class RobotFormulario:
    def __init__(self, 
                 modo_prueba=False,          # Por defecto: MODO REAL
                 velocidad=0.01,             # Por defecto: VELOCIDAD FLASH
                 log_path="logs/auditoria.txt", 
                 timelapse_segundos=0,       # Por defecto: SIN CÁMARA LENTA
                 titulo_ventana="(Com-Mae6)"): # Título específico de tu software
        
        self.modo_prueba = modo_prueba
        self.velocidad = velocidad
        self.log_path = log_path
        self.timelapse_segundos = timelapse_segundos 
        self.titulo_ventana = titulo_ventana
        
        self.pausado = False
        self.detenido = False
        
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        # Configuración de PyAutoGUI
        if not self.modo_prueba:
            pyautogui.PAUSE = self.velocidad
            pyautogui.FAILSAFE = True # Mueve el mouse a la esquina superior izquierda para abortar

        # Hotkeys de Control
        keyboard.add_hotkey('F8', self._toggle_pausa)
        keyboard.add_hotkey('esc', self._detener_emergencia)

        print("\n🤖 ROBOT INICIADO EN MODO PRODUCCIÓN")
        print("   [F8]  = PAUSAR / REANUDAR")
        print("   [ESC] = DETENER EMERGENCIA")
        
        if self.titulo_ventana:
            print(f"   [🎯]  TARGET VENTANA: '{self.titulo_ventana}'")

        if self.timelapse_segundos > 0:
            print(f"   [🐢]  TIMELAPSE ACTIVO: {self.timelapse_segundos}s")
        else:
            print("   [⚡]  VELOCIDAD MÁXIMA ACTIVADA")
        print("-" * 40)

    def _toggle_pausa(self):
        self.pausado = not self.pausado
        estado = "⏸️ PAUSADO" if self.pausado else "▶️ REANUDANDO..."
        print(f"\n{estado}\n")

    def _detener_emergencia(self):
        self.detenido = True
        print("\n🛑 SEÑAL DE DETENCIÓN RECIBIDA (ESC)\n")

    def _chequear_estado(self):
        if self.detenido:
            raise KeyboardInterrupt("Detención manual por usuario (ESC)")

        if self.pausado:
            self._log("⏳ Robot en PAUSA... (Presiona F8 para continuar)")
            while self.pausado:
                if self.detenido:
                     raise KeyboardInterrupt("Detención manual durante pausa")
                time.sleep(0.5) 
            self._log("🚀 Reanudando operaciones...")

    def _log(self, mensaje):
        timestamp = datetime.now().strftime("%H:%M:%S")
        linea = f"[{timestamp}] {mensaje}"
        print(linea)
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(linea + "\n")

    def _enfocar_ventana(self):
        """Trae la ventana del software al frente."""
        if self.modo_prueba or not self.titulo_ventana or not PYWINAUTO_AVAILABLE:
            return

        try:
            handle = findwindows.find_window(title_re=f".*{self.titulo_ventana}.*", backend="win32")
            app = Application(backend="win32").connect(handle=handle)
            ventana = app.window(handle=handle)
            
            if ventana.get_show_state() == 2: # Si está minimizada
                 ventana.restore()
            
            ventana.set_focus()
        except Exception:
            # Silencioso en producción para no llenar el log, salvo que sea crítico
            pass

    def _ejecutar_timelapse(self, accion_detalle):
        if self.timelapse_segundos > 0:
            print(f"   🐢 [TIMELAPSE] {accion_detalle} ({self.timelapse_segundos}s...)")
            time.sleep(self.timelapse_segundos)

    def _escribir(self, texto):
        self._chequear_estado() 
        val = str(texto).strip()
        if val.lower() == 'nan' or val == '':
            return 
        
        self._ejecutar_timelapse(f"Escribir '{val}'")

        if self.modo_prueba:
            self._log(f"⌨️  ESCRIBIR: '{val}'")
        else:
            pyautogui.write(val, interval=0.01) # Ultra rápido para texto normal

    def _tab(self, cantidad=1):
        if cantidad == 0: return
        self._chequear_estado() 
        self._ejecutar_timelapse(f"TAB x{cantidad}")

        if self.modo_prueba:
            self._log(f"🔘 TAB (x{cantidad})")
        else:
            # Optimización para múltiples tabs
            if cantidad > 5:
                pyautogui.press('tab', presses=cantidad, interval=0.01)
            else:
                pyautogui.press('tab', presses=cantidad)

    def _enter(self):
        self._chequear_estado() 
        self._ejecutar_timelapse("ENTER")
        if self.modo_prueba:
            self._log("🔘 ENTER")
        else:
            pyautogui.press('enter')
    
    def _esperar_carga(self, tiempo=1.0):
        self._chequear_estado() 
        if self.modo_prueba:
            self._log(f"⏳ ESPERA TÉCNICA: {tiempo}s")
        else:
            time.sleep(tiempo)

    def _manejar_error_y_limpiar(self):
        if self.modo_prueba:
            return False

        imagen = "data/error.png"
        if not os.path.exists(imagen):
            return False

        try:
            # Grayscale=True hace la búsqueda más rápida
            if pyautogui.locateOnScreen(imagen, confidence=0.9, grayscale=True):
                self._log("👁️ ERROR DETECTADO. Limpiando...")
                self._enter()
                time.sleep(0.5) 
                pyautogui.hotkey('alt', 'o') # Botón Limpiar
                time.sleep(1.0) 
                return True 
        except Exception:
            pass
        return False

    def _ingresar_capacidad_mascara(self, valor_raw):
        self._chequear_estado()
        if str(valor_raw).lower() == 'nan' or str(valor_raw).strip() == '':
            return

        try:
            valor = float(str(valor_raw).replace(',', '.'))
        except ValueError:
            self._log(f"⚠️ Valor inválido: {valor_raw}")
            return

        flechas = 0
        texto = ""

        if valor >= 1:
            flechas = 0 
            texto = str(int(valor)) if valor.is_integer() else str(valor)
        else:
            # Lógica para < 1 (0.X, 0.0X, 0.00X)
            str_val = f"{valor:.3f}".split('.')[1]
            if valor >= 0.1:   flechas, texto = 1, str_val
            elif valor >= 0.01: flechas, texto = 2, str_val[1:]
            else:              flechas, texto = 3, str_val[2:]
            
            texto = texto.rstrip('0')

        self._ejecutar_timelapse(f"MASCARA: {valor} -> Flechas: {flechas}, Txt: {texto}")

        if not self.modo_prueba:
            pyautogui.press('delete') 
            time.sleep(0.05) # Pequeña pausa mecánica
            if flechas > 0:
                pyautogui.press('right', presses=flechas)
                time.sleep(0.05)
            if texto: 
                pyautogui.write(texto, interval=0.02)


    def procesar_producto(self, fila, indice):
        self._chequear_estado() 
        
        # 1. ENFOCAR VENTANA
        if self.titulo_ventana:
            self._enfocar_ventana()
            time.sleep(0.3) 
        
        self._log(f"🔹 PROCESANDO #{indice}: {fila.get('Descripcion', '')[:20]}...")

        # FASE 1: DATOS BÁSICOS
        self._esperar_carga(0.5) # Espera inicial reducida
        self._escribir(fila['Rubro'])
        self._escribir(fila['Clasificacion'])
        self._escribir(fila['Linea'])
        self._esperar_carga(0.5)
        self._tab(1) 

        # CHECK ERROR 1
        time.sleep(0.5)
        if self._manejar_error_y_limpiar():
            self._log(f"❌ ERROR FASE 1 (Rubro/Clas/Linea). Saltando.")
            return False
        
        # FASE 2: CÓDIGO BARRAS
        self._esperar_carga(0.5)
        self._escribir(fila['CodigoBarras'])
        # Tabs rápidos hasta descripción
        self._tab(1); self._esperar_carga(0.2)
        self._tab(1); self._esperar_carga(0.2)
        self._tab(1); self._esperar_carga(0.2)
        self._tab(1); self._esperar_carga(0.2)
        self._tab(1); self._esperar_carga(0.2)

        # CHECK ERROR 2 (Duplicado)
        time.sleep(0.8) # Un poco más de tiempo aquí es vital
        if self._manejar_error_y_limpiar():
            self._log(f"❌ DUPLICADO DETECTADO. Saltando.")
            return False

        # FASE 3: DETALLES
        desc = str(fila['Descripcion'])[:25]
        self._escribir(desc)
        self._tab(6)
        
        self._escribir(fila['Marca'])
        self._tab(6)

        # PESO (Con Máscara)
        self._ingresar_capacidad_mascara(fila['Peso'])
        self._tab(7)
        
        self._escribir(fila['Impuesto'])
        self._tab(4)

        self._escribir(fila['Compra'].upper())
        self._tab(1)
        self._escribir(fila['Venta'])
        
        # FASE 4: CAPACIDAD (Sincronización Crítica)
        self._tab(2) 
        self._esperar_carga(1.0) # <--- NO TOCAR ESTE 1.0s (Evita bug del delete)
        self._ingresar_capacidad_mascara(fila['Capacidad'])
        
        # FASE 5: FINAL
        self._tab(2)
        self._escribir(fila['Embalaje_1'])
        self._tab(1)
        self._escribir(fila['Embalaje_2'])
        self._tab(12)

        self._escribir("15")
        self._tab(41)
        
        self._enter() 
        self._esperar_carga(1.0) # Espera final para guardado
        
        return True