import pyautogui
import time
import os
import keyboard  # <--- NUEVA LIBRERÍA
from datetime import datetime

class RobotFormulario:
    def __init__(self, modo_prueba=True, velocidad=0.5, log_path="logs/auditoria.txt"):
        self.modo_prueba = modo_prueba
        self.velocidad = velocidad
        self.log_path = log_path
        
        # Estado del Robot
        self.pausado = False
        self.detenido = False
        
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        if not self.modo_prueba:
            pyautogui.PAUSE = self.velocidad
            pyautogui.FAILSAFE = True

        # Configurar los "Oídos" del robot (Hotkeys)
        # F8 alternará entre Pausa y Sigue
        keyboard.add_hotkey('F8', self._toggle_pausa)
        # ESC activará la bandera de detención
        keyboard.add_hotkey('esc', self._detener_emergencia)

        print("\n🎮 CONTROLES ACTIVADOS:")
        print("   [F8]  = PAUSAR / REANUDAR")
        print("   [ESC] = DETENER TODO\n")

    def _toggle_pausa(self):
        """Alterna el estado de pausa."""
        self.pausado = not self.pausado
        estado = "⏸️ PAUSADO" if self.pausado else "▶️ REANUDANDO..."
        print(f"\n{estado}\n")

    def _detener_emergencia(self):
        """Detiene el script de forma segura."""
        self.detenido = True
        print("\n🛑 SEÑAL DE DETENCIÓN RECIBIDA (ESC)\n")

    def _chequear_estado(self):
        """
        Esta función actúa como el 'Portero'.
        Se ejecuta antes de cada acción para ver si debe parar o esperar.
        """
        # 1. Si presionaron ESC, lanzamos error para romper el bucle
        if self.detenido:
            raise KeyboardInterrupt("Detención manual por usuario (ESC)")

        # 2. Si presionaron F8, entramos en un bucle infinito de espera
        if self.pausado:
            self._log("⏳ Robot en PAUSA... (Presiona F8 para continuar)")
            while self.pausado:
                # Si presionan ESC mientras está pausado, también salimos
                if self.detenido:
                     raise KeyboardInterrupt("Detención manual durante pausa")
                time.sleep(0.5) # Espera pasiva sin consumir CPU
            self._log("🚀 Reanudando operaciones...")

    def _log(self, mensaje):
        timestamp = datetime.now().strftime("%H:%M:%S")
        linea = f"[{timestamp}] {mensaje}"
        print(linea)
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(linea + "\n")

    def _escribir(self, texto):
        self._chequear_estado() # <--- VALIDACIÓN ANTES DE ACTUAR

        val = str(texto).strip()
        if val.lower() == 'nan' or val == '':
            return 

        if self.modo_prueba:
            self._log(f"⌨️  ESCRIBIR: '{val}'")
        else:
            pyautogui.write(val, interval=0.05)

    def _tab(self, cantidad=1):
        if cantidad == 0: return
        
        self._chequear_estado() # <--- VALIDACIÓN ANTES DE ACTUAR

        if self.modo_prueba:
            self._log(f"🔘 TAB (x{cantidad})")
        else:
            if cantidad > 5:
                original_pause = pyautogui.PAUSE
                pyautogui.PAUSE = 0.05 
                for _ in range(cantidad):
                    self._chequear_estado() # Chequeo incluso ENTRE tabs masivos
                    pyautogui.press('tab')
                pyautogui.PAUSE = original_pause
            else:
                pyautogui.press('tab', presses=cantidad)

    def _enter(self):
        self._chequear_estado() # <--- VALIDACIÓN ANTES DE ACTUAR
        if self.modo_prueba:
            self._log("🔘 ENTER")
        else:
            pyautogui.press('enter')
    
    def _esperar_carga(self, tiempo=1.0):
        """
        Detiene el robot X segundos para esperar que el sistema procese algo.
        """
        self._chequear_estado() # Verificar si hay pausa antes de esperar
        
        if self.modo_prueba:
            self._log(f"⏳ ESPERA TÉCNICA: {tiempo} segundos")
        else:
            time.sleep(tiempo)

    def procesar_producto(self, fila, indice):
        self._chequear_estado() # <--- VALIDACIÓN INICIAL
        
        self._log(f"--- PRODUCTO #{indice} ({fila.get('Descripcion', 'Sin Nombre')}) ---")

        # FASE 1
        self._esperar_carga(1)
        self._escribir(fila['Rubro'])
        self._escribir(fila['Clasificacion'])
        self._escribir(fila['Linea'])
        self._esperar_carga(1)
        self._tab(1) 

        # FASE 2
        self._esperar_carga(1)
        self._escribir(fila['CodigoBarras'])
        self._tab(1)
        self._esperar_carga(1)
        self._tab(1)
        self._esperar_carga(1)
        self._tab(1)
        self._esperar_carga(1)
        self._tab(1)
        self._esperar_carga(1)
        self._tab(1)
        self._esperar_carga(1)

        desc = str(fila['Descripcion'])
        if len(desc) > 25: desc = desc[:25]
        self._escribir(desc)
        self._tab(6)
        
        self._escribir(fila['Marca'])
        self._tab(6)

        # FASE 3
        self._escribir(fila['Peso'])
        self._tab(7)
        self._escribir(fila['Impuesto'])
        self._tab(4)

        # FASE 4
        self._escribir(fila['Compra'].upper())
        self._tab(1)
        self._escribir(fila['Venta'])
        self._tab(1)
        self._escribir(fila['Capacidad'])
        self._tab(3)
        self._escribir(fila['Embalaje_1'])
        self._tab(1)
        self._escribir(fila['Embalaje_2'])
        self._tab(12)

        # FASE 5
        self._escribir("15")
        self._tab(41)
        self._enter()
        self._esperar_carga(1)
        
        self._log(f"--- FIN PRODUCTO #{indice} ---\n")