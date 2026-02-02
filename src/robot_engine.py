import pyautogui
import time
import os
from datetime import datetime

class RobotFormulario:
    def __init__(self, modo_prueba=True, velocidad=0.5, log_path="logs/auditoria.txt"):
        """
        Inicializa el robot.
        :param modo_prueba: Si es True, NO interactúa con el PC, solo escribe logs.
        :param velocidad: Tiempo de espera entre acciones (segundos).
        :param log_path: Ruta donde guardar el historial de pruebas.
        """
        self.modo_prueba = modo_prueba
        self.velocidad = velocidad
        self.log_path = log_path
        
        # Crear carpeta de logs si no existe
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        # Configuración de seguridad para modo real
        if not self.modo_prueba:
            pyautogui.PAUSE = self.velocidad
            pyautogui.FAILSAFE = True

    def _log(self, mensaje):
        """Escribe en el archivo de texto y en la consola."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        linea = f"[{timestamp}] {mensaje}"
        print(linea)
        
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(linea + "\n")

    def _escribir(self, texto):
        """Escribe el contenido de una celda del Excel."""
        val = str(texto).strip()
        if val.lower() == 'nan' or val == '':
            return 

        if self.modo_prueba:
            self._log(f"⌨️  ESCRIBIR: '{val}'")
        else:                                                               
                   
            # CORRECCIÓN AQUÍ:
            # interval=0.05 significa "espera 50 milisegundos entre cada letra"
            # Esto elimina el efecto "ARRRROZZ"
            pyautogui.write(val, interval=0.08)

    def _presionar(self, tecla, veces=1):
        """Presiona una tecla N veces."""
        if self.modo_prueba:
            if veces > 1:
                self._log(f"🔘 PRESIONAR: '{tecla}' (x{veces} veces)")
            else:
                self._log(f"🔘 PRESIONAR: '{tecla}'")
        else:
            # En modo real, a veces es mejor un loop pequeño para asegurar estabilidad
            if tecla == 'tab' and veces > 10:
                # Optimización para saltos largos (como el de 41 tabs)
                for _ in range(veces):
                    pyautogui.press(tecla)
            else:
                pyautogui.press(tecla, presses=veces)

    def _esperar(self, segundos):
        """Pausa explicita."""
        if self.modo_prueba:
            self._log(f"⏳ ESPERA TÉCNICA: {segundos} seg")
        else:
            time.sleep(segundos)

    def procesar_producto(self, datos, indice):
        """
        Ejecuta el flujo completo para UN producto (una fila del Excel).
        :param datos: Diccionario o Serie de Pandas con los datos de la fila.
        :param indice: Número de fila para referencia.
        """
        self._log(f"--- INICIANDO PRODUCTO #{indice} ---")

        # --- FASE 1: Identificación y Clasificación ---
        # Asumo nombres de columnas del Excel basados en tu descripción
        self._escribir(datos.get('Rubro', ''))
        self._presionar('enter') # Usualmente enter o tab para pasar campo? Asumiré TAB por defecto abajo si no
        self._presionar('tab') 
        
        self._escribir(datos.get('Clasificacion', ''))
        self._presionar('tab')

        self._escribir(datos.get('Linea', ''))
        # Nota: El correlativo se autocompleta. Damos un respiro al sistema.
        self._esperar(1.0) 
        self._presionar('tab', veces=1) # Acción Fase 1 final

        # --- FASE 2: Datos de Identificación ---
        self._escribir(datos.get('CodigoBarras', '')) # Campo 5
        self._presionar('tab', veces=5) # Acción 6

        # Campo 7: Descripción con Validación (Max 25 chars)
        descripcion = str(datos.get('Descripcion', ''))
        if len(descripcion) > 25:
            self._log(f"⚠️ AVISO: Descripción truncada (era {len(descripcion)} chars)")
            descripcion = descripcion[:25]
        self._escribir(descripcion)
        
        self._presionar('tab', veces=6) # Acción 8

        self._escribir(datos.get('Marca', '')) # Campo 9
        self._presionar('tab', veces=5) # Acción 10

        # --- FASE 3: Propiedades Físicas y Tributarias ---
        self._escribir(datos.get('Peso', '')) # Campo 11
        self._presionar('tab', veces=6) # Acción 12

        # Campo 13: Impuesto (Condicional)
        impuesto = datos.get('Impuesto', '')
        # Pandas a veces pone 'nan' si está vacío, validamos eso
        if impuesto and str(impuesto).lower() != 'nan': 
            self._escribir(impuesto)
        else:
            self._log("ℹ️ Impuesto vacío en Excel, saltando escritura.")
        
        self._presionar('tab', veces=4) # Acción 14

        # --- FASE 4: Unidades y Logística ---
        self._escribir(datos.get('Compra', '')) # Campo 15 (KG, CJ, UN)
        self._presionar('tab') # Asumo tab entre compra y venta
        self._escribir(datos.get('Venta', ''))  # Campo 16
        self._presionar('tab', veces=1) # Acción 17

        self._escribir(datos.get('Capacidad', '')) # Campo 18
        self._presionar('tab', veces=2) # Acción 19

        self._escribir(datos.get('Embalaje', '')) # Campo 20
        self._presionar('tab', veces=12) # Acción 21

        # --- FASE 5: Finalización ---
        # Campo 22: Bruto (Valor fijo 15)
        self._escribir("15") 
        
        # Acción final masiva
        self._presionar('tab', veces=41) 
        self._presionar('enter')
        
        self._log(f"--- PRODUCTO #{indice} FINALIZADO ---\n")