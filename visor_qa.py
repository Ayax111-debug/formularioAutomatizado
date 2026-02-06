import tkinter as tk
from tkinter import font

class VisorQA:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ENTORNO DE PRUEBAS QA - DUMMY TARGET")
        self.root.geometry("800x600")
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#1e1e1e")

        self.custom_font = font.Font(family="Consolas", size=14)

        tk.Label(self.root, text="MODO AUTOMÁTICO ACTIVO", 
                 bg="#1e1e1e", fg="#00ff00", font=("Arial", 16, "bold")).pack(pady=10)

        self.text_area = tk.Text(self.root, height=20, width=80, 
                                 bg="black", fg="white", font=self.custom_font,
                                 insertbackground="white")
        self.text_area.pack(padx=20, pady=10)
        
        # Escuchar teclas
        self.text_area.bind("<Key>", self.registrar_evento)
        self.text_area.focus_set()

    def registrar_evento(self, event):
        tecla = event.keysym
        
        # Ignorar teclas de sistema aisladas
        if tecla in ["Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"]:
            return "break"

        # --- LOGICA DE TECLAS ESPECIALES ---

        if tecla == "Tab":
            self.text_area.insert(tk.END, " [TAB] ", "tag_tab")
            self.text_area.tag_config("tag_tab", foreground="#FFFF00") # Amarillo
            
        elif tecla == "Return":
            self.text_area.insert(tk.END, "\n [ENTER] \n", "tag_enter")
            self.text_area.tag_config("tag_enter", foreground="#FF5555") # Rojo

        # --- AQUÍ ESTÁN LAS FLECHAS Y BORRADO ---
        elif tecla == "Right":
            self.text_area.insert(tk.END, " [→] ", "tag_arrow")
            self.text_area.tag_config("tag_arrow", foreground="#00FFFF") # Cyan
            
        elif tecla == "Left":
            self.text_area.insert(tk.END, " [←] ", "tag_arrow")
            self.text_area.tag_config("tag_arrow", foreground="#00FFFF")

        elif tecla == "Delete":
            self.text_area.insert(tk.END, " [DEL] ", "tag_del")
            self.text_area.tag_config("tag_del", foreground="#FF00FF") # Magenta

        elif tecla == "BackSpace":
            self.text_area.insert(tk.END, " [BACK] ", "tag_del")
            self.text_area.tag_config("tag_del", foreground="#FF00FF")

        else:
            # Solo insertar si es un caracter imprimible
            if event.char and len(event.char) > 0:
                 self.text_area.insert(tk.END, event.char)
            
        self.text_area.see(tk.END)
        
        return "break" 

if __name__ == "__main__":
    app = VisorQA()
    app.root.mainloop()