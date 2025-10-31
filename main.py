import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import xml.etree.ElementTree as ET

# ==================== LÓGICA DE PROCESAMIENTO ====================
class XMLProcessor:
    def __init__(self, excel_path, xml_path):
        self.excel_path = excel_path
        self.xml_path = xml_path

    def process(self, progress_callback, log_callback):
        # Leer Excel
        try:
            df = pd.read_excel(self.excel_path)
        except Exception as e:
            log_callback(f"❌ Error al leer Excel: {e}")
            return

        if df is None or len(df) == 0:
            log_callback("❌ El Excel está vacío o no contiene filas.")
            return

        # Leer XML
        try:
            tree = ET.parse(self.xml_path)
        except Exception as e:
            log_callback(f"❌ Error al leer XML: {e}")
            return

        root = tree.getroot()
        total = len(df)

        for index, row in df.iterrows():
            nombre = row.get("nombre", "")
            apellido = row.get("apellido", "")
            numero = row.get("numero", "")
            contacto = None

            # Buscar contacto por número
            for c in root.findall("Contact"):
                phone = c.find("Phone")
                if phone is not None:
                    phonenumber = phone.find("phonenumber")
                    if phonenumber is not None and phonenumber.text == str(numero):
                        contacto = c
                        break

            progreso = f"({index + 1} de {total})"
            if contacto is not None:
                # Obtener valores anteriores (seguridad si no existen nodos)
                fn_node = contacto.find("FirstName")
                ln_node = contacto.find("LastName")
                old_first = (fn_node.text if fn_node is not None else "") or ""
                old_last = (ln_node.text if ln_node is not None else "") or ""

                # Asegurar que los nodos existan antes de asignar
                if fn_node is None:
                    fn_node = ET.SubElement(contacto, "FirstName")
                if ln_node is None:
                    ln_node = ET.SubElement(contacto, "LastName")

                fn_node.text = "" if pd.isna(nombre) else str(nombre)
                ln_node.text = "" if pd.isna(apellido) else str(apellido)

                msg = (
                    f"✅ Actualizado contacto {numero} {progreso}\n"
                    f"   Nombre: '{old_first}' → '{fn_node.text}'\n"
                    f"   Apellido: '{old_last}' → '{ln_node.text}'"
                )
            else:
                msg = f"⚠️ No se encontró contacto con número {numero} {progreso}"

            log_callback(msg)
            progress_callback((index + 1) / total * 100)

        # Guardar XML en la misma ruta proporcionada (sobrescribe)
        try:
            output_path = os.path.abspath(self.xml_path)
            tree.write(output_path, encoding="UTF-8", xml_declaration=True)
            log_callback(f"\n💾 Archivo guardado como: {output_path}")
        except Exception as e:
            log_callback(f"❌ Error al guardar XML: {e}")


# ==================== PANTALLA PRINCIPAL ====================
class MainFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, style="Card.TFrame")
        self.controller = controller
        self.configure(padding=25)

        ttk.Label(self, text="Actualizador de Contactos", style="Title.TLabel").pack(pady=(5, 20))

        self.excel_var = tk.StringVar()
        self.xml_var = tk.StringVar()

        # Campos de entrada
        ttk.Label(self, text="Archivo Excel (.xlsx):", style="SubTitle.TLabel").pack(anchor="w")
        ttk.Entry(self, textvariable=self.excel_var, width=60).pack(pady=3)
        ttk.Button(self, text="Seleccionar", style="Accent.TButton", command=self.select_excel).pack(pady=(0, 10))

        ttk.Label(self, text="Archivo XML (.xml):", style="SubTitle.TLabel").pack(anchor="w")
        ttk.Entry(self, textvariable=self.xml_var, width=60).pack(pady=3)
        ttk.Button(self, text="Seleccionar", style="Accent.TButton", command=self.select_xml).pack(pady=(0, 15))

        # Botones
        btn_frame = ttk.Frame(self, style="Card.TFrame")
        btn_frame.pack(pady=10)
        self.clear_btn = ttk.Button(btn_frame, text="🧹 Limpiar", style="Secondary.TButton", command=self.clear_paths)
        self.clear_btn.grid(row=0, column=0, padx=6)
        self.process_btn = ttk.Button(btn_frame, text="▶ Procesar", style="Accent.TButton", command=self.start_processing)
        self.process_btn.grid(row=0, column=1, padx=6)
        ttk.Button(btn_frame, text="❌ Cerrar", style="Danger.TButton", command=self.controller.quit).grid(row=0, column=2, padx=6)

    def select_excel(self):
        path = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=[("Excel", "*.xlsx"), ("Excel (xls)", "*.xls")])
        if path:
            self.excel_var.set(path)

    def select_xml(self):
        path = filedialog.askopenfilename(title="Seleccionar XML", filetypes=[("XML", "*.xml")])
        if path:
            self.xml_var.set(path)

    def clear_paths(self):
        self.excel_var.set("")
        self.xml_var.set("")
        messagebox.showinfo("Limpieza", "Las rutas se han borrado correctamente.")

    def start_processing(self):
        excel_path = self.excel_var.get().strip()
        xml_path = self.xml_var.get().strip()

        # Validaciones básicas de existencia
        if not excel_path or not xml_path:
            messagebox.showwarning("Advertencia", "Selecciona ambos archivos antes de continuar.")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"No se encontró el archivo Excel:\n{excel_path}")
            return
        if not os.path.exists(xml_path):
            messagebox.showerror("Error", f"No se encontró el archivo XML:\n{xml_path}")
            return

        # Desactivar botón procesar para evitar doble ejec.
        self.process_btn.state(["disabled"])
        self.controller.show_frame("ProcessFrame")
        # iniciar hilo de procesamiento
        self.controller.process_frame.start_process_thread(excel_path, xml_path, on_done=self.on_processing_done)

    def on_processing_done(self):
        # Reactivar botón procesar (se ejecuta desde hilo principal via .after)
        self.process_btn.state(["!disabled"])
        messagebox.showinfo("Listo", "Procesamiento finalizado.")


# ==================== PANTALLA DE PROCESAMIENTO ====================
class ProcessFrame(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, style="Card.TFrame")
        self.controller = controller
        self.configure(padding=25)

        ttk.Label(self, text="Procesando contactos...", style="Title.TLabel").pack(pady=(5, 15))

        self.text = tk.Text(self, width=80, height=20, state="disabled", wrap="word", bg="#f1f3f8", relief="flat")
        self.text.pack(pady=10)

        self.progress = ttk.Progressbar(self, length=500, mode="determinate", style="Accent.Horizontal.TProgressbar")
        self.progress.pack(pady=10)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="⬅ Atrás", style="Secondary.TButton", command=lambda: controller.show_frame("MainFrame")).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="❌ Cerrar", style="Danger.TButton", command=controller.quit).grid(row=0, column=1, padx=6)

    def log(self, msg):
        # Insertar sin bloquear UI
        def _append():
            self.text.config(state="normal")
            self.text.insert("end", msg + "\n")
            self.text.config(state="disabled")
            self.text.see("end")
        self.after(0, _append)

    def update_progress(self, value):
        def _upd():
            self.progress["value"] = value
            self.update_idletasks()
        self.after(0, _upd)

    def start_process(self, excel_path, xml_path, on_done=None):
        processor = XMLProcessor(excel_path, xml_path)
        processor.process(self.update_progress, self.log)
        # callback when done: run on main thread
        if on_done:
            self.after(0, on_done)

    def start_process_thread(self, excel_path, xml_path, on_done=None):
        # limpiar log y progress antes de iniciar
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.config(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self.start_process, args=(excel_path, xml_path, on_done), daemon=True).start()


# ==================== APP PRINCIPAL ====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Actualizador de Contactos AMPYME")
        self.configure(bg="#cfe2ff")
        self.resizable(False, False)

        width, height = 760, 620
        x = int((self.winfo_screenwidth() - width) / 2)
        y = int((self.winfo_screenheight() - height) / 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 14), background="#ffffff", foreground="#2b2d42")
        style.configure("SubTitle.TLabel", font=("Segoe UI", 10), background="#ffffff", foreground="#495057")
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
        style.configure("Accent.TButton", background="#0078d7", foreground="white")
        style.map("Accent.TButton", background=[("active", "#005a9e")])
        style.configure("Danger.TButton", background="#e63946", foreground="white")
        style.map("Danger.TButton", background=[("active", "#c71c28")])
        style.configure("Secondary.TButton", background="#adb5bd", foreground="white")
        style.map("Secondary.TButton", background=[("active", "#868e96")])
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#dee2e6", background="#0078d7")

        container = ttk.Frame(self, style="Card.TFrame")
        container.pack(fill="both", expand=True)

        self.main_frame = MainFrame(container, self)
        self.process_frame = ProcessFrame(container, self)

        for frame in (self.main_frame, self.process_frame):
            frame.grid(row=0, column=0, sticky="nsew")

        self.frames = {"MainFrame": self.main_frame, "ProcessFrame": self.process_frame}

        self.show_frame("MainFrame")

    def show_frame(self, name):
        frame = self.frames[name]
        frame.tkraise()


if __name__ == "__main__":
    app = App()
    app.mainloop()
