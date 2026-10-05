import csv
import datetime
import os
import shutil
import sqlite3
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


class MesaPartesApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Mesa de Partes Digital")
        self.geometry("1050x680")
        self.minsize(950, 600)
        self.configure(bg="#f8fafc")

        #Configuración de Archivos y Base de Datos
        self.DIR_BASE = os.getcwd()
        self.DIR_ADJUNTOS = os.path.join(self.DIR_BASE, "adjuntos")
        self.DB_PATH = os.path.join(self.DIR_BASE, "mesapartes.db")
        os.makedirs(self.DIR_ADJUNTOS, exist_ok=True)

        self.ruta_adjunto_seleccionado = ""
        self.estados_lista = ["Pendiente", "En Proceso", "Atendido"]
        self.tipos_doc = [
            "Factura",
            "Guía de Remisión",
            "Nota de Crédito",
            "Solicitud",
            "Oficio / Carta",
        ]
        self.areas_destino = [
            "Gerencia",
            "Administración",
            "Ventas",
            "Logística",
            "Contabilidad",
        ]

        #Inicializar base de datos y cargar datos
        self._init_db()

        self._setup_styles()
        self._build_ui()
        self._cargar_datos_db()

    #Gestion de datos
    def _init_db(self):
        """Crea la tabla de expedientes e incluye el campo de estado."""
        with sqlite3.connect(self.DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS expedientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo TEXT UNIQUE NOT NULL,
                    fecha TEXT NOT NULL,
                    doc_id TEXT NOT NULL,
                    remitente TEXT NOT NULL,
                    tipo TEXT NOT NULL,
                    num_doc TEXT,
                    area TEXT NOT NULL,
                    folios INTEGER NOT NULL,
                    asunto TEXT NOT NULL,
                    archivo_nombre TEXT,
                    archivo_ruta TEXT,
                    estado TEXT DEFAULT 'Pendiente'
                )
            """)
            
            cursor.execute("PRAGMA table_info(expedientes)")
            columnas = [column[1] for column in cursor.fetchall()]
            if "estado" not in columnas:
                cursor.execute("ALTER TABLE expedientes ADD COLUMN estado TEXT DEFAULT 'Pendiente'")

            conn.commit()

    def _obtener_siguiente_secuencia(self, fecha_hoy_str):
        with sqlite3.connect(self.DB_PATH) as conn:
            cursor = conn.cursor()
            patron = f"EXP-{fecha_hoy_str}-%"
            cursor.execute("SELECT COUNT(*) FROM expedientes WHERE codigo LIKE ?", (patron,))
            count = cursor.fetchone()[0]
            return count + 1

    def _cargar_datos_db(self):
        self.expedientes = []
        with sqlite3.connect(self.DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM expedientes ORDER BY id DESC")
            rows = cursor.fetchall()
            for row in rows:
                self.expedientes.append(dict(row))

        self._actualizar_tabla()
        self._actualizar_reporte()

    #Estilos e interfaz
    def _setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        BG = "#f8fafc"
        CARD_BG = "#ffffff"
        TEXT = "#0f172a"
        TEXT_MUTED = "#64748b"
        PRIMARY = "#0f172a"
        SUCCESS = "#15803d"

        self.style.configure(".", background=BG, foreground=TEXT, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=BG)
        self.style.configure("Card.TFrame", background=CARD_BG, relief="solid", borderwidth=1)

        self.style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"), foreground=TEXT)
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 9), foreground=TEXT_MUTED)
        self.style.configure("FormLabel.TLabel", font=("Segoe UI", 9, "bold"), foreground="#334155")

        self.style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background=PRIMARY, foreground="#ffffff", padding=(12, 6))
        self.style.map("Primary.TButton", background=[("active", "#334155")])

        self.style.configure("Success.TButton", font=("Segoe UI", 9, "bold"), background=SUCCESS, foreground="#ffffff", padding=(10, 5))
        self.style.map("Success.TButton", background=[("active", "#166534")])

        self.style.configure("TNotebook", background=BG, borderwidth=0)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 9, "bold"), padding=[12, 6], background="#e2e8f0", foreground=TEXT_MUTED)
        self.style.map("TNotebook.Tab", background=[("selected", CARD_BG)], foreground=[("selected", TEXT)])

        self.style.configure("Treeview", rowheight=26, font=("Segoe UI", 9), background=CARD_BG, fieldbackground=CARD_BG)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#f1f5f9", foreground=TEXT)

    def _build_ui(self):
        header = ttk.Frame(self, padding=(20, 15, 20, 10))
        header.pack(fill="x")
        ttk.Label(header, text="Mesa de Partes Digital", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="Distribuidora Fabri S.A.C.", style="Subtitle.TLabel").pack(anchor="w")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.tab_registro = ttk.Frame(self.notebook, padding=15)
        self.tab_busqueda = ttk.Frame(self.notebook, padding=15)
        self.tab_reporte = ttk.Frame(self.notebook, padding=15)

        self.notebook.add(self.tab_registro, text="Registrar")
        self.notebook.add(self.tab_busqueda, text="Expedientes")
        self.notebook.add(self.tab_reporte, text="Métricas y Exportación")

        self._build_tab_registro()
        self._build_tab_busqueda()
        self._build_tab_reporte()

    #Registro
    def _build_tab_registro(self):
        card = ttk.Frame(self.tab_registro, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True)

        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)

        fields = [
            ("DNI / RUC Remitente:*", "entry_doc_id", 0, 0),
            ("Nombre / Razón Social:*", "entry_remitente", 0, 2),
            ("Tipo de Documento:*", "combo_tipo", 1, 0),
            ("N° Documento Emitido:", "entry_num_doc", 1, 2),
            ("Área Destino:*", "combo_area", 2, 0),
            ("Número de Folios:*", "entry_folios", 2, 2),
            ("Asunto / Descripción:*", "entry_asunto", 3, 0),
        ]

        for label_text, var_name, r, c in fields:
            ttk.Label(card, text=label_text, style="FormLabel.TLabel").grid(row=r*2, column=c, sticky="w", padx=10, pady=(10, 2))
            
            if "combo" in var_name:
                vals = self.tipos_doc if "tipo" in var_name else self.areas_destino
                widget = ttk.Combobox(card, values=vals, state="readonly")
                widget.current(0)
            else:
                widget = ttk.Entry(card)

            if var_name == "entry_asunto":
                widget.grid(row=r*2+1, column=c, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
            else:
                widget.grid(row=r*2+1, column=c, sticky="ew", padx=10, pady=(0, 10))

            setattr(self, var_name, widget)

        ttk.Label(card, text="Documento Adjunto (PDF, DOCX, Imagen):", style="FormLabel.TLabel").grid(row=8, column=0, sticky="w", padx=10, pady=(10, 2))
        
        frame_adjunto = ttk.Frame(card)
        frame_adjunto.grid(row=9, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        
        self.lbl_adjunto_path = ttk.Label(frame_adjunto, text="Ningún archivo seleccionado", foreground="#64748b")
        self.lbl_adjunto_path.pack(side="left", fill="x", expand=True)

        btn_seleccionar = ttk.Button(frame_adjunto, text="Seleccionar Archivo...", command=self._seleccionar_archivo)
        btn_seleccionar.pack(side="right")

        btn_guardar = ttk.Button(card, text="Registrar Expediente", style="Primary.TButton", command=self._registrar)
        btn_guardar.grid(row=10, column=3, sticky="e", padx=10, pady=15)

    def _seleccionar_archivo(self):
        tipos = [("Todos los archivos", "*.*"), ("PDF", "*.pdf"), ("Imágenes", "*.png;*.jpg;*.jpeg"), ("Word", "*.docx")]
        path = filedialog.askopenfilename(title="Seleccionar documento adjunto", filetypes=tipos)
        if path:
            self.ruta_adjunto_seleccionado = path
            self.lbl_adjunto_path.config(text=os.path.basename(path), foreground="#0f172a")

    def _registrar(self):
        doc_id = self.entry_doc_id.get().strip()
        remitente = self.entry_remitente.get().strip()
        asunto = self.entry_asunto.get().strip()
        folios_str = self.entry_folios.get().strip()

        if not all([doc_id, remitente, asunto, folios_str]):
            messagebox.showwarning("Atención", "Complete todos los campos obligatorios (*).")
            return

        if not folios_str.isdigit():
            messagebox.showwarning("Atención", "El campo folios debe ser un número entero.")
            return

        now = datetime.datetime.now()
        fecha_hoy_str = now.strftime('%Y%m%d')
        secuencia = self._obtener_siguiente_secuencia(fecha_hoy_str)
        codigo = f"EXP-{fecha_hoy_str}-{secuencia:04d}"
        
        ruta_guardada = ""
        nombre_archivo = "Sin adjunto"
        if self.ruta_adjunto_seleccionado:
            ext = os.path.splitext(self.ruta_adjunto_seleccionado)[1]
            nombre_archivo = f"{codigo}_adjunto{ext}"
            ruta_guardada = os.path.join(self.DIR_ADJUNTOS, nombre_archivo)
            shutil.copy(self.ruta_adjunto_seleccionado, ruta_guardada)

        fecha_formateada = now.strftime("%d/%m/%Y %H:%M")
        num_doc = self.entry_num_doc.get().strip() or "S/N"
        tipo = self.combo_tipo.get()
        area = self.combo_area.get()
        folios = int(folios_str)
        estado_inicial = "Pendiente"

        try:
            with sqlite3.connect(self.DB_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO expedientes (codigo, fecha, doc_id, remitente, tipo, num_doc, area, folios, asunto, archivo_nombre, archivo_ruta, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (codigo, fecha_formateada, doc_id, remitente, tipo, num_doc, area, folios, asunto, nombre_archivo, ruta_guardada, estado_inicial))
                conn.commit()

            messagebox.showinfo("Éxito", f"Expediente registrado con éxito:\n\nCódigo: {codigo}")
            self._limpiar_formulario()
            self._cargar_datos_db()

        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo guardar el expediente:\n{e}")

    def _limpiar_formulario(self):
        self.entry_doc_id.delete(0, tk.END)
        self.entry_remitente.delete(0, tk.END)
        self.entry_num_doc.delete(0, tk.END)
        self.entry_folios.delete(0, tk.END)
        self.entry_asunto.delete(0, tk.END)
        self.combo_tipo.current(0)
        self.combo_area.current(0)
        self.ruta_adjunto_seleccionado = ""
        self.lbl_adjunto_path.config(text="Ningún archivo seleccionado", foreground="#64748b")

    #Consultas, tablas y estados
    def _build_tab_busqueda(self):
        top_frame = ttk.Frame(self.tab_busqueda)
        top_frame.pack(fill="x", pady=(0, 10))

        #Buscador
        ttk.Label(top_frame, text="Buscar:").pack(side="left", padx=(0, 5))
        self.entry_buscar = ttk.Entry(top_frame, width=25)
        self.entry_buscar.pack(side="left", padx=5)
        self.entry_buscar.bind("<KeyRelease>", lambda e: self._filtrar_tabla())

        #Botón ver adjunto
        btn_abrir_adjunto = ttk.Button(top_frame, text="Ver Adjunto", command=self._abrir_adjunto_seleccionado)
        btn_abrir_adjunto.pack(side="right", padx=(10, 0))

        #Controles para actualizar el estado
        btn_cambiar_estado = ttk.Button(top_frame, text="Cambiar Estado", command=self._cambiar_estado)
        btn_cambiar_estado.pack(side="right", padx=(5, 0))

        self.combo_nuevo_estado = ttk.Combobox(top_frame, values=self.estados_lista, state="readonly", width=12)
        self.combo_nuevo_estado.current(0)
        self.combo_nuevo_estado.pack(side="right", padx=5)

        ttk.Label(top_frame, text="Nuevo Estado:").pack(side="right", padx=(10, 0))

        cols = ("codigo", "fecha", "remitente", "tipo", "area", "folios", "estado", "asunto", "adjunto")
        self.tabla = ttk.Treeview(self.tab_busqueda, columns=cols, show="headings", selectmode="browse")

        headers = [
            ("codigo", "Código", 120),
            ("fecha", "Fecha/Hora", 110),
            ("remitente", "Remitente", 140),
            ("tipo", "Tipo Doc.", 85),
            ("area", "Área Destino", 95),
            ("folios", "Folios", 50),
            ("estado", "Estado", 90),
            ("asunto", "Asunto", 160),
            ("adjunto", "Adjunto", 110)
        ]

        for col, title, width in headers:
            self.tabla.heading(col, text=title)
            self.tabla.column(col, width=width, anchor="w" if col in ["remitente", "asunto"] else "center")

        scroll = ttk.Scrollbar(self.tab_busqueda, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscroll=scroll.set)

        self.tabla.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _actualizar_tabla(self, datos=None):
        for row in self.tabla.get_children():
            self.tabla.delete(row)

        dataset = datos if datos is not None else self.expedientes
        for item in dataset:
            self.tabla.insert("", "end", values=(
                item["codigo"],
                item["fecha"],
                f"{item['remitente']} ({item['doc_id']})",
                item["tipo"],
                item["area"],
                item["folios"],
                item.get("estado", "Pendiente"),
                item["asunto"],
                item["archivo_nombre"]
            ))

    def _filtrar_tabla(self):
        query = self.entry_buscar.get().strip().lower()
        if not query:
            self._actualizar_tabla()
            return

        filtrados = [
            exp for exp in self.expedientes
            if query in exp["codigo"].lower() or query in exp["remitente"].lower() or query in exp["doc_id"].lower()
        ]
        self._actualizar_tabla(filtrados)

    def _cambiar_estado(self):
        selected_item = self.tabla.selection()
        if not selected_item:
            messagebox.showwarning("Atención", "Seleccione un expediente de la tabla para cambiar su estado.")
            return

        item_values = self.tabla.item(selected_item)["values"]
        codigo_exp = item_values[0]
        nuevo_estado = self.combo_nuevo_estado.get()

        try:
            with sqlite3.connect(self.DB_PATH) as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE expedientes SET estado = ? WHERE codigo = ?", (nuevo_estado, codigo_exp))
                conn.commit()

            messagebox.showinfo("Éxito", f"Estado del expediente {codigo_exp} actualizado a: {nuevo_estado}")
            self._cargar_datos_db()

        except sqlite3.Error as e:
            messagebox.showerror("Error", f"No se pudo actualizar el estado:\n{e}")

    def _abrir_adjunto_seleccionado(self):
        selected_item = self.tabla.selection()
        if not selected_item:
            messagebox.showwarning("Atención", "Seleccione un expediente de la tabla para ver su adjunto.")
            return

        item_values = self.tabla.item(selected_item)["values"]
        codigo_exp = item_values[0]

        exp = next((e for e in self.expedientes if e["codigo"] == codigo_exp), None)
        if exp and exp["archivo_ruta"] and os.path.exists(exp["archivo_ruta"]):
            if sys.platform == "win32":
                os.startfile(exp["archivo_ruta"])
            elif sys.platform == "darwin":
                subprocess.call(["open", exp["archivo_ruta"]])
            else:
                subprocess.call(["xdg-open", exp["archivo_ruta"]])
        else:
            messagebox.showinfo("Información", "El expediente seleccionado no posee un archivo adjunto accesible.")

    #Metricas y exportacion
    def _build_tab_reporte(self):
        card = ttk.Frame(self.tab_reporte, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True)

        self.lbl_total = ttk.Label(card, text="Total Registrados en DB: 0", style="Title.TLabel")
        self.lbl_total.pack(anchor="w", pady=(0, 15))

        #Métricas por estado
        ttk.Label(card, text="Resumen por Estado", style="FormLabel.TLabel").pack(anchor="w", pady=(5, 5))
        self.metricas_estado_labels = {}
        for est in self.estados_lista:
            frame_linea = ttk.Frame(card)
            frame_linea.pack(fill="x", pady=2)
            ttk.Label(frame_linea, text=f"{est}:", width=18, font=("Segoe UI", 9, "bold")).pack(side="left")
            lbl_val = ttk.Label(frame_linea, text="0 expedientes")
            lbl_val.pack(side="left")
            self.metricas_estado_labels[est] = lbl_val

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=15)

        #Métricas por área
        ttk.Label(card, text="Documentos por Área Destino", style="FormLabel.TLabel").pack(anchor="w", pady=(5, 5))
        self.metricas_labels = {}
        for area in self.areas_destino:
            frame_linea = ttk.Frame(card)
            frame_linea.pack(fill="x", pady=2)
            ttk.Label(frame_linea, text=f"{area}:", width=18, font=("Segoe UI", 9, "bold")).pack(side="left")
            lbl_val = ttk.Label(frame_linea, text="0 expedientes")
            lbl_val.pack(side="left")
            self.metricas_labels[area] = lbl_val

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=15)
        ttk.Label(card, text="Exportar Información", style="FormLabel.TLabel").pack(anchor="w", pady=(0, 10))

        btn_exportar = ttk.Button(card, text="Exportar Expedientes a Excel", style="Success.TButton", command=self._exportar_excel)
        btn_exportar.pack(anchor="w")

    def _actualizar_reporte(self):
        total = len(self.expedientes)
        self.lbl_total.config(text=f"Total Registrados en DB: {total}")

        #Conteo de estados
        conteo_est = {est: 0 for est in self.estados_lista}
        for exp in self.expedientes:
            est = exp.get("estado", "Pendiente")
            if est in conteo_est:
                conteo_est[est] += 1
        for est, count in conteo_est.items():
            self.metricas_estado_labels[est].config(text=f"{count} expedientes")

        #Conteo de áreas
        conteo_area = {area: 0 for area in self.areas_destino}
        for exp in self.expedientes:
            if exp["area"] in conteo_area:
                conteo_area[exp["area"]] += 1
        for area, count in conteo_area.items():
            self.metricas_labels[area].config(text=f"{count} expedientes")

    def _exportar_excel(self):
        if not self.expedientes:
            messagebox.showwarning("Atención", "No hay expedientes en la base de datos para exportar.")
            return

        if HAS_PANDAS:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if path:
                df = pd.DataFrame(self.expedientes)
                if "id" in df.columns:
                    df = df.drop(columns=["id"])
                df.to_excel(path, index=False)
                messagebox.showinfo("Éxito", f"Datos exportados correctamente a:\n{path}")
        else:
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV (Excel)", "*.csv")])
            if path:
                campos = ["codigo", "fecha", "doc_id", "remitente", "tipo", "num_doc", "area", "folios", "estado", "asunto", "archivo_nombre", "archivo_ruta"]
                with open(path, mode="w", newline="", encoding="utf-8-sig") as file:
                    writer = csv.DictWriter(file, fieldnames=campos, extrasaction="ignore")
                    writer.writeheader()
                    writer.writerows(self.expedientes)
                messagebox.showinfo("Éxito", f"Datos exportados como CSV compatible con Excel en:\n{path}")


if __name__ == "__main__":
    app = MesaPartesApp()
    app.mainloop()