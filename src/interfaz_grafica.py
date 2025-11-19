#!/usr/bin/env python3
"""
Interfaz gráfica para automatización de documentos Word
Permite seleccionar plantillas, rellenar campos y generar documentos
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
from datetime import datetime
import json

# Importar el procesador de plantillas
from procesador_plantillas import ProcesadorPlantillas, generar_nombre_archivo


class AplicacionAutomatizacion(tk.Tk):
    """Aplicación principal de automatización de documentos"""

    def __init__(self):
        super().__init__()

        self.title("Automatizador de Documentos Word")
        self.geometry("900x700")
        self.minsize(800, 600)

        # Variables
        self.ruta_plantilla = tk.StringVar()
        self.campos_entrada = {}
        self.procesador = None
        self.marcadores = []

        # Colores y estilos
        self.config(bg="#f0f0f0")
        self.estilo = ttk.Style()
        self.estilo.theme_use('clam')

        # Configurar estilos personalizados
        self.estilo.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        self.estilo.configure('Header.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        self.estilo.configure('Info.TLabel', font=('Arial', 10), foreground='#7f8c8d')
        self.estilo.configure('Accion.TButton', font=('Arial', 10, 'bold'), padding=10)

        # Crear interfaz
        self.crear_interfaz()

        # Centrar ventana
        self.centrar_ventana()

    def centrar_ventana(self):
        """Centra la ventana en la pantalla"""
        self.update_idletasks()
        ancho = self.winfo_width()
        alto = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.winfo_screenheight() // 2) - (alto // 2)
        self.geometry(f'{ancho}x{alto}+{x}+{y}')

    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""

        # Frame principal con scroll
        main_frame = ttk.Frame(self, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Título
        titulo = ttk.Label(main_frame, text="Automatizador de Documentos Word",
                          style='Title.TLabel')
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Sección 1: Selección de plantilla
        self.crear_seccion_plantilla(main_frame, row=1)

        # Sección 2: Campos del formulario
        self.frame_campos = ttk.LabelFrame(main_frame, text="Datos del Documento",
                                           padding="15")
        self.frame_campos.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S),
                              pady=(20, 0))

        # Mensaje inicial
        self.label_info = ttk.Label(self.frame_campos,
                                    text="Seleccione una plantilla para comenzar",
                                    style='Info.TLabel')
        self.label_info.grid(row=0, column=0, pady=40)

        # Sección 3: Botones de acción
        self.crear_seccion_acciones(main_frame, row=3)

        # Configurar expansión
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

    def crear_seccion_plantilla(self, parent, row):
        """Crea la sección de selección de plantilla"""
        frame = ttk.LabelFrame(parent, text="1. Seleccionar Plantilla", padding="15")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # Campo de ruta
        ttk.Label(frame, text="Plantilla:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        entry = ttk.Entry(frame, textvariable=self.ruta_plantilla, width=60, state='readonly')
        entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        btn_examinar = ttk.Button(frame, text="Examinar...", command=self.seleccionar_plantilla)
        btn_examinar.grid(row=0, column=2)

        btn_cargar = ttk.Button(frame, text="Cargar Campos", command=self.cargar_campos_plantilla,
                               style='Accion.TButton')
        btn_cargar.grid(row=0, column=3, padx=(10, 0))

        frame.columnconfigure(1, weight=1)

    def crear_seccion_acciones(self, parent, row):
        """Crea la sección de botones de acción"""
        frame = ttk.Frame(parent, padding="15")
        frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 0))

        # Botones
        btn_generar = ttk.Button(frame, text="Generar Documento",
                                command=self.generar_documento,
                                style='Accion.TButton')
        btn_generar.pack(side=tk.LEFT, padx=(0, 10))

        btn_limpiar = ttk.Button(frame, text="Limpiar Campos",
                                command=self.limpiar_campos)
        btn_limpiar.pack(side=tk.LEFT, padx=(0, 10))

        btn_guardar_config = ttk.Button(frame, text="Guardar Configuración",
                                       command=self.guardar_configuracion)
        btn_guardar_config.pack(side=tk.LEFT, padx=(0, 10))

        btn_cargar_config = ttk.Button(frame, text="Cargar Configuración",
                                      command=self.cargar_configuracion)
        btn_cargar_config.pack(side=tk.LEFT)

    def seleccionar_plantilla(self):
        """Abre diálogo para seleccionar archivo de plantilla"""
        ruta = filedialog.askopenfilename(
            title="Seleccionar Plantilla Word",
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")],
            initialdir="templates"
        )

        if ruta:
            self.ruta_plantilla.set(ruta)

    def cargar_campos_plantilla(self):
        """Carga los campos de la plantilla seleccionada"""
        ruta = self.ruta_plantilla.get()

        if not ruta:
            messagebox.showwarning("Advertencia", "Por favor seleccione una plantilla primero")
            return

        try:
            # Cargar plantilla
            self.procesador = ProcesadorPlantillas(ruta)
            self.procesador.cargar_documento()

            # Extraer marcadores
            self.marcadores = sorted(self.procesador.extraer_marcadores())

            if not self.marcadores:
                messagebox.showinfo("Información",
                                  "No se encontraron marcadores en la plantilla.\n\n"
                                  "Asegúrese de que la plantilla contenga marcadores en formato {{CAMPO}}")
                return

            # Limpiar frame de campos
            for widget in self.frame_campos.winfo_children():
                widget.destroy()

            self.campos_entrada.clear()

            # Crear campos dinámicamente
            ttk.Label(self.frame_campos, text="Complete los siguientes campos:",
                     style='Header.TLabel').grid(row=0, column=0, columnspan=2,
                                                 sticky=tk.W, pady=(0, 15))

            # Canvas con scrollbar para muchos campos
            canvas = tk.Canvas(self.frame_campos, bg='white', highlightthickness=0)
            scrollbar = ttk.Scrollbar(self.frame_campos, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Crear campos de entrada para cada marcador
            for idx, marcador in enumerate(self.marcadores):
                # Etiqueta
                label_text = marcador.replace('_', ' ').title()
                ttk.Label(scrollable_frame, text=f"{label_text}:",
                         font=('Arial', 10)).grid(row=idx, column=0,
                                                 sticky=tk.W, pady=8, padx=(10, 20))

                # Campo de entrada
                entrada = ttk.Entry(scrollable_frame, width=50, font=('Arial', 10))
                entrada.grid(row=idx, column=1, sticky=(tk.W, tk.E), pady=8, padx=(0, 10))

                self.campos_entrada[marcador] = entrada

            # Posicionar canvas y scrollbar
            canvas.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
            scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S), pady=(0, 10))

            self.frame_campos.columnconfigure(0, weight=1)
            self.frame_campos.rowconfigure(1, weight=1)

            messagebox.showinfo("Éxito",
                              f"Se cargaron {len(self.marcadores)} campos de la plantilla")

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar plantilla:\n{str(e)}")

    def generar_documento(self):
        """Genera el documento Word con los datos ingresados"""
        if not self.procesador:
            messagebox.showwarning("Advertencia",
                                 "Por favor cargue una plantilla primero")
            return

        # Recopilar datos de los campos
        datos = {}
        campos_vacios = []

        for marcador, entrada in self.campos_entrada.items():
            valor = entrada.get().strip()
            if valor:
                datos[marcador] = valor
            else:
                campos_vacios.append(marcador.replace('_', ' ').title())

        # Advertir sobre campos vacíos
        if campos_vacios:
            respuesta = messagebox.askyesno("Campos Vacíos",
                                           f"Los siguientes campos están vacíos:\n\n" +
                                           "\n".join(f"- {campo}" for campo in campos_vacios) +
                                           "\n\n¿Desea continuar de todas formas?")
            if not respuesta:
                return

        # Solicitar nombre y ubicación del archivo
        nombre_default = generar_nombre_archivo("documento_generado")
        ruta_salida = filedialog.asksaveasfilename(
            title="Guardar Documento Como",
            defaultextension=".docx",
            initialfile=nombre_default,
            initialdir="output",
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )

        if not ruta_salida:
            return

        try:
            # Procesar documento
            if self.procesador.procesar(datos, ruta_salida):
                messagebox.showinfo("Éxito",
                                  f"Documento generado exitosamente:\n\n{ruta_salida}")

                # Preguntar si desea abrir el documento
                if messagebox.askyesno("Abrir Documento",
                                      "¿Desea abrir el documento generado?"):
                    self.abrir_documento(ruta_salida)
            else:
                messagebox.showerror("Error", "No se pudo generar el documento")

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar documento:\n{str(e)}")

    def abrir_documento(self, ruta):
        """Abre el documento generado con la aplicación predeterminada"""
        import platform
        import subprocess

        try:
            if platform.system() == 'Windows':
                os.startfile(ruta)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', ruta])
            else:  # Linux
                subprocess.run(['xdg-open', ruta])
        except Exception as e:
            messagebox.showwarning("Advertencia",
                                 f"No se pudo abrir el documento automáticamente:\n{str(e)}")

    def limpiar_campos(self):
        """Limpia todos los campos de entrada"""
        for entrada in self.campos_entrada.values():
            entrada.delete(0, tk.END)

    def guardar_configuracion(self):
        """Guarda los valores actuales en un archivo JSON"""
        if not self.campos_entrada:
            messagebox.showwarning("Advertencia", "No hay campos para guardar")
            return

        # Recopilar datos
        datos = {marcador: entrada.get() for marcador, entrada in self.campos_entrada.items()}

        # Solicitar ubicación
        ruta = filedialog.asksaveasfilename(
            title="Guardar Configuración",
            defaultextension=".json",
            initialfile="configuracion.json",
            filetypes=[("Archivos JSON", "*.json"), ("Todos los archivos", "*.*")]
        )

        if ruta:
            try:
                with open(ruta, 'w', encoding='utf-8') as f:
                    json.dump(datos, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("Éxito", "Configuración guardada exitosamente")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar configuración:\n{str(e)}")

    def cargar_configuracion(self):
        """Carga valores desde un archivo JSON"""
        if not self.campos_entrada:
            messagebox.showwarning("Advertencia",
                                 "Por favor cargue una plantilla primero")
            return

        ruta = filedialog.askopenfilename(
            title="Cargar Configuración",
            filetypes=[("Archivos JSON", "*.json"), ("Todos los archivos", "*.*")]
        )

        if ruta:
            try:
                with open(ruta, 'r', encoding='utf-8') as f:
                    datos = json.load(f)

                # Cargar datos en los campos
                campos_cargados = 0
                for marcador, valor in datos.items():
                    if marcador in self.campos_entrada:
                        self.campos_entrada[marcador].delete(0, tk.END)
                        self.campos_entrada[marcador].insert(0, valor)
                        campos_cargados += 1

                messagebox.showinfo("Éxito",
                                  f"Se cargaron {campos_cargados} campos desde la configuración")
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar configuración:\n{str(e)}")


def main():
    """Función principal"""
    app = AplicacionAutomatizacion()
    app.mainloop()


if __name__ == "__main__":
    main()
