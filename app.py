import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import time

class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Análisis de Datos")
        
        # Obtener el tamaño de la pantalla y ajustar la ventana al 75% del tamaño
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.75)
        window_height = int(screen_height * 0.75)
        position_right = int((screen_width - window_width) / 2)
        position_down = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_right}+{position_down}")

        self.df = None
        self.nombre_archivo = None
        self.fecha_inicio = None
        self.fecha_fin = None

        # Crear botón para cargar archivo
        self.load_button = tk.Button(root, text="Cargar archivo", command=self.cargar_archivo)
        self.load_button.pack(pady=10)

        # Etiqueta para mostrar el nombre del archivo seleccionado
        self.file_label = tk.Label(root, text="No se ha seleccionado un archivo")
        self.file_label.pack()

        # Dropdown para la selección de hoja
        self.sheet_label = tk.Label(root, text="Selecciona una hoja:")
        self.sheet_label.pack()
        self.sheet_dropdown = ttk.Combobox(root, state="readonly")
        self.sheet_dropdown.pack()

        # Botón para seleccionar fechas
        self.filtrar_fechas_button = tk.Button(root, text="Seleccionar Fechas", command=self.seleccionar_fechas)
        self.filtrar_fechas_button.pack(pady=10)

        # Texto de estado
        self.status_label = tk.Label(root, text="")
        self.status_label.pack()

        # Botón para leer datos de la hoja seleccionada
        self.load_sheet_button = tk.Button(root, text="Cargar Hoja", command=self.cargar_hoja_con_progreso)
        self.load_sheet_button.pack(pady=10)

        # Barra de progreso (inicialmente oculta)
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")
        self.progress_bar.pack(pady=20)
        self.progress_bar.pack_forget()  # Ocultarla al inicio

        # Texto de estado 2
        self.status_label2 = tk.Label(root, text="")
        self.status_label2.pack()
        self.progress_bar.pack_forget()  # Ocultarla al inicio

        # Botón "Siguiente" (inicialmente oculto)
        self.next_button = tk.Button(root, text="Siguiente", command=self.mostrar_menu, state="disabled")
        self.next_button.pack(pady=10)
        self.next_button.pack_forget()  # Ocultar el botón al inicio


    def cargar_archivo(self):
        # Usar filedialog para seleccionar un archivo .xlsx
        archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        if archivo:
            self.file_label.config(text=archivo)
            self.nombre_archivo = archivo
            try:
                # Obtener nombres de las hojas
                self.hojas = pd.ExcelFile(archivo).sheet_names
                self.sheet_dropdown['values'] = self.hojas  # Actualizar opciones del combobox
                self.status_label.config(text="Archivo cargado con éxito.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo leer el archivo: {str(e)}")
        else:
            messagebox.showwarning("Advertencia", "No se seleccionó un archivo.")

    def seleccionar_fechas(self):
        hoja_seleccionada = self.sheet_dropdown.get()
        if hoja_seleccionada and self.nombre_archivo:  # Verificar si hay hoja seleccionada y archivo cargado
            self.abrir_filtro_fecha(hoja=hoja_seleccionada, archivo=self.nombre_archivo)
        else:
            messagebox.showwarning("Advertencia", "Por favor selecciona una hoja y carga un archivo antes de continuar.")

    def cargar_hoja_con_progreso(self):
        # Mostrar la barra de progreso y comenzar el movimiento
        self.progress_bar.pack(pady=20)
        self.progress_bar.start()  # Iniciar animación de la barra en modo "indeterminate"
        
        # Iniciar un hilo para cargar la hoja sin bloquear la interfaz
        threading.Thread(target=self.cargar_hoja_con_progreso_hilo).start()

    def cargar_hoja_con_progreso_hilo(self):
        hoja_seleccionada = self.sheet_dropdown.get()
        if hoja_seleccionada:
            try:
                # Simular tiempo de carga
                time.sleep(2)  # Simular que toma tiempo leer la hoja

                # Cargar la hoja real
                self.df = pd.read_excel(self.nombre_archivo, sheet_name=hoja_seleccionada)

                # Si hay fechas de filtro, filtrar por ellas
                if self.fecha_inicio and self.fecha_fin:
                    self.df = self.df[(self.df['Fecha'] >= self.fecha_inicio) & (self.df['Fecha'] <= self.fecha_fin)]
                
                # Detener el movimiento de la barra de progreso y llenarla completamente
                self.progress_bar.stop()  # Detener animación "indeterminate"
                self.progress_bar["mode"] = "determinate"  # Cambiar a modo "determinate"
                self.progress_bar["value"] = 100  # Llenar completamente la barra
                self.status_label2.config(text=f"Hoja '{hoja_seleccionada}' cargada con éxito.")
                self.status_label2.pack()
                self.status_label.pack_forget()  # Ocultar el estado anterior
                
                # Habilitar y volver a mostrar el botón "Siguiente"
                self.next_button.config(state="normal")
                self.next_button.pack()  # Volver a mostrar el botón
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar la hoja: {str(e)}")
        else:
            messagebox.showwarning("Advertencia", "No has seleccionado ninguna hoja.")

    def cargar_hoja(self):
        hoja_seleccionada = self.sheet_dropdown.get()
        if hoja_seleccionada():
            self.cargar_hoja_con_progreso()
        else:
            messagebox.showwarning("Advertencia", "Por favor selecciona una hoja antes de continuar.")

    def abrir_filtro_fecha(self, hoja, archivo):
        self.df = pd.read_excel(archivo, sheet_name=hoja)
        if self.df.empty:
            messagebox.showwarning("Advertencia", "No has seleccionado ninguna hoja.")
            return
        
        self.fecha_inicio = self.df['Fecha'].iloc[0]  # Primer valor
        self.fecha_fin = self.df['Fecha'].iloc[-1]    # Último valor

        self.rango_fechas = pd.date_range(start=self.fecha_inicio, end=self.fecha_fin).strftime('%d/%m/%Y').to_numpy()

        # Crear una nueva ventana para la selección de fechas
        self.fecha_window = tk.Toplevel(self.root)
        self.fecha_window.title("Seleccionar Fechas")
        self.fecha_window.geometry("400x400")

        # Dropdown para fecha de inicio
        self.label_inicio = tk.Label(self.fecha_window, text="Selecciona la fecha inicial:")
        self.label_inicio.pack(pady=10)
        self.fecha_inicio_dropdown = ttk.Combobox(self.fecha_window, state="readonly")
        self.fecha_inicio_dropdown['values'] = [str(fecha) for fecha in self.rango_fechas]
        self.fecha_inicio_dropdown.pack(pady=10)

        # Dropdown para fecha final
        self.label_fin = tk.Label(self.fecha_window, text="Selecciona la fecha final:")
        self.label_fin.pack(pady=10)
        self.fecha_fin_dropdown = ttk.Combobox(self.fecha_window, state="readonly")
        self.fecha_fin_dropdown['values'] = [str(fecha) for fecha in self.rango_fechas]
        self.fecha_fin_dropdown.pack(pady=10)

        # Botón "Filtrar"
        self.filtrar_button = tk.Button(self.fecha_window, text="Filtrar", state="disabled", command=self.aplicar_filtro_fechas)
        self.filtrar_button.pack(pady=10)

        # Habilitar el botón cuando ambas fechas estén seleccionadas
        self.fecha_inicio_dropdown.bind("<<ComboboxSelected>>", self.verificar_seleccion)
        self.fecha_fin_dropdown.bind("<<ComboboxSelected>>", self.verificar_seleccion)

    def verificar_seleccion(self, event):
        # Verificar si ambas fechas han sido seleccionadas
        if self.fecha_inicio_dropdown.get() and self.fecha_fin_dropdown.get():
            self.filtrar_button.config(state="normal")

    def aplicar_filtro_fechas(self):
        # Obtener las fechas seleccionadas
        self.fecha_inicio = self.fecha_inicio_dropdown.get()
        self.fecha_fin = self.fecha_fin_dropdown.get()
        self.fecha_window.destroy()  # Cerrar la ventana de selección de fechas
        self.status_label.config(text=f"Fechas seleccionadas: {self.fecha_inicio} a {self.fecha_fin}")

    def mostrar_menu(self):
        # Limpiar la ventana y mostrar botones de menú
        for widget in self.root.winfo_children():
            widget.pack_forget()

        # Ejemplo de botones de menú
        menu_label = tk.Label(self.root, text="Menú Principal", font=("Helvetica", 16))
        menu_label.pack(pady=20)

        # Botones del menú
        button1 = tk.Button(self.root, text="Análisis Univariado", command=self.opcion1)
        button1.pack(pady=10)

        button2 = tk.Button(self.root, text="Opción 2", command=self.opcion2)
        button2.pack(pady=10)

        button3 = tk.Button(self.root, text="Opción 3", command=self.opcion3)
        button3.pack(pady=10)

    def opcion1(self):
        # Limpiar la ventana y mostrar el análisis univariado
        for widget in self.root.winfo_children():
            widget.pack_forget()

        self.root.geometry("800x600")  # Ajustar tamaño de la ventana

        # titulo
        menu_label = tk.Label(self.root, text="Análisis Univariado", font=("Helvetica", 16))
        menu_label.pack(pady=20)

        # dropdown para seleccionar columna
        self.label_fin = tk.Label(self.root, text="Selecciona la variable:")
        columnas = self.df.columns[self.df.columns != 'Fecha']
        self.column_dropdown = ttk.Combobox(self.root, state="readonly")
        self.column_dropdown['values'] = columnas
        self.column_dropdown.pack()

        # al seleccionar variable, comprobar valores NaN en la columna de dicha variable
        self.column_dropdown.bind("<<ComboboxSelected>>", self.comprobar_nan)

        # dropdown para seleccionar unidad de medida
        self.label_fin = tk.Label(self.root, text="Selecciona la unidad de medida:")
        unidades = ['unidad', 'Kg', 'g', 'ton', 'Pa', 'bar', 'psi', 'm/s', 'km/h', 'rpm', '°C', '°F', 'K',
                    '%', 'TPH', 'm3', 'm3/h', 'm3/día']
        self.unidad_dropdown = ttk.Combobox(self.root, state="readonly")
        self.unidad_dropdown['values'] = unidades
        self.unidad_dropdown.pack()

        # dropdown para seleccionar tipo de outlier identification
        self.label_fin = tk.Label(self.root, text="Selecciona el tipo de identificación de outliers:")
        tipos = ['MCD', 'KNN', 'Ninguno']
        self.tipo_dropdown = ttk.Combobox(self.root, state="readonly")
        self.tipo_dropdown['values'] = tipos
        self.tipo_dropdown.pack()

        # Campo de texto para indicar contaminacion
        self.contaminacion_label = tk.Label(self.root, text="Contaminación:")
        self.contaminacion_label.pack()
        self.contaminacion_entry = tk.Entry(self.root)
        self.contaminacion_entry.pack()

        #  Botón "Limpiar Datos"
        self.limpiar_button = tk.Button(self.root, text="Limpiar Datos", command=self.limpiar_datos)
        self.limpiar_button.pack(pady=10)
        self.limpiar_button.pack_forget()  # Ocultar el botón inicialmente

        # Boton de Outliers
        self.outliers_button = tk.Button(self.root, text="Identificar Outliers", command=self.identificar_outliers)
        self.outliers_button.pack(pady=10)

        # Etiqueta de estado
        self.status_label4 = tk.Label(self.root, text="")
        self.status_label4.pack()
        self.status_label4.pack_forget()  # Ocultar el estado inicialmente

        # Botón "Graficar"
        self.filtrar_button = tk.Button(self.root, text="Graficar", command=self.graficar)
        self.filtrar_button.pack(pady=10)
        self.filtrar_button.pack_forget()  # Ocultar el botón al inicio

    def comrobar_nan(self, event):
        self.variable = self.column_dropdown.get()
        num_nan = self.df[self.variable].isna().sum()
        if num_nan > 0:
            messagebox.showwarning("Advertencia", f"La variable '{self.variable}' tiene {num_nan} valores NaN.")
            self.limpiar_button.pack()  # Mostrar el botón "Limpiar Datos"

    def limpiar_datos(self):
        self.clean_window = tk.Toplevel(self.root)
        self.clean_window.title("Limpieza de Datos")
        self.clean_window.geometry("400x400")

        # Dropdown para seleccionar método de limpieza
        self.label_inicio = tk.Label(self.clean_window, text="Selecciona el método de limpieza:")
        self.label_inicio.pack(pady=10)

        self.metodos = ['Eliminar', 'Rellenar con Media', 'Imputar con Mediana']
        self.metodo_dropdown = ttk.Combobox(self.clean_window, state="readonly")
        self.metodo_dropdown['values'] = self.metodos
        self.metodo_dropdown.pack(pady=10)

        # Botón "Aplicar"
        self.filtrar_button = tk.Button(self.clean_window, text="Aplicar", command=self.aplicar_limpieza)
        self.filtrar_button.pack(pady=10)

        # Texto de estado
        self.status_label3 = tk.Label(self.clean_window, text="")
        self.status_label3.pack()

        # Botón Terminar de limpiar
        self.terminar_button = tk.Button(self.clean_window, text="Terminar", command=self.clean_window.destroy)
        self.terminar_button.pack(pady=10)
        self.terminar_button.pack_forget()  # Ocultar el botón al inicio

    def aplicar_limpieza(self):
        # Obtener el método de limpieza seleccionado
        metodo = self.metodo_dropdown.get()

        if metodo == 'Eliminar':
            self.df = self.df.dropna(subset=[self.variable])
            self.status_label3.config(text="Se han eliminado los valores NaN.")
        elif metodo == 'Rellenar con Media':
            media = self.df[self.variable].mean()
            self.df[self.variable].fillna(media, inplace=True)
            self.status_label3.config(text=f"Se han rellenado los valores NaN con la media: {media}.")
        elif metodo == 'Imputar con Mediana':
            mediana = self.df[self.variable].median()
            self.df[self.variable].fillna(mediana, inplace=True)
            self.status_label3.config(text=f"Se han imputado los valores NaN con la mediana: {mediana}.")
        else:
            self.status_label3.config(text="Método no válido. Por favor, elija una opción válida.")

        self.terminar_button.pack()  # Mostrar el botón "Terminar"

    def identificar_outliers(self):
        # Obtener el método de identificación de outliers seleccionado
        metodo = self.tipo_dropdown.get()

        if metodo == 'MCD':
            # Identificar outliers usando Minimum Covariance Determinant
            from sklearn.covariance import EllipticEnvelope
            contaminacion = self.contaminacion_entry.get()
            envelope = EllipticEnvelope(contamination=float(contaminacion))
            self.df['Outlier'] = envelope.fit_predict(self.df[[self.variable]])

            # Si valor de variable es 0, no se considera inlier
            self.df['Outlier'] = self.df['Outlier'].apply(lambda x: 1 if x == -1 else 0)
            # Cantidad de outliers identificados
            num_outliers = self.df['Outlier'].sum()
            self.status_label4.config(text=f"Outliers identificados usando MCD: {num_outliers}.")

        elif metodo == 'KNN':
            self.status_label4.config(text="Outliers identificados usando KNN.")
        else:
            self.status_label4.config(text="Método no válido. Por favor, elija una opción válida.")

    def opcion2(self):
        messagebox.showinfo("Opción 2", "Has seleccionado la Opción 2.")

    def opcion3(self):
        messagebox.showinfo("Opción 3", "Has seleccionado la Opción 3.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
