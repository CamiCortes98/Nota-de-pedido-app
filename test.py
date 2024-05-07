import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pandastable import Table
import openpyxl
from tkcalendar import DateEntry
import json
from ttkwidgets.autocomplete import AutocompleteCombobox
from datetime import datetime
import locale

class SumarizadorNotasPedido:
    def __init__(self, root):
        self.root = root
        self.root.title("Sumarizador de Notas de Pedido")
        self.root.geometry("600x300")
        
        self.df = None
        self.df_sumarizado = None
        
        # Cargar laboratorios desde archivo JSON
        with open("laboratorios.json", encoding='utf-8') as f:
            data = json.load(f)
            self.laboratorios = sorted(set(d["Laboratorio"] for d in data))  # Ordenar alfabéticamente
        
        # Botones
        self.btn_cargar = tk.Button(self.root, text="Cargar archivos", command=self.cargar_archivos)
        self.btn_cargar.pack(pady=10)

        # Calendarios para seleccionar el rango de fechas
        self.lbl_fecha_inicio = tk.Label(self.root, text="Seleccione fecha de inicio:")
        self.lbl_fecha_inicio.pack()
        self.cal_fecha_inicio = DateEntry(self.root, width=12, background='darkblue', foreground='white', borderwidth=2, locale='es_ES')
        self.cal_fecha_inicio.pack()

        self.lbl_fecha_fin = tk.Label(self.root, text="Seleccione fecha de fin:")
        self.lbl_fecha_fin.pack()
        self.cal_fecha_fin = DateEntry(self.root, width=12, background='darkblue', foreground='white', borderwidth=2, locale='es_ES')
        self.cal_fecha_fin.pack()

        # Desplegable para seleccionar el laboratorio con campo de búsqueda
        self.lbl_laboratorio = tk.Label(self.root, text="Seleccione el laboratorio:")
        self.lbl_laboratorio.pack()
        self.selected_laboratorio = tk.StringVar()
        self.entry_laboratorio = AutocompleteCombobox(self.root, textvariable=self.selected_laboratorio, completevalues=self.laboratorios)
        self.entry_laboratorio.pack()

        # Botones Sumarizar y Descargar Resultados
        self.frame_botones = tk.Frame(self.root)
        self.frame_botones.pack(pady=10)
        
        self.btn_sumarizar = tk.Button(self.frame_botones, text="Sumarizar", command=self.sumarizar)
        self.btn_sumarizar.pack(side=tk.LEFT, padx=5)
        
        self.btn_descargar = tk.Button(self.frame_botones, text="Descargar Resultados", command=self.descargar_resultados)
        self.btn_descargar.pack(side=tk.LEFT, padx=5)

        # Etiqueta de resultados
        self.lbl_resultado = tk.Label(self.root, text="")
        self.lbl_resultado.pack(pady=10)

    def cargar_archivos(self):
        files_selected = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if files_selected:
            try:
                dfs = []  # Lista para almacenar los DataFrames cargados de los archivos

                for file_path in files_selected:
                    print("Procesando archivo:", file_path)  # Mensaje de depuración

                    try:
                        # Leer todas las hojas del archivo Excel
                        xl = pd.ExcelFile(file_path)
                        for sheet_name in xl.sheet_names:
                            try:
                                # Leer el contenido de cada hoja
                                df = pd.read_excel(xl, sheet_name=sheet_name, skiprows=9)

                                # Verificar si las columnas necesarias están presentes en el DataFrame
                                if "Can" not in df.columns or "Imp. Total" not in df.columns:
                                    print(f"No se encontraron las columnas necesarias en la hoja '{sheet_name}' del archivo '{file_path}'. Saltando esta hoja...")  # Mensaje de depuración
                                    continue

                                # Reemplazar los valores vacíos en la columna 'Codebar' con una cadena vacía
                                df["Codebar"] = df["Codebar"].fillna('').astype(str)

                                # Leer la fecha del pedido de la celda I2
                                wb = openpyxl.load_workbook(file_path)
                                ws = wb[sheet_name]

                                # Leer la fecha del pedido de la celda I2
                                fecha_pedido = ws['I2'].value

                                # Convertir la columna "Fecha del pedido" a formato de fecha
                                df["Fecha del pedido"] = pd.to_datetime(fecha_pedido, errors='coerce')

                                # Leer el nombre del laboratorio de la celda combinada E2
                                laboratorio = ws['E2'].value

                                # Leer la droguería por donde llega de la celda combinada E4
                                drogueria = ws['E4'].value if ws['E4'].value else "No asignado"

                                # Leer el comprador de la celda combinada C6
                                comprador = ws['C6'].value if ws['C6'].value else "No asignado"

                                # Asignar el nombre del laboratorio, droguería y comprador a las respectivas columnas
                                df["Laboratorio"] = laboratorio
                                df["Droguería por donde Llega"] = drogueria
                                df["Comprador"] = comprador

                                # Agregar el DataFrame procesado a la lista
                                dfs.append(df)
                            except Exception as e:
                                print(f"Error al procesar la hoja '{sheet_name}' del archivo '{file_path}': {e}")  # Mensaje de depuración
                    except Exception as e:
                        print(f"Error al procesar el archivo {file_path}: {e}")  # Mensaje de depuración

                if not dfs:
                    messagebox.showerror("Error", "No se encontraron archivos Excel seleccionados.")
                    return

                # Fusionar los DataFrames en uno solo
                self.df = pd.concat(dfs)

                messagebox.showinfo("Éxito", "Archivos cargados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron cargar los archivos: {e}")

    def sumarizar(self):
        if self.df is None:
            messagebox.showerror("Error", "Primero carga un archivo.")
            return

        try:
            # Convertir la columna "Can" a tipo float
            self.df["Can"] = pd.to_numeric(self.df["Can"], errors="coerce")
            # Convertir la columna "Imp. Total" a tipo float
            self.df["Imp. Total"] = pd.to_numeric(self.df["Imp. Total"], errors="coerce")

            # Obtener el rango de fechas seleccionado
            fecha_inicio = self.cal_fecha_inicio.get_date()  # Objeto datetime.date
            fecha_fin = self.cal_fecha_fin.get_date()  # Objeto datetime.date
            laboratorio = self.selected_laboratorio.get()
        
            # Convertir las fechas a datetime
            fecha_inicio = datetime.combine(fecha_inicio, datetime.min.time())
            fecha_fin = datetime.combine(fecha_fin, datetime.max.time())
        
            # Filtrar por rango de fechas y laboratorio
            df_filtrado = self.df[
                (self.df["Fecha del pedido"] >= fecha_inicio) &
                (self.df["Fecha del pedido"] <= fecha_fin) &
                (self.df["Laboratorio"].str.lower() == laboratorio.lower())
            ]

            # Verificar si se encontraron datos para el rango de fechas y laboratorio ingresados
            if df_filtrado.empty:
                messagebox.showerror("Error", "No se encontraron datos para el rango de fechas y laboratorio ingresados.")
                return

            # Agrupar por código de barras, producto, droguería y comprador y sumarizar cantidades e importes totales
            resumen = df_filtrado.groupby(["Codebar", "Producto", "Droguería por donde Llega", "Comprador"]).agg({'Can': 'sum', 'Imp. Total': 'sum'})

            # Calcular el total de cantidades
            total_cantidades = resumen["Can"].sum()
            # Calcular el total de importes totales
            total_importes = resumen["Imp. Total"].sum()

            # Crear un DataFrame con los resultados sumarizados y los totales
            totales = pd.DataFrame({"Codebar": ["TOTAL"], "Producto": [None], "Droguería por donde Llega": [None], "Comprador": [None], "Can": [total_cantidades], "Imp. Total": [total_importes]})

            # Concatenar el DataFrame de resumen y el DataFrame de totales
            resumen_con_total = pd.concat([resumen.reset_index(), totales], ignore_index=True)  

            # Eliminar el .0 del final en el Codebar
            resumen_con_total['Codebar'] = resumen_con_total['Codebar'].astype(str).str.replace('.0', '')
            
            self.df_sumarizado = resumen_con_total

            # Mostrar el resultado en una nueva ventana
            self.mostrar_resultado(resumen_con_total)

            messagebox.showinfo("Éxito", "Resultados sumarizados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo sumarizar los datos: {e}")

    def mostrar_resultado(self, df):
        # Crear una ventana secundaria para mostrar la tabla
        self.top = tk.Toplevel(self.root)
        self.top.title("Resultados Sumarizados")

        # Mostrar la tabla en la ventana secundaria
        pt = Table(self.top, dataframe=df)
        pt.show()

    def descargar_resultados(self):
        if self.df_sumarizado is None:
            messagebox.showerror("Error", "No hay datos sumarizados para descargar.")
            return

        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        if filename:
            try:
                with pd.ExcelWriter(filename) as writer:
                    self.df_sumarizado.to_excel(writer, index=False, sheet_name="Datos Sumarizados")
                    workbook = writer.book
                    worksheet = writer.sheets["Datos Sumarizados"]
                    header_format = workbook.add_format({'bold': True, 'align': 'center'})
                    worksheet.write(0, 0, "Codebar", header_format)
                    worksheet.write(0, 1, "Producto", header_format)
                    worksheet.write(0, 2, "Droguería por donde Llega", header_format)
                    worksheet.write(0, 3, "Comprador", header_format)
                    worksheet.write(0, 4, "Cantidad", header_format)
                    worksheet.write(0, 5, "Imp. Total", header_format)
                    worksheet.set_column('A:G', 20)
                    messagebox.showinfo("Éxito", "Resultados sumarizados descargados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo descargar los resultados sumarizados: {e}")

# Configurar la localización en español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

# Crear la ventana principal de la aplicación
root = tk.Tk()
app = SumarizadorNotasPedido(root)
root.mainloop()