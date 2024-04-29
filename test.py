import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pandastable import Table
import openpyxl
from tkcalendar import DateEntry
import json
from ttkwidgets.autocomplete import AutocompleteCombobox
from datetime import datetime

class SumarizadorNotasPedido:
    def __init__(self, root):
        self.root = root
        self.root.title("Sumarizador de Notas de Pedido")
        
        self.df = None
        self.df_sumarizado = None
        
        # Cargar laboratorios desde archivo JSON
        with open("laboratorios.json", encoding='utf-8') as f:
            data = json.load(f)
            self.laboratorios = sorted(set(d["Laboratorio"] for d in data))  # Ordenar alfabéticamente
        
        # Botones
        self.btn_cargar = tk.Button(self.root, text="Cargar archivos", command=self.cargar_archivo_carpeta)
        self.btn_cargar.pack(pady=10)
        
        self.btn_sumarizar = tk.Button(self.root, text="Sumarizar", command=self.sumarizar)
        self.btn_sumarizar.pack(pady=5)

        self.btn_descargar = tk.Button(self.root, text="Descargar Resultados", command=self.descargar_resultados)
        self.btn_descargar.pack(pady=5)

        # Calendarios para seleccionar el rango de fechas
        self.lbl_fecha_inicio = tk.Label(self.root, text="Seleccione fecha de inicio:")
        self.lbl_fecha_inicio.pack()
        self.cal_fecha_inicio = DateEntry(self.root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.cal_fecha_inicio.pack()

        self.lbl_fecha_fin = tk.Label(self.root, text="Seleccione fecha de fin:")
        self.lbl_fecha_fin.pack()
        self.cal_fecha_fin = DateEntry(self.root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.cal_fecha_fin.pack()

        # Desplegable para seleccionar el laboratorio con campo de búsqueda
        self.lbl_laboratorio = tk.Label(self.root, text="Seleccione el laboratorio:")
        self.lbl_laboratorio.pack()
        self.selected_laboratorio = tk.StringVar()
        self.entry_laboratorio = AutocompleteCombobox(self.root, textvariable=self.selected_laboratorio, completevalues=self.laboratorios)
        self.entry_laboratorio.pack()

        # Etiqueta de resultados
        self.lbl_resultado = tk.Label(self.root, text="")
        self.lbl_resultado.pack(pady=10)

    def cargar_archivo_carpeta(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            try:
                dfs = []  # Lista para almacenar los DataFrames cargados de los archivos

                for filename in os.listdir(folder_selected):
                    if filename.endswith(".xlsx"):
                        print("Procesando archivo:", filename)  # Mensaje de depuración

                        file_path = os.path.join(folder_selected, filename)

                        try:
                            # Leer el archivo de Excel sin saltar las primeras filas
                            df = pd.read_excel(file_path, skiprows=9)

                            # Convertir el código de barras a enteros
                            df["Codebar"] = df["Codebar"].astype(str)

                            # Leer la fecha del pedido de la celda I2
                            wb = openpyxl.load_workbook(file_path)
                            ws = wb.active

                            # Leer la fecha del pedido de la celda I2
                            fecha_pedido = ws['I2'].value

                            # Convertir la columna "Fecha del pedido" a formato de fecha
                            df["Fecha del pedido"] = pd.to_datetime(fecha_pedido, errors='coerce')

                            # Leer el nombre del laboratorio de la celda combinada E2
                            laboratorio = ws['E2'].value

                            # Asignar el nombre del laboratorio a la columna "Laboratorio"
                            df["Laboratorio"] = laboratorio

                            # Agregar el DataFrame procesado a la lista
                            dfs.append(df)
                        except Exception as e:
                            print(f"Error al procesar el archivo {filename}: {e}")  # Mensaje de depuración

                if not dfs:
                    messagebox.showerror("Error", "No se encontraron archivos Excel en la carpeta seleccionada.")
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
            # Convertir la columna "Imp. Total" a tipo float
            self.df["Imp. Total"] = pd.to_numeric(self.df["Imp. Total"], errors="coerce")
            # Redondear los importes totales a 2 decimales
            self.df["Imp. Total"] = self.df["Imp. Total"].round(2)

            # Obtener el rango de fechas seleccionado
            fecha_inicio = self.cal_fecha_inicio.get_date()  # Objeto datetime.date
            fecha_fin = self.cal_fecha_fin.get_date()  # Objeto datetime.date
            
            # Convertir a objetos datetime
            fecha_inicio = datetime.combine(fecha_inicio, datetime.min.time())
            fecha_fin = datetime.combine(fecha_fin, datetime.max.time())

            laboratorio = self.selected_laboratorio.get()
        
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

            # Agrupar por código de barras y producto y sumarizar importes
            resumen = df_filtrado.groupby(["Codebar", "Producto"]).agg({'Cantidad': 'sum', 'Imp. Total': 'sum'})

            # Calcular el importe total de todos los productos
            importe_total_total = resumen["Imp. Total"].sum()

            # Calcular el total de cantidades
            total_cantidades = resumen["Cantidad"].sum()

            totales = pd.DataFrame({"Codebar": ["TOTAL"], "Producto": [None], "Cantidad": [total_cantidades], "Imp. Total": [importe_total_total]})

            # Concatenar el DataFrame de resumen y el DataFrame de totales
            resumen_con_total = pd.concat([resumen.reset_index(), totales], ignore_index=True)
            resumen_con_total['Codebar'] = resumen_con_total['Codebar'].astype(str).str.replace('.0', '')
            # Crear un DataFrame con los resultados sumarizados y el total de cantidades
            
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
        pt = Table(self.top, dataframe=df.reset_index(drop=True))
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
                    worksheet.write(0, 2, "Cantidad", header_format)
                    worksheet.write(0, 3, "Imp. Total", header_format)
                    worksheet.set_column('A:D', 20)
                    messagebox.showinfo("Éxito", "Resultados sumarizados descargados correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo descargar los resultados sumarizados: {e}")

# Crear la ventana principal de la aplicación
root = tk.Tk()
app = SumarizadorNotasPedido(root)
root.mainloop()
