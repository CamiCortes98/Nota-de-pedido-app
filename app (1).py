import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

archivos_cargados = []

def buscar_archivos():
    directorio = filedialog.askdirectory()
    archivos = [os.path.join(directorio, archivo) for archivo in os.listdir(directorio) if archivo.endswith('.xlsx') and not archivo.startswith('~$')]
    for archivo in archivos:
        print(archivo)
    return archivos

def cargar_datos():
    global archivos_cargados
    archivos_cargados = buscar_archivos()
    for archivo in archivos_cargados:
        df = pd.read_excel(archivo)
        # valor_a2 = df.iloc[0, 0]  # Obtener el valor de la celda A2 (fila 1, columna 1)
        # print(valor_a2)

        
    
# def filtrar_datos(archivos_cargados, laboratorio, fecha_inicio, fecha_fin):
#     datos_totales = []
#     # Iterar sobre cada archivo cargado
#     for archivo in archivos_cargados:
#         # Leer el archivo
#         df = pd.read_excel(archivo)
#         # Convertir la columna 'Laboratorio' a cadena (str) si no lo es
#         df['Laboratorio'] = df['Laboratorio'].astype(str)
#         # Convertir la columna 'Fecha' a formato de fecha y hora (datetime) si no lo es
#         df['Fecha'] = pd.to_datetime(df['Fecha'])
#         # Filtrar por laboratorio y rango de fechas
#         filtro = (df['Laboratorio'] == laboratorio) & (df['Fecha'].between(fecha_inicio, fecha_fin))

#         datos_filtrados = df[filtro]
#         datos_totales.append(datos_filtrados)
#     # Concatenar los datos filtrados en un único DataFrame
#     return pd.concat(datos_totales)
def filtrar_datos(archivos_cargados, laboratorio, fecha_inicio, fecha_fin):
    datos_totales = []
    # Iterar sobre cada archivo cargado
    for archivo in archivos_cargados:
        # Leer el archivo
        df = pd.read_excel(archivo)
        # Convertir la columna 'Laboratorio' a cadena (str) si no lo es
        df['Laboratorio'] = df['Laboratorio'].astype(str)
        # Convertir la columna 'Fecha' a formato de fecha y hora (datetime) si no lo es
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        # Filtrar por laboratorio y rango de fechas
        filtro = (df['Laboratorio'] == laboratorio) & (df['Fecha'].between(fecha_inicio, fecha_fin))

        datos_filtrados = df[filtro]
        datos_totales.append(datos_filtrados)
    # Concatenar los datos filtrados en un único DataFrame
    return pd.concat(datos_totales)


def calcular_suma_y_cantidad(datos):
    suma_total = datos['Importe'].sum()
    cantidad_total = datos['Unidades'].sum()
    return suma_total, cantidad_total

def mostrar_previsualizacion():
    global archivos_cargados
    if not archivos_cargados:
        tk.messagebox.showerror("Error", "Debe cargar una carpeta primero.")
        return
    
    laboratorio = entry_laboratorio.get()
    fecha_inicio = entry_fecha_inicio.get()
    fecha_fin = entry_fecha_fin.get()
    datos_filtrados = filtrar_datos(archivos_cargados, laboratorio, fecha_inicio, fecha_fin)
    ventana_previsualizacion = tk.Toplevel(ventana_principal)
    ventana_previsualizacion.title("Previsualización de datos")
    previsualizacion_texto = tk.Text(ventana_previsualizacion)
    previsualizacion_texto.insert(tk.END, datos_filtrados.to_string(index=False))
    previsualizacion_texto.pack()

def mostrar_resultados(suma_total, cantidad_total, laboratorio, fecha_inicio, fecha_fin):
    ventana_resultados = tk.Toplevel(ventana_principal)
    ventana_resultados.title("Resultados")
    
    tk.Label(ventana_resultados, text="Laboratorio:").grid(row=0, column=0)
    tk.Label(ventana_resultados, text=laboratorio).grid(row=0, column=1)
    
    tk.Label(ventana_resultados, text="Rango de fechas:").grid(row=1, column=0)
    tk.Label(ventana_resultados, text=f"{fecha_inicio} - {fecha_fin}").grid(row=1, column=1)
    
    tk.Label(ventana_resultados, text="Suma total:").grid(row=2, column=0)
    tk.Label(ventana_resultados, text=suma_total).grid(row=2, column=1)
    
    tk.Label(ventana_resultados, text="Cantidad total:").grid(row=3, column=0)
    tk.Label(ventana_resultados, text=cantidad_total).grid(row=3, column=1)

def procesar_datos():
    global archivos_cargados
    if not archivos_cargados:
        tk.messagebox.showerror("Error", "Debe cargar una carpeta primero.")
        return
    
    laboratorio = entry_laboratorio.get()
    fecha_inicio = entry_fecha_inicio.get()
    fecha_fin = entry_fecha_fin.get()
    datos_filtrados = filtrar_datos(archivos_cargados, laboratorio, fecha_inicio, fecha_fin)
    if datos_filtrados.empty:
        tk.messagebox.showinfo("Información", "No hay datos disponibles para el filtro especificado.")
    else:
        suma_total, cantidad_total = calcular_suma_y_cantidad(datos_filtrados)
        mostrar_resultados(suma_total, cantidad_total, laboratorio, fecha_inicio, fecha_fin)

# Crear la ventana principal
ventana_principal = tk.Tk()
ventana_principal.title("Análisis de datos")

# Crear widgets
tk.Label(ventana_principal, text="Laboratorio:").grid(row=0, column=0)
entry_laboratorio = tk.Entry(ventana_principal)
entry_laboratorio.grid(row=0, column=1)

tk.Label(ventana_principal, text="Fecha inicio (YYYY-MM-DD):").grid(row=1, column=0)
entry_fecha_inicio = tk.Entry(ventana_principal)
entry_fecha_inicio.grid(row=1, column=1)

tk.Label(ventana_principal, text="Fecha fin (YYYY-MM-DD):").grid(row=2, column=0)
entry_fecha_fin = tk.Entry(ventana_principal)
entry_fecha_fin.grid(row=2, column=1)

boton_cargar = tk.Button(ventana_principal, text="Cargar carpeta", command=cargar_datos)
boton_cargar.grid(row=3, column=0, columnspan=2)

boton_previsualizar = tk.Button(ventana_principal, text="Previsualizar datos", command=mostrar_previsualizacion)
boton_previsualizar.grid(row=4, column=0, columnspan=2)

boton_procesar = tk.Button(ventana_principal, text="Procesar", command=procesar_datos)
boton_procesar.grid(row=5, column=0, columnspan=2)

boton_cerrar = tk.Button(ventana_principal, text="Salir", command=ventana_principal.destroy)
boton_cerrar.grid(row=6, column=1, columnspan=2)

# Ejecutar la ventana
ventana_principal.mainloop()
