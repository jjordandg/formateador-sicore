import tkinter as tk
from tkinter import filedialog
import pandas as pd

#establecemos el largo de cada bloque
long_bloq1 = 28
long_bloq2 = 27
long_bloq3 = 43
long_bloq4 = 37

def cargar_excel():
    archivo_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if archivo_excel:
        entry_ruta_excel.delete(0, tk.END)
        entry_ruta_excel.insert(0, archivo_excel)

def procesar_excel():
    ruta_excel = entry_ruta_excel.get()
    if not ruta_excel:
        label_resultado.config(text="Por favor, selecciona el archivo Excel para Sicore.")
        return

    archivo_salida = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if not archivo_salida:
        label_resultado.config(text="Por favor, selecciona una ubicación para guardar el archivo de salida.")
        return

    df = pd.read_excel(ruta_excel, dtype={
        "Cod.1": str,
        "Fecha Comprobante": str,
        "Cod. Comprobante": str,
        "Honorarios brutos": str,
        "Cod.2": str,
        "Base de calculo": str,
        "Fecha fin de mes": str,
        "Cod.3": str,
        "Retencion": str,
        "Cod.4": str,
        "Cod.5": str,
        "CUIL": str,
        "Digito1": str,
        "Cod.6": str,
        "Año": str,
        "Digito2": str,
        "Digito3": str
    })

    def custom_float_format(value):
        return float(value.replace(',', '.'))

    df['Honorarios brutos'] = df['Honorarios brutos'].apply(custom_float_format)
    df['Retencion'] = df['Retencion'].apply(custom_float_format)
    df['Base de calculo'] = df['Base de calculo'].apply(custom_float_format)

    df['Honorarios brutos'] = df['Honorarios brutos'].apply(lambda x: '{:.2f}'.format(x))
    df['Retencion'] = df['Retencion'].apply(lambda x: '{:.2f}'.format(x))
    df['Base de calculo'] = df['Base de calculo'].apply(lambda x: '{:.2f}'.format(x))

    with open(archivo_salida, "w") as file:
        for index, row in df.iterrows():
            bloq1 = str(row['Cod.1']) + str(row['Fecha Comprobante']) + str(row['Cod. Comprobante'])
            bloq2 = str(row['Honorarios brutos']).replace('.', ',') + str(row['Cod.2'])
            bloq3 = str(row['Base de calculo']).replace('.', ',') + str(row['Fecha fin de mes']).replace(' 00:00:00','').replace('-', '/') + str(row['Cod.3'])
            bloq4 = str(row['Retencion']).replace('.', ',') + '  ' + str(row['Cod.4']) + '          ' + str(row['Cod.5']) + str(row['CUIL'])
            bloq5 = str(row['Digito1']) + str(row['Cod.6']) + str(row['Año']) + str(row['Digito2']) + str(row['Digito3'])

            espacio_bloq1 = ' ' * (long_bloq1 - len(bloq2))
            espacio_bloq2 = ' ' * (long_bloq2 - len(bloq3))
            espacio_bloq3 = ' ' * (long_bloq3 - len(bloq4))
            espacio_bloq4 = ' ' * (long_bloq4 - len(bloq5))

            linea_formateada = (
                    bloq1 + espacio_bloq1 +
                    bloq2 + espacio_bloq2 +
                    bloq3 + espacio_bloq3 +
                    bloq4 + espacio_bloq4 +
                    bloq5
            )

            file.write(linea_formateada + "\n")

        label_resultado.config(text=f"Los datos se han escrito en {archivo_salida}")

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Procesador de formato para Sicore")

# Crear elementos de la GUI
label_instrucciones = tk.Label(ventana, text="Selecciona un archivo Excel para procesar:")
button_cargar_excel = tk.Button(ventana, text="Cargar Excel", command=cargar_excel)
entry_ruta_excel = tk.Entry(ventana)
button_procesar = tk.Button(ventana, text="Procesar y Guardar", command=procesar_excel)
label_resultado = tk.Label(ventana, text="La conversión ha sido procesada con éxito")

# Colocar elementos en la ventana
label_instrucciones.pack()
button_cargar_excel.pack()
entry_ruta_excel.pack()
button_procesar.pack()
label_resultado.pack()

# Iniciar la aplicación
ventana.mainloop()




