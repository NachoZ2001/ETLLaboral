import pandas as pd
from tkinter import Tk, filedialog, StringVar, OptionMenu, simpledialog, Label, Button
from datetime import datetime

# Función para seleccionar el Excel de Holistor
def seleccionar_archivo():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos de Excel", "*.xlsxclear")])
    return archivo

# Función para seleccionar la ubicación de destino para el archivo resultante
def seleccionar_destino():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    destino = filedialog.asksaveasfilename(title="Guardar archivo de resultado", defaultextension=".csv", filetypes=[("Archivos CSV", "*.csv")])
    return destino

# Función para mostrar un desplegable para seleccionar el mes
def seleccionar_periodo():
    root = Tk()
    root.title("Seleccionar Mes del Período")

    # Definir las opciones de meses
    opciones_mes = ["01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", 
                    "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto", 
                    "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre"]

    # Variable para almacenar el mes seleccionado
    mes_seleccionado = StringVar(root)
    mes_seleccionado.set(opciones_mes[0])  # Valor por defecto

    # Etiqueta y menú desplegable
    label = Label(root, text="Seleccione el mes del período:")
    label.pack(pady=10)
    desplegable = OptionMenu(root, mes_seleccionado, *opciones_mes)
    desplegable.pack(pady=10)

    # Botón para confirmar la selección
    def confirmar_seleccion():
        root.quit()

    boton = Button(root, text="Confirmar", command=confirmar_seleccion)
    boton.pack(pady=10)

    root.mainloop()

    # Obtener el mes seleccionado y cerrar la ventana
    mes = mes_seleccionado.get().split(" - ")[0]  # Extraer el código del mes
    root.destroy()

    año_actual = datetime.now().year
    return f"{año_actual}{mes}"

# Leer el archivo de Excel de Holistor desde la ruta seleccionada
archivo_holistor = seleccionar_archivo()
if not archivo_holistor:
    print("No se ha seleccionado ningún archivo.")
    exit()
try:
    holistor_df = pd.read_excel(archivo_holistor)
except FileNotFoundError as e:
    print(f"Error al leer el archivo de Excel: {e}")
    exit()

# Crear una lista para almacenar los nuevos registros
nuevos_registros = []

# Obtener el periodo
periodo = seleccionar_periodo()

# Transformar y copiar los datos de holistor a los nuevos registros
for index, row in holistor_df.iterrows():
   
    # Cuit sin guion
    cuil = row.get('ncuil')

    # Obtengo el nombre
    nombre = row.get('nombr')

    # Obtengo el concepto
    concepto = 2

    # Obtengo el salario
    remsindes = float(row.get('remsindes'))
    remcondes = float(row.get('remcondes'))
    salario = remsindes + remcondes

    # Creo el registro y lo agrego a la lista
    nuevo_registro = f"{cuil};{periodo};{nombre};{concepto};{salario}"
    nuevos_registros.append(nuevo_registro)

# Seleccionar el destino del archivo de salida
archivo_destino = seleccionar_destino()
if archivo_destino:
    try:
        # Guardar los registros en un archivo CSV
        with open(archivo_destino, 'w', newline='') as file:
            file.write("CUIL;PERIODO;NOMBRE;CONCEPTO;SALARIO\n")  # Escribir encabezados
            for registro in nuevos_registros:
                file.write(registro + '\n')
        print("Archivo CSV generado correctamente.")
    except Exception as e:
        print(f"Error al escribir el archivo CSV: {e}")
else:
    print("No se ha seleccionado ningún destino.")
