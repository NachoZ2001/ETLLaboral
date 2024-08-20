import pandas as pd
from tkinter import Tk, filedialog
from datetime import datetime

# Función para seleccionar el Excel de Holistor
def seleccionar_archivo():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    archivo = filedialog.askopenfilename(title="Seleccionar archivo de Holistor", filetypes=[("Archivos de Excel", "*.xlsx")])
    return archivo

# Función para seleccionar la ubicación de destino para el archivo resultante
def seleccionar_destino():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    destino = filedialog.asksaveasfilename(title="Guardar archivo de resultado", defaultextension=".txt", filetypes=[("Archivos de Texto", "*.txt")])
    return destino

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

# Transformar y copiar los datos de holistor a los nuevos registros
for index, row in holistor_df.iterrows():
    # Cuit sin guion
    cuil = row.get('ncuil')
    cuil_final = str(cuil).replace("-","")

    # Verificar que el CUIT tenga exactamente 11 posiciones
    if len(cuil_final) != 11:
        raise ValueError(f"El CUIT '{cuil_final}' no tiene 11 dígitos. Por favor, verifica el valor.")

    # Gremio para ver si es comercio o no
    gremio = row.get('cdgre')

    # Sueldo de 10 posiciones sin coma
    sueldo = round(float(row.get('remcondes')), 2)
    sueldo_sin_coma = str(sueldo).replace(".","")
    sueldo_final = sueldo_sin_coma.zfill(10)

    # Verificar que el SUELDO tenga exactamente 10 posiciones
    if len(sueldo_final) != 10:
        raise ValueError(f"El SUELDO '{sueldo_final}' no tiene 10 dígitos. Por favor, verifica el valor.")

    importe_acuerdo = round(float(row.get('remsindes')), 2)
    importe_acuerdo_sin_coma = str(importe_acuerdo).replace(".","")
    importe_acuerdo_final = importe_acuerdo_sin_coma.zfill(10)

    # Sueldo de jornada parcial de 10 posiciones sin coma
    sueldo_jornada_parcial = 0
    sueldo_jornada_parcial_final = str(sueldo_jornada_parcial).zfill(10)

    # Verificar que el SUELDO JORNADA PARCIAL tenga exactamente 10 posiciones
    if len(sueldo_jornada_parcial_final) != 10:
        raise ValueError(f"El SUELDO JORNADA PARCIAL'{sueldo_jornada_parcial_final}' no tiene 10 dígitos. Por favor, verifica el valor.")
    
    # Verificar que sea del gremio de comercio 
    if gremio == 'COM':
        nuevo_registro = f"{cuil_final}{sueldo_final}{importe_acuerdo_final}{sueldo_jornada_parcial_final}"
        nuevos_registros.append(nuevo_registro)

# Convertir las listas de registros a DataFrames
resultado_df = pd.DataFrame(nuevos_registros)

# Seleccionar el destino del archivo de salida
archivo_destino = seleccionar_destino()
if archivo_destino:
    try:
        # Guardar los registros en un archivo de texto
        with open(archivo_destino, 'w') as file:
            for registro in nuevos_registros:
                file.write(registro + '\n')
        print("Archivo TXT generado correctamente.")
    except Exception as e:
        print(f"Error al escribir el archivo TXT: {e}")
else:
    print("No se ha seleccionado ningún destino.")