import pandas as pd
from tkinter import Tk, filedialog
from datetime import datetime

# Función para seleccionar el archivo de Holistor
def seleccionar_archivo():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    archivo = filedialog.askopenfilename(title="Seleccionar archivo de Holistor", filetypes=[("Archivos de Excel", "*.xlsx")])
    return archivo

# Función para seleccionar la ubicación de destino para el archivo resultante
def seleccionar_destino():
    Tk().withdraw()  # Ocultar la ventana principal de tkinter
    destino = filedialog.asksaveasfilename(title="Guardar archivo de resultado", defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
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


# Diccionarios de mapeo
diccionario_nacionalidades = {
    'ARG': 'Argentina',
    'BRA': 'Brasil',
    'CHI': 'Chile',
}

diccionario_codigos_categoria_convenio = {
    'ADM-A': 'ADMIN A',
    'ADM-B': 'ADMIN B',
    'ADM-C': 'ADMIN C',
    'ADM-D': 'ADMIN D',
    'ADM-E': 'ADMIN E',
    'ADM-F': 'ADMIN F',
    'AUX-A': 'AUX A',
    'AUX-B': 'AUX B',
    'AUX-C': 'AUX C',
    'CAJ-A': 'CAJERO A',
    'CAJ-B': 'CAJERO B',
    'CAJ-C': 'CAJERO C',
    'ESP-A': 'ESPECIAL A',
    'ESP-B': 'ESPECIAL B',
    'MAE-A': 'MAEST A',
    'MAE-B': 'MAEST B',
    'MAE-C': 'MAEST C',
    'VEN-A': 'VEND A',
    'VEN-B': 'VEND B',
    'VEN-C': 'VEND C',
    'VEN-D': 'VEND D',
    'F/CON': 'FUERAC'
}

diccionario_codigos_categoria_madereros = {
    'AYUD' : 'AYUD',
    'MO' : 'MED OF',
    'OF ES' : 'OF ESP',
    'OF GE' : 'OF GEN',
    'OF MU' : 'OF MULT',
    'OP AI' : 'OP ACT IND'
}

diccionario_codigos_categoria_metalurgicos ={
    'ADM1º' : 'ADMIN 1',
    'ADM2º' : 'ADMIN 2',
    'ADM3º' : 'ADMIN 3',
    'ADM4º' : 'ADMIN 4',
    'AUX 1' : 'AUX 1',
    'TEC3º' : 'TECNICO 3',
    'TEC' : 'TECNICO 4',
    'TEC6º' : 'TECNICO 6',
    'INGR' : 'INGRESANTE',
    'OPCAL' : 'OPERARIO CALIFICADO',
    'MEDOF' : 'MEDIO OFICIAL',
    'MO' : 'MEDIO OFICIAL',
    'OPESP' : 'OPERARIO ESPECIALIZADO',
    'OEM' : 'OP ESPEC MULTIPLE',
    'OPESM' : 'OP ESPEC MULTIPLE',
    'OFIC' : 'OFICIAL',
    'OFMUL' : 'OFICIAL MULTIPLE',
    'OFMS' : 'OFICIAL MULT SUPERIOR',
    'TEC1º' : 'SUPERVISOR TECNICO 1',
    'SUPER' : 'SUPERVISOR DE FABRICA 3'
}


diccionario_sexo = {
    'M': 'Masculino',
    'F': 'Femenino',
}

diccionario_condiciones = {
    'Mes': 'Mensualizado',
    'Quincena' : 'Quincena',
    'Hora' : 'Quincena'
}

diccionario_codigos_convenio = {
    'F/C': 'FUERAC',
    'COM': 'COMERCIO',
    'MET' : 'METALURGICOS',
    'MADE' : 'MADEREROS'
}

diccionario_codigos_sindicato = {
    'F/C': ' ',
    'COM': 'CEC',
    'MET' : 'METALURGICOS',
    'MADE' : 'MADEREROS'
}

# Funciones de transformación
def extraer_nombre(nombre_completo):
    partes = nombre_completo.split(maxsplit=1)
    return partes[1] if len(partes) > 1 else ''

def extraer_apellido(nombre_completo):
    partes = nombre_completo.split()
    return partes[0] if partes else ''

def obtener_nacionalidad(codigo):
    return diccionario_nacionalidades.get(codigo, ' ')

def convertir_formato(fecha):
    if not fecha or fecha.strip() == '-   -':
        return ' '
    try:
        # Dividir la fecha por guion
        fecha_dividida = fecha.split('-')
        dia = int(fecha_dividida[0])
        mes_texto = fecha_dividida[1]
        año = int(fecha_dividida[2])

        # Mapeo de los meses a sus números correspondientes
        meses = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
            'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
            'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        }
        
        # Convertir el mes de texto a número
        mes = meses[mes_texto]

        date = datetime.now()
        year = str(date.year)
        year = year[-2:]

        if año < int(year):
            año += 2000
        else:
            año += 1900

        # Retornar la fecha en el formato "día/mes/año"
        return f"{dia:02d}/{mes}/{año}"
    except ValueError:
        return ' '

def obtener_sexo(codigo):
    return diccionario_sexo.get(codigo, ' ')

def obtener_condicion(codigo):
    return diccionario_condiciones.get(codigo, ' ')

def obtener_codigo_convenio(codigo):
    return diccionario_codigos_convenio.get(codigo, ' ')

def obtener_codigo_categoria(codigo, codigo_convenio):
    if codigo_convenio == 'COM':
        return diccionario_codigos_categoria_convenio.get(codigo, ' ')
    elif codigo_convenio == 'MET':
        return diccionario_codigos_categoria_metalurgicos.get(codigo, ' ')
    elif codigo_convenio == 'MADE':
        return diccionario_codigos_categoria_madereros.get(codigo, '')
    else:
        return 'FUERAC'

def obtener_codigo_sindicato(codigo):
    return diccionario_codigos_sindicato.get(codigo, ' ')

def es_fuera_convenio(codigo):
    return 'Si' if codigo == 'FUERAC' else 'No'

def es_afiliado_sindicato(codigo):
    return 'No' if codigo == ' ' or codigo == '' else 'Si'

def con_cobertura_seguro_vida(codigo):
    return 'SI' if codigo == 'S' else 'NO'

def obtener_tipo_telefono(codigo):
    if codigo == 1:
        return 'Particular'
    return ' '

def obtener_telefono(codigo_area, numero):
    return f"{str(codigo_area)}-{str(numero)}"

def normalizar_valor(valor):
    # Convertir el valor a string y eliminar ceros a la izquierda
    return str(int(valor))

def verificar_valor(valor, lista_valores):
    # Normalizar el valor de entrada
    valor_normalizado = normalizar_valor(valor)
    
    # Crear un diccionario de valores normalizados a valores originales
    diccionario_valores = {normalizar_valor(v): v for v in lista_valores}
    
    # Verificar si el valor normalizado está en el diccionario
    if valor_normalizado in diccionario_valores:
        return diccionario_valores[valor_normalizado]
    else:
        return None
    
# Definir los valores de comparación
valores_comparacion_codigo_situacion_revista = ['01', '10', '11', '12', '20', '21', '22', '23', '24', '50']
valores_comparacion_codigo_situacion_revista_1 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38','39','40','41','42','99']
valores_comparacion_codigo_condicion = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
valores_comparacion_codigo_actividad = ['00', '01', '02', '03', '04', '05', '06', '07', '08', '09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34','35','36','37','38','39','40','41','42','43','44','45','46','47','48','49','50','51','52','53','54','55','56','57','58','59','60','61','62','63','64','65','66','67','68','69','70','71','72','73','74','75','76','77','78','79','80','81','82','83','84','85','86','87','88','89','90','91','92','93','94','95','96','97','98','100','101','120','122']
valores_comparacion_codigo_contratacion = ['000', '001', '002', '003', '004', '005', '006', '007', '008', '009','010','011','012','013','014','015','016','017','018','019','020','021','022','023','024','025','026','027','028','029','030','031','032','033','034','035','036','037','038','039','040','041','042','043','044','045','046','047','048','049','050','099','100','102','110','111','201','202','203','211','212','213','221','222','223','301','302','303','304','305','306','307','308','309','310','311','312','313','314','315','999']

def obtener_codigo_situacion_revista(codigo):
    return verificar_valor(codigo, valores_comparacion_codigo_situacion_revista)

def obtener_codigo_situacion_revista_1(codigo):
    return verificar_valor(codigo, valores_comparacion_codigo_situacion_revista_1)

def obtener_codigo_condicion(codigo):
    return verificar_valor(codigo, valores_comparacion_codigo_condicion)

def obtener_codigo_actividad(codigo):
    return verificar_valor(codigo, valores_comparacion_codigo_actividad)

def obtener_codigo_modalidad_contratacion(codigo):
    return verificar_valor(codigo, valores_comparacion_codigo_contratacion)
    
# Crear una lista para almacenar los nuevos registros
nuevos_registros = []

# Transformar y copiar los datos de holistor a los nuevos registros
for index, row in holistor_df.iterrows():
    apellido = extraer_apellido(row['nombr'])
    nombre = extraer_nombre(row['nombr'])
    nacionalidad = obtener_nacionalidad(row.get("cdnac_a"))
    fecha_nacimiento = convertir_formato(row.get('fenac'))
    sexo = obtener_sexo(row.get('cdsex_a'))
    condicion = obtener_condicion(row.get('cdescrip_f'))
    codigo_convenio = obtener_codigo_convenio(row.get('cdgre_c'))
    codigo_categoria = obtener_codigo_categoria(row.get('cdcat_a'), row.get('cdgre_c'))
    codigo_sindicato = obtener_codigo_sindicato(row.get('cdgre_a'))
    cobertura_seguro_vida = con_cobertura_seguro_vida(row.get('csegobli'))
    codigo_situacion_revista = obtener_codigo_situacion_revista(row.get('sijsi_a'))
    codigo_situacion_revista_1 = obtener_codigo_situacion_revista_1(row.get('sijsi_a'))
    codigo_condicion = obtener_codigo_condicion(row.get('sijco_a'))
    codigo_actividad = obtener_codigo_actividad(row.get('sijac_a'))
    codigo_contratacion = obtener_codigo_modalidad_contratacion(row.get('sijmo_a'))

    nuevo_registro = {
        'Código de tipos de servicio AFIP': '',
        'Número': row.get('legaj'),
        'Fecha de ingreso': convertir_formato(row.get('feing')),
        'Fecha de egreso de último tramo': '',
        'Apellido': apellido,
        'Apellido materno': '',
        'Nombre': nombre,
        'Código de tipo de documento': 'DU',
        'Número de documento': row.get('nrodoc'),
        'CUIL': row.get('ncuil'),
        'Nacionalidad': nacionalidad,
        'Fecha de nacimiento': fecha_nacimiento,
        'Sexo': sexo,
        'Estado civil': row.get('civil'),
        'Apellido del cónyuge': '',
        'Habilitado para Sueldos': 'Si',
        'Vinculado a Sueldos': 'Si',
        'Habilitado para Tango Empleados': 'No',
        'Habilitado para Tango Reportes': 'Si',
        'Confidencial': 'No',
        'Calle': row.get('dirca'),
        'Número de domicilio': row.get('dirnu'),
        'Piso': '',
        'Departamento': '',
        'Entre calles': '',
        'Torre': '',
        'Bloque': '',
        'Localidad': row.get('local'),
        'Código postal': row.get('cdpos_b'),
        'Provincia': '',
        'Correo electrónico': row.get('cemail'),
        'Correo electrónico personal': '',
        'Condición': condicion,
        'Inicio tiempo de servicio': '',
        'Tarea': '',
        'Código de puesto desempeñado': '',
        'Código de grupo jerárquico': '',
        'Código de convenio': codigo_convenio,
        'Código de categoría': codigo_categoria,
        'Fuera de convenio': es_fuera_convenio(codigo_convenio),
        'Código de situación de revista': codigo_situacion_revista,
        'Horas por día': '',
        'Días adicionales': '',
        'Administra vigencia de adicionales para vacaciones':'No',
        'Años de vigencia con beneficio de adicionales para vacaciones':'',
        'Lunes': 'Si',
        'Martes': 'Si',
        'Miércoles': 'Si',
        'Jueves': 'Si',
        'Viernes': 'Si',
        'Sábado':'No',
        'Domingo':'No',
        'Salud': 'Normal',
        'Incapacidad': '',
        'Código de obra social': '',
        'Número de afiliación de obra social': '',
        'Código de obra social a cargo de la empresa': '',
        'Código de sindicato': codigo_sindicato,
        'Número de afiliación de sindicato': '',
        'Es afiliado a sindicato': es_afiliado_sindicato(codigo_sindicato),
        'Sueldo en rango para alta': '',
        'Código de ART': '',
        'Última revisión': '',
        'Forma de pago': '',
        'Código de lugar de pago': '',
        'Código de banco': '',
        'Sucursal': '',
        'CBU': '',
        'Tipo de cuenta': '',
        'Número de cuenta': '',
        'Dígito verificador de cuenta': '',
        'Contribuyente cumplidor (Ley 27.260 / 2016)': 'No',
        'Código de modelo de asientos de sueldos': '',
        'Liquida impuesto a las ganancias': 'Si',
        'Beneficiario Ley 27.549': 'No',
        'Es personal de pozo':'No',
        'Fecha desde régimen teletrabajo': '',
        'Fecha hasta régimen teletrabajo': '',
        'Código de departamento': '',
        'Número de legajo del jefe': '',
        'Afecta archivo ASCII': 'Si',
        'Es legajo principal':'Si',
        'Día de situación de revista 1': '1',
        'Código de situación de revista 1': codigo_situacion_revista_1,
        'Día de situación de revista 2': '',
        'Código de situación de revista 2': '',
        'Día de situación de revista 3': '',
        'Código de situación de revista 3': '',
        'Código de situación de revista general': codigo_situacion_revista_1,
        'Código de condición': codigo_condicion,
        'Código de actividad': codigo_actividad,
        'Código de modalidad de contratación': codigo_contratacion,
        'Código de lugar de trabajo': '',
        'Código de jurisdicción': '13',
        'Observaciones para libro de sueldos digital': '',
        'Corresponde reducción': 'No',
        'Código de siniestrado': '00',
        'Capital de LRT': '0,00',
        'Con cobertura de seguro colectivo de vida obligatorio': cobertura_seguro_vida,
        'Días trabajados': '',
        'Horas trabajadas': '',
        'Aporte adicional': '',
        'Aporte voluntario': '',
        'Excedente de seguridad social': '0,00',
        'Número de concepto para importe adicional de obra social': '',
        'Número de concepto para aporte adicional de obra social': '',
        'Número de concepto de ajuste de seguridad social':'',
        'Número de concepto para incremento salarial Dtos 14/2020 y 56/2020':'',
        'Excedente de obra social': '0,00',
        'Método de cálculo de seguridad social': 'Por liquidación',
        'Importe de seguridad social': '0,00',
        'Aplica tope mínimo de seguridad social': 'Si',
        'Aplica tope máximo de seguridad social': 'Si',
        'Método de cálculo de obra social': 'Por liquidación',
        'Importe de obra social': '0,00',
        'Aplica tope mínimo de obra social': 'Si',
        'Aplica tope máximo de obra social': 'Si',
        'Rectificación de remuneración': '0,00',
        'Contribución tarea diferencial': '0,00',
        'Observaciones': '',
        'Número de concepto para jornada completa de sueldo': '',
        'Número de concepto para jornada completa de SAC': '',
        'Número de concepto para jornada completa de vacaciones': '',
    }
    nuevos_registros.append(nuevo_registro)

# Crear una lista para almacenar los registros de teléfonos
telefonos_registros = []

# Extraer y copiar los datos de teléfono a los registros de teléfonos
for index, row in holistor_df.iterrows():
    codigo_area = row.get('ctelarea')
    numero_telefono = row.get('ctelefono')
    ban = 0

    if pd.isna(codigo_area):
        codigo_area = '0'
        ban = 1
    if pd.isna(numero_telefono):
        numero_telefono = '0'
        ban = 1
    
    if ban == 0:
        registro_telefono = {
            'Número': row.get('legaj'),
            'Tipo de teléfono': obtener_tipo_telefono(row.get('ctipotelef_a')),
            'Teléfono': obtener_telefono(str(int(codigo_area)), str(int(numero_telefono)))
        }
    else:
        registro_telefono = {
            'Número': row.get('legaj'),
            'Tipo de teléfono': obtener_tipo_telefono(row.get('ctipotelef_a')),
            'Teléfono': ''
        }

    telefonos_registros.append(registro_telefono)

# Convertir las listas de registros a DataFrames
resultado_df = pd.DataFrame(nuevos_registros)
telefonos_df = pd.DataFrame(telefonos_registros)

# Guardar el DataFrame en el archivo de destino seleccionado
archivo_destino = seleccionar_destino()
if archivo_destino:
    try:
        with pd.ExcelWriter(archivo_destino) as writer:
            resultado_df.to_excel(writer, sheet_name='Legajos de sueldos', index=False)
            telefonos_df.to_excel(writer, sheet_name='Telefonos', index=False)
        print("Archivo de Excel generado correctamente.")
    except Exception as e:
        print(f"Error al escribir el archivo de Excel: {e}")
else:
    print("No se ha seleccionado ningún destino.")