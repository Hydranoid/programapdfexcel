from tkinter import *    # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import PyPDF2
import xlsxwriter
import os

archivo_texto = 'texto_plano.txt'

def convertir_archivo(txt):


    Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    my_types = [('PDF Files','.pdf')]
    filename = askopenfilename()
    pdfFile = open(filename,'rb')

    pdfReader = PyPDF2.PdfFileReader(pdfFile)

    num_paginas = pdfReader.getNumPages()

#Abriendo txt y concatenando contenido del pdf
    f = open(txt, "x")

    for i in range(num_paginas):
        pagina = pdfReader.getPage(i)
        texto = pagina.extractText()
        f.write(texto)

    pdfFile.close()
    f.close()



#Definiendo la función principal para buscar cadenas en el archivo txt

def cadenas_multiples(archivo,lista_cadenas):
    lista_resultados = []
    with open(archivo , 'r') as read_obj:
        for line in read_obj:
            for cadena in lista_cadenas:
                if cadena in line:

                    lista_resultados.append(line.rstrip())

    return lista_resultados


def obteniendo_fechas_emision(archivo):
    lista_resultados = []
    with open(archivo , 'r') as read_obj:
        for line in read_obj:
            if 'Emisión' in line:
                    lista_resultados.append(next(read_obj))
    return lista_resultados



def obteniendo_rfcs_emisores(archivo):
    lista_resultados = []
    with open(archivo , 'r') as read_obj:
        for line in read_obj:
            if 'Emisor' in line:
                lista_resultados.append(next(read_obj))
    return lista_resultados
      
def obteniendo_rfcs_receptores(archivo):
    lista_resultados = []
    with open(archivo , 'r') as read_obj:
        for line in read_obj:
            if 'Receptor' in line:
                lista_resultados.append(next(read_obj))
    return lista_resultados 

def quitando_espacios(archivo):
    clean_lines = []
    with open(archivo, "r") as f:
        lines = f.readlines()
        clean_lines = [l.strip() for l in lines if l.strip()]

    with open(archivo, "w") as f:
        f.writelines('\n'.join(clean_lines))

def obteniendo_razones(archivo):
    """
    Funcion para obtener total o parcialmente las razones sociales
    """
    lista = []
    with open(archivo, 'r') as read_obj:
        for line in read_obj:

            if 'Social:' in line:
                lista.append(next(read_obj))

    lista_razones = [s.replace('\n','') for s in lista]
    lista_razones = [s.replace('RFC Receptor:','') for s in lista_razones]   
    lista_razones = [s.replace('RFC Emisor:','') for s in lista_razones]   
    return lista_razones


#Obteniendo las distintas listas de cadenas

#Lista de Leyenda Comprobante
convertir_archivo(archivo_texto)

#Estado del Comprobante
resultados_comprobante = cadenas_multiples(archivo_texto,['Vigente','Cancelado'])

#Lista de Precios

resultados_monto = cadenas_multiples(archivo_texto,['$'])
resultados_monto=[s.replace('$','') for s in resultados_monto]
resultados_monto=[s.replace(',','') for s in resultados_monto]
totales = []
for elem in resultados_monto:
    totales.append((float(elem)))


#Lista de Efectos

resultados_efecto = cadenas_multiples(archivo_texto,['Ingreso','Egreso','Nómina','Pago'])

#Quitando espacios al archivo txt para poder obtener los RFC's Emisores y Receptores, así como las Razones Sociales
quitando_espacios(archivo_texto)

#Lista de razones sociales
lista_razones = obteniendo_razones(archivo_texto)
razones_emisoras = lista_razones[0::2]
razones_receptoras = lista_razones[1::2]


#RFC's Emisores y Receptores

rfcs_emisores = obteniendo_rfcs_emisores(archivo_texto)
rfcs_receptores = obteniendo_rfcs_receptores(archivo_texto)


#Lista de Fechas

resultados_fecha = obteniendo_fechas_emision(archivo_texto)
resultados_fecha=[s.replace('T',' ') for s in resultados_fecha]



#Comprobación de datos obtenidos de funciones

# print(resultados_comprobante)
# print(resultados_efecto)
# print(resultados_fecha)
# print(totales)
# print(resultados_rfcs)
# print(razones_emisoras)
# print(razones_receptoras)


# Eliminando archivo txt
if os.path.exists(archivo_texto):
  os.remove(archivo_texto)
else:
  print("The file does not exist")



#----------------Creando archivo de Excel y asignando valores de listas---------------#

#Creando archivo nuevo

# raiz.mainloop()

i=0
while os.path.exists('tabla_%s.xlsx' % i):
    i+=1
workbook = xlsxwriter.Workbook('tabla_%s.xlsx' % i)
# workbook = xlsxwriter.Workbook(nombre)


#Creando hoja
worksheet = workbook.add_worksheet()
#Declarando el estilo de letra negrita
bold = workbook.add_format({'bold': True})
#Declarando el formato tipo moneda
money = workbook.add_format({'num_format': '$#,##0.00'})
#Añadiendo la cabeceera al archivo principal
worksheet.write('A1', 'Fecha de Emisión', bold)
worksheet.write('B1', 'Razón Social Emisor', bold)
worksheet.write('C1', 'RFC Emisor', bold)
worksheet.write('D1', 'Razón Social Receptor', bold)
worksheet.write('E1', 'RFC Receptor', bold)
worksheet.write('F1', 'Estado del Comprobante', bold)
worksheet.write('G1', 'Efecto del Comprobante', bold)
worksheet.write('H1', 'Total por Factura', bold)

#Definiendo funciones para agregar elementos al archivo excel

#Función para elementos Cadena
def insertar_elementos_excel(lista,col):
    row=1
    for elem in lista:
        worksheet.write(row, col,  elem  )
        row += 1

#Función para elementos numéricos
def insertar_montos_excel(lista,col):
    row=1
    for elem in lista:
        worksheet.write(row, col,  elem , money  )
        row += 1

#Invocando funciones e insertando elementos
insertar_elementos_excel(resultados_fecha,0)
insertar_elementos_excel(razones_emisoras,1)
insertar_elementos_excel(rfcs_emisores,2)
insertar_elementos_excel(razones_receptoras,3)
insertar_elementos_excel(rfcs_receptores,4)
insertar_elementos_excel(resultados_comprobante,5)
insertar_elementos_excel(resultados_efecto,6)
insertar_montos_excel(totales,7)

#Cerrando el libro recien creado
workbook.close()

