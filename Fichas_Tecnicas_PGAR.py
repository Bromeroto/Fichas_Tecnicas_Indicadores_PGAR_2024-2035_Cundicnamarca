# -*- coding: utf-8 -*-
"""
@author: Bernardo Romero-Torres
"""
%reset -f #Limpia el enviroment de Python

#Se cargan las librerias#
#-----------------------#

#Libreria de manejo de los datos#
import pandas as pd 

#Libreria para la creación de documentos PDF personalizados#
from reportlab.lib.pagesizes import letter #COnfigura el tamaño de página.
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph, Spacer, Image #clases necesarias para crear el documento PDF
from reportlab.lib import colors # para definir colores
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle #para obtener estilos de párrafo predefinidos
from reportlab.lib.units import inch #para especificar medidas en pulgadas
from reportlab.pdfgen import canvas #para generar el PDF y las páginas
from reportlab.platypus import PageTemplate, BaseDocTemplate, Frame #para crear plantillas de página y marcos
import PyPDF2 #para leer archivos PDF y concatenarlos

#Otras librerias
import os #Para determinar la ruta del proyecto
import glob #para trabajar con rutas de archivos y búsqueda de archivos



#Se configura la carpeta del proyecto: donde estan los inputs y se exportaran las fichas
os.chdir('H:/.shortcut-targets-by-id/1OY_UdTEsmb3rO2hLySPNadFfOD_jBt8p/Research & Public Policy/CAR  S&E PGAR') # Ruta del proyecto


#1. Se preparan los Datos#
#------------------------#

#Se Carga el Excel#
indicadores = pd.read_excel('Linea base prospectiva PGAR - fichas tecnicas PR.xlsx', 
                            sheet_name='Línea base', header = 0, dtype=str)


indicadores.drop(indicadores.columns[-1], axis=1, inplace= True) #Se elimina la Ultima columna que no nos sirve (En este caso)

#Se cambia el nombre de un indicador
indicadores.loc[indicadores['Nombre del Indicador']=='Índice de vulnerabilidad al desabasticimiento hídrico para año seco', 'Nombre del Indicador'] = 'Índice de vulnerabilidad al desabasticimiento hídrico' 


#Se configuran algunas Características del texto contenido en las fichas
style_column1 = ParagraphStyle(name='CustomStyle', fontName='Helvetica-Bold') #Definimos un estilo de párrafo llamado para la columna 1 con fuente Helvetica en negrilla
style_column2 = ParagraphStyle(name='CustomStyle', fontName='Helvetica', alignment=1) #Definimos un estilo de párrafo llamado para la columna 2 con fuente Helvetica
 
#2. Creacion de las fichas#
#-------------------------#
for i in range(indicadores.shape[0]):
    
    #Se crea una función que agrega una imagen en el encbezado de la hojael titulo.
    def agregar_titulo(canvas, doc):
        canvas.saveState()  # Guarda el estado actual del lienzo
        
        # Inserta la imagen como encabezado
        image_path = 'Header_CAR.png'  # La ruta de la imagen para el encabezado
        image_width = 6.3 * inch  # Ancho de la imagen en pulgadas
        image_height = 0.433 * inch  # Alto de la imagen en pulgadas
        x_position = 0.9842519685 * inch # Distancia desde el margen izquierdo en pulgadas
        y_position = letter[1] - 0.395 * inch - image_height  # Distancia desde el margen superior en pulgadas
        
        image = Image(image_path, width=image_width, height=image_height) # Creamos un objeto de imagen utilizando la ruta de la imagen, ancho y alto especificados
        image.drawOn(canvas, x_position, y_position) # Dibujamos la imagen en el lienzo especificando su posición en las coordenadas (x_position, y_position)
        
        #Se inserta el titulo en cada hoja
        canvas.setFont('Helvetica-Bold', 14) #Configura el estilo de la letra y el tamaño del texto de encabezado           
        canvas.drawCentredString(letter[0] / 2, letter[1] - 1.1 * inch, "Plan de Gestión Ambiental Regional (PGAR) 2024-2035") # Coloca texto centrada en el eje horizontal a la mitad de la página 
        canvas.drawCentredString(letter[0] / 2, letter[1] - 1.3 * inch, "Indicadores del Sistema de Seguimiento y Evaluación") # Coloca texto centrada en el eje horizontal a la mitad de la página 

        canvas.restoreState()  # Restaura el estado previo del lienzo
         
    row = indicadores.iloc[i][1:].rename_axis('variable').to_frame().reset_index().fillna('').rename(columns={i: 'valor'}) #Se toma cada la fila (indicador) para generar la ficha
    
    file = f"Ficha_tecnica_LB_{i+1}.pdf" # Se genera la ruta y el nombre del archivo pdf de cada ficha de indicador
    pdf_file = SimpleDocTemplate(file, pagesize=letter) # Creamos el archivo PDF de la ficha donde se creará la tabla con la información del indicador.
    
    ficha = [] #Se crea una lista vcía que contendrá la información de la ficha.
    
    # Se agrega una página de título
    frame = Frame(0, 0, letter[0], letter[1]) # Creamos un marco (frame) que abarca toda la página 
    template = PageTemplate(id='title', frames=[frame], onPage=agregar_titulo) # Definimos una plantilla de página  que utiliza el marco definido y llama a la función 'agregar_titulo' en cada página
    pdf_file.addPageTemplates([template]) # Agregamos la plantilla de página al documento PDF
    ficha.append(Spacer(1, inch+35)) # Agregamos un espacio en blanco en el contenido del documento
    
    # Creamos una lista de listas llamada 'data' para definir los contenidos de la tabla #
    # Agregamos filas con Paragraphs formateados con los estilos 'style_column1' y 'style_column2'
    # Cada fila corresponde a los datos del dataframe "row" que contoiene la información del indicador.
    data = [['PLAN DE GESTIÓN AMBIENTAL REGIONAL (PGAR) 2024-2035 CAR CUNDINAMARCA', ''],
            [Paragraph(row.iloc[0,0], style_column1),  Paragraph(row.iloc[0,1], style_column2)], 
            [Paragraph(row.iloc[1,0], style_column1),  Paragraph(row.iloc[1,1], style_column2)], 
            [Paragraph(row.iloc[2,0], style_column1),  Paragraph(row.iloc[2,1], style_column2)],
            [Paragraph(row.iloc[3,0], style_column1),  Paragraph(row.iloc[3,1], style_column2)],
            [Paragraph(row.iloc[4,0], style_column1),  Paragraph(row.iloc[4,1], style_column2)],
            [Paragraph(row.iloc[5,0], style_column1),  Paragraph(row.iloc[5,1], style_column2)],
            [Paragraph(row.iloc[6,0], style_column1),  Paragraph(row.iloc[6,1], style_column2)],
            [Paragraph(row.iloc[7,0], style_column1),  Paragraph(row.iloc[7,1], style_column2)],
            [Paragraph(row.iloc[8,0], style_column1),  Paragraph('Año ' + row.iloc[8,1], style_column2)],
            [Paragraph(row.iloc[9,0], style_column1),  Paragraph(row.iloc[9,1], style_column2)],
            [Paragraph(row.iloc[10,0], style_column1), Paragraph(row.iloc[10,1], style_column2)],
            [Paragraph(row.iloc[11,0], style_column1), Paragraph(row.iloc[11,1], style_column2)],
            [Paragraph(row.iloc[12,0], style_column1), Paragraph(row.iloc[12,1], style_column2)],
            [Paragraph(row.iloc[13,0], style_column1), Paragraph(row.iloc[13,1], style_column2)],
            [Paragraph(row.iloc[14,0], style_column1), Paragraph(row.iloc[14,1], style_column2)],
            [Paragraph(row.iloc[15,0], style_column1), Paragraph(row.iloc[15,1], style_column2)],
            [Paragraph(row.iloc[16,0], style_column1), Paragraph(row.iloc[16,1], style_column2)],
            [Paragraph(row.iloc[17,0], style_column1), Paragraph(row.iloc[17,1], style_column2)]]
    
    cuadro = Table(data, colWidths=[185, 380]) # Definimos el ancho de las columnas en la tabla
    
    # Definimos el estilo de la tabla utilizando la lista de tuplas 'TableStyle'
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),  # Fondo blanco en toda la tabla
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),  # Texto negro en toda la tabla
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Líneas de cuadrícula negras con un grosor de 1
        ('SPAN', (0, 0), (-1, 0)),  # Une las dos columnas de la primera fila
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),  # Fondo de la primera fila de color azul claro
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Estilo de letra negrita para la primera fila
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alineación al centro para la primera fila
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold', 13),  # Estilo de letra negrita para la primera columna a partir de la segunda fila
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica', 13),  # Estilo de letra para la segunda columna a partir de la segunda fila
        ('ALIGN', (1, 1), (1, -1), 'CENTER'),  # Alineación al centro para la segunda columna a partir de la segunda fila
    ])
    
    
    cuadro.setStyle(style) # Aplicamos el estilo a la tabla 'cuadro'
    ficha.append(cuadro) # Agregamos la tabla 'cuadro' al contenido del documento
    
    pdf_file.build(ficha) # Genera el documento PDF final utilizando la estructura y contenido definidos en 'ficha'


#3. Union de los Archivos en un solo archivo #
#--------------------------------------------#
files_fichas = [i.replace('\\', '/')  for i in glob.glob('Outputs/PDF_Fichas/Ficha_tecnica_*')] # Crea una lista de archivos que coinciden con el patrón 'Ficha_tecnica_*'. Archivos creado con las fichas.

pdf_merger = PyPDF2.PdfMerger() # Creamos un objeto PdfMerger de PyPDF2 para fusionar archivos PDF

# Iteramos a través de la lista de rutas de archivos PDF en 'files_fichas'
for pdf in files_fichas:
    pdf_merger.append(pdf) # Agregamos cada archivo PDF al objeto PdfMerger

pdf_merger.write('Outputs/PDF_Fichas/Ficha_Indicadores_LB_PGAR.pdf') # Escribimos el PDF fusionado en el archivo 'Ficha_Indicadores_LB_PGAR.pdf'

pdf_merger.close() # Cerramos el objeto PdfMerger después de haber creado el PDF fusionado


#4. Se eliminan los archivos creados#
#-----------------------------------#
# Iteramos a través de la lista de rutas de archivos PDF en 'files_fichas'
for f in files_fichas: 
    if os.path.exists(f): # Verificamos si el archivo existe
        os.remove(f) # Si existe, lo eliminamos
    else:
        print('Archivo no existe') # Si el archivo no existe, mostramos un mensaje indicando que no se encontró




