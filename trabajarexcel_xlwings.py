
# -*- coding: utf-8 -*-


"""
Spyder Editor

This is a temporary script file
"""
# from docxtpl import DocxTemplate 
import pandas as pd
import math
import xlwings as xw    

# Definición de la API
from fastapi import FastAPI
app = FastAPI()

# LIBRERIAS de documentos y sistema operativo
import os
import os.path
from docxtpl import DocxTemplate

# para poder llamar a Outlook
import win32com.client
from jinja2 import Environment, FileSystemLoader

#pip install docx2pdf
from docx2pdf import convert
import base64
import glob

# move files
import shutil
import datetime

from pathlib import Path

cif = ""
global cifglobal 
global email
global nombre_solicitante
global opcionmenu
global proyectos

# esta es la ruta que se utiliza para el cif que estamos tratando
global rutaendiscoglobal 
global rutadescargas
rutadescargas = "C:\\Users\\Alfonso\\Downloads"
       
global programa
fecha_documento = ''
max_score = 100

outlook = win32com.client.Dispatch('Outlook.Application')

doc = DocxTemplate('Anexo 03. Ficha de Empresa 2022-v2.docx')
CF = pd.DataFrame()


os.chdir('C:\\desarrollo\\genanexcam')
path = 'C:\\desarrollo\\genanexcam'

# plantillas
doc = DocxTemplate('my_word_template.docx')


# almacenará el numero de fila en la pestaña alta_usuario   
global numero_fila
numero_fila = 0

# CREAMOS DATAFRAMES de EMPRESAS Y USUARIOS ASOCIADOS
global data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME

# ====================================== UTILIDAD TRABAJAR CON EXCEL XLWINGS ===========================================
############## Esta es la parte principal #######
#### con xlwings vamos a importar un excel
## seleccionamos el rango usado
## y lo convertimos a dataframe
## Después de darle vueltas, esta parece la forma de trabajarlo 17/01/2024

book = xw.Book(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')
hoja =  book.sheets['2'].used_range.value
celdas =  book.sheets['2']
df = pd.DataFrame(hoja)
df.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\resultado_xlwings.xlsx', index=False, header=True)
row_num = 2
col_num = 2
celdas.range(row_num, col_num).value = '01/01/2026'

# ===============================================================================================


# print(df)
pulsartecla = input("Pulse una tecla ...")

#def main():    
print ('saliendo ')
if __name__ == '__main__':
            # menu_principal(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
    print("hola")
        
# return "terminada la ejecucion"
        



    

