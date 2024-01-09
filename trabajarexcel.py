
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
book = xw.Book(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')
#data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME = pd.read_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0, converters= { 'fecha_documento_inicio_diagnostico': pd.to_datetime, 'fecha_diagnostico': pd.to_datetime, 'cp': pd.to_numeric, 'telefono_solicitante': pd.to_numeric})
# pestana_LISTA_ASESORE_GENERACION_DOCUME = workbook.sheets['LISTA-ASESORE-GENERACION-DOCUME']
pestana_LISTA_ASESORE_GENERACION_DOCUME = book.sheets['LISTA-ASESORE-GENERACION-DOCUME']
#data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME = book.sheets['LISTA-ASESORE-GENERACION-DOCUME'].used_range.value
#df = pd.DataFrame(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME, index=None) 
data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME = pestana_LISTA_ASESORE_GENERACION_DOCUME.range('A1').options(pd.DataFrame, expand='table').value
data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME =pd.read_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)
df = pd.DataFrame(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
#df = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME

# data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME =pd.read_excel(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)
# data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME =pd.read_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)
# data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDE = pd.read_excel('crusos-espana-emprende.xlsx', sheet_name='HOJA1', usecols = 'A:R',header = 0, converters= { 'Marca temporal': pd.to_datetime, 'Fecha de nacimiento': pd.to_datetime})
# data_oap_tab_alta_usuario = pd.read_excel('10044_Huelva_Registro_Actividades_OAs_202112_v2.5.xlsx', sheet_name='Alta_Usuario', usecols = 'E:N',header = 3)
#data_oap_tab_actividades_individuales = pd.read_excel('10044_Huelva_Registro_Actividades_OAs_202112_v2.5.xlsx', sheet_name='Actividades individuales', usecols = 'E:T',header = 3)
print(df)

# al exportar a excel elimina la primera fila con header = False

# funciona, genera un excel con el dataframe, sirve para ver que está bien formado :
# df.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\result.xlsx', index=False, header=False)

# vamos a poner excel en ONEDRIVE para poder sincronizar las cambios online
wb = xw.Book(r'c:\\users\\alfonso\\onedrive\\prueba\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')   #open your workbook

#Specify the value of the cell of the worksheet
# Data=wb.sheets['LISTA-ASESORE-GENERACION-DOCUME'].range()  
# Data=pd.DataFrame(Data)
# sheet1 = wb.sheets['LISTA-ASESORE-GENERACION-DOCUME'].used_range.value
# df = pd.DataFrame(sheet1)

# print(df)
pulsartecla = input("Pulse una tecla ...")

def prueba_modificar():
    df_tab_empresa_fila_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cifglobal]   
    print(" EL CIFFFFFF ES" + df_tab_empresa_fila_buscada)
    # pestana_LISTA_ASESORE_GENERACION_DOCUME.range('A1').value = df
    aaaa = input("seguir progreso como camareero ")
    return


# =================================== leer excel ===========
def read_data_pandas():
    data = pd.read_excel('example.xlsx')
    return data


def read_data_pandas_oap():
        # extraer datos a variable data1 de la oap
        # en la siguiente línea header, indica en qué linea empiezan los indices.
      
        print(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        return data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME, # data_oap_tab_alta_usuario, data_oap_tab_actividades_individuales

    

def crear_draftcorreo_19():
    send_account = None
    ruta = "C:/desarrollo/genanexcam/generated_anexo19_" + cifglobal + ".docx"
    convert(ruta, "C:/desarrollo/genanexcam/anexo19.pdf")
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'alfonso.raggio@camarahuelva.com':
            send_account = account
            break

    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add('alfonso.raggio@camarahuelva.com')
    mail_item.Subject = 'Test sending using particular account'
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = "C:/desarrollo/genanexcam/anexo19.pdf"
    attachment2 = "C:/desarrollo/genanexcam/ppi.pdf"
    mail_item.Attachments.Add(attachment)
    mail_item.Attachments.Add(attachment2)
    filename = "innocamaras.png"                   
    attachment3 =  'C:\\desarrollo\\genanexcam\\img\\innocamaras.png'
    

    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    


    mail_item.HTMLBody ='''
     <html><body>
	<p></p>
	<p>


</p><p>Buenos
días.</p>

<p>Adjunto les remito la documentación a firmar para que puedan comenzar la implantación
de los proyectos que nos indicó.</p>

<p>Ruego comprueben bien los datos antes de su firma ya que en
ella se estipulan los proyectos y cuantías económicas autorizadas por esta
Cámara. <b><u>Una vez firmado por ambas partes, podrán comenzar a ejecutar
gastos. Ningún gasto realizado con anterioridad a la firma de los documentos
por ambas partes, no será admitido.</u></b></p>

<p>Los
documentos&nbsp; en formato PDF pueden remitirlos firmados con <b><u>su
certificado digital. </u></b></p>

<p>A continuación recibirá un email indicando las obligaciones de publicidad de
la UE que deben cumplir.</p>

<p>Les recordamos que los plazas de ejecución y justificación son los
siguientes.</p>

<table style="background-color:#FFFFE0;" border="1">
 <tbody><tr>
  <td valign="top">
  <p><b>FECHA FIN DE
  EJECUCIÓN Y PAGOS</b></p>
  </td>
  <td valign="top">
  <p><b>31 de agosto
  de 2023</b></p>
  </td>
 </tr>
 <tr>
  <td valign="top">
  <p><b>PLAZO LÍMITE
  DE JUSTIFICACIÓN</b></p>
  </td>
  <td valign="top">
  <p><b>15 de
  septiembre de 2023</b></p>
  </td>
 </tr>
</tbody></table>

<p>Le adjuntamos también el documento donde se detalla algunas consideraciones
a tener en cuenta para el pago y justificación de las facturas. Recomendamos su
lectura para evitar incidencias en la justificación.</p>

<p>Saludos y gracias de antemano. </p>
''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return


def crear_draft_anexo18(dirdatos):
   global rutaendiscoglobal
   global proyectos
   rutafiles = rutaendiscoglobal
   adjuntos = []
   send_account = None
   # global rutaendiscoglobal 
   print ("RUTAS DISCO GLOBAL ??? " +  rutaendiscoglobal )

   """ 
   listddp = glob.glob(f"{rutafiles}/*ddp*.pdf")
   listdiag =  glob.glob(f"{rutafiles}/*diagnostico*.pdf")
   listanexo18 =  glob.glob(f"{rutafiles}/*anexo18*.docx") """

   listado_anexo18 = os.path.join(str(rutafiles), "*anexo18*.docx")
   listado_ddp = os.path.join(str(rutafiles), "*ddp*.pdf")
   listado_diag = os.path.join(str(rutafiles), "*diag*.pdf")
   listddp = glob.glob(listado_ddp)
   listdiag =  glob.glob(listado_diag)
   listanexo18 =  glob.glob(listado_anexo18)
   ruta_yaexiste_anexo18 = os.path.join(str(rutafiles), "*Anexo 18*.docx")
   anexo18_existente = glob.glob(ruta_yaexiste_anexo18)
   

   #print("listado de ddp " + listddp)
   #print("Diagnostico " + listdiag[0])
   print("ruta de los anexo 18 " + rutaendiscoglobal)
   wait = input("Press Enter to continue." )
   print("something")
   # shutil.move(rutaorigen + listdiag[0], rutaendiscoglobal + '/' + listdiag[0])
   for file in listddp:
       print("ARCHIVOS PARA DRAFT" + "  " + file + " " )
       adjuntos.append(file)    
  
   for account in outlook.Session.Accounts:
        print("nombre cuenta " + account.DisplayName)
        if account.DisplayName == 'alfonso.raggio@camarahuelva.com':
            print ("account.DisplayName "  + account.DisplayName )
            send_account = account
            break

   mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
   mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

   mail_item.Recipients.Add('alfonso.raggio@camarahuelva.com')
   if programa == "TICCAMARAS":
        mail_item.Subject = ' Diagnostico TICCámaras 2023' + nombre_solicitante
   elif programa == "INNOCAMARAS":
       mail_item.Subject = ' Diagnostico INNOCámaras 2023' + nombre_solicitante
       

       
   
   mail_item.BodyFormat = 2   # 2: Html format 
   for file in adjuntos:      
      mail_item.Attachments.Add(file)
   for file in listdiag:
      mail_item.Attachments.Add(file)
   print ( " LISTAAAAAAAAAAAAAAAAAAA ANEXO 18 ")
   print (listanexo18)
   for file in listanexo18:
     os.rename(file, f"{rutafiles}\\Anexo 18. Registro Prestación Servicio Fase I.docx")      
     
     print(f"RUTAR DEL ANEXO 18.===================. {rutafiles}/Anexo 18. Registro Prestación Servicio Fase I_a.docx")
  # si por lo que sea ya se habia renombrado se adjunta. solo se renombra si hay algo en listanexo18
   print (" anexo 18 existente " )
   a = input ("pulse una tecla...")
   ruta_yaexiste_anexo18 = os.path.join(str(rutafiles), "*Anexo 18*.docx")
   for file in anexo18_existente:
        mail_item.Attachments.Add(file)

   if programa == "INNOCAMARAS":
        filename = "\\inno\\innocamaras.png"                   
       
   elif programa == "TICCAMARAS" :
        filename = "\\tic\\ticcamaras.jpg"                 
   else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")
        return
           
    
   attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo' + filename

   print ( "attachment 3 " + attachment3)
   attach = mail_item.Attachments.Add(attachment3)  
    

    # data:image/png;base64,

   with open(attachment3, "rb") as image:
     image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

   mail_item.HTMLBody ='''
      <html><body> <p>


</p><p>Buenos
días.</p>

<p>&nbsp;Adjunto
le remito el diagnostico en el que hemos indicado como prioritarios los
siguientes proyectos:</p>

<ul type="disc">
 {{ proyectos }}
</ul>

<p>Le adjunto los DDP de cada proyecto
que son los documentos de Definición del Proyecto donde puede consultar los
detalles del mismo.</p>

<p>También
le adjunto el Anexo de Registro Prestación Servicio Fase I en formato editable
para que pueda rellenar y devolverme <b><u>con la encuesta cumplimentada</u></b>&nbsp;
en formato PDF y firmado preferiblemente<b><u> firmarlo con su certificado
digital.</u></b></p>

<p>Una
vez me remita la documentación firmada, podemos iniciar los trámites de la Fase
II.</p>

<p>El
siguiente paso será que nos envíen los presupuestos de los proyectos que
decidan poner en marcha. Los presupuestos deben ser acordes con las
especificaciones de los proyectos que se indican en los Documentos de&nbsp;
Definición de Proyectos adjuntos (DDP). Una vez recibidos, firmaremos la
participación en la Fase II y a partir de la firma pueden iniciar los
proyectos.</p>

<p>Saludos y gracias de antemano. </p>

<p>&nbsp;</p>

    ''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


    # <p></p><p></p><p>&nbsp;</p> </body></html>    
   mail_item.Save()
   return






# anexo-19-bueno
def crear_draft_anexo19():
    send_account = None
    ruta19 = rutaendiscoglobal + '\\' + 'generated_anexo19_' + cifglobal + '.docx'
    print( "fichero anexo 19 " + ruta19 )
    rutappi = rutaendiscoglobal + '\\' + 'ppi.docx'
    listppi = glob.glob(str(rutappi))
    print( "rutar anexo ppi " + rutappi )
    for file in listppi:
     print( "fichero a convertir " + file)
     a = input ( "fichero a convertir ")
     os.rename(file, f"{rutaendiscoglobal}\\ppi.docx")           
     
    
    #convert('generated_anexo19_'+ cifglobal + '.docx', rutaendiscoglobal + '/' + 'anexo19.pdf')
    convert(ruta19, rutaendiscoglobal + '/' + 'anexo19.pdf')
    convert(rutappi, rutaendiscoglobal + '/' + 'PPI.pdf')
    print(" PROGRAMA " + programa )
 
    if programa == "INNOCAMARAS":
        for account in outlook.Session.Accounts:   
          if account.DisplayName == 'innocamaras@camarahuelva.com':
            send_account = account
            break
    if programa == "TICCAMARAS":        
        for account in outlook.Session.Accounts:
           if account.DisplayName == 'ticcamaras@camarahuelva.com':           
            send_account = account
            break
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")      
        b = input(" pulse una tecla....")
           

    
    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Documentación Inicio Fase II ' + programa + ' 2023 - ' + nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = rutaendiscoglobal + '/' + 'anexo19.pdf'
    mail_item.Attachments.Add(attachment)
    attachment21 = rutaendiscoglobal + '/' + 'PPI.pdf'
    mail_item.Attachments.Add(attachment21)

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)    
    ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\tic\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_inno.pdf"
   
    path = str(os.getcwd())

    if programa == "INNOCAMARAS":
        filename = "\\inno\\innocamaras.png"            
        ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\inno\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_inno.pdf"
        # mail_item.Attachments.Add(path + "\\" + attachment22)               
        mail_item.Attachments.Add(ruta_gasto_elegible)               
       
    elif programa == "TICCAMARAS" :
        filename = "\\tic\\ticcamaras.jpg"     
        # attachment22 = rutaendiscoglobal + '/' + '05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_tic.pdf'
        ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\tic\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_tic.pdf"
        # mail_item.Attachments.Add(path + "\\" + attachment22)               
        mail_item.Attachments.Add(ruta_gasto_elegible)            
        
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")
        return
           
    
    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo' + filename
    

    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    




    mail_item.HTMLBody = '''
      <html><body>
	<p></p>
	<p>


</p><p>Buenos
días.</p>

<p>Adjunto les remito la documentación a firmar para que puedan comenzar la implantación
de los proyectos que nos indicó.</p>

<p>Ruego comprueben bien los datos antes de su firma ya que en
ella se estipulan los proyectos y cuantías económicas autorizadas por esta
Cámara. <b><u>Una vez firmado por ambas partes, podrán comenzar a ejecutar
gastos. Ningún gasto realizado con anterioridad a la firma de los documentos
por ambas partes, no será admitido.</u></b></p>

<p>Los
documentos&nbsp; en formato PDF pueden remitirlos firmados con <b><u>su
certificado digital. </u></b></p>

<p>A continuación recibirá un email indicando las obligaciones de publicidad de
la UE que deben cumplir.</p>

<p>Les recordamos que los plazas de ejecución y justificación son los
siguientes.</p>

<table style="background-color:#FFFFE0;" border="1">
 <tbody><tr>
  <td valign="top">
  <p><b>FECHA FIN DE
  EJECUCIÓN Y PAGOS</b></p>
  </td>
  <td valign="top">
  <p><b>31 de agosto
  de 2023</b></p>
  </td>
 </tr>
 <tr>
  <td valign="top">
  <p><b>PLAZO LÍMITE
  DE JUSTIFICACIÓN</b></p>
  </td>
  <td valign="top">
  <p><b>15 de
  septiembre de 2023</b></p>
  </td>
 </tr>
</tbody></table>

<p>Le adjuntamos también el documento donde se detalla algunas consideraciones
a tener en cuenta para el pago y justificación de las facturas. Recomendamos su
lectura para evitar incidencias en la justificación.</p>

<p>Saludos y gracias de antemano. </p>
''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return







def creardraft_anexo20anexo21():
    send_account = None
    ruta20 = rutaendiscoglobal + '\\' + 'Anexo 20.docx'
    ruta20_new = rutaendiscoglobal + '\\' + 'Anexo 20. Registro Prestación Servicio y Seguimiento - Fase II.docx'
    os.rename(str(ruta20), str(ruta20_new))
    ruta20 = ruta20_new
    print( "fichero anexo 20 " + ruta20 )
    ruta21 = rutaendiscoglobal + '\\' + 'Anexo 21.docx'
    ruta21_new = rutaendiscoglobal + '\\' + 'Anexo 21. Memoria Ejecución Proyecto y Cuestionarios.docx'
    os.rename(str(ruta21), str(ruta21_new))
    listppi = glob.glob(str(ruta21))
    print( "rutar anexo ppi " + ruta21 )
    """  for file in listppi:
     print( "fichero a convertir " + file)
     a = input ( "fichero a convertir ")
     os.rename(file, f"{rutaendiscoglobal}\\ppi.docx")           
      """
    
    #convert('generated_anexo19_'+ cifglobal + '.docx', rutaendiscoglobal + '/' + 'anexo19.pdf')
    convert(ruta20, rutaendiscoglobal + '/' + 'Anexo 20.pdf')
    #convert(rutappi, rutaendiscoglobal + '/' + 'PPI.pdf')
    print(" PROGRAMA " + programa )
 
    if programa == "INNOCAMARAS":
        for account in outlook.Session.Accounts:   
          if account.DisplayName == 'innocamaras@camarahuelva.com':
            send_account = account
            break
    if programa == "TICCAMARAS":        
        for account in outlook.Session.Accounts:
           if account.DisplayName == 'ticcamaras@camarahuelva.com':           
            send_account = account
            break
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")      
        b = input(" pulse una tecla....")
           

    
    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Justificación Programa InnoCámaras 2023 - Memoria final y Registro de prestación de servicios - PLAZOS Y PROCEDIMIENTO - MUY IMPORTANTE - ' + nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = rutaendiscoglobal + '/' + 'Anexo 20. Registro Prestación Servicio y Seguimiento - Fase II.docx'
    
    mail_item.Attachments.Add(attachment)
    attachment21 = rutaendiscoglobal + '/' + 'Anexo 21. Memoria Ejecución Proyecto y Cuestionarios.docx'
    mail_item.Attachments.Add(attachment21)

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    #print("la ruta de trabajo es " + path)    
    #ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\tic\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_inno.pdf"
   
    path = str(os.getcwd())


  #doc = DocxTemplate('anexo18_T.docx')    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo11b\\tic\\"
        attachment4 = ruta + 'Anexo 11_1 TIC Declaración jurada ayudas y gastos.docx'   
        mail_item.Attachments.Add(attachment4) 
        ruta = path + "\\" + "template_doc\\anexo29_ccc\\tic\\"
        attachment6 = ruta + 'Anexo 29. Formulario CCC - EMPRESAS .docx'
        mail_item.Attachments.Add(attachment6) 
        
       
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo11b\\inno\\"
        attachment4 = ruta + 'Anexo 11_1 INNO Declaración jurada ayudas y gastos.docx'
        mail_item.Attachments.Add(attachment4) 
        ruta = path + "\\" + "template_doc\\anexo29_ccc\\inno\\"
        attachment6 = ruta + 'Anexo 29. Formulario CCC - EMPRESAS .docx'
        mail_item.Attachments.Add(attachment6) 
       

    attachment5 = path + "\\" + "Justific@-Guia_de_usuario.pdf"
    mail_item.Attachments.Add(attachment5) 

    
    if programa == "INNOCAMARAS":
        filename = "\\inno\\innocamaras.png"  
    elif programa == "TICCAMARAS" :
        filename = "\\tic\\ticcamaras.jpg"     
                
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")
        return
            

    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo' + filename
    

    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    




    mail_item.HTMLBody = '''

<html><body> <p>



</p><p>Estimado
empresari@: </p>

<p>Como ya saben, hoy 31 de agosto finaliza el plazo para la <b><u>ejecución,
facturación y pago</u></b> de los proyectos puestos en marcha y hay una serie <b><u>documentación
y plazos</u></b> muy importantes que hay que cumplir y unas pautas a seguir
para el cumplimiento del procedimiento del programa.</p>

<p>Este año tenemos un plazo muy reducido para la justificación, por lo que hay
que ser especialmente riguroso con las fechas que tienen que cumplir.</p>

<p>En este email se le adjuntan en formato editable la siguiente documentación:</p>

<ul type="disc">
 <li>Anexo 20. Registro Prestación Servicio y Seguimiento -
     Fase II. Deben revisar la información y si es necesario pueden
     modificarla. Deben remitírnosla por correo electrónico en formato PDF
     firmada con el certificado digital.</li>
 <li>Anexo 21. Memoria Ejecución Proyecto y Cuestionarios.
     Deben revisar la información y cumplimentar los apartados de
     &quot;Encuesta de satisfacción y evaluación del impacto&quot; así como el
     de Cuestionario de caso de éxito&quot; Deben remitírnosla por correo
     electrónico en formato PDF firmada con el certificado digital.</li>
 <li>Anexo 11.1. Declaración jurada ayudas y gastos. Deben
     cumplimentarla y remitírnosla por correo electrónico en formato PDF firmada
     con el certificado digital.</li>
</ul>

<p>Pasamos a describirles el procedimiento de la justificación y plazos
máximos.</p>

<ul type="disc">
 <li><b>4 de septiembre de 2023</b>, finalización del plazo
     para <b>remitirnos por email los anexos 20, 21 y declaración jurada</b>
     firmados con certificado digital.</li>
 <li><b>El 15 de septiembre de 2023</b>, finaliza el plazo para
     subir toda la documentación a la plataforma Justific@.&nbsp; <b><u>Una vez
     finalizado este plazo, no será posible aportar nueva documentación, solo
     será posible subsanar la ya aportada si es requerido por parte de los
     auditores. </u></b>Este es un proceso fácil pero que no podéis dejar para
     última hora ya que tenemos poco tiempo y cualquier problema informático os
     puede dejar fuera de plazo. </li>
 <li>Ya podéis ir aportando documentación, deberían comenzar
     a subirla e ir probando la herramienta, además en ella iréis consultando
     el estado de la revisión por si fuera necesario subsanar alguna
     documentación.&nbsp;</li>
 <li>El enlace a la herramienta Justific@ es<b><i>&nbsp; <a href="https://justifica.camaras.es/ayudas">https://justifica.camaras.es/ayudas</a>
     </i></b>A efectos<b><i> </i></b>de conocer cómo se utiliza la mencionada
     plataforma se pone a disposición un video donde se explica el
     procedimiento. También adjuntamos a este correo el manual de usuario.<b><i>
     </i></b>Puede visualizar el video en <a href="https://www.youtube.com/watch?v=V3HJNMhvLc8">https://www.youtube.com/watch?v=V3HJNMhvLc8</a>
     </li>
</ul>

<p>A modo de resumen, la documentación que tenéis que aportar en todo el
proceso de justificación es la siguiente:</p>

<ul type="disc">
 <li>Por email antes del 4 de septiembre de 2023</li>
 <ul type="circle">
  <li>Anexo 20. Registro
      Prestación Servicio y Seguimiento - Fase II.</li>
  <li>&nbsp;Anexo 21. Memoria
      Ejecución Proyecto y Cuestionarios.</li>
  <li>Anexo 11.1. Declaración
      jurada ayudas y gastos.</li>
 </ul>
 <li>A través de la herramienta Justific@ antes del 15 de
     septiembre de 2023</li>
 <ul type="circle">
  <li>Facturas.</li>
  <li>Justificantes de pago.
      Orden de transferencia + Extracto bancario. Puede ser sustituido por
      certificado bancario de la transferencia realizada o cualquier otro
      documento que acredite el pago.</li>
  <li>Evidencias de los
      gastos realizados.</li>
  <li>Evidencias del
      cumplimiento de la publicidad del FEDER. Texto en la página web y cartel
      A3. (Se envió email explicativo)</li>
  <li>Documento de
      identificación financiera. Certificado de titularidad de la cuenta
      bancaria de abono de la ayuda. Se adjunta un modelo, aunque se puede
      aportar el certificado que os proporcione el banco.</li>
 </ul>
</ul>

<p>Como siempre nos tiene a tu disposición para resolver cualquier consulta o
duda.</p>

<p>Un cordial saludo.</p>




 
''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return










def enviar_anexo19():
    print("opcion seleccionada enviar anexo 19")
        
    # mail = outlook.CreateItem(0) 
    mail =  outlook.CreateItem(0) 
    #mail =  outlook.CreateItemFromTemplate('C:\\temp\\prueba-oft.oft')
                                                

    mail.To = 'alfonso.raggio@camarahuelva.com'
    mail.Subject = ' asunto de correo '


    mail.HTMLBody =  """
  <html><body>
	<p></p>
	<p>


</p><p>Buenos
días.</p>

<p>AdjuntoRRRRRRRRRRRRRRRRRRRRRRR 
les remito la documentación a firmar para que puedan comenzar la implantación
de los proyectos que nos indicó.</p>

<p>Ruego comprueben bien los datos antes de su firma ya que en
ella se estipulan los proyectos y cuantías económicas autorizadas por esta
Cámara. <b><u>Una vez firmado por ambas partes, podrán comenzar a ejecutar
gastos. Ningún gasto realizado con anterioridad a la firma de los documentos
por ambas partes, no será admitido.</u></b></p>

<p>Los
documentos&nbsp; en formato PDF pueden remitirlos firmados con <b><u>su
certificado digital. </u></b></p>

<p>A continuación recibirá un email indicando las obligaciones de publicidad de
la UE que deben cumplir.</p>

<p>Les recordamos que los plazas de ejecución y justificación son los
siguientes.</p>

<table style="background-color:#FFFFE0;" border="1">
 <tbody><tr>
  <td valign="top">
  <p><b>FECHA FIN DE
  EJECUCIÓN Y PAGOS</b></p>
  </td>
  <td valign="top">
  <p><b>31 de agosto
  de 2023</b></p>
  </td>
 </tr>
 <tr>
  <td valign="top">
  <p><b>PLAZO LÍMITE
  DE JUSTIFICACIÓN</b></p>
  </td>
  <td valign="top">
  <p><b>15 de
  septiembre de 2023</b></p>
  </td>
 </tr>
</tbody></table>

<p>Le adjuntamos también el documento donde se detalla algunas consideraciones
a tener en cuenta para el pago y justificación de las facturas. Recomendamos su
lectura para evitar incidencias en la justificación.</p>

<p>Saludos y gracias de antemano. </p>

<!-- HTML Code -->
<img src="C://desarrollo//genanexcam//img//innocamaras.png">
  </figure><p></p><p></p> </body></html>

    """

    mail.BodyFormat = 2

    mail.Send()


    return
    

def crear_draft_anexo20():
    send_account = None
    #ruta20 = rutaendiscoglobal + '\\' + 'anexo 20'  + '.docx'
    #print( "fichero anexo 19 " + ruta20 )
    rutaanexo20 = rutaendiscoglobal + '\\' + 'anexo 20.docx'
    listanexo20 = glob.glob(str(rutaanexo20))
    print( "rutar anexo ppi " + rutaanexo20 )
    for file in listanexo20:
     print( "fichero a convertir " + file)
     a = input ( "fichero a convertir ")
     os.rename(file, f"{rutaendiscoglobal}\\ppi.docx")           
     
    
    #convert('generated_anexo19_'+ cifglobal + '.docx', rutaendiscoglobal + '/' + 'anexo19.pdf')
    convert(rutaanexo20, rutaendiscoglobal + '/' + 'anexo19.pdf')
    #convert(rutappi, rutaendiscoglobal + '/' + 'PPI.pdf')
    print(" PROGRAMA " + programa )
 
    if programa == "INNOCAMARAS":
        for account in outlook.Session.Accounts:   
          if account.DisplayName == 'innocamaras@camarahuelva.com':
            send_account = account
            break
    if programa == "TICCAMARAS":        
        for account in outlook.Session.Accounts:
           if account.DisplayName == 'ticcamaras@camarahuelva.com':           
            send_account = account
            break
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")      
        b = input(" pulse una tecla....")
           

    
    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Justificación Programa TICCámaras 2023 - Memoria final y Registro de prestación de servicios - ' +  nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = rutaendiscoglobal + '/' + 'Anexo 20.pdf'
    mail_item.Attachments.Add(attachment)
    attachment21 = rutaendiscoglobal + '/' + 'Anexo 21.docx'
    mail_item.Attachments.Add(attachment21)

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)    
    ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\tic\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_inno.pdf"
   
    path = str(os.getcwd())

    if programa == "INNOCAMARAS":
        filename = "\\inno\\innocamaras.png"            
        ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\inno\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_inno.pdf"
        # mail_item.Attachments.Add(path + "\\" + attachment22)               
        mail_item.Attachments.Add(ruta_gasto_elegible)               
       
    elif programa == "TICCAMARAS" :
        filename = "\\tic\\ticcamaras.jpg"     
        # attachment22 = rutaendiscoglobal + '/' + '05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_tic.pdf'
        ruta_gasto_elegible = path + "\\" + "template_doc\\gasto_elegible\\tic\\05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas_tic.pdf"
        # mail_item.Attachments.Add(path + "\\" + attachment22)               
        mail_item.Attachments.Add(ruta_gasto_elegible)            
        
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")
        return
           
    
    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo' + filename
    

    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    




    mail_item.HTMLBody = '''

<html><body> <p>


</p><p>Buenos
días. </p>

<p>De cara a la justificación de los gastos presentados, le adjunto la
siguiente documentación que nos debe entregar a la mayor brevedad posible. Esta
documentación no pueden subir a la herramienta Justific@, por lo que deben
remitírnosla por email.</p>



<ol type="1">
 <li>Memoria
     Ejecución Fase II y cuestionarios. En este documento se plasman los gastos
     y cuantías presentadas para la justificación.<u> <b>Es necesario que se
     rellenen la encuesta de satisfacción y cuestionario de caso de éxito.</b></u>&nbsp;
     Se lo enviamos en formado editable pero<b><u> debe remitírnoslo en PDF
     firmado con el certificado digital antes del 31 de enero de 2023.</u></b></li>
 <li>Registro
     de prestación de servicio y seguimiento FII.&nbsp; documento en el que
     aparecen las acciones realizadas durante toda la Fase II. <b><u>Debe&nbsp;remitírnoslo en PDF firmado con el certificado digital antes del 31 de agosto de 2023.</u></b></li>
</ol>

<p>Un cordial saludo.</p>


<p></p></body></html>

''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'




# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return










def valuestovar_anexo19_marcado_ara_borrar(df):
    df_tab_empresa_fila_buscada = df
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    documento_solicitante = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    tratamiento_tecnico  = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    

    
     

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'tratamiento_tecnico' : tratamiento_tecnico,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                

    


                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
                
    }
   

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)

    print("la ruta de trabajo es " + path)    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo19\\tic\\anexo19_T.docx")    
    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo19\\tic\\"
        doc = DocxTemplate(ruta + 'anexo19_T.docx')   
        doc.render(context)         
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo19\\inno\\"
        doc = DocxTemplate(ruta + 'anexo19_I.docx')        
        doc.render(context)  
           
    generado = rutaendiscoglobal + "\\"
    # concatenear cadenas
    anexo19_pdfname =  f"generated_anexo19_{cifglobal}.pdf"
    anexo19_wordname = f"generated_anexo19_{cifglobal}.docx"
    rutacompletaendiscoword = generado + anexo19_wordname
    rutacompletaendiscopdf = generado + anexo19_pdfname
    
    #file_target = f"{rutadisco}\\{file_no_ext}.pdf"
    #convert(f"{rutadisco}\\{os.path.basename(listdiag[0])}", file_target)
    print ( " ruta word " + rutacompletaendiscoword)
    print ( " ruta pdf " + rutacompletaendiscopdf)

    a = input (" Pulse una tecla ....  " )

    # CONVERTIR A PDF el ANEXO 19
    # onvert(f"{rutacompletaendiscoword}", f"{rutacompletaendiscopdf}")    
    
    #shutil.move(str(generado), f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(rutacompletaendiscoword)           
    return





def valuestovar_anexo19(df):
    """ df_tab_empresa_fila_buscada = df
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    documento_solicitante = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    tratamiento_tecnico  = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    

    
     

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'tratamiento_tecnico' : tratamiento_tecnico,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                

    






                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
                
    }
   
 """
    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)

    print("la ruta de trabajo es " + path)    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo19\\tic\\anexo19_T.docx")    
    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo19\\tic\\"
        doc = DocxTemplate(ruta + 'anexo19_T.docx')   
        doc.render(df)         
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo19\\inno\\"
        doc = DocxTemplate(ruta + 'anexo19_I.docx')        
        doc.render(df)  
           
    generado = rutaendiscoglobal + "\\"
    # concatenear cadenas
    anexo19_pdfname =  f"generated_anexo19_{cifglobal}.pdf"
    anexo19_wordname = f"generated_anexo19_{cifglobal}.docx"
    rutacompletaendiscoword = generado + anexo19_wordname
    rutacompletaendiscopdf = generado + anexo19_pdfname
    
    #file_target = f"{rutadisco}\\{file_no_ext}.pdf"
    #convert(f"{rutadisco}\\{os.path.basename(listdiag[0])}", file_target)
    print ( " ruta word " + rutacompletaendiscoword)
    print ( " ruta pdf " + rutacompletaendiscopdf)

    a = input (" Pulse una tecla ....  " )

    # CONVERTIR A PDF el ANEXO 19
    # onvert(f"{rutacompletaendiscoword}", f"{rutacompletaendiscopdf}")    
    
    #shutil.move(str(generado), f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(rutacompletaendiscoword)           
    return


def valuestovar_anexo20(df):
    
    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    
    print("la ruta de trabajo es " + path)    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo20\\tic\\anexo20_T.docx")    
    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo20\\tic\\"
        doc = DocxTemplate(ruta + 'anexo20_T.docx')   
        doc.render(df)         
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo20\\inno\\"
        doc = DocxTemplate(ruta + 'anexo20_I.docx')        
        doc.render(df)  
           
    generado = rutaendiscoglobal + "\\"
    # concatenear cadenas
    #anexo20_pdfname =  f"G_anexo20_{cifglobal}.pdf"
    anexo20_pdfname =  f"Anexo 20.pdf"
    #anexo20_wordname = f"G_anexo20_{cifglobal}.docx"
    anexo20_wordname = f"Anexo 20.docx"
    rutacompletaendiscoword = generado + anexo20_wordname
    rutacompletaendiscopdf = generado + anexo20_pdfname
    
    #file_target = f"{rutadisco}\\{file_no_ext}.pdf"
    #convert(f"{rutadisco}\\{os.path.basename(listdiag[0])}", file_target)
    print ( " ruta word " + rutacompletaendiscoword)
    print ( " ruta pdf " + rutacompletaendiscopdf)

    a = input (" Pulse una tecla ....  " )

    # CONVERTIR A PDF el ANEXO 19
    # onvert(f"{rutacompletaendiscoword}", f"{rutacompletaendiscopdf}")    
    
    #shutil.move(str(generado), f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(rutacompletaendiscoword)  

    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo20\\tic\\anexo21_T.docx")    
    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo21\\tic\\"
        doc = DocxTemplate(ruta + 'anexo21_T.docx')   
        doc.render(df)         
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo21\\inno\\"
        doc = DocxTemplate(ruta + 'anexo21_I.docx')        
        doc.render(df)  
           
    generado = rutaendiscoglobal + "\\"
    # concatenear cadenas
    anexo21_pdfname =  f"Anexo 21.pdf"
    anexo21_wordname = f"Anexo 21.docx"
    rutacompletaendiscoword = generado + anexo21_wordname
    rutacompletaendiscopdf = generado + anexo21_pdfname
    
    #file_target = f"{rutadisco}\\{file_no_ext}.pdf"
    #convert(f"{rutadisco}\\{os.path.basename(listdiag[0])}", file_target)
    print ( " ruta word " + rutacompletaendiscoword)
    print ( " ruta pdf " + rutacompletaendiscopdf)

    a = input (" Pulse una tecla ....  " )

    # CONVERTIR A PDF el ANEXO 19
    # onvert(f"{rutacompletaendiscoword}", f"{rutacompletaendiscopdf}")    
    
    #shutil.move(str(generado), f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(rutacompletaendiscoword)           

    return





def generar_ficha_anexo18(dirdatos):   
    # cif=input()
    cif_a_buscar = cifglobal
 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   

    valuestovar_anexo18(df_tab_empresa_fila_buscada)
    """  
    # se localiza por NIF en el dataframe que corresponde al tab ALTA USUARIO
    # df_tab_usuario_buscado = df_tab_alta_usuarios.loc[df_tab_alta_usuarios['documento_solicitante'] == cif_a_buscar]
    
    #print(" kkkk" + df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_string())

    print("empresa encontrada para imprimir ")
    #print(df_tab_empresa_fila_buscada)
    print("USUARIO BUSCADO ============> usuario encontrada para imprimir ")
    # print(df_tab_usuario_buscado)
    # buscar valor en un dataframe por columna
    #
    # con at df1.at[0,'randomcolumn'] => el AT NO FUNCIONA
    # He probado con iloc y funciona
    #==========================================
    #================= DATOS DE EMPRESA ==========
    # ============================================
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                

    






                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
                
    }

   

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)
    doc = DocxTemplate('anexo19.docx')        
    doc.render(context)
    doc.save('generated_anexo19_'+ cifglobal + '.docx')           
    return
 """

def generar_ficha_anexo18_2(dirdatos):   
    # cif=input()
    #cif_a_buscar = cifglobal
    print( " denttro del ANEXO 18 2  ")
    print(dirdatos)

    a = input ( "en generar anexo 18 2 ")
    #context = {'mi_nombre': 'Fran tarci'}
    #df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    #print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    #df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   

    valuestovar_anexo18_2(dirdatos)
    

def valuestovar_deca(df):
    df_tab_empresa_fila_buscada = df
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']

    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    fasesqueparticipa = ""
    if fases == "FI+FII":
        fasesqueparticipa = "Fase de Diagnóstico y Fase de implantación"
    elif fases == "FI":
        fasesqueparticipa = "Fase de Diagnóstico Asistido"
    elif fases == "FII": 
        fasesqueparticipa = "Fase de implantación"
   
       

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))
    # r_admision_formateada = ""
    r_admision_formateada = r_admision.strftime("%d/%m/%Y")
    r_admision = r_admision_formateada
    # print( " R. admision " + r_admision)
    #print( " R. admision formateado " + r_admision_formateada)


    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                'r_admision' : r_admision,
                'fases_que_participa' : fasesqueparticipa

                

                
                
    }
   
    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)



    doc = DocxTemplate('deca_T.docx')    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\deca\\tic\\"
        doc = DocxTemplate(ruta + 'Deca_T.docx')    
        doc.render(context)  
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\deca\\inno\\"
        rutaparainno = f"{ruta}Deca_I.docx"
        print ( " ruta para inno  " +  rutaparainno )
        doc = DocxTemplate(ruta + 'Deca_I.docx')   
        doc.render(context)  
    
    #print(f"{ruta}Deca_I.docx")
        
    

    generado = rutaendiscoglobal + "\\"
    generado = generado + f"deca_{cifglobal}.docx"
    doc.save(generado)           
    # Fase de Asesoramiento individualizado y Fase de Implantación
   
    doc = DocxTemplate('Anexo16.docx')    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\deca\\tic\\"
        """  doc = DocxTemplate(ruta + 'Deca_T.docx')    
        doc.render(context)  
        doc.save(generado)   """          
        doc = DocxTemplate(ruta + 'Anexo16_T.docx')    
        doc.render(context)  
               
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\deca\\inno\\"
        """  doc = DocxTemplate(ruta + 'Deca_I.docx')
        doc.render(context)  
        doc.save(generado)  """           
        doc = DocxTemplate(ruta + 'Anexo16_I.docx')    
        doc.render(context)  

    
    generado = rutaendiscoglobal + "\\"
    generado = generado + f"Anexo_16{cifglobal}.docx"
        
    
    print (" GENERADO " + generado)

    a = input ( " Pulse una tecla ...") 

    doc.save(generado)           
    return



def valuestovar_anexo18(df):
    df_tab_empresa_fila_buscada = df
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    documento_solicitante = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    tratamiento_tecnico  = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))
    print ( ' email - correo electronico ' + str(email))

    fecha_inicio_diagnostico_formateada = fecha_documento_inicio_diagnostico.strftime("%d/%m/%Y")
    fecha_fin_diagnostico_formateada = fecha_diagnostico.strftime("%d/%m/%Y")
    print (' FECHA FORMATEADA '   + fecha_fin_diagnostico_formateada  )    


    a = input ( ' pulsa una tecla ')
    dni_formateado = dni_tecnico[:2] + "." + dni_tecnico[2:5] + "." + dni_tecnico[5:8] + "-" + dni_tecnico[8]
    print ( " DNI FORMATEADO " + dni_formateado)
    A = input( "PULSE UNA TECLA ....")


    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'documento_solicitante' : cif_empresa,
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,
                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : dni_formateado,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,                
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_formateado,
                'tratamiento_tecnico' : tratamiento_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_inicio_diagnostico_formateada,
                'fecha_diagnostico' : fecha_fin_diagnostico_formateada,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_fin_diagnostico_formateada,
                 # str(fecha_registro_participacion_fase_i_anexo_18)
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19': fecha_registro_participacion_en_fase_ii_anexo_19,
                
                 #.strftime('%d/%m/%Y'),
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                'num_expediente' : num_expediente,
                'r_admision' : r_admision


      

    






                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
    
    }

    

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo18\\tic\\anexo18_T.docx")
   
    #doc = DocxTemplate('anexo18_T.docx')    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo18\\tic\\"
        doc = DocxTemplate(ruta + 'anexo18_T.docx')    
        doc.render(context)  
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo18\\inno\\"
        doc = DocxTemplate(ruta + 'anexo18_I.docx')
        doc.render(context)  

    generado = rutaendiscoglobal + "\\"
    generado = generado + f"anexo18_{cifglobal}.docx"
   
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(generado)           
    return

def generar_ficha_anexo19_pendiente_borrar_duplicado():   
    # cif=input()
    cif_a_buscar = cifglobal

 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   

    valuestovar_anexo19(df_tab_empresa_fila_buscada)
    

#######################################################
## ------------------ GENERAR ANEXO 18 ----------------
#######################################################
def valuestovar_anexo18_2(df):

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)    
    doc = DocxTemplate(path + "\\" + "template_doc\\anexo18\\tic\\anexo18_T.docx")
   
    #doc = DocxTemplate('anexo18_T.docx')    
    if programa == "TICCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo18\\tic\\"
        doc = DocxTemplate(ruta + 'anexo18_T.docx')    
        doc.render(df)  
    if programa == "INNOCAMARAS":
        ruta = path + "\\" + "template_doc\\anexo18\\inno\\"
        doc = DocxTemplate(ruta + 'anexo18_I.docx')
        doc.render(df)  

    generado = rutaendiscoglobal + "\\"
    generado = generado + f"anexo18_{cifglobal}.docx"
   
    
    print (" GENERADO " + generado)
    a = input ( " Pulse una tecla ...") 

    doc.save(generado)           
    return


#######################################################
## ------------------ GENERAR ANEXO 19 ----------------
#######################################################

def generar_ficha_anexo19(dirdatos):   
    # cif=input()
    cif_a_buscar = cifglobal

 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    # df_tab_empresa_fila_buscada
    valuestovar_anexo19(dirdatos)
    # df_tab_empresa_fila_buscada['poblacion'] = 'Aljaraquerrrrr'
    # write again the excel file

    

#######################################################
## ------------------ GENERAR ANEXO 20 ----------------
#######################################################

def generar_ficha_anexo20(dirdatos):   
    # cif=input()
    cif_a_buscar = cifglobal
    

 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    # df_tab_empresa_fila_buscada
    valuestovar_anexo20(dirdatos)
    
def procbuscarcif2():
    print('Introduzca CIF ')
    
    # cif=input()
    #cif = '44237153B'
    global cifglobal
    cifglobal = input(" Introducir CIF ")
    
    # ========= Si lo encontramos, creamos un dataframe con esa fila para la pestaña EMPRESA y USUARIO  
    # == Cargamos en DF datos TAB EMPRESA
    print('')
    #df_empresa_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cif]      
    
    print('El CIF es ' + cif)

    print( "" )
    # dfAltaEmpresas, dfAltaUsuario = modify_data_pandas(cif)
    #,data_oap_tab_alta_usuario)
    #update_spreadsheet(path, CF, 1, 1, 'Alta_Usuario') #Write to sheet1 starting from row 20 and column 3 / column C
    cif_a_buscar = cifglobal
 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['poblacion'] = 'Aljaraquerrrrr'
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )


    #valuestovar(df_tab_empresa_fila_buscada)


    global programa
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    global nombre_solicitante
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    global email    
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    global proyectos
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    


    global rutaendiscoglobal
    rutaendiscoglobal = ruta_en_disco
    print ( "ruta en disco global dentro " + rutaendiscoglobal )

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                'ruta_en_disco' :  ruta_en_disco
                

    






                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
                
    }

   



    return context






# no usado #####################################
# no usado =======================================
# no usado #####################################

# ==============================================================
# ==============================================================
# ==============================================================


# este es la primera funcion que se ejecuta y captura los datos a partir del cif




def procbuscarcif():
    fecha_inicio_proyecto = ""
    print('Introduzca CIF ')
    
    # cif=input()
    #cif = '44237153B'
    global cifglobal
    cifglobal = input(" Introducir CIF ")
    
    # ========= Si lo encontramos, creamos un dataframe con esa fila para la pestaña EMPRESA y USUARIO  
    # == Cargamos en DF datos TAB EMPRESA
    print('')
    #df_empresa_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cif]      
    
    a = input('El CIF es ...........' + cifglobal)

    print( "" )

    cif_a_buscar = cifglobal
 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME 

    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))
    # df_fila_cero = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[0,:]
    print ( " fila cero ....")
    #print( df_fila_cero)
    pulsatecla = input("pulse una tecla...")
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\salida.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )




 
    b = input ( " HE IMPRIMIDO DATOS LINEA DF ... PULSE UNA TECLA. .... ")
    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    #df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['poblacion'] = 'Aljaraquerrrrr'
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\salida.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )
    print ( "fila buscada " + str(df_tab_empresa_fila_buscada))
    pulsetecla = input(" Pulse una tecla ...")


    #valuestovar(df_tab_empresa_fila_buscada)
    global programa
    global proyectos
    global nombre_solicitante
    global email    
    global rutaendiscoglobal

    # OBSERVAR SI FALTA AL ANEXIONAR DE ANEXO 18 INCLUIR ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
                
    # OBSERVAR SI FALTA AL ANEXIONANR DE ANEXO 18 INCLUIR -  rutaendiscoglobal = ruta_en_disco

    
       
    #df_tab_empresa_fila_buscada = df    
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    sector_empresa = df_tab_empresa_fila_buscada.iloc[0]['sector_empresa']
    pagina_web = df_tab_empresa_fila_buscada.iloc[0]['pagina_web']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    documento_solicitante = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']

    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    cp = str(math.trunc(cp))

    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']

    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    telefono_solicitante =  cp = str(math.trunc(telefono_solicitante))

    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    tratamiento_tecnico  = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    rutaendiscoglobal = ruta_en_disco    
    nivel_madurez = df_tab_empresa_fila_buscada.iloc[0]['nivel_madurez']

    segui_1 = df_tab_empresa_fila_buscada.iloc[0]['segui_1']
    segui_2 = df_tab_empresa_fila_buscada.iloc[0]['segui_2']
    segui_3 = df_tab_empresa_fila_buscada.iloc[0]['segui_3']
    segui_4 = df_tab_empresa_fila_buscada.iloc[0]['segui_4']
    segui_5 = df_tab_empresa_fila_buscada.iloc[0]['segui_5']
    segui_6 = df_tab_empresa_fila_buscada.iloc[0]['segui_6']
    fecha_fin_proyecto = df_tab_empresa_fila_buscada.iloc[0]['fecha_fin_proyecto']



    
    
    print ( ' leyendo datos   .................................... ' )

    proyectos_seleccionados = df_tab_empresa_fila_buscada.iloc[0]['proyectos_seleccionados'] 
    nombre_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto1']

    proveedor_proyecto1	= df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto1']
    cif_proveedor_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto1']
    proveedor2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto1']
    cif_proveedor2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto1']
    proveedor3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto1']
    cif_proveedor3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto1']
    proveedor4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto1']
    cif_proveedor4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto1']
    
    duracion_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto1']

    concepto_importe1_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto1']
    concepto_importe2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto1']
    concepto_importe3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto1']
    concepto_importe4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto1']

    fecha_importe1_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto1']
    fecha_importe2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto1']
    fecha_importe3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto1']
    fecha_importe4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto1']

    importe1_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto1']
    importe2_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto1']
    importe3_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto1']
    importe4_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto1']
    total_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto1']
    nombre_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto2']

    proveedor_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto2']
    cif_proveedor_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto2']
    proveedor2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto2']
    cif_proveedor2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto2']        
    proveedor3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto2']
    cif_proveedor3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto2']
    proveedor4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto2']
    cif_proveedor4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto2']

      

    duracion_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto2']

    concepto_importe1_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto2']
    concepto_importe2_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto2']
    concepto_importe3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto2']
    concepto_importe4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto2']
    

    fecha_importe1_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto2']
    fecha_importe2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto2']
    fecha_importe3_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto2']
    fecha_importe4_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto2']


    importe1_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto2']
    importe2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto2']
    importe3_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto2']
    importe4_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto2']
    total_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto2']    
    nombre_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto3']

    proveedor_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto3']
    cif_proveedor_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto3']
    proveedor2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto3']
    cif_proveedor2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto3']
    proveedor3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto3']
    cif_proveedor3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto3']
    proveedor4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto3']
    cif_proveedor4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto3']


    duracion_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto3']

    concepto_importe1_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto3']
    concepto_importe2_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto3']
    concepto_importe3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto3']
    concepto_importe4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto3']
    

    fecha_importe1_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto3']
    fecha_importe2_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto3']
    fecha_importe3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto3']
    fecha_importe4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto3']

    importe1_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto3']
    importe2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto3']
    importe3_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto3']
    importe4_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto3']
    total_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto3']
    nombre_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto4']

    proveedor_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto4']
    cif_proveedor_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto4']
    proveedor2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto4']
    cif_proveedor2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto4']    
    proveedor3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto4']
    cif_proveedor3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto4']
    proveedor4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto4']
    cif_proveedor4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto4']

    duracion_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto4']

    concepto_importe1_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto4']
    concepto_importe2_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto4']
    concepto_importe3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto4']
    concepto_importe4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto4']
    



    fecha_importe1_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto4']
    fecha_importe2_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto4']
    fecha_importe3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto4']
    fecha_importe4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto4']


    importe1_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto4']
    importe2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto4']
    importe3_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto4']
    importe4_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto4']

    total_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto4']

    totalproyecto = df_tab_empresa_fila_buscada.iloc[0]['totalproyecto']  
    descripcion_actividades_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto1']  
    descripcion_actividades_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto2']  
    descripcion_actividades_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto3']  
    descripcion_actividades_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto4']  
    detalle_de_las_soluciones_implantadas_pr1 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr1']  
    detalle_de_las_soluciones_implantadas_pr2 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr2']  
    detalle_de_las_soluciones_implantadas_pr3 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr3']  
    detalle_de_las_soluciones_implantadas_pr4 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr4']  


   

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
  
    print('codigo postal ' + cp)
    
    print('telefono ' + telefono_solicitante)





    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))
    print ( ' email - correo electronico ' + str(email))

    fecha_inicio_diagnostico_formateada = fecha_documento_inicio_diagnostico.strftime("%d/%m/%Y")
    fecha_fin_diagnostico_formateada = fecha_diagnostico.strftime("%d/%m/%Y")

    print (' FECHA FORMATEADA '   + str(segui_1 ))    
    print (' FECHA FORMATEADA '   + str(segui_2 ))    
    print (' FECHA FORMATEADA  5'   + str(segui_5 ))    
    print (' FECHA FORMATEADA 6 '   + str(segui_6 ))    




    a = input ( ' pulsa una tecla ')
    dni_formateado = documento_representante[:2] + "." + documento_representante[2:5] + "." + documento_representante[5:8] + "-" + documento_representante[8]
    print ( " DNI FORMATEADO " + dni_formateado)
    A = input( "PULSE UNA TECLA ....")

    
    fecha_diagnostico = fecha_diagnostico.strftime('%d/%m/%Y')
    
    if fases == "FI + FII":
        fecha_inicio_proyecto_base = fecha_firma_ppi
        fecha_inicio_proyecto = fecha_inicio_proyecto_base + datetime.timedelta(days=1)
        fecha_firma_ppi = fecha_firma_ppi.strftime('%d/%m/%Y')      
        envio_ppi = envio_ppi.strftime('%d/%m/%Y')
        fecha_inicio_proyecto = fecha_inicio_proyecto.strftime('%d/%m/%Y')
        fecha_fin_proyecto = fecha_fin_proyecto.strftime('%d/%m/%Y')
        fecha_recepcion_presupuesto = fecha_recepcion_presupuesto.strftime('%d/%m/%Y')
        

    if  segui_1 == segui_1:
        segui_1 = segui_1.strftime('%d/%m/%Y')
    if  segui_2 == segui_2:
        segui_2 = segui_2.strftime('%d/%m/%Y')
    if  segui_3 == segui_3:
        segui_3 = segui_3.strftime('%d/%m/%Y')    
    if  segui_4 == segui_4:
        segui_4 = segui_4.strftime('%d/%m/%Y')    
    if  segui_5 == segui_5:
        segui_5 = segui_5.strftime('%d/%m/%Y')
    if  segui_6 == segui_6:
        segui_6 = segui_6.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto1 == fecha_importe1_proyecto1:
        fecha_importe1_proyecto1 = fecha_importe1_proyecto1.strftime('%d/%m/%Y')
        print ( " DENTRO DEL IF FECHA IMPORTE1 PROYECTO1  " + str(fecha_importe1_proyecto1) )
        A = input( "PULSE UNA TECLA ....")
    if  fecha_importe2_proyecto1 == fecha_importe2_proyecto1:
            fecha_importe2_proyecto1 = fecha_importe2_proyecto1.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto1 == fecha_importe3_proyecto1:
            fecha_importe3_proyecto1 = fecha_importe3_proyecto1.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto1 == fecha_importe4_proyecto1:
            fecha_importe4_proyecto1 = fecha_importe4_proyecto1.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto2 == fecha_importe1_proyecto2:
            fecha_importe1_proyecto2 = fecha_importe1_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto2 == fecha_importe2_proyecto2:
            fecha_importe2_proyecto2 = fecha_importe2_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto2 == fecha_importe3_proyecto2:
            fecha_importe3_proyecto2 = fecha_importe3_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto2 == fecha_importe4_proyecto2:
            fecha_importe4_proyecto2 = fecha_importe4_proyecto2.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto3 == fecha_importe1_proyecto3:
            fecha_importe1_proyecto3 = fecha_importe1_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto3 == fecha_importe2_proyecto3:
            fecha_importe2_proyecto3 = fecha_importe2_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto3 == fecha_importe3_proyecto3:
            fecha_importe3_proyecto3 = fecha_importe3_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto3 == fecha_importe4_proyecto3:
            fecha_importe4_proyecto3 = fecha_importe4_proyecto3.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto4 == fecha_importe1_proyecto4:
            fecha_importe1_proyecto4 = fecha_importe1_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto4 == fecha_importe2_proyecto4:
            fecha_importe2_proyecto4 = fecha_importe2_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto4 == fecha_importe3_proyecto4:
            fecha_importe3_proyecto4 = fecha_importe3_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto4 == fecha_importe4_proyecto4:
            fecha_importe4_proyecto4 = fecha_importe4_proyecto4.strftime('%d/%m/%Y')


    
    

    print( "IMPORTE IMPORTE2_PROYECTO1 " + str(importe2_proyecto1 ))
    a= input("poulse una tecla para continuar ....  ")
    if importe1_proyecto1 != importe1_proyecto1:
        importe1_proyecto1 = ""
    if importe2_proyecto1 != importe2_proyecto1:
        importe2_proyecto1 = ""
    if importe3_proyecto1 != importe3_proyecto1:
        importe3_proyecto1 = ""
    if importe4_proyecto1 != importe4_proyecto1:
        importe4_proyecto1 = ""
    if importe1_proyecto2 != importe1_proyecto2:
        importe1_proyecto2 = ""
    if importe2_proyecto2 != importe2_proyecto2:
        importe2_proyecto2 = ""
    if importe3_proyecto2 != importe3_proyecto2:
        importe3_proyecto2 = ""
    if importe4_proyecto2 != importe4_proyecto2:
        importe4_proyecto2 = ""
    if importe1_proyecto3 != importe1_proyecto3:
        importe1_proyecto3 = ""
    if importe2_proyecto3 != importe2_proyecto3:
        importe2_proyecto3 = ""
    if importe3_proyecto3 != importe3_proyecto3:
        importe3_proyecto3 = ""    
    if importe4_proyecto3 != importe4_proyecto3:
        importe4_proyecto3 = ""
    if importe1_proyecto4 != importe1_proyecto4:
        importe1_proyecto4 = ""        
    if importe2_proyecto4 != importe2_proyecto4:
        importe2_proyecto4 = ""
    if importe3_proyecto4 != importe3_proyecto4:
        importe3_proyecto4 = "" 
    if importe4_proyecto4 != importe4_proyecto4:
        importe4_proyecto4 = "" 
    if proveedor_proyecto1 != proveedor_proyecto1:
        proveedor_proyecto1 = ""
    if proveedor2_proyecto1 != proveedor2_proyecto1:
        proveedor2_proyecto1 = ""
    if proveedor3_proyecto1 != proveedor3_proyecto1:
        proveedor3_proyecto1 = ""
    if proveedor4_proyecto1 != proveedor4_proyecto1:
        proveedor4_proyecto1 = ""
    if proveedor_proyecto2 != proveedor_proyecto2:
        proveedor_proyecto2 = ""
    if proveedor2_proyecto2 != proveedor2_proyecto2:
        proveedor2_proyecto2 = ""
    if proveedor3_proyecto2 != proveedor3_proyecto2:
        proveedor3_proyecto2 = ""
    if proveedor4_proyecto2 != proveedor3_proyecto2:
        proveedor4_proyecto2 = ""
    if proveedor_proyecto3 != proveedor_proyecto3:
        proveedor_proyecto3 = ""
    if proveedor2_proyecto3 != proveedor2_proyecto3:
        proveedor2_proyecto3 = ""
    if proveedor3_proyecto3 != proveedor3_proyecto3:
        proveedor3_proyecto3 = ""
    if proveedor4_proyecto3 != proveedor3_proyecto3:
        proveedor4_proyecto3 = ""
    if proveedor_proyecto4 != proveedor_proyecto4:
        proveedor_proyecto4 = ""        
    if proveedor2_proyecto4 != proveedor2_proyecto4:
        proveedor2_proyecto4 = ""     
    if proveedor3_proyecto4 != proveedor3_proyecto4:
        proveedor3_proyecto4 = ""
    if proveedor4_proyecto4 != proveedor4_proyecto4:
        proveedor4_proyecto4 = ""           
    print( "IMPORTE IMPORTE2_PROYECTO1  DESPUES DEL IF " +  str(importe2_proyecto1 ) + ' ' + str(importe1_proyecto1 ))
    a= input("poulse una tecla para continuar ....  ")

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'documento_solicitante' : cif_empresa,
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,
                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'sector_empresa' : sector_empresa,
                'pagina_web' : pagina_web,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : dni_formateado,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,                
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_formateado,
                'tratamiento_tecnico' : tratamiento_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_inicio_diagnostico_formateada,
                'fecha_diagnostico' : fecha_fin_diagnostico_formateada,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_fin_diagnostico_formateada,
                 # str(fecha_registro_participacion_fase_i_anexo_18)
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19': fecha_registro_participacion_en_fase_ii_anexo_19,
                'nivel_madurez' : nivel_madurez,
                'segui_1' : segui_1,
                'segui_2' : segui_2,
                'segui_3' : segui_3,
                'segui_4' : segui_4,
                'segui_5' : segui_5,
                'segui_6' : segui_6,


                
                 #.strftime('%d/%m/%Y'),
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                'num_expediente' : num_expediente,
                'r_admision' : r_admision,      
                'proyectos_seleccionados' : proyectos_seleccionados,
                'nombre_proyecto1' : nombre_proyecto1,
                'proveedor_proyecto1'	: proveedor_proyecto1,
                'cif_proveedor_proyecto1'	 : cif_proveedor_proyecto1,
                'proveedor2_proyecto1' : proveedor2_proyecto1,
                'cif_proveedor2_proyecto1' : cif_proveedor2_proyecto1,
                'proveedor3_proyecto1' : proveedor3_proyecto1,
                'cif_proveedor3_proyecto1' : cif_proveedor3_proyecto1,
                'proveedor4_proyecto1' : proveedor4_proyecto1,
                'cif_proveedor4_proyecto1' : cif_proveedor4_proyecto1,

                'duracion_proyecto1' : duracion_proyecto1,

                

                'concepto_importe1_proyecto1' : concepto_importe1_proyecto1,
                'concepto_importe2_proyecto1' : concepto_importe2_proyecto1,
                'concepto_importe3_proyecto1' : concepto_importe3_proyecto1,
                'concepto_importe4_proyecto1' : concepto_importe4_proyecto1,


                'proyecto_mas_concepto1_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe1_proyecto1),
                'proyecto_mas_concepto2_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe2_proyecto1),
                'proyecto_mas_concepto3_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe3_proyecto1),
                'proyecto_mas_concepto4_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe4_proyecto1),



                'fecha_importe1_proyecto1' : fecha_importe1_proyecto1,
                'fecha_importe2_proyecto1' : fecha_importe2_proyecto1,
                'fecha_importe3_proyecto1' : fecha_importe3_proyecto1,
                'fecha_importe4_proyecto1' : fecha_importe4_proyecto1,
                'importe1_proyecto1'	 : importe1_proyecto1,
                'importe2_proyecto1'	 : importe2_proyecto1,
                'importe3_proyecto1'	 : importe3_proyecto1,
                'importe4_proyecto1'	 : importe4_proyecto1,
                'total_proyecto1'	 : total_proyecto1,
                'nombre_proyecto2'	 : nombre_proyecto2,
                'proveedor_proyecto2'	 : proveedor_proyecto2,
                'cif_proveedor_proyecto2'	 : cif_proveedor_proyecto2,
                'proveedor2_proyecto2'	 : proveedor2_proyecto2,
                'cif_proveedor2_proyecto2' : cif_proveedor2_proyecto2,
		        'proveedor3_proyecto2' : proveedor3_proyecto2,
                'cif_proveedor3_proyecto2' : cif_proveedor3_proyecto2,
                'proveedor4_proyecto2' : proveedor4_proyecto2,
                'cif_proveedor4_proyecto2' : cif_proveedor4_proyecto2,

                'duracion_proyecto2'	 : duracion_proyecto2,

                'concepto_importe1_proyecto2' : concepto_importe1_proyecto2,
                'concepto_importe2_proyecto2' : concepto_importe2_proyecto2,
                'concepto_importe3_proyecto2' : concepto_importe3_proyecto2,
                'concepto_importe4_proyecto2' : concepto_importe4_proyecto2,

                
                'proyecto_mas_concepto1_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe1_proyecto2),
                'proyecto_mas_concepto2_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe2_proyecto2),
                'proyecto_mas_concepto3_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe3_proyecto2),
                'proyecto_mas_concepto4_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe4_proyecto2),

                'fecha_importe1_proyecto2' : fecha_importe1_proyecto2,
                'fecha_importe2_proyecto2' : fecha_importe2_proyecto2,
                'fecha_importe3_proyecto2' : fecha_importe3_proyecto2,
                'fecha_importe4_proyecto2' : fecha_importe4_proyecto2,


                'importe1_proyecto2'	 :  importe1_proyecto2,
                'importe2_proyecto2'	 :  importe2_proyecto2,
                'importe3_proyecto2'	 :  importe3_proyecto2,
                'importe4_proyecto2'	 :  importe4_proyecto2,
                'total_proyecto2'	 : total_proyecto2,
                'nombre_proyecto3'	 : nombre_proyecto3,
                'proveedor_proyecto3'	 :  proveedor_proyecto3,
                'cif_proveedor_proyecto3'	 : cif_proveedor_proyecto3,
                'proveedor2_proyecto3'	 : proveedor2_proyecto3,
                'cif_proveedor2_proyecto3'	 : cif_proveedor2_proyecto3,
                'proveedor3_proyecto3' : proveedor3_proyecto3,
                'cif_proveedor3_proyecto3' : cif_proveedor3_proyecto3,
                'proveedor4_proyecto3' : proveedor4_proyecto3,
                'cif_proveedor4_proyecto3' : cif_proveedor4_proyecto3,

                'duracion_proyecto3'	 : duracion_proyecto3,

                'concepto_importe1_proyecto3' : concepto_importe1_proyecto3,
                'concepto_importe2_proyecto3' : concepto_importe2_proyecto3,
                'concepto_importe3_proyecto3' : concepto_importe3_proyecto3,
                'concepto_importe4_proyecto3' : concepto_importe4_proyecto3,

                'proyecto_mas_concepto1_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe1_proyecto3),
                'proyecto_mas_concepto2_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe2_proyecto3),
                'proyecto_mas_concepto3_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe3_proyecto3),
                'proyecto_mas_concepto4_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe4_proyecto3),
    



                'fecha_importe1_proyecto3' : fecha_importe1_proyecto3,
                'fecha_importe2_proyecto3' : fecha_importe2_proyecto3,
                'fecha_importe3_proyecto3' : fecha_importe3_proyecto3,
                'fecha_importe4_proyecto3' : fecha_importe4_proyecto3,
                'importe1_proyecto3'	 :  importe1_proyecto3,
                'importe2_proyecto3'	 :  importe2_proyecto3,
                'importe3_proyecto3'	 :  importe3_proyecto3,
                'importe4_proyecto3'	 :  importe4_proyecto3,
                'total_proyecto3'	 : total_proyecto3,
                'nombre_proyecto4'	 : nombre_proyecto4,
                'proveedor_proyecto4'	 :  proveedor_proyecto4,
                'cif_proveedor_proyecto4'	 : cif_proveedor_proyecto4,
                'proveedor2_proyecto4'	 :  proveedor2_proyecto4,
                'cif_proveedor2_proyecto4'	 : cif_proveedor2_proyecto4,
		        'proveedor3_proyecto4' : proveedor3_proyecto4,
                'cif_proveedor3_proyecto4' : cif_proveedor3_proyecto4,
                'proveedor4_proyecto4' : proveedor4_proyecto4,
                'cif_proveedor4_proyecto4' : cif_proveedor4_proyecto4,

                'duracion_proyecto4'	 : duracion_proyecto4,

                'concepto_importe1_proyecto4' : concepto_importe1_proyecto4,
                'concepto_importe2_proyecto4' : concepto_importe2_proyecto4,
                'concepto_importe3_proyecto4' : concepto_importe3_proyecto4,
                'concepto_importe4_proyecto4' : concepto_importe4_proyecto4,

                'proyecto_mas_concepto1_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe1_proyecto4),
                'proyecto_mas_concepto2_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe2_proyecto4),
                'proyecto_mas_concepto3_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe3_proyecto4),
                'proyecto_mas_concepto4_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe4_proyecto4),
    

                'fecha_importe1_proyecto4' : fecha_importe1_proyecto4,
                'fecha_importe2_proyecto4' : fecha_importe2_proyecto4,
                'fecha_importe3_proyecto4' : fecha_importe3_proyecto4,
                'fecha_importe4_proyecto4' : fecha_importe4_proyecto4,


                'importe1_proyecto4'	 :  importe1_proyecto4,
                'importe2_proyecto4'	 :  importe2_proyecto4,
                'importe3_proyecto4'	 :  importe3_proyecto4,
                'importe4_proyecto4'	 :  importe4_proyecto4,
                'total_proyecto4'	 : total_proyecto4,
                'totalproyecto' : totalproyecto,
                'descripcion_actividades_proyecto1' : descripcion_actividades_proyecto1,
                'descripcion_actividades_proyecto2' : descripcion_actividades_proyecto2,
                'descripcion_actividades_proyecto3' : descripcion_actividades_proyecto3,
                'descripcion_actividades_proyecto4' : descripcion_actividades_proyecto4,
                'detalle_de_las_soluciones_implantadas_pr1' : detalle_de_las_soluciones_implantadas_pr1,
                'detalle_de_las_soluciones_implantadas_pr2' : detalle_de_las_soluciones_implantadas_pr2,                 
                'detalle_de_las_soluciones_implantadas_pr3' : detalle_de_las_soluciones_implantadas_pr3,
                'detalle_de_las_soluciones_implantadas_pr4' : detalle_de_las_soluciones_implantadas_pr4,                
                'fecha_inicio_proyecto' : fecha_inicio_proyecto,
                'fecha_fin_proyecto' : fecha_fin_proyecto

            
          
    }

    
    print ( " descripcion PROVEEDOR 3 Y 4    " + str( proveedor3_proyecto1) + str(proveedor4_proyecto1))



    A = input( "PULSE UNA TECLA PARA GRABBAR LOS DATOS")



    
    

    return context

    def buscar_cif_oap():
        fecha_inicio_proyecto = ""
        print('Introduzca CIF ')
        
        # cif=input()
        #cif = '44237153B'
        global cifglobal
        cifglobal = input(" Introducir CIF ")
        
        # ========= Si lo encontramos, creamos un dataframe con esa fila para la pestaña EMPRESA y USUARIO  
        # == Cargamos en DF datos TAB EMPRESA
        print('')
        #df_empresa_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cif]      
        
        print('El CIF es ' + cif)

        print( "" )
        # dfAltaEmpresas, dfAltaUsuario = modify_data_pandas(cif)
        #,data_oap_tab_alta_usuario)
        #update_spreadsheet(path, CF, 1, 1, 'Alta_Usuario') #Write to sheet1 starting from row 20 and column 3 / column C
        cif_a_buscar = cifglobal
    
        #context = {'mi_nombre': 'Fran tarci'}
        
        data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs = data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDE
        print (('en generar ficha '+ str(data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs)))

        
        # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
        df_tab_empresa_fila_buscada = data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs.loc[data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs['documento_solicitante'] == cif_a_buscar]   

        #valuestovar(df_tab_empresa_fila_buscada)
        global programa
        global proyectos
        global nombre_solicitante
        global email    
        global rutaendiscoglobal

        # OBSERVAR SI FALTA AL ANEXIONAR DE ANEXO 18 INCLUIR ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
                    
        # OBSERVAR SI FALTA AL ANEXIONANR DE ANEXO 18 INCLUIR -  rutaendiscoglobal = ruta_en_disco

        
        
        #df_tab_empresa_fila_buscada = df    
        fecha_asesoramiento = df_tab_empresa_fila_buscada.iloc[0]['Marca temporal']
        razon_social = df_tab_empresa_fila_buscada.iloc[0]['Nombre']
        nombre = df_tab_empresa_fila_buscada.iloc[0]['nombre']
        apellidos = df_tab_empresa_fila_buscada.iloc[0]['apellidos']
        

        fecha_nacimiento = df_tab_empresa_fila_buscada.iloc[0]['Fecha de nacimiento']
        documento_representante = df_tab_empresa_fila_buscada.iloc[0]['NIF/NIE']
        # documento_representante = df_tab_empresa_fila_buscada.iloc[0]['documento_representante']
        domicilio = df_tab_empresa_fila_buscada.iloc[0]['Domicilio']
        localidad = df_tab_empresa_fila_buscada.iloc[0]['Localidad']
        # poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
        cp = df_tab_empresa_fila_buscada.iloc[0]['Codigo Postal']
        telefono = df_tab_empresa_fila_buscada.iloc[0]['Teléfono']
        email = df_tab_empresa_fila_buscada.iloc[0]['Email']
        
        perfil_solicitante = df_tab_empresa_fila_buscada.iloc[0]['Situacion Laboral']
        
        
               
        fecha_nacimiento_formateada = fecha_nacimiento.strftime("%d/%m/%Y")
        fecha_fin_diagnostico_formateada = fecha_diagnostico.strftime("%d/%m/%Y")
        print (' FECHA nacimiento formateada '   + str(fecha_nacimiento_formateada ))    
        
        # ficha _oap ---- pestaña Alta_Empresa
        nombre_o_razon_social = nombre + apellidos
        empleados_empresa = "< 10 empleados"
        inicio_de_la_actividad = "Últimos 5 años"
        sector = "GRUPO S: Otros servicios"
        nif_empresa = documento_representante
        fecha_asesoramiento = fecha_asesoramiento
        ccaa = "Andalucía"
        provincia = "Huelva"
        localidad = localidad
        cp = cp
        tipo_doc = "NIF/NIE"

        # pestaña Alta_Usuario
        nif_empresa = documento_representante
        fecha_alta_usuario = fecha_asesoramiento

        tratamiento_representante
        # nombre_
        # primer apellido
        # segundo apellido
        nif = documento_representante
        # cargo = cargo
        email = email







        # pestaña Actividades Individuales










        a = input ( ' pulsa una tecla ')
        dni_formateado = documento_representante[:2] + "." + documento_representante[2:5] + "." + documento_representante[5:8] + "-" + documento_representante[8]
        print ( " DNI FORMATEADO " + dni_formateado)
        A = input( "PULSE UNA TECLA ....")

        
        fecha_asesoramiento = fecha_asesoramiento.strftime('%d/%m/%Y')

        
                          
        

        context = {
                    'fecha_documento' : fecha_documento,
                    # =============== FICHA DE EMPRESA
                    'nif_de_empresa' : cif_empresa,
                    'nombre_representante' : nombre_representante.lower().title(),
                    'documento_solicitante' : cif_empresa,
                    'razon_social' : nombre_o_razon_social       , #revisar si intruso
                    'tecnico_justificar' :  tecnico_justificar,
                    'programa' : programa,
                    'fases' : fases,
                    'nombre_solicitante' : nombre_solicitante,
                    'sector_empresa' : sector_empresa,
                    'pagina_web' : pagina_web,
                    'provincia' : provincia,
                    'poblacion' : poblacion,
                    'cp' : cp,
                    'email' : email,
                    'direccion' : direccion,
                    'telefono_solicitante' : telefono_solicitante,
                    'documento_representante' : dni_formateado,
                    'email_representante' : email_representante,
                    'cargo' : cargo,
                    'tratamiento_representante' : tratamiento_representante,
                    'fases2' : fases2,                
                    'tecnico_justificar' : tecnico_justificar,
                    'dni_tecnico' : dni_formateado,
                    'tratamiento_tecnico' : tratamiento_tecnico,
                    'fecha_documento_inicio_diagnostico' : fecha_inicio_diagnostico_formateada,
                    'fecha_diagnostico' : fecha_fin_diagnostico_formateada,
                    'anexo_18_firmado' : anexo_18_firmado,
                    'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                    'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                    'envio_ppi' : envio_ppi,
                    'fecha_firma_ppi' : fecha_firma_ppi,
                    'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                    'duracion_del_plan' : duracion_del_plan,
                    'encuesta_satisfaccion' : encuesta_satisfaccion,
                    'proyectos' : proyectos,
                    'descripcion_empresa' : descripcion_empresa,
                    'fecha_registro_participacion_fase_i_anexo_18' : fecha_fin_diagnostico_formateada,
                    # str(fecha_registro_participacion_fase_i_anexo_18)
                    'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                    'fecha_registro_participacion_en_fase_ii_anexo_19': fecha_registro_participacion_en_fase_ii_anexo_19,
                    'nivel_madurez' : nivel_madurez,
                    'segui_1' : segui_1,
                    'segui_2' : segui_2,
                    'segui_3' : segui_3,
                    'segui_4' : segui_4,
                    'segui_5' : segui_5,
                    'segui_6' : segui_6,


                    
                    #.strftime('%d/%m/%Y'),
                    'cif_empresa' :cif_empresa,
                    'nombre_o_razon_social' :  nombre_o_razon_social,
                    'num_expediente' : num_expediente,
                    'r_admision' : r_admision,      
                    'proyectos_seleccionados' : proyectos_seleccionados,
                    'nombre_proyecto1' : nombre_proyecto1,
                    'proveedor_proyecto1'	: proveedor_proyecto1,
                    'cif_proveedor_proyecto1'	 : cif_proveedor_proyecto1,
                    'proveedor2_proyecto1' : proveedor2_proyecto1,
                    'cif_proveedor2_proyecto1' : cif_proveedor2_proyecto1,
                    'proveedor3_proyecto1' : proveedor3_proyecto1,
                    'cif_proveedor3_proyecto1' : cif_proveedor3_proyecto1,
                    'proveedor4_proyecto1' : proveedor4_proyecto1,
                    'cif_proveedor4_proyecto1' : cif_proveedor4_proyecto1,

                    'duracion_proyecto1' : duracion_proyecto1,

                    

                    'concepto_importe1_proyecto1' : concepto_importe1_proyecto1,
                    'concepto_importe2_proyecto1' : concepto_importe2_proyecto1,
                    'concepto_importe3_proyecto1' : concepto_importe3_proyecto1,
                    'concepto_importe4_proyecto1' : concepto_importe4_proyecto1,


                    'proyecto_mas_concepto1_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe1_proyecto1),
                    'proyecto_mas_concepto2_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe2_proyecto1),
                    'proyecto_mas_concepto3_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe3_proyecto1),
                    'proyecto_mas_concepto4_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe4_proyecto1),



                    'fecha_importe1_proyecto1' : fecha_importe1_proyecto1,
                    'fecha_importe2_proyecto1' : fecha_importe2_proyecto1,
                    'fecha_importe3_proyecto1' : fecha_importe3_proyecto1,
                    'fecha_importe4_proyecto1' : fecha_importe4_proyecto1,
                    'importe1_proyecto1'	 : importe1_proyecto1,
                    'importe2_proyecto1'	 : importe2_proyecto1,
                    'importe3_proyecto1'	 : importe3_proyecto1,
                    'importe4_proyecto1'	 : importe4_proyecto1,
                    'total_proyecto1'	 : total_proyecto1,
                    'nombre_proyecto2'	 : nombre_proyecto2,
                    'proveedor_proyecto2'	 : proveedor_proyecto2,
                    'cif_proveedor_proyecto2'	 : cif_proveedor_proyecto2,
                    'proveedor2_proyecto2'	 : proveedor2_proyecto2,
                    'cif_proveedor2_proyecto2' : cif_proveedor2_proyecto2,
                    'proveedor3_proyecto2' : proveedor3_proyecto2,
                    'cif_proveedor3_proyecto2' : cif_proveedor3_proyecto2,
                    'proveedor4_proyecto2' : proveedor4_proyecto2,
                    'cif_proveedor4_proyecto2' : cif_proveedor4_proyecto2,

                    'duracion_proyecto2'	 : duracion_proyecto2,

                    'concepto_importe1_proyecto2' : concepto_importe1_proyecto2,
                    'concepto_importe2_proyecto2' : concepto_importe2_proyecto2,
                    'concepto_importe3_proyecto2' : concepto_importe3_proyecto2,
                    'concepto_importe4_proyecto2' : concepto_importe4_proyecto2,

                    
                    'proyecto_mas_concepto1_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe1_proyecto2),
                    'proyecto_mas_concepto2_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe2_proyecto2),
                    'proyecto_mas_concepto3_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe3_proyecto2),
                    'proyecto_mas_concepto4_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe4_proyecto2),

                    'fecha_importe1_proyecto2' : fecha_importe1_proyecto2,
                    'fecha_importe2_proyecto2' : fecha_importe2_proyecto2,
                    'fecha_importe3_proyecto2' : fecha_importe3_proyecto2,
                    'fecha_importe4_proyecto2' : fecha_importe4_proyecto2,


                    'importe1_proyecto2'	 :  importe1_proyecto2,
                    'importe2_proyecto2'	 :  importe2_proyecto2,
                    'importe3_proyecto2'	 :  importe3_proyecto2,
                    'importe4_proyecto2'	 :  importe4_proyecto2,
                    'total_proyecto2'	 : total_proyecto2,
                    'nombre_proyecto3'	 : nombre_proyecto3,
                    'proveedor_proyecto3'	 :  proveedor_proyecto3,
                    'cif_proveedor_proyecto3'	 : cif_proveedor_proyecto3,
                    'proveedor2_proyecto3'	 : proveedor2_proyecto3,
                    'cif_proveedor2_proyecto3'	 : cif_proveedor2_proyecto3,
                    'proveedor3_proyecto3' : proveedor3_proyecto3,
                    'cif_proveedor3_proyecto3' : cif_proveedor3_proyecto3,
                    'proveedor4_proyecto3' : proveedor4_proyecto3,
                    'cif_proveedor4_proyecto3' : cif_proveedor4_proyecto3,

                    'duracion_proyecto3'	 : duracion_proyecto3,

                    'concepto_importe1_proyecto3' : concepto_importe1_proyecto3,
                    'concepto_importe2_proyecto3' : concepto_importe2_proyecto3,
                    'concepto_importe3_proyecto3' : concepto_importe3_proyecto3,
                    'concepto_importe4_proyecto3' : concepto_importe4_proyecto3,

                    'proyecto_mas_concepto1_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe1_proyecto3),
                    'proyecto_mas_concepto2_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe2_proyecto3),
                    'proyecto_mas_concepto3_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe3_proyecto3),
                    'proyecto_mas_concepto4_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe4_proyecto3),
        



                    'fecha_importe1_proyecto3' : fecha_importe1_proyecto3,
                    'fecha_importe2_proyecto3' : fecha_importe2_proyecto3,
                    'fecha_importe3_proyecto3' : fecha_importe3_proyecto3,
                    'fecha_importe4_proyecto3' : fecha_importe4_proyecto3,
                    'importe1_proyecto3'	 :  importe1_proyecto3,
                    'importe2_proyecto3'	 :  importe2_proyecto3,
                    'importe3_proyecto3'	 :  importe3_proyecto3,
                    'importe4_proyecto3'	 :  importe4_proyecto3,
                    'total_proyecto3'	 : total_proyecto3,
                    'nombre_proyecto4'	 : nombre_proyecto4,
                    'proveedor_proyecto4'	 :  proveedor_proyecto4,
                    'cif_proveedor_proyecto4'	 : cif_proveedor_proyecto4,
                    'proveedor2_proyecto4'	 :  proveedor2_proyecto4,
                    'cif_proveedor2_proyecto4'	 : cif_proveedor2_proyecto4,
                    'proveedor3_proyecto4' : proveedor3_proyecto4,
                    'cif_proveedor3_proyecto4' : cif_proveedor3_proyecto4,
                    'proveedor4_proyecto4' : proveedor4_proyecto4,
                    'cif_proveedor4_proyecto4' : cif_proveedor4_proyecto4,

                    'duracion_proyecto4'	 : duracion_proyecto4,

                    'concepto_importe1_proyecto4' : concepto_importe1_proyecto4,
                    'concepto_importe2_proyecto4' : concepto_importe2_proyecto4,
                    'concepto_importe3_proyecto4' : concepto_importe3_proyecto4,
                    'concepto_importe4_proyecto4' : concepto_importe4_proyecto4,

                    'proyecto_mas_concepto1_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe1_proyecto4),
                    'proyecto_mas_concepto2_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe2_proyecto4),
                    'proyecto_mas_concepto3_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe3_proyecto4),
                    'proyecto_mas_concepto4_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe4_proyecto4),
        

                    'fecha_importe1_proyecto4' : fecha_importe1_proyecto4,
                    'fecha_importe2_proyecto4' : fecha_importe2_proyecto4,
                    'fecha_importe3_proyecto4' : fecha_importe3_proyecto4,
                    'fecha_importe4_proyecto4' : fecha_importe4_proyecto4,


                    'importe1_proyecto4'	 :  importe1_proyecto4,
                    'importe2_proyecto4'	 :  importe2_proyecto4,
                    'importe3_proyecto4'	 :  importe3_proyecto4,
                    'importe4_proyecto4'	 :  importe4_proyecto4,
                    'total_proyecto4'	 : total_proyecto4,
                    'totalproyecto' : totalproyecto,
                    'descripcion_actividades_proyecto1' : descripcion_actividades_proyecto1,
                    'descripcion_actividades_proyecto2' : descripcion_actividades_proyecto2,
                    'descripcion_actividades_proyecto3' : descripcion_actividades_proyecto3,
                    'descripcion_actividades_proyecto4' : descripcion_actividades_proyecto4,
                    'detalle_de_las_soluciones_implantadas_pr1' : detalle_de_las_soluciones_implantadas_pr1,
                    'detalle_de_las_soluciones_implantadas_pr2' : detalle_de_las_soluciones_implantadas_pr2,                 
                    'detalle_de_las_soluciones_implantadas_pr3' : detalle_de_las_soluciones_implantadas_pr3,
                    'detalle_de_las_soluciones_implantadas_pr4' : detalle_de_las_soluciones_implantadas_pr4,                
                    'fecha_inicio_proyecto' : fecha_inicio_proyecto,
                    'fecha_fin_proyecto' : fecha_fin_proyecto

                
            
        }

        
        print ( " descripcion PROVEEDOR 3 Y 4    " + str( proveedor3_proyecto1) + str(proveedor4_proyecto1))
        A = input( "PULSE UNA TECLA ....")

        
        

        return context





def procbuscarcif_api():
    fecha_inicio_proyecto = ""
    print('Introduzca CIF ')
    
    # cif=input()
    #cif = '44237153B'
    global cifglobal
    cifglobal = input(" Introducir CIF ")
    
    # ========= Si lo encontramos, creamos un dataframe con esa fila para la pestaña EMPRESA y USUARIO  
    # == Cargamos en DF datos TAB EMPRESA
    print('')
    #df_empresa_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cif]      
    
    a = input('El CIF es ' + cifglobal)

    print( "" )

    cif_a_buscar = cifglobal
 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME 

    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))
    # df_fila_cero = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[0,:]
    print ( " fila cero ....")
    #print( df_fila_cero)
    pulsatecla = input("pulse una tecla...")
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\salida.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )




 
    b = input ( " HE IMPRIMIDO DATOS LINEA DF ... PULSE UNA TECLA. .... ")
    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    #df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['poblacion'] = 'Aljaraquerrrrr'
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\salida.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )
    print ( "fila buscada " + str(df_tab_empresa_fila_buscada))
    pulsetecla = input(" Pulse una tecla ...")


    #valuestovar(df_tab_empresa_fila_buscada)
    global programa
    global proyectos
    global nombre_solicitante
    global email    
    global rutaendiscoglobal

    # OBSERVAR SI FALTA AL ANEXIONAR DE ANEXO 18 INCLUIR ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
                
    # OBSERVAR SI FALTA AL ANEXIONANR DE ANEXO 18 INCLUIR -  rutaendiscoglobal = ruta_en_disco

    
       
    #df_tab_empresa_fila_buscada = df    
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    sector_empresa = df_tab_empresa_fila_buscada.iloc[0]['sector_empresa']
    pagina_web = df_tab_empresa_fila_buscada.iloc[0]['pagina_web']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    documento_solicitante = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']

    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    cp = str(math.trunc(cp))

    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']

    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    telefono_solicitante =  cp = str(math.trunc(telefono_solicitante))

    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    tratamiento_tecnico  = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']
    num_expediente = df_tab_empresa_fila_buscada.iloc[0]['num_expediente']
    ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
    r_admision = df_tab_empresa_fila_buscada.iloc[0]['r_admision']
    rutaendiscoglobal = ruta_en_disco    
    nivel_madurez = df_tab_empresa_fila_buscada.iloc[0]['nivel_madurez']

    segui_1 = df_tab_empresa_fila_buscada.iloc[0]['segui_1']
    segui_2 = df_tab_empresa_fila_buscada.iloc[0]['segui_2']
    segui_3 = df_tab_empresa_fila_buscada.iloc[0]['segui_3']
    segui_4 = df_tab_empresa_fila_buscada.iloc[0]['segui_4']
    segui_5 = df_tab_empresa_fila_buscada.iloc[0]['segui_5']
    segui_6 = df_tab_empresa_fila_buscada.iloc[0]['segui_6']
    fecha_fin_proyecto = df_tab_empresa_fila_buscada.iloc[0]['fecha_fin_proyecto']



    
    
    print ( ' leyendo datos   .................................... ' )

    proyectos_seleccionados = df_tab_empresa_fila_buscada.iloc[0]['proyectos_seleccionados'] 
    nombre_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto1']

    proveedor_proyecto1	= df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto1']
    cif_proveedor_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto1']
    proveedor2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto1']
    cif_proveedor2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto1']
    proveedor3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto1']
    cif_proveedor3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto1']
    proveedor4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto1']
    cif_proveedor4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto1']
    
    duracion_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto1']

    concepto_importe1_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto1']
    concepto_importe2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto1']
    concepto_importe3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto1']
    concepto_importe4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto1']

    fecha_importe1_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto1']
    fecha_importe2_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto1']
    fecha_importe3_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto1']
    fecha_importe4_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto1']

    importe1_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto1']
    importe2_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto1']
    importe3_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto1']
    importe4_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto1']
    total_proyecto1	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto1']
    nombre_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto2']

    proveedor_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto2']
    cif_proveedor_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto2']
    proveedor2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto2']
    cif_proveedor2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto2']        
    proveedor3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto2']
    cif_proveedor3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto2']
    proveedor4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto2']
    cif_proveedor4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto2']

      

    duracion_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto2']

    concepto_importe1_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto2']
    concepto_importe2_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto2']
    concepto_importe3_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto2']
    concepto_importe4_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto2']
    

    fecha_importe1_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto2']
    fecha_importe2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto2']
    fecha_importe3_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto2']
    fecha_importe4_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto2']


    importe1_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto2']
    importe2_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto2']
    importe3_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto2']
    importe4_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto2']
    total_proyecto2	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto2']    
    nombre_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto3']

    proveedor_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto3']
    cif_proveedor_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto3']
    proveedor2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto3']
    cif_proveedor2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto3']
    proveedor3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto3']
    cif_proveedor3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto3']
    proveedor4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto3']
    cif_proveedor4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto3']


    duracion_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto3']

    concepto_importe1_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto3']
    concepto_importe2_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto3']
    concepto_importe3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto3']
    concepto_importe4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto3']
    

    fecha_importe1_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto3']
    fecha_importe2_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto3']
    fecha_importe3_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto3']
    fecha_importe4_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto3']

    importe1_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto3']
    importe2_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto3']
    importe3_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto3']
    importe4_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto3']
    total_proyecto3	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto3']
    nombre_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['nombre_proyecto4']

    proveedor_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor_proyecto4']
    cif_proveedor_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor_proyecto4']
    proveedor2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['proveedor2_proyecto4']
    cif_proveedor2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor2_proyecto4']    
    proveedor3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['proveedor3_proyecto4']
    cif_proveedor3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor3_proyecto4']
    proveedor4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['proveedor4_proyecto4']
    cif_proveedor4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['cif_proveedor4_proyecto4']

    duracion_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['duracion_proyecto4']

    concepto_importe1_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe1_proyecto4']
    concepto_importe2_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe2_proyecto4']
    concepto_importe3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe3_proyecto4']
    concepto_importe4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['concepto_importe4_proyecto4']
    



    fecha_importe1_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe1_proyecto4']
    fecha_importe2_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe2_proyecto4']
    fecha_importe3_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe3_proyecto4']
    fecha_importe4_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['fecha_importe4_proyecto4']


    importe1_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe1_proyecto4']
    importe2_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe2_proyecto4']
    importe3_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe3_proyecto4']
    importe4_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['importe4_proyecto4']

    total_proyecto4	 = df_tab_empresa_fila_buscada.iloc[0]['total_proyecto4']

    totalproyecto = df_tab_empresa_fila_buscada.iloc[0]['totalproyecto']  
    descripcion_actividades_proyecto1 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto1']  
    descripcion_actividades_proyecto2 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto2']  
    descripcion_actividades_proyecto3 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto3']  
    descripcion_actividades_proyecto4 = df_tab_empresa_fila_buscada.iloc[0]['descripcion_actividades_proyecto4']  
    detalle_de_las_soluciones_implantadas_pr1 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr1']  
    detalle_de_las_soluciones_implantadas_pr2 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr2']  
    detalle_de_las_soluciones_implantadas_pr3 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr3']  
    detalle_de_las_soluciones_implantadas_pr4 = df_tab_empresa_fila_buscada.iloc[0]['detalle_de_las_soluciones_implantadas_pr4']  


   

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
  
    print('codigo postal ' + cp)
    
    print('telefono ' + telefono_solicitante)





    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))
    print ( ' email - correo electronico ' + str(email))

    fecha_inicio_diagnostico_formateada = fecha_documento_inicio_diagnostico.strftime("%d/%m/%Y")
    fecha_fin_diagnostico_formateada = fecha_diagnostico.strftime("%d/%m/%Y")

    print (' FECHA FORMATEADA '   + str(segui_1 ))    
    print (' FECHA FORMATEADA '   + str(segui_2 ))    
    print (' FECHA FORMATEADA  5'   + str(segui_5 ))    
    print (' FECHA FORMATEADA 6 '   + str(segui_6 ))    




    a = input ( ' pulsa una tecla ')
    dni_formateado = documento_representante[:2] + "." + documento_representante[2:5] + "." + documento_representante[5:8] + "-" + documento_representante[8]
    print ( " DNI FORMATEADO " + dni_formateado)
    A = input( "PULSE UNA TECLA ....")

    
    fecha_diagnostico = fecha_diagnostico.strftime('%d/%m/%Y')
    
    if fases == "FI + FII":
        fecha_inicio_proyecto_base = fecha_firma_ppi
        fecha_inicio_proyecto = fecha_inicio_proyecto_base + datetime.timedelta(days=1)
        fecha_firma_ppi = fecha_firma_ppi.strftime('%d/%m/%Y')      
        envio_ppi = envio_ppi.strftime('%d/%m/%Y')
        fecha_inicio_proyecto = fecha_inicio_proyecto.strftime('%d/%m/%Y')
        fecha_fin_proyecto = fecha_fin_proyecto.strftime('%d/%m/%Y')
        fecha_recepcion_presupuesto = fecha_recepcion_presupuesto.strftime('%d/%m/%Y')
        

    if  segui_1 == segui_1:
        segui_1 = segui_1.strftime('%d/%m/%Y')
    if  segui_2 == segui_2:
        segui_2 = segui_2.strftime('%d/%m/%Y')
    if  segui_3 == segui_3:
        segui_3 = segui_3.strftime('%d/%m/%Y')    
    if  segui_4 == segui_4:
        segui_4 = segui_4.strftime('%d/%m/%Y')    
    if  segui_5 == segui_5:
        segui_5 = segui_5.strftime('%d/%m/%Y')
    if  segui_6 == segui_6:
        segui_6 = segui_6.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto1 == fecha_importe1_proyecto1:
        fecha_importe1_proyecto1 = fecha_importe1_proyecto1.strftime('%d/%m/%Y')
        print ( " DENTRO DEL IF FECHA IMPORTE1 PROYECTO1  " + str(fecha_importe1_proyecto1) )
        A = input( "PULSE UNA TECLA ....")
    if  fecha_importe2_proyecto1 == fecha_importe2_proyecto1:
            fecha_importe2_proyecto1 = fecha_importe2_proyecto1.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto1 == fecha_importe3_proyecto1:
            fecha_importe3_proyecto1 = fecha_importe3_proyecto1.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto1 == fecha_importe4_proyecto1:
            fecha_importe4_proyecto1 = fecha_importe4_proyecto1.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto2 == fecha_importe1_proyecto2:
            fecha_importe1_proyecto2 = fecha_importe1_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto2 == fecha_importe2_proyecto2:
            fecha_importe2_proyecto2 = fecha_importe2_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto2 == fecha_importe3_proyecto2:
            fecha_importe3_proyecto2 = fecha_importe3_proyecto2.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto2 == fecha_importe4_proyecto2:
            fecha_importe4_proyecto2 = fecha_importe4_proyecto2.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto3 == fecha_importe1_proyecto3:
            fecha_importe1_proyecto3 = fecha_importe1_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto3 == fecha_importe2_proyecto3:
            fecha_importe2_proyecto3 = fecha_importe2_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto3 == fecha_importe3_proyecto3:
            fecha_importe3_proyecto3 = fecha_importe3_proyecto3.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto3 == fecha_importe4_proyecto3:
            fecha_importe4_proyecto3 = fecha_importe4_proyecto3.strftime('%d/%m/%Y')

    if  fecha_importe1_proyecto4 == fecha_importe1_proyecto4:
            fecha_importe1_proyecto4 = fecha_importe1_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe2_proyecto4 == fecha_importe2_proyecto4:
            fecha_importe2_proyecto4 = fecha_importe2_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe3_proyecto4 == fecha_importe3_proyecto4:
            fecha_importe3_proyecto4 = fecha_importe3_proyecto4.strftime('%d/%m/%Y')
    if  fecha_importe4_proyecto4 == fecha_importe4_proyecto4:
            fecha_importe4_proyecto4 = fecha_importe4_proyecto4.strftime('%d/%m/%Y')


    
    

    print( "IMPORTE IMPORTE2_PROYECTO1 " + str(importe2_proyecto1 ))
    a= input("poulse una tecla para continuar ....  ")
    if importe1_proyecto1 != importe1_proyecto1:
        importe1_proyecto1 = ""
    if importe2_proyecto1 != importe2_proyecto1:
        importe2_proyecto1 = ""
    if importe3_proyecto1 != importe3_proyecto1:
        importe3_proyecto1 = ""
    if importe4_proyecto1 != importe4_proyecto1:
        importe4_proyecto1 = ""
    if importe1_proyecto2 != importe1_proyecto2:
        importe1_proyecto2 = ""
    if importe2_proyecto2 != importe2_proyecto2:
        importe2_proyecto2 = ""
    if importe3_proyecto2 != importe3_proyecto2:
        importe3_proyecto2 = ""
    if importe4_proyecto2 != importe4_proyecto2:
        importe4_proyecto2 = ""
    if importe1_proyecto3 != importe1_proyecto3:
        importe1_proyecto3 = ""
    if importe2_proyecto3 != importe2_proyecto3:
        importe2_proyecto3 = ""
    if importe3_proyecto3 != importe3_proyecto3:
        importe3_proyecto3 = ""    
    if importe4_proyecto3 != importe4_proyecto3:
        importe4_proyecto3 = ""
    if importe1_proyecto4 != importe1_proyecto4:
        importe1_proyecto4 = ""        
    if importe2_proyecto4 != importe2_proyecto4:
        importe2_proyecto4 = ""
    if importe3_proyecto4 != importe3_proyecto4:
        importe3_proyecto4 = "" 
    if importe4_proyecto4 != importe4_proyecto4:
        importe4_proyecto4 = "" 
    if proveedor_proyecto1 != proveedor_proyecto1:
        proveedor_proyecto1 = ""
    if proveedor2_proyecto1 != proveedor2_proyecto1:
        proveedor2_proyecto1 = ""
    if proveedor3_proyecto1 != proveedor3_proyecto1:
        proveedor3_proyecto1 = ""
    if proveedor4_proyecto1 != proveedor4_proyecto1:
        proveedor4_proyecto1 = ""
    if proveedor_proyecto2 != proveedor_proyecto2:
        proveedor_proyecto2 = ""
    if proveedor2_proyecto2 != proveedor2_proyecto2:
        proveedor2_proyecto2 = ""
    if proveedor3_proyecto2 != proveedor3_proyecto2:
        proveedor3_proyecto2 = ""
    if proveedor4_proyecto2 != proveedor3_proyecto2:
        proveedor4_proyecto2 = ""
    if proveedor_proyecto3 != proveedor_proyecto3:
        proveedor_proyecto3 = ""
    if proveedor2_proyecto3 != proveedor2_proyecto3:
        proveedor2_proyecto3 = ""
    if proveedor3_proyecto3 != proveedor3_proyecto3:
        proveedor3_proyecto3 = ""
    if proveedor4_proyecto3 != proveedor3_proyecto3:
        proveedor4_proyecto3 = ""
    if proveedor_proyecto4 != proveedor_proyecto4:
        proveedor_proyecto4 = ""        
    if proveedor2_proyecto4 != proveedor2_proyecto4:
        proveedor2_proyecto4 = ""     
    if proveedor3_proyecto4 != proveedor3_proyecto4:
        proveedor3_proyecto4 = ""
    if proveedor4_proyecto4 != proveedor4_proyecto4:
        proveedor4_proyecto4 = ""           
    print( "IMPORTE IMPORTE2_PROYECTO1  DESPUES DEL IF " +  str(importe2_proyecto1 ) + ' ' + str(importe1_proyecto1 ))
    a= input("poulse una tecla para continuar ....  ")

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'documento_solicitante' : cif_empresa,
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,
                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'sector_empresa' : sector_empresa,
                'pagina_web' : pagina_web,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : dni_formateado,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,                
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_formateado,
                'tratamiento_tecnico' : tratamiento_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_inicio_diagnostico_formateada,
                'fecha_diagnostico' : fecha_fin_diagnostico_formateada,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_fin_diagnostico_formateada,
                 # str(fecha_registro_participacion_fase_i_anexo_18)
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19': fecha_registro_participacion_en_fase_ii_anexo_19,
                'nivel_madurez' : nivel_madurez,
                'segui_1' : segui_1,
                'segui_2' : segui_2,
                'segui_3' : segui_3,
                'segui_4' : segui_4,
                'segui_5' : segui_5,
                'segui_6' : segui_6,


                
                 #.strftime('%d/%m/%Y'),
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                'num_expediente' : num_expediente,
                'r_admision' : r_admision,      
                'proyectos_seleccionados' : proyectos_seleccionados,
                'nombre_proyecto1' : nombre_proyecto1,
                'proveedor_proyecto1'	: proveedor_proyecto1,
                'cif_proveedor_proyecto1'	 : cif_proveedor_proyecto1,
                'proveedor2_proyecto1' : proveedor2_proyecto1,
                'cif_proveedor2_proyecto1' : cif_proveedor2_proyecto1,
                'proveedor3_proyecto1' : proveedor3_proyecto1,
                'cif_proveedor3_proyecto1' : cif_proveedor3_proyecto1,
                'proveedor4_proyecto1' : proveedor4_proyecto1,
                'cif_proveedor4_proyecto1' : cif_proveedor4_proyecto1,

                'duracion_proyecto1' : duracion_proyecto1,

                

                'concepto_importe1_proyecto1' : concepto_importe1_proyecto1,
                'concepto_importe2_proyecto1' : concepto_importe2_proyecto1,
                'concepto_importe3_proyecto1' : concepto_importe3_proyecto1,
                'concepto_importe4_proyecto1' : concepto_importe4_proyecto1,


                'proyecto_mas_concepto1_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe1_proyecto1),
                'proyecto_mas_concepto2_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe2_proyecto1),
                'proyecto_mas_concepto3_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe3_proyecto1),
                'proyecto_mas_concepto4_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe4_proyecto1),



                'fecha_importe1_proyecto1' : fecha_importe1_proyecto1,
                'fecha_importe2_proyecto1' : fecha_importe2_proyecto1,
                'fecha_importe3_proyecto1' : fecha_importe3_proyecto1,
                'fecha_importe4_proyecto1' : fecha_importe4_proyecto1,
                'importe1_proyecto1'	 : importe1_proyecto1,
                'importe2_proyecto1'	 : importe2_proyecto1,
                'importe3_proyecto1'	 : importe3_proyecto1,
                'importe4_proyecto1'	 : importe4_proyecto1,
                'total_proyecto1'	 : total_proyecto1,
                'nombre_proyecto2'	 : nombre_proyecto2,
                'proveedor_proyecto2'	 : proveedor_proyecto2,
                'cif_proveedor_proyecto2'	 : cif_proveedor_proyecto2,
                'proveedor2_proyecto2'	 : proveedor2_proyecto2,
                'cif_proveedor2_proyecto2' : cif_proveedor2_proyecto2,
		        'proveedor3_proyecto2' : proveedor3_proyecto2,
                'cif_proveedor3_proyecto2' : cif_proveedor3_proyecto2,
                'proveedor4_proyecto2' : proveedor4_proyecto2,
                'cif_proveedor4_proyecto2' : cif_proveedor4_proyecto2,

                'duracion_proyecto2'	 : duracion_proyecto2,

                'concepto_importe1_proyecto2' : concepto_importe1_proyecto2,
                'concepto_importe2_proyecto2' : concepto_importe2_proyecto2,
                'concepto_importe3_proyecto2' : concepto_importe3_proyecto2,
                'concepto_importe4_proyecto2' : concepto_importe4_proyecto2,

                
                'proyecto_mas_concepto1_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe1_proyecto2),
                'proyecto_mas_concepto2_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe2_proyecto2),
                'proyecto_mas_concepto3_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe3_proyecto2),
                'proyecto_mas_concepto4_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe4_proyecto2),

                'fecha_importe1_proyecto2' : fecha_importe1_proyecto2,
                'fecha_importe2_proyecto2' : fecha_importe2_proyecto2,
                'fecha_importe3_proyecto2' : fecha_importe3_proyecto2,
                'fecha_importe4_proyecto2' : fecha_importe4_proyecto2,


                'importe1_proyecto2'	 :  importe1_proyecto2,
                'importe2_proyecto2'	 :  importe2_proyecto2,
                'importe3_proyecto2'	 :  importe3_proyecto2,
                'importe4_proyecto2'	 :  importe4_proyecto2,
                'total_proyecto2'	 : total_proyecto2,
                'nombre_proyecto3'	 : nombre_proyecto3,
                'proveedor_proyecto3'	 :  proveedor_proyecto3,
                'cif_proveedor_proyecto3'	 : cif_proveedor_proyecto3,
                'proveedor2_proyecto3'	 : proveedor2_proyecto3,
                'cif_proveedor2_proyecto3'	 : cif_proveedor2_proyecto3,
                'proveedor3_proyecto3' : proveedor3_proyecto3,
                'cif_proveedor3_proyecto3' : cif_proveedor3_proyecto3,
                'proveedor4_proyecto3' : proveedor4_proyecto3,
                'cif_proveedor4_proyecto3' : cif_proveedor4_proyecto3,

                'duracion_proyecto3'	 : duracion_proyecto3,

                'concepto_importe1_proyecto3' : concepto_importe1_proyecto3,
                'concepto_importe2_proyecto3' : concepto_importe2_proyecto3,
                'concepto_importe3_proyecto3' : concepto_importe3_proyecto3,
                'concepto_importe4_proyecto3' : concepto_importe4_proyecto3,

                'proyecto_mas_concepto1_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe1_proyecto3),
                'proyecto_mas_concepto2_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe2_proyecto3),
                'proyecto_mas_concepto3_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe3_proyecto3),
                'proyecto_mas_concepto4_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe4_proyecto3),
    



                'fecha_importe1_proyecto3' : fecha_importe1_proyecto3,
                'fecha_importe2_proyecto3' : fecha_importe2_proyecto3,
                'fecha_importe3_proyecto3' : fecha_importe3_proyecto3,
                'fecha_importe4_proyecto3' : fecha_importe4_proyecto3,
                'importe1_proyecto3'	 :  importe1_proyecto3,
                'importe2_proyecto3'	 :  importe2_proyecto3,
                'importe3_proyecto3'	 :  importe3_proyecto3,
                'importe4_proyecto3'	 :  importe4_proyecto3,
                'total_proyecto3'	 : total_proyecto3,
                'nombre_proyecto4'	 : nombre_proyecto4,
                'proveedor_proyecto4'	 :  proveedor_proyecto4,
                'cif_proveedor_proyecto4'	 : cif_proveedor_proyecto4,
                'proveedor2_proyecto4'	 :  proveedor2_proyecto4,
                'cif_proveedor2_proyecto4'	 : cif_proveedor2_proyecto4,
		        'proveedor3_proyecto4' : proveedor3_proyecto4,
                'cif_proveedor3_proyecto4' : cif_proveedor3_proyecto4,
                'proveedor4_proyecto4' : proveedor4_proyecto4,
                'cif_proveedor4_proyecto4' : cif_proveedor4_proyecto4,

                'duracion_proyecto4'	 : duracion_proyecto4,

                'concepto_importe1_proyecto4' : concepto_importe1_proyecto4,
                'concepto_importe2_proyecto4' : concepto_importe2_proyecto4,
                'concepto_importe3_proyecto4' : concepto_importe3_proyecto4,
                'concepto_importe4_proyecto4' : concepto_importe4_proyecto4,

                'proyecto_mas_concepto1_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe1_proyecto4),
                'proyecto_mas_concepto2_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe2_proyecto4),
                'proyecto_mas_concepto3_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe3_proyecto4),
                'proyecto_mas_concepto4_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe4_proyecto4),
    

                'fecha_importe1_proyecto4' : fecha_importe1_proyecto4,
                'fecha_importe2_proyecto4' : fecha_importe2_proyecto4,
                'fecha_importe3_proyecto4' : fecha_importe3_proyecto4,
                'fecha_importe4_proyecto4' : fecha_importe4_proyecto4,


                'importe1_proyecto4'	 :  importe1_proyecto4,
                'importe2_proyecto4'	 :  importe2_proyecto4,
                'importe3_proyecto4'	 :  importe3_proyecto4,
                'importe4_proyecto4'	 :  importe4_proyecto4,
                'total_proyecto4'	 : total_proyecto4,
                'totalproyecto' : totalproyecto,
                'descripcion_actividades_proyecto1' : descripcion_actividades_proyecto1,
                'descripcion_actividades_proyecto2' : descripcion_actividades_proyecto2,
                'descripcion_actividades_proyecto3' : descripcion_actividades_proyecto3,
                'descripcion_actividades_proyecto4' : descripcion_actividades_proyecto4,
                'detalle_de_las_soluciones_implantadas_pr1' : detalle_de_las_soluciones_implantadas_pr1,
                'detalle_de_las_soluciones_implantadas_pr2' : detalle_de_las_soluciones_implantadas_pr2,                 
                'detalle_de_las_soluciones_implantadas_pr3' : detalle_de_las_soluciones_implantadas_pr3,
                'detalle_de_las_soluciones_implantadas_pr4' : detalle_de_las_soluciones_implantadas_pr4,                
                'fecha_inicio_proyecto' : fecha_inicio_proyecto,
                'fecha_fin_proyecto' : fecha_fin_proyecto

            
          
    }

    
    print ( " descripcion PROVEEDOR 3 Y 4    " + str( proveedor3_proyecto1) + str(proveedor4_proyecto1))



    A = input( "PULSE UNA TECLA PARA GRABBAR LOS DATOS")



    
    

    return context

    def buscar_cif_oap():
        fecha_inicio_proyecto = ""
        print('Introduzca CIF ')
        
        # cif=input()
        #cif = '44237153B'
        global cifglobal
        cifglobal = input(" Introducir CIF ")
        
        # ========= Si lo encontramos, creamos un dataframe con esa fila para la pestaña EMPRESA y USUARIO  
        # == Cargamos en DF datos TAB EMPRESA
        print('')
        #df_empresa_buscada = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME.loc[data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME['documento_solicitante'] == cif]      
        
        print('El CIF es ' + cif)

        print( "" )
        # dfAltaEmpresas, dfAltaUsuario = modify_data_pandas(cif)
        #,data_oap_tab_alta_usuario)
        #update_spreadsheet(path, CF, 1, 1, 'Alta_Usuario') #Write to sheet1 starting from row 20 and column 3 / column C
        cif_a_buscar = cifglobal
    
        #context = {'mi_nombre': 'Fran tarci'}
        
        data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs = data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDE
        print (('en generar ficha '+ str(data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs)))

        
        # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
        df_tab_empresa_fila_buscada = data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs.loc[data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDEs['documento_solicitante'] == cif_a_buscar]   

        #valuestovar(df_tab_empresa_fila_buscada)
        global programa
        global proyectos
        global nombre_solicitante
        global email    
        global rutaendiscoglobal

        # OBSERVAR SI FALTA AL ANEXIONAR DE ANEXO 18 INCLUIR ruta_en_disco = df_tab_empresa_fila_buscada.iloc[0]['ruta_en_disco']
                    
        # OBSERVAR SI FALTA AL ANEXIONANR DE ANEXO 18 INCLUIR -  rutaendiscoglobal = ruta_en_disco

        
        
        #df_tab_empresa_fila_buscada = df    
        fecha_asesoramiento = df_tab_empresa_fila_buscada.iloc[0]['Marca temporal']
        razon_social = df_tab_empresa_fila_buscada.iloc[0]['Nombre']
        nombre = df_tab_empresa_fila_buscada.iloc[0]['nombre']
        apellidos = df_tab_empresa_fila_buscada.iloc[0]['apellidos']
        

        fecha_nacimiento = df_tab_empresa_fila_buscada.iloc[0]['Fecha de nacimiento']
        documento_representante = df_tab_empresa_fila_buscada.iloc[0]['NIF/NIE']
        # documento_representante = df_tab_empresa_fila_buscada.iloc[0]['documento_representante']
        domicilio = df_tab_empresa_fila_buscada.iloc[0]['Domicilio']
        localidad = df_tab_empresa_fila_buscada.iloc[0]['Localidad']
        # poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
        cp = df_tab_empresa_fila_buscada.iloc[0]['Codigo Postal']
        telefono = df_tab_empresa_fila_buscada.iloc[0]['Teléfono']
        email = df_tab_empresa_fila_buscada.iloc[0]['Email']
        
        perfil_solicitante = df_tab_empresa_fila_buscada.iloc[0]['Situacion Laboral']
        
        
               
        fecha_nacimiento_formateada = fecha_nacimiento.strftime("%d/%m/%Y")
        fecha_fin_diagnostico_formateada = fecha_diagnostico.strftime("%d/%m/%Y")
        print (' FECHA nacimiento formateada '   + str(fecha_nacimiento_formateada ))    
        
        # ficha _oap ---- pestaña Alta_Empresa
        nombre_o_razon_social = nombre + apellidos
        empleados_empresa = "< 10 empleados"
        inicio_de_la_actividad = "Últimos 5 años"
        sector = "GRUPO S: Otros servicios"
        nif_empresa = documento_representante
        fecha_asesoramiento = fecha_asesoramiento
        ccaa = "Andalucía"
        provincia = "Huelva"
        localidad = localidad
        cp = cp
        tipo_doc = "NIF/NIE"

        # pestaña Alta_Usuario
        nif_empresa = documento_representante
        fecha_alta_usuario = fecha_asesoramiento

        tratamiento_representante
        # nombre_
        # primer apellido
        # segundo apellido
        nif = documento_representante
        # cargo = cargo
        email = email







        # pestaña Actividades Individuales










        a = input ( ' pulsa una tecla ')
        dni_formateado = documento_representante[:2] + "." + documento_representante[2:5] + "." + documento_representante[5:8] + "-" + documento_representante[8]
        print ( " DNI FORMATEADO " + dni_formateado)
        A = input( "PULSE UNA TECLA ....")

        
        fecha_asesoramiento = fecha_asesoramiento.strftime('%d/%m/%Y')

        
                          
        

        context = {
                    'fecha_documento' : fecha_documento,
                    # =============== FICHA DE EMPRESA
                    'nif_de_empresa' : cif_empresa,
                    'nombre_representante' : nombre_representante.lower().title(),
                    'documento_solicitante' : cif_empresa,
                    'razon_social' : nombre_o_razon_social       , #revisar si intruso
                    'tecnico_justificar' :  tecnico_justificar,
                    'programa' : programa,
                    'fases' : fases,
                    'nombre_solicitante' : nombre_solicitante,
                    'sector_empresa' : sector_empresa,
                    'pagina_web' : pagina_web,
                    'provincia' : provincia,
                    'poblacion' : poblacion,
                    'cp' : cp,
                    'email' : email,
                    'direccion' : direccion,
                    'telefono_solicitante' : telefono_solicitante,
                    'documento_representante' : dni_formateado,
                    'email_representante' : email_representante,
                    'cargo' : cargo,
                    'tratamiento_representante' : tratamiento_representante,
                    'fases2' : fases2,                
                    'tecnico_justificar' : tecnico_justificar,
                    'dni_tecnico' : dni_formateado,
                    'tratamiento_tecnico' : tratamiento_tecnico,
                    'fecha_documento_inicio_diagnostico' : fecha_inicio_diagnostico_formateada,
                    'fecha_diagnostico' : fecha_fin_diagnostico_formateada,
                    'anexo_18_firmado' : anexo_18_firmado,
                    'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                    'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                    'envio_ppi' : envio_ppi,
                    'fecha_firma_ppi' : fecha_firma_ppi,
                    'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                    'duracion_del_plan' : duracion_del_plan,
                    'encuesta_satisfaccion' : encuesta_satisfaccion,
                    'proyectos' : proyectos,
                    'descripcion_empresa' : descripcion_empresa,
                    'fecha_registro_participacion_fase_i_anexo_18' : fecha_fin_diagnostico_formateada,
                    # str(fecha_registro_participacion_fase_i_anexo_18)
                    'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                    'fecha_registro_participacion_en_fase_ii_anexo_19': fecha_registro_participacion_en_fase_ii_anexo_19,
                    'nivel_madurez' : nivel_madurez,
                    'segui_1' : segui_1,
                    'segui_2' : segui_2,
                    'segui_3' : segui_3,
                    'segui_4' : segui_4,
                    'segui_5' : segui_5,
                    'segui_6' : segui_6,


                    
                    #.strftime('%d/%m/%Y'),
                    'cif_empresa' :cif_empresa,
                    'nombre_o_razon_social' :  nombre_o_razon_social,
                    'num_expediente' : num_expediente,
                    'r_admision' : r_admision,      
                    'proyectos_seleccionados' : proyectos_seleccionados,
                    'nombre_proyecto1' : nombre_proyecto1,
                    'proveedor_proyecto1'	: proveedor_proyecto1,
                    'cif_proveedor_proyecto1'	 : cif_proveedor_proyecto1,
                    'proveedor2_proyecto1' : proveedor2_proyecto1,
                    'cif_proveedor2_proyecto1' : cif_proveedor2_proyecto1,
                    'proveedor3_proyecto1' : proveedor3_proyecto1,
                    'cif_proveedor3_proyecto1' : cif_proveedor3_proyecto1,
                    'proveedor4_proyecto1' : proveedor4_proyecto1,
                    'cif_proveedor4_proyecto1' : cif_proveedor4_proyecto1,

                    'duracion_proyecto1' : duracion_proyecto1,

                    

                    'concepto_importe1_proyecto1' : concepto_importe1_proyecto1,
                    'concepto_importe2_proyecto1' : concepto_importe2_proyecto1,
                    'concepto_importe3_proyecto1' : concepto_importe3_proyecto1,
                    'concepto_importe4_proyecto1' : concepto_importe4_proyecto1,


                    'proyecto_mas_concepto1_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe1_proyecto1),
                    'proyecto_mas_concepto2_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe2_proyecto1),
                    'proyecto_mas_concepto3_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe3_proyecto1),
                    'proyecto_mas_concepto4_proyecto1' : str(nombre_proyecto1) + ' ref: ' + str(concepto_importe4_proyecto1),



                    'fecha_importe1_proyecto1' : fecha_importe1_proyecto1,
                    'fecha_importe2_proyecto1' : fecha_importe2_proyecto1,
                    'fecha_importe3_proyecto1' : fecha_importe3_proyecto1,
                    'fecha_importe4_proyecto1' : fecha_importe4_proyecto1,
                    'importe1_proyecto1'	 : importe1_proyecto1,
                    'importe2_proyecto1'	 : importe2_proyecto1,
                    'importe3_proyecto1'	 : importe3_proyecto1,
                    'importe4_proyecto1'	 : importe4_proyecto1,
                    'total_proyecto1'	 : total_proyecto1,
                    'nombre_proyecto2'	 : nombre_proyecto2,
                    'proveedor_proyecto2'	 : proveedor_proyecto2,
                    'cif_proveedor_proyecto2'	 : cif_proveedor_proyecto2,
                    'proveedor2_proyecto2'	 : proveedor2_proyecto2,
                    'cif_proveedor2_proyecto2' : cif_proveedor2_proyecto2,
                    'proveedor3_proyecto2' : proveedor3_proyecto2,
                    'cif_proveedor3_proyecto2' : cif_proveedor3_proyecto2,
                    'proveedor4_proyecto2' : proveedor4_proyecto2,
                    'cif_proveedor4_proyecto2' : cif_proveedor4_proyecto2,

                    'duracion_proyecto2'	 : duracion_proyecto2,

                    'concepto_importe1_proyecto2' : concepto_importe1_proyecto2,
                    'concepto_importe2_proyecto2' : concepto_importe2_proyecto2,
                    'concepto_importe3_proyecto2' : concepto_importe3_proyecto2,
                    'concepto_importe4_proyecto2' : concepto_importe4_proyecto2,

                    
                    'proyecto_mas_concepto1_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe1_proyecto2),
                    'proyecto_mas_concepto2_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe2_proyecto2),
                    'proyecto_mas_concepto3_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe3_proyecto2),
                    'proyecto_mas_concepto4_proyecto2' : str(nombre_proyecto2) + ' ref: ' + str(concepto_importe4_proyecto2),

                    'fecha_importe1_proyecto2' : fecha_importe1_proyecto2,
                    'fecha_importe2_proyecto2' : fecha_importe2_proyecto2,
                    'fecha_importe3_proyecto2' : fecha_importe3_proyecto2,
                    'fecha_importe4_proyecto2' : fecha_importe4_proyecto2,


                    'importe1_proyecto2'	 :  importe1_proyecto2,
                    'importe2_proyecto2'	 :  importe2_proyecto2,
                    'importe3_proyecto2'	 :  importe3_proyecto2,
                    'importe4_proyecto2'	 :  importe4_proyecto2,
                    'total_proyecto2'	 : total_proyecto2,
                    'nombre_proyecto3'	 : nombre_proyecto3,
                    'proveedor_proyecto3'	 :  proveedor_proyecto3,
                    'cif_proveedor_proyecto3'	 : cif_proveedor_proyecto3,
                    'proveedor2_proyecto3'	 : proveedor2_proyecto3,
                    'cif_proveedor2_proyecto3'	 : cif_proveedor2_proyecto3,
                    'proveedor3_proyecto3' : proveedor3_proyecto3,
                    'cif_proveedor3_proyecto3' : cif_proveedor3_proyecto3,
                    'proveedor4_proyecto3' : proveedor4_proyecto3,
                    'cif_proveedor4_proyecto3' : cif_proveedor4_proyecto3,

                    'duracion_proyecto3'	 : duracion_proyecto3,

                    'concepto_importe1_proyecto3' : concepto_importe1_proyecto3,
                    'concepto_importe2_proyecto3' : concepto_importe2_proyecto3,
                    'concepto_importe3_proyecto3' : concepto_importe3_proyecto3,
                    'concepto_importe4_proyecto3' : concepto_importe4_proyecto3,

                    'proyecto_mas_concepto1_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe1_proyecto3),
                    'proyecto_mas_concepto2_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe2_proyecto3),
                    'proyecto_mas_concepto3_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe3_proyecto3),
                    'proyecto_mas_concepto4_proyecto3' : str(nombre_proyecto3) + ' ref: ' + str(concepto_importe4_proyecto3),
        



                    'fecha_importe1_proyecto3' : fecha_importe1_proyecto3,
                    'fecha_importe2_proyecto3' : fecha_importe2_proyecto3,
                    'fecha_importe3_proyecto3' : fecha_importe3_proyecto3,
                    'fecha_importe4_proyecto3' : fecha_importe4_proyecto3,
                    'importe1_proyecto3'	 :  importe1_proyecto3,
                    'importe2_proyecto3'	 :  importe2_proyecto3,
                    'importe3_proyecto3'	 :  importe3_proyecto3,
                    'importe4_proyecto3'	 :  importe4_proyecto3,
                    'total_proyecto3'	 : total_proyecto3,
                    'nombre_proyecto4'	 : nombre_proyecto4,
                    'proveedor_proyecto4'	 :  proveedor_proyecto4,
                    'cif_proveedor_proyecto4'	 : cif_proveedor_proyecto4,
                    'proveedor2_proyecto4'	 :  proveedor2_proyecto4,
                    'cif_proveedor2_proyecto4'	 : cif_proveedor2_proyecto4,
                    'proveedor3_proyecto4' : proveedor3_proyecto4,
                    'cif_proveedor3_proyecto4' : cif_proveedor3_proyecto4,
                    'proveedor4_proyecto4' : proveedor4_proyecto4,
                    'cif_proveedor4_proyecto4' : cif_proveedor4_proyecto4,

                    'duracion_proyecto4'	 : duracion_proyecto4,

                    'concepto_importe1_proyecto4' : concepto_importe1_proyecto4,
                    'concepto_importe2_proyecto4' : concepto_importe2_proyecto4,
                    'concepto_importe3_proyecto4' : concepto_importe3_proyecto4,
                    'concepto_importe4_proyecto4' : concepto_importe4_proyecto4,

                    'proyecto_mas_concepto1_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe1_proyecto4),
                    'proyecto_mas_concepto2_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe2_proyecto4),
                    'proyecto_mas_concepto3_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe3_proyecto4),
                    'proyecto_mas_concepto4_proyecto4' : str(nombre_proyecto4) + ' ref: ' + str(concepto_importe4_proyecto4),
        

                    'fecha_importe1_proyecto4' : fecha_importe1_proyecto4,
                    'fecha_importe2_proyecto4' : fecha_importe2_proyecto4,
                    'fecha_importe3_proyecto4' : fecha_importe3_proyecto4,
                    'fecha_importe4_proyecto4' : fecha_importe4_proyecto4,


                    'importe1_proyecto4'	 :  importe1_proyecto4,
                    'importe2_proyecto4'	 :  importe2_proyecto4,
                    'importe3_proyecto4'	 :  importe3_proyecto4,
                    'importe4_proyecto4'	 :  importe4_proyecto4,
                    'total_proyecto4'	 : total_proyecto4,
                    'totalproyecto' : totalproyecto,
                    'descripcion_actividades_proyecto1' : descripcion_actividades_proyecto1,
                    'descripcion_actividades_proyecto2' : descripcion_actividades_proyecto2,
                    'descripcion_actividades_proyecto3' : descripcion_actividades_proyecto3,
                    'descripcion_actividades_proyecto4' : descripcion_actividades_proyecto4,
                    'detalle_de_las_soluciones_implantadas_pr1' : detalle_de_las_soluciones_implantadas_pr1,
                    'detalle_de_las_soluciones_implantadas_pr2' : detalle_de_las_soluciones_implantadas_pr2,                 
                    'detalle_de_las_soluciones_implantadas_pr3' : detalle_de_las_soluciones_implantadas_pr3,
                    'detalle_de_las_soluciones_implantadas_pr4' : detalle_de_las_soluciones_implantadas_pr4,                
                    'fecha_inicio_proyecto' : fecha_inicio_proyecto,
                    'fecha_fin_proyecto' : fecha_fin_proyecto

                
            
        }

        
        print ( " descripcion PROVEEDOR 3 Y 4    " + str( proveedor3_proyecto1) + str(proveedor4_proyecto1))
        A = input( "PULSE UNA TECLA ....")

        
        

        return context








# no usado #####################################
# no usado =======================================
# no usado #####################################

# ==============================================================
# ==============================================================
# ==============================================================


def fechadocumento():
    print('Fecha documento a generar ')
    
    # cif=input()
    #cif = '44237153B'
    global fecha_documento
    fecha_documento = input(" Introducir fecha documento ")
    
def listar_ddp(dirdatos):
   
   # enlace trabajar fichero
   #  https://stackoverflow.com/questions/2953834/how-should-i-write-a-windows-path-in-a-python-string-literal
   rutaorigen  = 'C:/desarrollo/genanexcam/'
   global rutaendiscoglobal

   print (" razon_social_global  " + str(dirdatos))
   
   descargas = "c:\\users\\alfonso\\donwloads"
   
   # rutaendiscoglobal = descargas
   print ( "ruta en disco global fuera " + rutaendiscoglobal ) 
   print ( "ruta descargas  " + descargas) 

   # concatenear cadenas
   
   rutacompletaddp =  "c:/users/alfonso/downloads/ddp*.docx"
   rutacompletadiag =  "c:/users/alfonso/downloads/diag*.docx"


   
   #rutacompletaddp1 =  f"{descargas1}/ddp*.docx"
   #rutacompletadiag =  f"{descargas}\\diagnostico*.docx"

   print ("ruta para desccargar "  + rutacompletaddp  )
   print ("ruta para desccargar "  + rutacompletadiag  )
   
   
   
   listddp = glob.glob(rutacompletaddp)   
   listdiag = glob.glob(rutacompletadiag)

   # imprimir lista
   print('listddp {}'.format(listddp))
   # print('listddp1 {}'.format(listddp1))
   # imprimir lista
   print('listddp {}'.format(listdiag))


   hayddp = "yes"
   haydiag = "yes"
   if listddp: 
     print('listddp {}'.format(listddp))
   else:
        hayddp = "no"
        user_input = input(" No hay ddp para mover, seguir ?")
        if user_input.lower() == 'yes':
            print('user typed yes')           
        elif user_input.lower() == 'no':
            return
        listddp.append("no ddp")
   if listdiag:
     print(listdiag[0])
   else: 
     haydiag = "no"
     user_input = input(" No hay DIAG para mover, seguir ?")
     listdiag.append("no diag")
   
     if user_input.lower() == 'yes':
            print('user typed yes')           
     elif user_input.lower() == 'no':
            return

   print(rutaendiscoglobal)
   wait = input("Press Enter to continue." )
   print("something")

   
   rutadisco = rutaendiscoglobal

   if haydiag == "yes":    
        shutil.move( listdiag[0], f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
   if hayddp == "yes":    
        for file in listddp:
            print(file)
            print("directorio origen"  + str(rutaorigen) + str(file) )
            shutil.move( file,  f"{rutadisco}\\{os.path.basename(file)}")


def mover_ppi():
   
   # enlace trabajar fichero
   #  https://stackoverflow.com/questions/2953834/how-should-i-write-a-windows-path-in-a-python-string-literal
   rutaorigen  = 'C:/desarrollo/genanexcam/'
   global rutaendiscoglobal
   
   descargas = "c:\\users\\alfonso\\donwloads"
   
   # rutaendiscoglobal = descargas
   print ( "ruta en disco global fuera " + rutaendiscoglobal ) 
   print ( "ruta descargas  " + descargas) 

   # concatenear cadenas
   
   rutacompletappi =  "c:/users/alfonso/downloads/*ppi*.docx"
   


   
   #rutacompletaddp1 =  f"{descargas1}/ddp*.docx"
   #rutacompletadiag =  f"{descargas}\\diagnostico*.docx"

   print ("ruta para desccargar "  + rutacompletappi  )
   
   
   
   
   listppi = glob.glob(rutacompletappi)   
   

   # imprimir lista
   print ( "LISTA COMPLETA PPI ")
   print('listppi {}'.format(listppi))
   a = input ( " PUlse una tecla ...")
   # print('listddp1 {}'.format(listddp1))
   # imprimir lista
   


   hayppi = "yes"
   haydiag = "yes"
   if listppi: 
     print('listppi {}'.format(listppi))
   else:
        hayppi = "no"
        user_input = input(" No hay ppi para mover, seguir ?")
        if user_input.lower() == 'yes':
            print('user typed yes')           
        elif user_input.lower() == 'no':
            return
        listppi.append("no ppi")
   

   print(rutaendiscoglobal)
   wait = input("Press Enter to continue." )
   print("something")

   
   rutadisco = rutaendiscoglobal


   if hayppi == "yes":    
        for file in listppi:
            print(file)
            print("directorio origen"  + str(rutaorigen) + str(file) )
            shutil.move( file,  f"{rutadisco}\\{os.path.basename(file)}")

def preparar_escenario_ddp(): 
   rutaorigen  = 'C:/desarrollo/genanexcam'
   ruta_ddp_diag_test  = r'C:\desarrollo\genanexcam\ddp-diag-test'
   global rutaendiscoglobal
   buscar = ruta_ddp_diag_test +  '\\*ddp*.docx'
   listddp = glob.glob( buscar )
   print(listddp)
   wait = input("Press Enter to continue." )
   
   listdiag = glob.glob(ruta_ddp_diag_test +  '/diagnostico*.docx')
   for file in listddp:
        print("fichero a MOVER " + file)
        wait = input("Press Enter to continue " )
        print("something")
        print("directorio origen"  + str(rutaorigen) + str(file) )
        shutil.copy(file,  rutaorigen + '/' +  os.path.basename(file))
   for file in listdiag:
        print(file)
        print("directorio origen"  + str(rutaorigen) + str(file) )
        shutil.copy(file,  rutaorigen + '/' +  os.path.basename(file))





def recordatorio_inno():
    send_account = None
    
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'alfonso.raggio@camarahuelva.com':
            send_account = account
            break

    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add('alfonso.raggio@camarahuelva.com')
    mail_item.Subject = 'IMPORTATNTE - RECORDATORIO - Programa InnoCámaras 2023 - Plazos de ejecución y tramites a seguir'
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = "C:/desarrollo/genanexcam/doc-convocatoria/Justific@-Guia_de_usuario.pdf"
    attachment2 = "C:/desarrollo/genanexcam/doc-convocatoria/05-Anexo IV. Condiciones de participación, justificación y gastos elegibles Fase de Ayudas.pdf"
    mail_item.Attachments.Add(attachment)
    mail_item.Attachments.Add(attachment2)
    filename = "innocamaras.png"                   
    attachment3 =  'C:\\desarrollo\\genanexcam\\img\\innocamaras.png'
    
    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    


    mail_item.HTMLBody ='''
<html><body>


<p>Nos encontramos próximos a la recta final de la implantación de proyectos y
sois muy pocas las empresas que habéis comenzado a implantar las soluciones
elegidas. Por este motivo, os recordamos los plazos de ejecución y
justificación, así como algunas consideraciones de interés.</p>

<ol type="1">
 <li>En primer lugar os recordamos los pasos a seguir una
     vez hayáis finalizado la fase de diagnóstico Fase I y entregado firmado el
     Anexo 18 - Acta recepción y cuestionario de satisfacción Fase I.</li>
 <ol type="1">
  <li>Entregar los
      presupuestos de los proyectos que queráis poner en marcha.</li>
  <li>Firmar la documentación
      de inicio de la fase de implantación que os proporcionará Cámara de
      Comercio una vez recibidos y dado el visto bueno a los presupuestos.</li>
  <li>Hasta que no se firme
      esta documentación no se podrá comenzar a ejecutar gastos del proyecto.</li>
 </ol>
</ol>

<ol type="1">
 <li>Los plazos a cumplir son los siguientes:</li>
 <ol type="1">
  <li><b>15 de agosto de 2023</b> - Fecha tope para
      enviar a la Cámara de Comercio las facturas de los gastos a justificar.
      Serán enviadas por email a la Cámara de Comercio de Huelva a través de la
      dirección <a href="mailto:innocamaras@camarahuelva.com">innocamaras@camarahuelva.com</a></li>
  <li><b>31 de agosto de 2023</b> - Fecha tope para
      realizar los pagos de todos los gastos que se vayan a presentar en la
      justificación. Se recomienda la consulta del&nbsp; anexo IV a la
      convocatoria.</li>
  <li>Una vez recibidas las
      facturas, la Cámara le remitirá los anexos 20 (Registro Prestación
      Servicio y Seguimiento - Fase II) y 21 (Memoria Ejecución Proyecto y
      Cuestionarios) que deberán ser cumplimentados y firmados para ser
      remitidos a la Cámara antes del <b>15 de septiembre de 2023 </b>a través
      de la dirección <a href="mailto:innocamaras@camarahuelva.com">innocamaras@camarahuelva.com</a></li>
  <li>Toda la documentación
      para la justificación deberá ser aportada antes del <b>15 de septiembre
      de 2023</b>. A partir de esta fecha no se podrá aportar documentación salvo
      en el caso de que se le comunique la necesidad de subsanar alguna de la
      documentación ya aportada. Esta documentación debe ser subida a la
      plataforma justific@.</li>
  <li>La relación de
      documentación a aportar a través de la plataforma justific@ es la
      siguiente:</li>
  <ol type="1">
   <li>Facturas.</li>
   <li>Justificantes de pago.
       Orden de transferencia + Extracto bancario. Puede ser sustituido por
       certificado bancario de la transferencia realizada.</li>
   <li>Evidencias de los
       gastos realizados.</li>
   <li>Evidencias del
       cumplimiento de la publicidad del FEDER</li>
   <li>Documento de
       identificación financiera. Certificado de titularidad de la cuenta
       bancaria de abono de la ayuda.</li>
  </ol>
  <li>Además de la
      documentación que deben subir a la plataforma justific@, deberán aportar
      los anexos 20 (Registro Prestación Servicio y Seguimiento - Fase II) y 21
      (Memoria Ejecución Proyecto y Cuestionarios) debidamente cumplimentados y
      firmados, así como otros anexos relativos al asesoramiento recibido. Como
      se indicaba con anterioridad, estos anexos serán proporcionados por la
      Cámara de Comercio una vez recibamos las facturas.<b><i> </i></b></li>
 </ol>
</ol>

<p>Para poder explicaros todo el proceso de justificación, tuvo lugar una 
jornada en la que se explicaba todo el proceso. En realidad la herramienta es muy fácil y el proceso es sencillo, además siempre contareis con el apoyo
de los técnicos que os están gestionando vuestro expediente.</p>

El vídeo fue una charla &nbsp;de aproximadamente 45 minutos. &nbsp;Servirá para
entender lo más importante a la hora de subir los documentos de justificación.</u></p>

<p>Enlace para vídeo</p>

<p><a href="https://us06web.zoom.us/meeting/register/tZIkd-mhrDwrE9CMR1FDfXZpLsABjvvIsP8h">https://us06web.zoom.us/meeting/register/tZIkd-mhrDwrE9CMR1FDfXZpLsABjvvIsP8h</a></p>

<p>Como siempre nos tienen a su disposición para resolver cualquier consulta o
duda.</p>

<p>Un cordial saludo.</p>

<p>&nbsp;</p>

&nbsp;
&nbsp;&nbsp;</body></html>

''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'



# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return







def publicidad_tic():
    global email
    global nombre_solicitante
    send_account = None
    
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'alfonso.raggio@camarahuelva.com':
            send_account = account
            break

    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
    
    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Cumplir con la obligación de publicidad del programa TICCamaras 2023' + ' - ' + nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = "C:/desarrollo/genanexcam/carteles/carteles-tic/Bandera UE.jpg"
    attachment2 = "C:/desarrollo/genanexcam/carteles/carteles-tic/Cartel A3 TICCámaras.pdf"
    mail_item.Attachments.Add(attachment)
    mail_item.Attachments.Add(attachment2)
    filename = "innocamaras.png"                   
    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo\\tic\\ticcamaras.jpg'
    
    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    


    mail_item.HTMLBody ='''
<html><body>

<html><body> <p>


</p><p>Estimado Sr./ Sra.</p>

<p>La normativa marcada por FEDER obliga a cumplir una serie de
obligaciones en materia de publicidad del Programa TICCámaras. A continuación
se describen las acciones a realizar.</p>

<p>&nbsp;</p>

<p>1.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Las
empresas beneficiarias de la Fase II del Programa TICCámaras, deben colgar en
sus instalaciones un cartel en tamaño A3 del cual se adjunta archivo PDF en
este email.Las
empresas beneficiarias de la Fase II del Programa TICCámaras, deben colgar en
sus instalaciones un cartel en tamaño A3 del cual se adjunta archivo PDF en
este email.</p>

<p>2.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Las empresas beneficiarias de la
Fase II del Programa TICCámaras, Incorporarán en la página web o sitio de
Internet, si lo tuviera, una breve descripción de la operación con sus
objetivos y resultados, y destacando el apoyo financiero de la Unión. Cuando
sea manifiesta la imposibilidad de cumplimiento estricto de lo indicado en este
punto, la pyme o autónomo entregará un documento acreditando la imposibilidad
de cumplimiento.Las empresas beneficiarias de la
Fase II del Programa TICCámaras, Incorporarán en la página web o sitio de
Internet, si lo tuviera, una breve descripción de la operación con sus
objetivos y resultados, y destacando el apoyo financiero de la Unión. Cuando
sea manifiesta la imposibilidad de cumplimiento estricto de lo indicado en este
punto, la pyme o autónomo entregará un documento acreditando la imposibilidad
de cumplimiento.</p>

<p>En este 
apartado, figurará el logotipo de la Unión Europea (se adjunta en este email en
formato JPG), referencia al Fondo y lema junto con la siguiente frase:</p>

<p>“[Nombre de la
empresa] ha sido beneficiaria del Fondo Europeo de Desarrollo Regional cuyo
objetivo es [frase objetivo temático] y gracias al que ha [descripción de la
operación] para [objetivo del programa]. [Fecha de la acción]. Para ello ha
contado con el apoyo del [nombre del programa] de la Cámara de Comercio de
[nombre de la Cámara].”</p>

<p>Una manera de hacer Europa</p>


<p style="color:red;">Ejemplo: <i>“Mavents
Eventos Especiales ha sido beneficiaria del Fondo Europeo de Desarrollo
Regional cuyo objetivo es mejorar el uso y la calidad de las tecnologías de la
información y de las comunicaciones y el acceso a las mismas y gracias al que
ha podido optimizar su sistema de gestión y el contacto con sus clientes a
través de su nueva aplicación móvil. Esta acción ha tenido lugar durante 2023.
Para ello ha contado con el apoyo del programa TICCámaras de la Cámara de
Huelva”. </i></p>


<p style = "color:red;">Una manera de hacer Europa</p>

<p>&nbsp;</p>

<p>La empresa deberá Justificar ante la Cámara el cumplimiento de esta normativa aportando la siguiente
documentación.</p>

<p>1.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Fotografía del cartel en A3 en algún lugar visible de su edificio.Fotografía del cartel en A3 en algún lugar visible de su edificio.</p>

<p>2.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Pantallazo de la página web con lo referido en el punto 2, o documento
de manifiesta imposibilidad de cumplimiento al no tener página web, todos ellos
debidamente fechados.Pantallazo de la página web con lo referido en el punto 2, o documento
de manifiesta imposibilidad de cumplimiento al no tener página web, todos ellos
debidamente fechados.</p>

<p>3.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Estas evidencias se adjuntarán
al anexo 13, el cual la Cámara de Comercio pondrá a su disposición una vez se
finalizada la ejecución del Plan de Innovación.Estas evidencias se adjuntarán
al anexo 13, el cual la Cámara de Comercio pondrá a su disposición una vez se
finalizada la ejecución del Plan de Innovación.</p>

<p>Un cordial saludo</p>


<p>&nbsp;</p>

&nbsp;
&nbsp;&nbsp;</body></html>

''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


    
# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return


def publicidad_inno():
    global email
    global nombre_solicitante
    send_account = None

    
    for account in outlook.Session.Accounts:
        if account.DisplayName == 'innocamaras@camarahuelva.com':
            send_account = account
            break

    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Cumplir con la obligación de publicidad del programa INNOCamaras 2023' + ' - ' + nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format
    attachment = "C:/desarrollo/genanexcam/carteles/carteles-inno/Bandera UE.jpg"
    attachment2 = "C:/desarrollo/genanexcam/carteles/carteles-inno/Cartel A3 Innocamaras.pdf"
    mail_item.Attachments.Add(attachment)
    mail_item.Attachments.Add(attachment2)
    filename = "innocamaras.png"                   
    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo\\inno\\innocamaras.png'
    
    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
    

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    


    mail_item.HTMLBody ='''
<html><body>

<html><body> <p>


</p><p>Estimado Sr./ Sra.</p>

<p>La normativa marcada por FEDER obliga a cumplir una serie de
obligaciones en materia de publicidad del Programa INNOCámaras. A continuación
se describen las acciones a realizar.</p>

<p>&nbsp;</p>

<p>1.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Las
empresas beneficiarias de la Fase II del Programa INNOCámaras, deben colgar en
sus instalaciones un cartel en tamaño A3 del cual se adjunta archivo PDF en
este email.Las
empresas beneficiarias de la Fase II del Programa INNOCámaras, deben colgar en
sus instalaciones un cartel en tamaño A3 del cual se adjunta archivo PDF en
este email.</p>

<p>2.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Las empresas beneficiarias de la
Fase II del Programa INNOCámaras, Incorporarán en la página web o sitio de
Internet, si lo tuviera, una breve descripción de la operación con sus
objetivos y resultados, y destacando el apoyo financiero de la Unión. Cuando
sea manifiesta la imposibilidad de cumplimiento estricto de lo indicado en este
punto, la pyme o autónomo entregará un documento acreditando la imposibilidad
de cumplimiento.Las empresas beneficiarias de la
Fase II del Programa INNOCámaras, Incorporarán en la página web o sitio de
Internet, si lo tuviera, una breve descripción de la operación con sus
objetivos y resultados, y destacando el apoyo financiero de la Unión. Cuando
sea manifiesta la imposibilidad de cumplimiento estricto de lo indicado en este
punto, la pyme o autónomo entregará un documento acreditando la imposibilidad
de cumplimiento.</p>

<p>En este 
apartado, figurará el logotipo de la Unión Europea (se adjunta en este email en
formato JPG), referencia al Fondo y lema junto con la siguiente frase:</p>

<p>“[Nombre de la
empresa] ha sido beneficiaria del Fondo Europeo de Desarrollo Regional cuyo
objetivo es [frase objetivo temático] y gracias al que ha [descripción de la
operación] para [objetivo del programa]. [Fecha de la acción]. Para ello ha
contado con el apoyo del [nombre del programa] de la Cámara de Comercio de
[nombre de la Cámara].”</p>

<p>Una manera de hacer Europa</p>


<p style="color:red;">Ejemplo: <i>“Mavents
Eventos EspecialesEjemplo: “Productos La Higuera ha sido beneficiaria del Fondo Europeo de Desarrollo Regional 
cuyo objetivo es potenciar la investigación, el desarrollo tecnológico y la innovación y gracias al que ha podido
incorporar la innovación en sus procesos al adquirir inmovilizado de la fábrica y promocionarse internacionalmente 
para apoyar la creación y consolidación de empresas innovadoras. Esta acción ha tenido lugar durante 2023.
Para ello ha contado con el apoyo del programa InnoCámaras de la Cámara de Huelva”.”. </i></p>


<p style = "color:red;">Una manera de hacer Europa</p>

<p>&nbsp;</p>

<p>La empresa deberá Justificar ante la Cámara el cumplimiento de esta normativa aportando la siguiente
documentación.</p>

<p>1.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Fotografía del cartel en A3 en algún lugar visible de su edificio.Fotografía del cartel en A3 en algún lugar visible de su edificio.</p>

<p>2.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Pantallazo de la página web con lo referido en el punto 2, o documento
de manifiesta imposibilidad de cumplimiento al no tener página web, todos ellos
debidamente fechados.Pantallazo de la página web con lo referido en el punto 2, o documento
de manifiesta imposibilidad de cumplimiento al no tener página web, todos ellos
debidamente fechados.</p>

<p>3.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Estas evidencias se adjuntarán
al anexo 13, el cual la Cámara de Comercio pondrá a su disposición una vez se
finalizada la ejecución del Plan de Innovación.Estas evidencias se adjuntarán
al anexo 13, el cual la Cámara de Comercio pondrá a su disposición una vez se
finalizada la ejecución del Plan de Innovación.</p>

<p>Un cordial saludo</p>


<p>&nbsp;</p>

&nbsp;
&nbsp;&nbsp;</body></html>

''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'

    
# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return


def generar_deca():   
    # cif=input()
    cif_a_buscar = cifglobal
 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   

    valuestovar_deca(df_tab_empresa_fila_buscada)
    """  
    # se localiza por NIF en el dataframe que corresponde al tab ALTA USUARIO
    # df_tab_usuario_buscado = df_tab_alta_usuarios.loc[df_tab_alta_usuarios['documento_solicitante'] == cif_a_buscar]
    
    #print(" kkkk" + df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_string())

    print("empresa encontrada para imprimir ")
    #print(df_tab_empresa_fila_buscada)
    print("USUARIO BUSCADO ============> usuario encontrada para imprimir ")
    # print(df_tab_usuario_buscado)
    # buscar valor en un dataframe por columna
    #
    # con at df1.at[0,'randomcolumn'] => el AT NO FUNCIONA
    # He probado con iloc y funciona
    #==========================================
    #================= DATOS DE EMPRESA ==========
    # ============================================
    programa = df_tab_empresa_fila_buscada.iloc[0]['programa']
    fases = df_tab_empresa_fila_buscada.iloc[0]['Fases']
    nombre_solicitante = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    provincia = df_tab_empresa_fila_buscada.iloc[0]['provincia']
    poblacion = df_tab_empresa_fila_buscada.iloc[0]['poblacion']
    cp = df_tab_empresa_fila_buscada.iloc[0]['cp']
    email = df_tab_empresa_fila_buscada.iloc[0]['email']
    direccion = df_tab_empresa_fila_buscada.iloc[0]['direccion']
    telefono_solicitante = df_tab_empresa_fila_buscada.iloc[0]['telefono_solicitante']
    documento_representante = df_tab_empresa_fila_buscada.iloc[0]['docoumento_representante']
    email_representante = df_tab_empresa_fila_buscada.iloc[0]['email_representante']
    cargo = df_tab_empresa_fila_buscada.iloc[0]['cargo']
    tratamiento_representante = df_tab_empresa_fila_buscada.iloc[0]['tratamiento_representante']
    fases2 = df_tab_empresa_fila_buscada.iloc[0]['fases2']
    tecnico_justificar = df_tab_empresa_fila_buscada.iloc[0]['tecnico_justificar']
    dni_tecnico = df_tab_empresa_fila_buscada.iloc[0]['dni_tecnico']
    fecha_documento_inicio_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_documento_inicio_diagnostico']
    fecha_diagnostico = df_tab_empresa_fila_buscada.iloc[0]['fecha_diagnostico']
    anexo_18_firmado = df_tab_empresa_fila_buscada.iloc[0]['anexo_18_firmado']
    fecha_recepcion_presupuesto = df_tab_empresa_fila_buscada.iloc[0]['fecha_recepcion_presupuesto']
    fecha_envio_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_envio_anexo_19']
    envio_ppi = df_tab_empresa_fila_buscada.iloc[0]['envio_ppi']
    fecha_firma_ppi = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_ppi']
    fecha_firma_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_firma_anexo_19']
    duracion_del_plan = df_tab_empresa_fila_buscada.iloc[0]['duracion_del_plan']
    encuesta_satisfaccion = df_tab_empresa_fila_buscada.iloc[0]['encuesta_satisfaccion']
    proyectos = df_tab_empresa_fila_buscada.iloc[0]['proyectos']
    descripcion_empresa = df_tab_empresa_fila_buscada.iloc[0]['descripcion_empresa']
    fecha_registro_participacion_fase_i_anexo_18 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_fase_i_anexo_18']
    enviado_email_publicidad_ue = df_tab_empresa_fila_buscada.iloc[0]['enviado_email_publicidad_ue']
    fecha_registro_participacion_en_fase_ii_anexo_19 = df_tab_empresa_fila_buscada.iloc[0]['fecha_registro_participacion_en_fase_ii_anexo_19']    
    cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']
    nombre_representante = df_tab_empresa_fila_buscada.iloc[0]['nombre_representante']

    
    #empleados_empresa = df_tab_empresa_fila_buscada.iloc[0]['Empleados empresa']
    

    # ====================================
    # ==============   datos ALTA USAURIO 
    # ====================================
    # Razón social  - calculado
    # nifempresaE  = df_tab_usuario_buscado.iloc[0]['documento_solicitante']       
      # Estado NIF - calculado
    #tratamiento =   df_tab_usuario_buscado.iloc[0]['Tratamiento']
   
   
     

    ##### cif_empresa = df_tab_empresa_fila_buscada.iloc[0]['documento_solicitante']
    ##### nombre_o_razon_social = df_tab_empresa_fila_buscada.iloc[0]['nombre_solicitante']     
    print('cif de empresa ' + str(cif_empresa))
    print('nombre de empresa  ' + nombre_o_razon_social) 
    
    print ('Fecha del documento --> ' + str(fecha_documento))
    print ( ' REPRESENTANTE ' + str(nombre_representante.lower()))

    context = {
                'fecha_documento' : fecha_documento,
                 # =============== FICHA DE EMPRESA
                'nif_de_empresa' : cif_empresa,
                'nombre_representante' : nombre_representante.lower().title(),
                'razon_social' : nombre_o_razon_social       , #revisar si intruso
                'tecnico_justificar' :  tecnico_justificar,

                'programa' : programa,
                'fases' : fases,
                'nombre_solicitante' : nombre_solicitante,
                'provincia' : provincia,
                'poblacion' : poblacion,
                'cp' : cp,
                'email' : email,
                'direccion' : direccion,
                'telefono_solicitante' : telefono_solicitante,
                'documento_representante' : documento_representante,
                'email_representante' : email_representante,
                'cargo' : cargo,
                'tratamiento_representante' : tratamiento_representante,
                'fases2' : fases2,
                'tecnico_justificar' : tecnico_justificar,
                'dni_tecnico' : dni_tecnico,
                'fecha_documento_inicio_diagnostico' : fecha_documento_inicio_diagnostico,
                'fecha_diagnostico' : fecha_diagnostico,
                'anexo_18_firmado' : anexo_18_firmado,
                'fecha_recepcion_presupuesto' : fecha_recepcion_presupuesto,
                'fecha_envio_anexo_19' : fecha_envio_anexo_19,
                'envio_ppi' : envio_ppi,
                'fecha_firma_ppi' : fecha_firma_ppi,
                'fecha_firma_anexo_19' : fecha_firma_anexo_19,
                'duracion_del_plan' : duracion_del_plan,
                'encuesta_satisfaccion' : encuesta_satisfaccion,
                'proyectos' : proyectos,
                'descripcion_empresa' : descripcion_empresa,
                'fecha_registro_participacion_fase_i_anexo_18' : fecha_registro_participacion_fase_i_anexo_18,
                'enviado_email_publicidad_ue' : enviado_email_publicidad_ue,
                'fecha_registro_participacion_en_fase_ii_anexo_19':   fecha_registro_participacion_en_fase_ii_anexo_19,
                'cif_empresa' :cif_empresa,
                'nombre_o_razon_social' :  nombre_o_razon_social,
                

    






                # empleados que es un ccheckbox
                # 'inicio_de actividad'  son cuadros
                # sector actividad   son cuadros
                #'domicilio_social' :  'pendiente',
                #'codigo_postal' : codigopostal,
                #'localidad' : localidad,
                #'provincia' : provincia,
                #'email_general'  : correoelectronico,
                # 'pagina web' : paginaweb,
                # ================= DATOS DE CONSULTA
                #'nombre' : nombre,
                #'Primer_apellido' : primerapellido,
                #'Segundo_apellido' : segundoapellido,
                #'fechaasesoramiento' : str(fechaasesoramiento),
                #'dni' : nif,
                #'cargo' : cargo,
                #'email_contacto' : correoelectronico,
                #'fecha_consulta' : fechaasesoramiento,
                #'como_nos_ha_conocido' : comonoshaconocido,
                #'num_assistentes' : 'n/a',
                #'canalservicio' : correoelectronico,
                #'tematica' : 'tematica a tratar',
                #'observaciones' : 'Incluir observaciones de la consulta 
                
                
    }

   

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    print("la ruta de trabajo es " + path)
    doc = DocxTemplate('anexo19.docx')        
    doc.render(context)
    doc.save('generated_anexo19_'+ cifglobal + '.docx')           
    return
 """








def convertir_pdf():
    
    global rutaendiscoglobal
    global rutadescargas

    ruta = "C:/desarrollo/genanexcam/generated_anexo19_" + cifglobal + ".docx"
    


    rutacompletaddp =  f"{rutaendiscoglobal}\\*ddp*.docx"
    rutacompletadiag =  f"{rutaendiscoglobal}\\*diagnostico*.docx"

    print ("ruta para desccargar "  + rutacompletaddp  )
    print ("ruta para desccargar "  + rutacompletadiag  )
   
    listddp = glob.glob(rutacompletaddp)
    listdiag = glob.glob(rutacompletadiag)

    hayddp = "yes"
    haydiag = "yes"
    
    if listddp: 
      print('listddp {}'.format(listddp))
    else:
        hayddp = "no"
        user_input = input(" No hay ddp para mover, seguir ?")
        if user_input.lower() == 'yes':
            print('user typed yes')           
        elif user_input.lower() == 'no':
            return
        listddp.append("no ddp")
    if listdiag:
      print(listdiag[0])
    else: 
      haydiag = "no"
      user_input = input(" No hay DIAG para mover, seguir ?")
      listdiag.append("no diag")
    
      if user_input.lower() == 'yes':
            print('user typed yes')           
      elif user_input.lower() == 'no':
            return

    print(rutaendiscoglobal)
    wait = input("Press Enter to continue." )
    print("something")


    rutadisco = rutaendiscoglobal

    if haydiag == "yes":    
            #shutil.move( listdiag[0], f"{rutadisco}\\{os.path.basename(listdiag[0])}"   ) 
            file_base_name = os.path.basename(listdiag[0])            
            file_no_ext = file_base_name.rsplit('.', 1)[0]

            file_target = f"{rutadisco}\\{file_no_ext}.pdf"
            
            print (" file target DIAG "  + file_target  )
            print (" file origen " + f"{rutadisco}\\{os.path.basename(listdiag[0])}" )
            convert(f"{rutadisco}\\{os.path.basename(listdiag[0])}", file_target)
    if hayddp == "yes":    
            for file in listddp:
                print(file)
                print("directorio origen"  + str(rutadescargas) + str(file) )
                #shutil.move( file,  f"{rutadisco}\\{os.path.basename(file)}")
                file_base_name = os.path.basename(file)    
                file_no_ext = file_base_name.rsplit('.', 1)[0]        
                print("filchero si extension " + file_base_name.rsplit('.', 1)[0])   
                file_target = f"{rutadisco}\\{file_no_ext}.pdf"
                print (" file target DDP "  + file_target  )
                print (" file origen " + f"{rutadisco}\\{os.path.basename(listdiag[0])}" )
                convert(f"{rutadisco}\\{os.path.basename(file)}", file_target)



def crear_draft_ejecutar_gasto(dirdatos):
    global rutaendiscoglobal
    global proyectos
    rutafiles = rutaendiscoglobal
    adjuntos = []
    send_account = None
  
       
   
    if programa == "INNOCAMARAS":
        for account in outlook.Session.Accounts:   
          if account.DisplayName == 'alfonso.raggio@camarahuelva.com':
            send_account = account
            break
    if programa == "TICCAMARAS":        
        for account in outlook.Session.Accounts:
           if account.DisplayName == 'alfonso.raggio@camarahuelva.com':           
            send_account = account
            break
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")      
        b = input(" pulse una tecla....")
           

    ruta_anexo19 = os.path.join(str(rutafiles), "*anexo19*.pdf")
    ruta_ppi = os.path.join(str(rutafiles), "*ppi*.pdf")
    listppi = glob.glob(ruta_anexo19)
    listanexo19 =  glob.glob(ruta_ppi)
    mail_item = outlook.CreateItem(0)   # 0: olMailItem

    # mail_item.SendUsingAccount = send_account not working
    # the following statement performs the function instead
    mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

    mail_item.Recipients.Add(email)
    mail_item.Subject = 'Documentación Inicio Fase II InnoCámaras 2023 - ' + programa + ' 2023 - ' + nombre_solicitante
    mail_item.BodyFormat = 2   # 2: Html format

    for file in listppi:      
        mail_item.Attachments.Add(file)
    for file in listanexo19:
        mail_item.Attachments.Add(file)

    path = str(os.getcwd())
    # si quisieramos cambiar la ruta os.getcwd(path)
    
    if programa == "INNOCAMARAS":
        filename = "\\inno\\innocamaras.png"            
        
       
    elif programa == "TICCAMARAS" :
        filename = "\\tic\\ticcamaras.jpg"     
        
    else:
        print( " NOMBRE DE PROGRAMA INCORRECTO ")
        return
    
    attachment3 =  'C:\\desarrollo\\genanexcam\\firma-correo' + filename
    
    attach = mail_item.Attachments.Add(attachment3)  
    attach.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # data:image/png;base64,

    with open(attachment3, "rb") as image:
            image_string = base64.b64encode(image.read())
            # imagede_string = base64.b64decode(image_string))

    

    if programa == "INNOCAMARAS":
    
        mail_item.HTMLBody = '''
<html><body> <p> 




</p><p>Buenos días,</p>

<p>&nbsp;A partir de este
momento, podéis ejecutar el proyecto y facturar con proveedores y realizar los pagos a los mismos con las fechas límites que marcan el programa. </p>
	<p>Adjuntamos en este correo:</p>
	<p>Anexo 19 y PPI firmados por nuestra parte</p>
	
	<p><b>Tened en cuenta por favor toda la
     documentación del Programa</b> en cuanto a plazos, pagos y formas de
     ejecutar el proyecto;<b> especialmente la información y plazos</b> del proyecto
     incluidos y detallados en el Anexo 10 DECA y Anexo 19. La información de este correo es solo un extracto, La informaicón completa y obligaciones vienen en la convocatoria</p>


<p>IMPORTANTE
(Recordar a vuestro proveedor para que incluya, es altamente recomendado de cara a auditoría):</p><ul type="disc">
 <li>En las <b>facturas</b>, incluir en alguna en alguna parte:</li>
 <ul type="circle">
  <ul type="square">
   <ul type="disc">
    <li>El nombre del proyecto (que sería el
        <b><u>nombre del proyecto que viene en el archivo DDP</u></b> del
        proyecto que se factura, ejemplo “Proyecto: Soluciones de comercio
        electrónico”)</li>
					<li>Incluir una nota o texto : &quot;Duración del proyecto del dd/mm/yyyy al 31/08/2023&quot; &nbsp; siendo dd/mm/yyyy la fecha en el que PPI y anexo 19 están firmados por ambas partes</li>
    <li>Incluir un texto o leyenda la factura, podría ser TEXTUALMENTE algo como: &nbsp;
&nbsp;</li></ul></ul></ul></ul><p><b>&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp; **&nbsp;Programa INNOCámaras cofinanciado por el Fondo Europeo de Desarrollo Regional
(FEDER) de la unión Europeo,&nbsp;</b></p>
	<p><b>&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	dentro del Programa Operativo de Crecimiento Inteligente
FEDER 2014-2020 (POCint)&nbsp;</b> &nbsp;</p><ul type="disc">
 <li>Los archivos de publicidad del programa que aparecen en la web, se enviarán a continuación en correo separadamente al presente. </li><li>Logos de la UE en la web corporativa o, en caso de no tener web, aportar certificado de no tener web (preguntar al Técnico de la Cámara)</li>
		<li>Cartel FEDER en el la entrada al edificio, en la oficina, etc....,deberá colocarse a partir de este momento.</li></ul>
	

<p>&nbsp;</p>


</body></html>

    ''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'


    
    elif programa == "TICCAMARAS" :
        mail_item.HTMLBody = '''
<html><body> <p> 




</p><p>Buenos días,</p>

<p>&nbsp;A partir de este
momento, podéis ejecutar el proyecto y facturar con proveedores y realizar los pagos a los mismos con las fechas límites que marcan el programa. </p>
	<p>Adjuntamos en este correo:</p>
	<p>Anexo 19 y PPI firmados por nuestra parte</p>
	
	<p><b>Tened en cuenta por favor toda la
     documentación del Programa</b> en cuanto a plazos, pagos y formas de
     ejecutar el proyecto;<b> especialmente la información y plazos</b> del proyecto
     incluidos y detallados en el Anexo 10 DECA y Anexo 19. La información de este correo es solo un extracto, La informaicón completa y obligaciones vienen en la convocatoria</p>

     
<p>IMPORTANTE
(Recordar a vuestro proveedor para que incluya, es altamente recomendado de cara a auditoría):</p><ul type="disc">
 <li>En las <b>facturas</b>, incluir en alguna en alguna parte:</li>
 <ul type="circle">
  <ul type="square">
   <ul type="disc">
    <li>El nombre del proyecto (que sería el
        <b><u>nombre del proyecto que viene en el archivo DDP</u></b> del
        proyecto que se factura, ejemplo “Proyecto: Soluciones de comercio
        electrónico”)</li>
					<li>Incluir una nota o texto : &quot;Duración del proyecto del dd/mm/yyyy al 31/08/2023&quot; &nbsp; siendo dd/mm/yyyy la fecha en el que PPI y anexo 19 están firmados por ambas partes</li>
    <li>Incluir un texto o leyenda la factura, podría ser TEXTUALMENTE algo como: &nbsp;
&nbsp;</li></ul></ul></ul></ul><p><b>&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp; **&nbsp;Programa TICCámaras cofinanciado por el Fondo Europeo de Desarrollo Regional
(FEDER) de la unión Europeo,&nbsp;</b></p>
	<p><b>&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	&nbsp;	dentro del Programa Operativo de Crecimiento Inteligente
FEDER 2014-2020 (POCint)&nbsp;</b> &nbsp;</p><ul type="disc">
 <li>Los archivos de publicidad del programa que deberán aparecer en la web y en las creaciones de los proyectos (consultar con el técnico de Cámara). Estos archivos se enviarán a continuación en correo separadamente al presente. </li><li>Logos de la UE en la web corporativa o, en caso de no tener web, aportar certificado de no tener web (preguntar al Técnico de la Cámara)</li>
		<li>Cartel FEDER en el la entrada al edificio, en la oficina o espacio de trabajo (aunque esté en el propio domicilio), etc....,deberá colocarse a partir de este momento.</li></ul>
	

<p>&nbsp;</p>


</body></html>

    ''' +  '<p>' +   '</p>  <img src="data:image/png;base64, ' + '' + str(image_string.decode('utf-8')) + '">'

# <p></p><p></p><p>&nbsp;</p> </body></html>    
    mail_item.Save()
    return



def mover_ppi_anexo19():

   rutacompletappi =  rutaendiscoglobal + '\\*ppi*.pdf'
   rutacompletaanexo19 =  rutaendiscoglobal + '\\*anexo19*.pdf'
   rutacompleta_move_anexoppi =  rutaendiscoglobal + '/otros/*ppi*.pdf'
   rutacompleta_move_anexo19 =  rutaendiscoglobal + '/otros/*anexo19*.pdf'
   rutaotros = rutaendiscoglobal + '\\otros'

   listppi = glob.glob(str(rutacompletappi))   
   listanexo19 = glob.glob(str(rutacompletaanexo19))

   print (" LIST PPI  ")
   print('listddp {}'.format(listppi))

   print (" LIST ANEXO 19  ")
   print('listddp {}'.format(listanexo19))

   for file in listppi:
        shutil.move( file,  str(rutaotros))
   for file in listanexo19:
        shutil.move( file, str(rutaotros))
   
   a = input ( "pulse una tecla .....")

          
def mostrar_menu(opciones):
    print('Seleccione una opción:')
    for clave in opciones:
        # print(f' {opciones[clave][0]}')
         print(f' {clave}){opciones[clave][0]}')
    
  

    
    
def ejecutar_opcion(opcion, opciones,dirdatos):
    print("DENTRO DE EJECUTAR OPCION ")
    print(dirdatos)
    b = input(" PRIMER PRIMER DENTRO DE EJECUTAR DATOS ")
   
    if opcion == '1':
        a = input ( " la opcion es uno ")
        dirdatos = procbuscarcif()
        print(" dir datos .................. "  )
        print(str(dirdatos))
        a = input ( " Estos son los datos del diccionario .....")                                                  
    elif opcion == '2':
        fechadocumento()
    elif opcion == '4':
        listar_ddp(dirdatos)       
    elif opcion == '6':
        convertir_pdf()  
    elif opcion == '14':
        crear_draft_ejecutar_gasto(dirdatos)
    elif opcion == '7':
        generar_ficha_anexo18_2(dirdatos)
    elif opcion == '8':
        crear_draft_anexo18(dirdatos)
    elif opcion == '9':
        generar_ficha_anexo19(dirdatos)
    elif opcion == '10':
        mover_ppi()
    elif opcion == '11' :
        crear_draft_anexo19()
    elif opcion == '14':
        crear_draft_ejecutar_gasto(dirdatos)
    elif opcion == '18':
        generar_ficha_anexo20(dirdatos)
    elif opcion == '20':
        creardraft_anexo20anexo21()
    elif opcion == '99':
        prueba_modificar()
    return dirdatos


def leer_opcion(opciones):
    while (a := input('Opción: ')) not in opciones:
        print('Opción incorrecta, vuelva a intentarlo.')
    return a
    
def generar_menu(opciones, opcion_salida, dfdatos):
    global opcionmenu
    opcion = None    
    opcionmenu = opcion
    cifglobal = ""
    
    
    while opcion != opcion_salida:
        mostrar_menu(opciones)
        opcion = leer_opcion(opciones)
        a = input ( "despues de leer_opcion  ... " )
        # dirdatos =  leer_datos(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        dirdatos = ejecutar_opcion(opcion, opciones, dfdatos)
        dfdatos = dirdatos
        print (" DF DATOS VALE ")
        print(dfdatos)
        a = input(" no es lo que hemos hecho ")
        # dirdatos =  leer_datos(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        #ejecutar_opcion(opcion, opciones,dirdatos)
                
        print(" el cif en el while es " + cifglobal)
        print () # se imprime opcion en blanco para clarificar salir de pantalla

        
def menu_principal(df):
    print("menu - el alor del cif es " + cif)
    opciones = {
        '1': ('1.- Buscar por CIF', procbuscarcif),
        '2': ('2.- Fecha Documento', fechadocumento),       
        '4': ('4.- Listar y mover Diagnostico y DDPs - reqquiere cif', listar_ddp),
        '6': ('6.- convertir DIAGNOSTICO y DDP a PDF ', convertir_pdf),
        '7': ('7.- Crear ANEXO 18  ',generar_ficha_anexo18_2),
        '8': ('8- Crear draft de correo 18 + diagnostico ',crear_draft_anexo18),  
        '9': ('9.- Generar Anexo 19', generar_ficha_anexo19),
        '10': ('10.- mover PPI', mover_ppi),
        '11': ('11.- Crear draft 19 + ppi ',crear_draft_anexo19),    
        '13': ('13.- Mover 19 + ppi NO FIRMADOS antes de guardar los FIRMADOS -a carpeta OTROS- ', mover_ppi_anexo19),    
        '14': ('14.- RESPONDER a correo conn PPI + ANEXO 19 firmados por nosotros ',crear_draft_ejecutar_gasto) ,
        '18': ('18.- Generar anexo 20 ',generar_ficha_anexo20) ,
        '20': ('20.- Crear draft anexo 20 y anexo 21', creardraft_anexo20anexo21),        
        '29': ('17.- NO SELECCIONAR - Enviar anexo 19', enviar_anexo19),  
        '30': ('20.- Preparar Test ddp ', preparar_escenario_ddp),
        '33': ('23.- recordatorio TIC', recordatorio_inno),
        '35': ('26.- Publicidad TIC ', publicidad_tic),
        '37': ('29.- Publicidad INNO ', publicidad_inno),
        '39': ('32.- deca inno ', generar_deca),
        '99': ('99.- PRUEBA MODIFICAR ',prueba_modificar),
        # '40': ('40.- CIF oap ', buscar_cif_oap),  
        #'41': ('41.- Salir', main)
    }    
    generar_menu(opciones, '25',df)


#def main():    
print ('saliendo ')
if __name__ == '__main__':
            # menu_principal(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        menu_principal(df)
# return "terminada la ejecucion"
        



    

