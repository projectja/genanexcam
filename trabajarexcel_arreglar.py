
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

# Primera pestaña
data_actividades_individuales = pd.read_excel('C:/Users/Alfonso/OneDrive/PRUEBA/registro-roto.xlsx', sheet_name='Actividades individuales_R', usecols = 'A:O',header = 0)
df_actividades_individuales=pd.DataFrame(data_actividades_individuales)

# Pestaña alta usuarios
data_alta_usuario = pd.read_excel('C:/Users/Alfonso/OneDrive/PRUEBA/registro-roto.xlsx', sheet_name='Alta_usuario_R', usecols = 'A:O',header = 0)
df_alta_usuario=pd.DataFrame(data_alta_usuario)

df_alta_empresa=""

# data_alta_empresa = pd.read_excel(book, sheet_name='Alta_empresa_R', usecols = 'A:GO',header = 0)
# data_alta_usuario = pd.read_excel(book, sheet_name='Alta_usuario_R', usecols = 'A:GO',header = 0)
# df_actividades_individuales = pd.DataFrame(data_actividades_individuales)
# df_alta_usuario = pd.DataFrame(data_alta_usuario)
# df_alta_empresa = pd.DataFrame(data_alta_empresa)
#df = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME

# data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME =pd.read_excel(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)
# data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME =pd.read_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)
# data_oap_tab_LISTA_CURSOS_ESPANA_EMPRENDE = pd.read_excel('crusos-espana-emprende.xlsx', sheet_name='HOJA1', usecols = 'A:R',header = 0, converters= { 'Marca temporal': pd.to_datetime, 'Fecha de nacimiento': pd.to_datetime})
# data_oap_tab_alta_usuario = pd.read_excel('10044_Huelva_Registro_Actividades_OAs_202112_v2.5.xlsx', sheet_name='Alta_Usuario', usecols = 'E:N',header = 3)
#data_oap_tab_actividades_individuales = pd.read_excel('10044_Huelva_Registro_Actividades_OAs_202112_v2.5.xlsx', sheet_name='Actividades individuales', usecols = 'E:T',header = 3)
print(df_actividades_individuales)
print(df_alta_usuario)



# ===========================================================================================================
# al exportar a excel elimina la primera fila con header = False
# funciona, genera un excel con el dataframe, sirve para ver que está bien formado :
# digamos que el dataframe lo exporta por si necesitamos exportar despues de maninuplasr datos
df_actividades_individuales.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\result.xlsx', index=False, header=False)
df_alta_usuario.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\result1.xlsx', index=False, header=False)
#=============================================================================================================



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


def crear_draft_anexo18(dirdatos_ae):
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





def generar_ficha_anexo18(dirdatos_ae):   
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

def generar_ficha_anexo18_2(dirdatos_ae):   
    # cif=input()
    #cif_a_buscar = cifglobal
    print( " denttro del ANEXO 18 2  ")
    print(dirdatos_ae)

    a = input ( "en generar anexo 18 2 ")
    #context = {'mi_nombre': 'Fran tarci'}
    #df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    #print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    #df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   

    valuestovar_anexo18_2(dirdatos_ae)
    

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

def generar_ficha_anexo19(dirdatos_ae):   
    # cif=input()
    cif_a_buscar = cifglobal

 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    # df_tab_empresa_fila_buscada
    valuestovar_anexo19(dirdatos_ae)
    # df_tab_empresa_fila_buscada['poblacion'] = 'Aljaraquerrrrr'
    # write again the excel file

    

#######################################################
## ------------------ GENERAR ANEXO 20 ----------------
#######################################################

def generar_ficha_anexo20(dirdatos_ae):   
    # cif=input()
    cif_a_buscar = cifglobal
    

 
    #context = {'mi_nombre': 'Fran tarci'}
    df_tab_LISTA_ASESORE_GENERACION_DOCUMEs = data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME   
    print (('en generar ficha '+ str(df_tab_LISTA_ASESORE_GENERACION_DOCUMEs)))

    
    # se localiza por NIF en el dataframe que corresponde al tab ALTA EMPRESAS    
    df_tab_empresa_fila_buscada = df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.loc[df_tab_LISTA_ASESORE_GENERACION_DOCUMEs['documento_solicitante'] == cif_a_buscar]   
    # df_tab_empresa_fila_buscada
    valuestovar_anexo20(dirdatos_ae)
    
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




def procbuscarcif(datos_ae, datos_au, datos_ai):
  
    
    b = input ( "Dentro de PRODCBUSCARCIF ")

    # i = 0
    # for row in datos_ai.itertuples():
       
    #     i = i + 1
 
    #     print(type(row))
    #     print(row)
    #     print('------')

    #     print(row[2])
       
    #     print('------\n')
    # a = input (" Pulse una tecla .... ")

    for index, row in datos_ai.iterrows():
        if row['Estado NIF'] == 'No encontrado':
             print("fila " + str(index))
             #df_tab_empresa_fila_buscada = datos_ai.loc[datos_ai['index'] == index]   
             df_tab_empresa_fila_buscada = datos_ai.loc[index]   
             print ( "fila buscada " + str(df_tab_empresa_fila_buscada))
             df_tab_empresa_fila_buscada.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\resultado.xlsx', index=True, header=True)

        


    # for column_name in datos_ai:
    #     print(type(column_name))
    #     print("NOMBRE DE COLUMNA " + column_name)
    #     print('------\n')
    # a = input (" Pulse una tecla .... ")
 



          
def mostrar_menu(opciones):
    print('Seleccione una opción:')
    for clave in opciones:
        # print(f' {opciones[clave][0]}')
         print(f' {clave}){opciones[clave][0]}')
    
  

    
    
def ejecutar_opcion(opcion, opciones,dirdatos_ae, dirdatos_au, dirdatos_ai):
    print("DENTRO DE la funcion EJECUTAR OPCION ")
    print(dirdatos_ae)
    
   
    if opcion == '1':
        a = input ( " la opcion es uno ")
        dirdatos_ae = procbuscarcif(dirdatos_ae, dirdatos_au, dirdatos_ai)
        print(" dir datos .................. "  )
        print(str(dirdatos_ae))
        a = input ( " Estos son los datos del diccionario .....")                                                  
    elif opcion == '2':
        fechadocumento()
    elif opcion == '4':
        listar_ddp(dirdatos_ae)       
    elif opcion == '6':
        convertir_pdf()  
    elif opcion == '14':
        crear_draft_ejecutar_gasto(dirdatos_ae)
    elif opcion == '7':
        generar_ficha_anexo18_2(dirdatos_ae)
    elif opcion == '8':
        crear_draft_anexo18(dirdatos_ae)
    elif opcion == '9':
        generar_ficha_anexo19(dirdatos_ae)
    elif opcion == '10':
        mover_ppi()
    elif opcion == '11' :
        crear_draft_anexo19()
    elif opcion == '14':
        crear_draft_ejecutar_gasto(dirdatos_ae)
    elif opcion == '18':
        generar_ficha_anexo20(dirdatos_ae)
    elif opcion == '20':
        creardraft_anexo20anexo21()
    elif opcion == '99':
        prueba_modificar()
    return dirdatos_ae


def leer_opcion(opciones):
    while (a := input('Opción: ')) not in opciones:
        print('Opción incorrecta, vuelva a intentarlo.')
    return a
    
def generar_menu(opciones, opcion_salida, df_ae, df_au, df_ai):
    global opcionmenu
    opcion = None    
    opcionmenu = opcion
    cifglobal = ""
    
    
    while opcion != opcion_salida:
        mostrar_menu(opciones)
        opcion = leer_opcion(opciones)
        a = input ( "despues de leer_opcion  ... " )
        # dirdatos_ae =  leer_datos(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        dirdatos_ae = ejecutar_opcion(opcion, opciones, df_ae, df_au, df_ai)
        dfdatos = dirdatos_ae
        # print (" DF DATOS VALE ")
        # print(dfdatos)
        a = input(" En la funcion GENERAR_MENU despues de ejectuar EJECUTAR_OPCION ")
        # dirdatos_ae =  leer_datos(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        #ejecutar_opcion(opcion, opciones,dirdatos_ae)
                
        print(" el cif en el while es " + cifglobal)
        print () # se imprime opcion en blanco para clarificar salir de pantalla

        
def menu_principal(df_ae,df_au,df_ai):
    print("menu - el alor del cif es " + cif)
    opciones = {
        '1': ('1.- Buscar registros huerfanos', procbuscarcif),
      

        # '40': ('40.- CIF oap ', buscar_cif_oap),  
        #'41': ('41.- Salir', main)
    }    
    generar_menu(opciones, '25',df_ae, df_au, df_ai)


#def main():    
print ('saliendo ')
if __name__ == '__main__':
            # menu_principal(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
        menu_principal(df_alta_empresa, df_alta_usuario, df_actividades_individuales)
# return "terminada la ejecucion"
        



    

