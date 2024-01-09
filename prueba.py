
from fastapi import FastAPI
import xlwings as xw
import pandas as pd

#book = xw.Book(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')
#book = xw.Book(r'C:\\Temp\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')
book = xw.Book(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx')


pestana_LISTA_ASESORE_GENERACION_DOCUME = book.sheets['LISTA-ASESORE-GENERACION-DOCUME']
data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME = pestana_LISTA_ASESORE_GENERACION_DOCUME.range('A1').options(pd.DataFrame, expand='table').value

# data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME = pd.read_excel('LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V3.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', usecols = 'A:GO',header = 0)

# app = FastAPI()
# @app.get("/")
# def salir():    
#     # cifpar = cif
#     print ('saliendo ')
#     # menu_principal(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME,cifpar)
#     #if __name__ == '__main__':
#     #    procbuscarcif_api(cif)
#  #    # menu_principal(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME,cifpar)

#     return {"mensaje",1}

if __name__ == '__main__':

    print(data_oap_tab_LISTA_ASESORE_GENERACION_DOCUME)
    print(book)