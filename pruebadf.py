import pandas as pd
import xlwings as xw

# workbook = xw.Book(r'c:\\users\\alfonso\\onedrive\\prueba\\entrada.xlsx')
# sheet1 = workbook.sheets['Hoja1'].used_range.value

sheet2 = xw.Book(r'c:\\users\\alfonso\\onedrive\\prueba\\entrada.xlsx').sheets("Hoja1")



book = xw.Book(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\LISTA-ASESORE-GENERACION-DOCUMENTOS-2022-BAK-ANTES-DE-ANEXO-20-Y-21-AÑO-2023-V32.xlsx')
###df = book.sheets['LISTA-ASESORE-GENERACION-DOCUME'].range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value
sheet1 = book.sheets['LISTA-ASESORE-GENERACION-DOCUME'].used_range.value

# me pongo la nota para recordar la linea siguiente a esta está comentada, definiia el dataframe a partir del sheet2 pero no deja,
# para que hay que hacerlo a partir de workbook.sheets..  pero no a partir de sheet2.range....
#####df = workbook.sheets['Hoja1'].range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value
#df = sheet2.range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value

# sheet1 = workbook.sheets['Hoja1'].used_range.value

df = pd.DataFrame(sheet1)

### esta solucion que pongo ahora fue para eliminar la fila 0 que contenia indices numerados
### lo que queriamos es que el indice fuera las columnas proyecto, fecha,,,,, etc
### tomado de https://stackoverflow.com/questions/31328861/replacing-header-with-top-row
new_header = df.iloc[0] #grab the first row for the header
df = df[1:] #take the data less the header row
df.columns = new_header #set the header row as the df header
print(df)
a = input(" pulse una tecla")
df.to_excel(r'C:\\Users\\Alfonso\\OneDrive\\PRUEBA\\salidar.xlsx', index=False, header=False)
a = input(" GENERAR SALIDA pulse una tecla")

# sheet2.range("A1").value = "nombre" eso solo es una prueba asignar valores directo a una celda y ver que actualiza excel ,
# en realidad no tiene mas que explicacoin en este ejemplo que probar como actualiza un excel sincronizado en onedrive
sheet2.range("A1").value = "nombre"

# Buscamos un valor concreto. Podria ser buscar por el CIF
df_tab_empresa_fila_buscada = df.loc[df['documento_solicitante'] == 'B42831594']

# Hay que chequear qu el dataframe no sea nulo para que no de error los siguientes pasos, validamos con una condición:
if not df_tab_empresa_fila_buscada.empty:
    
    # esto devuelve el número de fila
    indice = df.loc[df['documento_solicitante'] == 'B42831594'].index.item() 

    # esto es para debugear...
    a = input(" esto es el indice ===> "  + str(indice))
    print (" fila buscada " + str(df_tab_empresa_fila_buscada))

    # modificamos la fila encontrada con la condición de apellido y asignamos un valor para probar que actualiza el dataframe
    df.at[indice,'valor'] = 11
    # ahora asignamos el dataframe al excel
    sheet2['A1'].options(pd.DataFrame, header=1, index=False, expand='table').value = df
   
 
    print(df)



# cuando se coloca header quita los encabezamientos

# df_tab_LISTA_ASESORE_GENERACION_DOCUMEs.to_excel(r'c:\\users\\alfonso\\onedrive\\prueba\\salida.xlsx', sheet_name='LISTA-ASESORE-GENERACION-DOCUME', index=False )