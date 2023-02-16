from openpyxl import load_workbook
from datetime import datetime

rut=r'hojaDatos.xlsx'
def agregar (ruta:int, datos: dict):
    archivo_exccel =load_workbook(ruta)
    hoja_datos = archivo_exccel["tareas"]
    hoja_datos=hoja_datos["A2":"F"+str(hoja_datos.max_row+1)]
    hoja=archivo_exccel.active

    titulo=2
    descripcion=3
    estado=4 
    fecha_inicio=5
    fecha_finalizado=6
    for i in hoja_datos:

        if not(isinstance(i[0].value, int)):
            identificador=i[0].row
            hoja.cell(row=identificador,column=1).value=identificador-1
            hoja.cell(row=identificador,column=titulo).value=datos["titulo"]
            hoja.cell(row=identificador,column=descripcion).value=datos["descripcion"]
            hoja.cell(row=identificador,column=estado).value=datos["estado"]
            hoja.cell(row=identificador,column=fecha_inicio).value=datos["fecha_inicio"]
            hoja.cell(row=identificador,column=fecha_finalizado).value=datos["fecha_finalizacion"]
            break
    archivo_exccel.save(ruta) 
    return   

