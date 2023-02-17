from openpyxl import load_workbook
from datetime import datetime

rut=r"hojaDatos.xlsx"

def agregar_producto(rut:int,datos:dict):
    Archivo_Exccel=load_workbook(rut)
    Hoja_datos=Archivo_Exccel["productos"]
    Hoja_datos=Hoja_datos["A2":"E"+str(Hoja_datos.max_row+1)]
    hoja=Archivo_Exccel.active

    nombre=2
    categoria=3
    precio=4
    cantidad=5
    for i in Hoja_datos:

        if not(isinstance(i[0].value, int)):
            identificador=i[0].row
            hoja.cell(row=identificador, column=1).value=identificador-1
            hoja.cell(row=identificador, column=nombre).value=datos["Nombre"]
            hoja.cell(row=identificador, column=categoria).value=datos["Categoria"]
            hoja.cell(row=identificador, column=precio).value=datos["Precio"]
            hoja.cell(row=identificador, column=cantidad).value=datos["Cantidad"] 
            break
    Archivo_Exccel.save(rut)
    return