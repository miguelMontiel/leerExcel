import openpyxl
import tkinter.filedialog as Tkinter
import os.path

# Seleccionar el Excel a usar
archivo = Tkinter.askopenfilename(initialdir = "C:/Users/mainUser/PycharmProjects/leerExcel/", title = "Archivo Excel")

# Abrirlo usando la libreria openpyxl, en este caso 'keep_vba' es para mantener los codigos visual basic que se han utilizado
workbook = openpyxl.load_workbook(os.path.basename(str(archivo)), keep_vba = True)
sheet = workbook.active

#http://zetcode.com/articles/openpyxl/
'''
cells = sheet['A1':'J6']
for c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 in cells:
    print("{0:8} {1:8}".format(c1.value, c2.value))
'''

prueba = {
    'TOTransferenciasStub': ['ktoli1', 'ktolv1'],
    'llave2': ['valor2', 'valor22']
}


for row in sheet.iter_cols(min_row=1, min_col=1, max_row=500, max_col=10):
    for cell in row:
        if cell.value in prueba.keys():
            print(" ", cell.value)
            for valores in prueba[cell.value]:
                print("     ", valores)
                if valores[:5] == 'ktoli':
                    print("         Entrada")
                    print("G", cell.row)
                elif valores[:5] == 'ktolv':
                    print("         Salida")
                    print("H", cell.row)