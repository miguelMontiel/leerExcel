from xlrd import open_workbook
import tkinter.filedialog as Tkinter
import os.path

# Seleccionar el Excel a usar
archivo = Tkinter.askopenfilename(initialdir = "C:/Users/mainUser/PycharmProjects/leerExcel/", title = "Archivo Excel")

#Abrirlo usando la libreria xlrd, para usarse en el programa.
book = open_workbook(os.path.basename(str(archivo)), on_demand = True)
sheet = book.sheet_by_index(0)

for cell in sheet.col(0):
    print(cell.value)