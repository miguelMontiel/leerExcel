import openpyxl
import tkinter.filedialog as Tkinter
import os.path

prueba = {
    'TOTransferenciasStub': ['ktoli1', 'ktolv1'],
    'KtotranuStub':         ['ktoli2', 'ktolv2']
}
def ConseguirEntradaSalida():


def InsertarEntradaSalida():
    # Seleccionar el Excel a usar
    archivo = Tkinter.askopenfilename(initialdir = "C:/Users/IBM_ADMIN/PycharmProjects/laYesseniaDemo/", title = "Archivo Excel")

    # Abrirlo usando la libreria openpyxl, en este caso 'keep_vba' es para mantener los codigos visual basic que se han utilizado
    workbook = openpyxl.load_workbook(os.path.basename(str(archivo)), keep_vba = True)
    sheet = workbook.active

    for row in sheet.iter_cols(min_row = 1, min_col = 1, max_row = 2000, max_col = 11):
        for cell in row:
            if cell.value in prueba.keys():
                print(" ", cell.value)
                for valores in prueba[cell.value]:
                    print("     ", valores)
                    if valores[:5] == 'ktoli':
                        entrada = "H" + str(cell.row)
                        sheet[entrada] = "Hola"
                        print(entrada)
                    elif valores[:5] == 'ktolv':
                        salida = "I" + str(cell.row)
                        print(salida)
                        sheet[salida] = "Mundo"


#InsertarEntradaSalida()
workbook.save("lonotarteamiymetiraspasando.xlsm")
