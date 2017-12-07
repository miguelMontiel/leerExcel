import openpyxl
import tkinter.filedialog as Tkinter
import os.path
import codecs as loscodecs

def ConseguirEntradaSalida():
    archivoTXT = Tkinter.askopenfilename(initialdir = "C:/Users/IBM_ADMIN/Desktop/archivosLaYessenia/",
                                         title = "Archivos TXT")
    archivo = loscodecs.open(archivoTXT, "r")
    dicProgramasStubs = {}

    key = ""
    value = ""

    for linea in archivo:
        lineaStrip = linea.strip()

        if lineaStrip[:6] == "Stub: ":
            value = lineaStrip[6:]

        elif lineaStrip[:10] == "Programa: ":
            key = (lineaStrip[10:])

        dicProgramasStubs[key] = value
    print(dicProgramasStubs)
    archivo.close()

    InsertarEntradaSalida(dicProgramasStubs)

def InsertarEntradaSalida(programas):
    # Seleccionar el Excel a usar
    archivo = Tkinter.askopenfilename(initialdir = "C:/Users/IBM_ADMIN/PycharmProjects/laYesseniaDemo/", title = "Archivo Excel")

    # Abrirlo usando la libreria openpyxl, en este caso 'keep_vba' es para mantener los codigos visual basic que se han utilizado
    workbook = openpyxl.load_workbook(os.path.basename(str(archivo)), keep_vba = True)
    sheet = workbook["BD_Tesoreria"]

    for row in sheet.iter_cols(min_row = 1, min_col = 1, max_row = 4125, max_col = 11):
        for cell in row:
            programa = str(cell.value)
            if programa[:8] in programas.keys():
                programasKeys = programas[programa[:8]]
                stubs = "G" + str(cell.row)
                print(stubs, programasKeys)
                sheet[stubs] = programasKeys
                '''
                for valores in programas[cell.value]:
                    print("     ", valores)
                    if valores[:5] == 'ktoli':
                        entrada = "H" + str(cell.row)
                        sheet[entrada] = "Hola"
                        print(entrada)
                    elif valores[:5] == 'ktolv':
                        salida = "I" + str(cell.row)
                        print(salida)
                        sheet[salida] = "Mundo"
                        
                '''

    workbook.save("lonotarteamiymetiraspasando.xlsm")

ConseguirEntradaSalida()
#InsertarEntradaSalida()
