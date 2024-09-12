#Como dominar el mundo
 
#paqutes
import openpyxl
import pyautogui
import keyboard
import pyperclip
import time
124002044
 
 
# Ruta de tu archivo Excel
archivo_excel = 'BDEI.xlsx'
 
 
# Cargar el libro de trabajo de Excel
libro_trabajo = openpyxl.load_workbook(archivo_excel)
hoja = libro_trabajo.active
total_filas = 1188
 
for fila in range(1, total_filas + 1):
    valor_celda = hoja.cell(row=fila, column=1).value
 
    # Copia el valor de la celda al portapapeles
    pyperclip.copy(str(valor_celda))
 
    # Espera a que se presione Enter
    keyboard.wait('enter')
 
libro_trabajo.close()