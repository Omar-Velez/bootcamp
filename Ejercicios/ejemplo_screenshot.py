# pip install pillow openpyxl

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import ImageGrab
import os
from datetime import datetime

# Nombre del archivo que se va a crear
excel_path = "evidencia_ejecucion.xlsx"
screenshot_temp = "captura_temp.png"

# 1. Tomar captura de pantalla
print("Tomando captura de pantalla...")
captura = ImageGrab.grab()           # Captura toda la pantalla
captura.save(screenshot_temp)

# 2. Crear libro de Excel
wb = Workbook()

# Hoja "datos" (puedes dejarla vacía o con alguna información básica)
ws_datos = wb.active
ws_datos.title = "datos"
ws_datos['A1'] = "Ejecución del script"
ws_datos['A2'] = f"Fecha y hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
ws_datos['A3'] = "Este sheet puede usarse para poner resultados en el futuro"

# Hoja "evidencia"
ws_ev = wb.create_sheet("evidencia")
ws_ev['A1'] = "Captura de pantalla del escritorio al momento de ejecución"
ws_ev['A2'] = f"Fecha y hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

# Insertar la imagen (ajusta la celda de inicio si quieres)
img = Image(screenshot_temp)
ws_ev.add_image(img, "A4")   # A partir de A4 para dejar espacio al texto

# 3. Guardar el archivo
wb.save(excel_path)

# 4. Eliminar archivo temporal
try:
    os.remove(screenshot_temp)
except:
    pass

print(f"Archivo creado exitosamente: {excel_path}")
print("Hoja 'datos' → información básica")
print("Hoja 'evidencia' → captura de pantalla insertada")
