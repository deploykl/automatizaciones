import pandas as pd
import glob
import os

# Buscar archivos .xls en ipress/, excluyendo "ipress.xlsx"
archivos_excel = [
    f for f in glob.glob("ipress/*.xls")  # Cambiado a .xls
    if not os.path.basename(f) == "ipress.xlsx"
]

if not archivos_excel:
    raise FileNotFoundError("No se encontró ningún archivo .xls válido en 'ipress/' (se excluyó ipress.xlsx)")

archivo = archivos_excel[0]  # Tomar el primer archivo .xls válido

# Leer el archivo .xls (usando xlrd)c
try:
    df = pd.read_excel(archivo, engine='xlrd')  # Cambiado a xlrd
except ModuleNotFoundError:
    raise ModuleNotFoundError(
        "¡Falta la librería 'xlrd'! Ejecuta: pip install xlrd>=2.0.1"
    )
except Exception as e:
    raise ValueError(f"Error al leer el archivo: {e}. ¿Está corrupto o protegido?")

# Filtrar: MINSA y Provincia != LIMA
filtro_minsa = df["Institución"] == "MINSA"
filtro_no_lima = df["Provincia"] != "LIMA"
df_filtrado = df[filtro_minsa & filtro_no_lima].copy()

# Renombrar "MINSA" a "GOBIERNO REGIONAL"
df_filtrado["Institución"] = "GOBIERNO REGIONAL"

# Combinar con el resto de datos
df_final = pd.concat([df[~(filtro_minsa & filtro_no_lima)], df_filtrado])

# Guardar como "ipress.xlsx" (en la misma carpeta 'ipress/') con hoja llamada "ipress"
nombre_modificado = "ipress/ipress.xlsx"
try:
    df_final.to_excel(nombre_modificado, index=False, engine='openpyxl', sheet_name='ipress')
    print(f"¡Archivo procesado y guardado como: {nombre_modificado} con hoja 'ipress'!")
except Exception as e:
    raise ValueError(f"No se pudo guardar el archivo: {e}")