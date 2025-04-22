import pandas as pd
import os

# Ruta donde se encuentran los archivos Excel
ruta_archivos = r'd:\SGD-DATA'
nombres_archivos = ['DGOS.xls', 'DIEM.xls', 'DIMON.xls']  # Asegúrate de cambiar la extensión a .xls

# Lista para almacenar DataFrames
dataframes = []

# Procesar cada archivo
for nombre_archivo in nombres_archivos:
    # Cargar el archivo Excel desde la ruta completa y omitir las primeras 4 filas
    df = pd.read_excel(os.path.join(ruta_archivos, nombre_archivo), skiprows=4)  # Lee todas las columnas

    # Eliminar la columna 'J' solo si está vacía
    if 'J' in df.columns and df['J'].isnull().all():
        df = df.drop(columns='J')

    # Eliminar columnas que tengan nombres "Unnamed"
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # Agregar el DataFrame limpio a la lista
    dataframes.append(df)

# Concatenar todos los DataFrames en uno solo
df_combined = pd.concat(dataframes, ignore_index=True)

# Exportar el DataFrame combinado a un único archivo CSV
df_combined.to_csv(os.path.join(ruta_archivos, 'todos_limpios.csv'), index=False)

print("Todos los archivos han sido procesados, combinados y exportados a todos_limpios.csv.")
