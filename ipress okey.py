import pandas as pd
import glob
import os

# Buscar archivos .xls en ipress/, excluyendo "ipress.xlsx"
archivos_excel = [
    f for f in glob.glob("ipress/*.xls")
    if not os.path.basename(f) == "ipress.xlsx"
]

if not archivos_excel:
    raise FileNotFoundError("No se encontró ningún archivo .xls válido en 'ipress/' (se excluyó ipress.xlsx)")

archivo = archivos_excel[0]

# Leer el archivo .xls
try:
    df = pd.read_excel(archivo, engine='xlrd')
except ModuleNotFoundError:
    raise ModuleNotFoundError("¡Falta la librería 'xlrd'! Ejecuta: pip install xlrd>=2.0.1")
except Exception as e:
    raise ValueError(f"Error al leer el archivo: {e}")

# ========== NUEVO: AGREGAR COLUMNA CON DIFERENCIACIÓN DE LIMA ==========
def clasificar_lima(fila):
    """Clasifica entre Lima Metropolitana y Lima Región"""
    if fila["Departamento"] == "LIMA":
        # Si es Lima Metropolitana (Provincia = LIMA)
        if fila["Provincia"] == "LIMA":
            return "LIMA METROPOLITANA"
        else:
            # Si es Lima Región (Provincia diferente a LIMA, pero Departamento = LIMA)
            return "LIMA REGIÓN"
    else:
        # No es Lima, devolver el departamento original o lo que prefieras
        return fila["Departamento"]

# Aplicar la función para crear la nueva columna
df["REGION_DIF"] = df.apply(clasificar_lima, axis=1)

# Mostrar un resumen de la clasificación (opcional, para verificar)
print("\n=== RESUMEN DE CLASIFICACIÓN ===")
print(df["REGION_DIF"].value_counts())
print(f"\nTotal de registros: {len(df)}")
print(f"Lima Metropolitana: {len(df[df['REGION_DIF'] == 'LIMA METROPOLITANA'])}")
print(f"Lima Región: {len(df[df['REGION_DIF'] == 'LIMA REGIÓN'])}")

# ========== CONTINÚA CON TU LÓGICA ORIGINAL ==========
# Filtrar: MINSA y Provincia != LIMA
if "Institución" in df.columns and "Provincia" in df.columns:
    filtro_minsa = df["Institución"] == "MINSA"
    filtro_no_lima = df["Provincia"] != "LIMA"
    df_filtrado = df[filtro_minsa & filtro_no_lima].copy()
    
    # Renombrar "MINSA" a "GOBIERNO REGIONAL"
    df_filtrado["Institución"] = "GOBIERNO REGIONAL"
    
    # Combinar con el resto de datos
    df_final = pd.concat([df[~(filtro_minsa & filtro_no_lima)], df_filtrado])
else:
    print("\n⚠️ Advertenta: No se encontraron las columnas 'Institución' o 'Provincia'")
    print("   Se omite el filtrado y renombrado de MINSA")
    df_final = df.copy()

# Guardar como "ipress.xlsx"
nombre_modificado = "ipress/ipress.xlsx"
try:
    df_final.to_excel(nombre_modificado, index=False, engine='openpyxl', sheet_name='ipress')
    print(f"\n✓ Archivo procesado y guardado como: {nombre_modificado}")
    print(f"✓ Se agregó la columna 'REGION_DIF' con diferenciación Lima Metropolitana / Lima Región")
except Exception as e:
    raise ValueError(f"No se pudo guardar el archivo: {e}")