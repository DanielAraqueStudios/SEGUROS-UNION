import pandas as pd

def leer_excel(ruta_archivo, hoja=0):
    """
    Lee una hoja de un archivo XLSB y retorna un DataFrame.
    """
    try:
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, engine='pyxlsb')
        return df
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        return None

if __name__ == "__main__":
    ruta = "InformeProduccion.xlsb"
    df = leer_excel(ruta)
    if df is not None:
        print(df.head())
    else:
        print("No se pudo leer el archivo.")
