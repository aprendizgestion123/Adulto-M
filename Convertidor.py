import pandas as pd
import os 

def convertir_archivos(carpeta_descargas):
    for archivo in os.listdir(carpeta_descargas):
        if archivo.endswith(".xls"):
            xls_path = os.path.join(carpeta_descargas, archivo)
            xlsx_path = os.path.join(carpeta_descargas, archivo.replace(".xls", ".xlsx"))
            
            try:
                # Leer el archivo .xls
                df = pd.read_excel(xls_path)
                
                # Guardar como .xlsx
                df.to_excel(xlsx_path, index=False)
                
                # Eliminar el archivo original
                os.remove(xls_path)
                print(f"Convertido: {archivo} -> {os.path.basename(xlsx_path)}")
            except Exception as e:
                print(f"Error al convertir {archivo}: {e}")

# Usar la función
carpeta_descargas = r"C:\Users\aprendiz.gestion\OneDrive - GANA S.A\Desktop\Adulto M\Reportes"
convertir_archivos(carpeta_descargas)
