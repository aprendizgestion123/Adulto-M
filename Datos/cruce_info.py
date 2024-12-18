import win32com.client as win32
import pandas as pd
import os
from openpyxl import load_workbook
from Datos.pasos import Pasos

class GuardarDatos(Pasos):

    def __init__(self):
        #Inicializa una instancia de Excel para trabajar con archivos de Excel.

        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False  # Cambiar a True para ver el proceso en Excel.
        self.path_ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path_ResultadoArchivo= os.path.join(self.path_,'Resultado', 'Adulto_mayor.xlsx')

    def cargar_archivos(self):
        #Carga los archivos necesarios para el procesamiento.
        columnas_Pagos=['Fecha','Ciudad','Zona','Cédula Cajero/Vendedor','Cajero/Vendedor',
            'Canal', 'Reclama titular?', 'Empresa', 'Producto','Cod. Oficina', 'Oficina',
            'Cod. Sitio ','Sitio de venta','Tipo Doc. Titular', 'Identificación Titular',  'Titular',
            'Tipo Doc. Autorizado', 'Identificación Autorizado', 'Autorizado','Periodo', 'Valor Reportado', 
            'Valor Pagado', 'Valor Redondeo', 'Grupo Pago','Tipo Pago','Tipo Subsidio', 'Titular Comfama', 
            'Titular Comfama']
        self.reporte_pagos = pd.read_excel('reportes/reportePagos.xlsx',usecols=columnas_Pagos)
        print(self.reporte_pagos)

        columnas_Usuarios=['Identificacion','Cargo','Estado']
        self.reporte_usuarios = pd.read_excel('reportes/reporteUsuarios.xlsx',usecols=columnas_Usuarios)
        print(self.reporte_usuarios)

         #Procesa los datos cargados, separa fecha y hora, y selecciona las columnas de interés.
        # Separar la columna FechaHora en Fecha y Hora

        self.reporte_pagos.columns = self.reporte_pagos.columns.str.strip()
        self.reporte_pagos[['Fecha', 'Hora']] = self.reporte_pagos['Fecha'].str.split(' ', expand=True)
        print("Reporte Pagos después de separar Fecha y Hora:",self.reporte_pagos)
        print(self.reporte_pagos[['Fecha', 'Hora']].head())

    def procesar_datos(self):
        #Procesa los datos cargados, separa fecha y hora, y selecciona las columnas de interés.
        # Separar la columna FechaHora en Fecha y Hora

        self.reporte_pagos.columns = self.reporte_pagos.columns.str.strip()
        self.reporte_pagos[['Fecha', 'Hora']] = self.reporte_pagos['Fecha'].str.split(' ', expand=True)
        # Mapeo de columnas según tu especificación

        self.mapeo_columnas = {
            'A': 'Fecha',  'B': 'Hora',  'C': 'Ciudad',  'D': 'Zona','E': 'Cédula Cajero/Vendedor',  'F': 'Cajero/Vendedor',
            'H': 'Canal','I': 'Reclama titular?','J': 'Empresa',  'K': 'Producto','L': 'Cod. Oficina',  'M': 'Oficina',
            'N': 'Cod. Sitio','O': 'Sitio de venta','P': 'Tipo Doc. Titular','Q': 'Identificación Titular',  'R': 'Titular',
            'S': 'Tipo Doc. Autorizado','T': 'Identificación Autorizado','U': 'Autorizado','X': 'Periodo','Y': 'Valor Reportado', 
            'Z': 'Valor Pagado','AA': 'Valor Redondeo','AB': 'Grupo Pago','AC': 'Tipo Pago','AD': 'Tipo Subsidio','AE': 'Titular Comfama', 
            'AF': 'Titular Comfama'}
        # Mapeo de columnas de reporte_usuarios
        self.mapeo_columnas_usuarios = {
            'G':'Cargo','AG':'Estado',}

    def convertir_columna_a_indice(self, columna):
        #Convierte una columna en formato letra (A, B, AA) a un índice numérico.
        resultado = 0
        for char in columna:
            resultado = resultado * 26 + (ord(char.upper()) - ord('A') + 1)
        return resultado
    
    def procesar_excel(self):
        print(self.path_ResultadoArchivo)
        df = pd.read_excel(self.path_ResultadoArchivo)
        
        titulos_df = pd.DataFrame({'Titulos': df.columns})
        print(type(titulos_df))
        print(titulos_df)

        # Seleccionar solo las columnas que existen en df_destino 
        columnas_comunes = [col for col in titulos_df['Titulos']if col in self.reporte_pagos.columns]
        df=df[columnas_comunes] 
        #titulos_df = self.reporte_pagos[columnas_comunes]        
        print(df)

    def cargar_archivos1(self):
        #Carga los archivos necesarios para el procesamiento.
        columnas_Pagos=['Fecha','Ciudad','Zona','Cédula Cajero/Vendedor','Cajero/Vendedor',
            'Canal', 'Reclama titular?', 'Empresa', 'Producto','Cod. Oficina', 'Oficina',
            'Cod. Sitio ','Sitio de venta','Tipo Doc. Titular', 'Identificación Titular',  'Titular',
            'Tipo Doc. Autorizado', 'Identificación Autorizado', 'Autorizado','Periodo', 'Valor Reportado', 
            'Valor Pagado', 'Valor Redondeo', 'Grupo Pago','Tipo Pago','Tipo Subsidio', 'Titular Comfama', 
            'Titular Comfama']
        self.reporte_pagos = None
        self.reporte_pagos = load_workbook('reportes/reportePagos.xlsx')
        print(self.reporte_pagos)

        columnas_Usuarios=['Identificacion','Cargo','Estado']
        self.reporte_usuarios = None
        self.reporte_usuarios = load_workbook('reportes/reporteUsuarios.xlsx')
        print(self.reporte_usuarios)

         #Procesa los datos cargados, separa fecha y hora, y selecciona las columnas de interés.
        # Separar la columna FechaHora en Fecha y Hora

        """ self.reporte_pagos.columns = self.reporte_pagos.columns.str.strip()
        self.reporte_pagos[['Fecha', 'Hora']] = self.reporte_pagos['Fecha'].str.split(' ', expand=True) """

    def guardar_en_excel(self, hoja):
        #Escribe los datos procesados en el archivo Excel en la hoja especificada.
        wb = self.excel.Workbooks.Open(self.path_ResultadoArchivo)
        try:
            ws = wb.Worksheets(hoja)
        except Exception as e:
            print(f"Error al acceder a la hoja '{hoja}': {e}")
            wb.Close(False)
            self.excel.Quit()
            return

        # Obtener la última fila con datos
        last_row = ws.Cells(ws.Rows.Count, 1).End(win32.constants.xlUp).Row + 1

         # Escribir datos según el mapeo
        print("Escribiendo datos en ADULTO MAYOR...")
        for i, row in self.reporte_pagos.iterrows():
            for col_letra, col_nombre in self.mapeo_columnas.items():
                col_idx = self.convertir_columna_a_indice(col_letra)
                valor = row[col_nombre] if not pd.isnull(row[col_nombre]) else ""
                ws.Cells(last_row + i, col_idx).Value = valor

            # Duplicar columna M (Oficina) en columna W
            col_idx_m = self.convertir_columna_a_indice('M')
            col_idx_w = self.convertir_columna_a_indice('W')
            ws.Cells(last_row + i, col_idx_w).Value = ws.Cells(last_row + i, col_idx_m).Value
            
        # Escribir datos de reporte_usuarios
        for i, row in self.reporte_usuarios.iterrows():
            for col_letra, col_nombre in self.mapeo_columnas_usuarios.items():
                col_idx = self.convertir_columna_a_indice(col_letra)
                valor = row[col_nombre] if not pd.isnull(row[col_nombre]) else ""
                ws.Cells(last_row + i, col_idx).Value = valor


        # Guardar y cerrar el archivo
        wb.Save()
        wb.Close()

    def cerrar_excel(self):
        """Cierra la instancia de Excel."""
        self.excel.Quit()