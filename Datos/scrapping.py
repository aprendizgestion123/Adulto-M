import time
import os
from datetime import datetime
from os.path import exists
import time as t
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Emails.DescargaCorreo import DescargaCorreo
from Datos.pasos import Pasos

# Obtener las fechas del correo
class Scrapping(Pasos):

    def __init__(self):
        super().__init__()
        
        self.Adicionar_Usuarios = '/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/a/span'
        self.Adicionar_pagos = '/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/a/span'
    
    def main(self):

        self.crear_carpeta()

        descargador = DescargaCorreo()
        sw,fecha_inicio , fecha_final = descargador.obtener_fechas_desde_correo()
        print(fecha_inicio)
        print(fecha_final)       
        if sw:      
            self.scrapping(fecha_inicio,fecha_final)
            self.CopiarArchivo()

    def __esperar_descarga(self, nombre_archivo, tiempo_espera=1000):
    
        ruta_archivo = os.path.join(self.download_dir, nombre_archivo)
        tiempo_transcurrido = 0
        while tiempo_transcurrido < tiempo_espera:
            if  exists(ruta_archivo):
                # Verifica si el archivo está siendo utilizado (descarga aún en curso)
                try:
                    with open(ruta_archivo, 'rb'):
                        return True
                except PermissionError:
                    # Si está en uso, espera un poco más
                    time.sleep(20)
                    tiempo_transcurrido += 20
            else:
                # Si no existe, espera un poco y vuelve a verificar
                time.sleep(20)
                tiempo_transcurrido += 20
    
        # Si no se encuentra el archivo después del tiempo máximo, retorna False
        return False

    def __cambiar_a_iframe(self,driver, xpath):
        # Función para navegar entre los iframes
        try:
            driver.switch_to.frame(driver.find_element(By.XPATH, xpath))
        except Exception as e:
            print(f"Error al cambiar de iframe: {e}")
    # Función para seleccionar una opción en un menú desplegable
    def __seleccionar_opcion(self , driver, xpath, opcion_index):
        try:
            driver.find_element(By.XPATH, xpath).click()
            t.sleep(1)
            driver.find_element(By.XPATH, f"{xpath}/option[{opcion_index}]").click()
        except Exception as e:
            print(f"Error al seleccionar opción: {e}")
    # Función para hacer clic en un elemento
    def __hacer_click(self ,driver, xpath):
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
        except Exception as e:
            print(f"Error al hacer clic en el elemento: {e}")

    def scrapping(self,fecha_inicio,fecha_final):
        # Inicializar el navegador
        driver=super().iniciar_pagina()

        try:
            # Esperar e ingresar al iframe principal
            driver.implicitly_wait(30)
            self.__cambiar_a_iframe(driver, "/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[1]/iframe")

            # Navegar por los menús
            self.__hacer_click(driver, '//*[@id="masterdiv"]/div[19]')  # Desplegar "Pagos subsidios"
            self.__hacer_click(driver, '//*[@id="sub20"]/div')  # Ingresar en "Reporte pagos"

            # Cambiar al segundo iframe
            driver.switch_to.default_content()
            driver.implicitly_wait(30)
            self.__cambiar_a_iframe(driver, "/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/iframe")

            # Rellenar fecha desde 
            driver.find_element(By.XPATH, "/html/body/form/center/table/tbody/tr[6]/td/table/tbody/tr/td[1]/table/tbody/tr/td[3]/input").clear()
            driver.find_element(By.XPATH, "/html/body/form/center/table/tbody/tr[6]/td/table/tbody/tr/td[1]/table/tbody/tr/td[3]/input").send_keys(fecha_inicio)

            #Rellenar fecha hasta 
            driver.find_element(By.XPATH, "/html/body/form/center/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/input").clear()
            driver.find_element(By.XPATH, "/html/body/form/center/table/tbody/tr[6]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/input").send_keys(fecha_final)

            # Click en "Todos"
            self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[7]/td/table/tbody/tr/td/table/tbody/tr/td[3]/nobr[3]/input")

            cont1 = 12
            while cont1 > 5:
                self.__seleccionar_opcion(driver, f"/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", cont1)
                self.__hacer_click(driver, self.Adicionar_pagos)
                cont1 -= 1

            # Seleccionar otras opciones y generar el reporte
            opciones = [
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 1),  # Sitio de Venta
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 2),  # Oficina
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 8),  # Producto
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 2),  # Empresa
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 4),  # Reclama titular?
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 1),  # Canal
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 3),  # Cajero/Vendedor
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 2),  # Cédula Cajero/Vendedor
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 2),  # Zona
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 4),  # Ciudad
                ("/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 1)   # Fecha
            ]

            for xpath, opcion in opciones:
                self.__seleccionar_opcion(driver, xpath, opcion)
                self.__hacer_click(driver, self.Adicionar_pagos)

            # Seleccionar "Adulto Mayor" y generar el reporte
            self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div/select")
            self.__seleccionar_opcion(driver, "/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/div/select", 9)

            self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[11]/td/table/tbody/tr/td/table/tbody/tr/td[3]/nobr[2]/input")
            self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[12]/td/table/tbody/tr/td/table/tbody/tr/td[3]/nobr[1]/input")
            self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[13]/td/table/tbody/tr/td/table/tbody/tr/td[2]/a/span")
            des = self.__esperar_descarga('reportePagos.xls')
            if des:
                print("primer archivo descargado correctamente")
            
                # SEGUNDO SCRAPPING
                driver.get(self.SIGA_USUARIOS)

                # Realizar las selecciones para el segundo scraping
                self.__seleccionar_opcion(driver, "/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 10)
                self.__hacer_click(driver, self.Adicionar_Usuarios)

                self.__seleccionar_opcion(driver, "/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 4)
                self.__hacer_click(driver, self.Adicionar_Usuarios)

                self.__seleccionar_opcion(driver, "/html/body/form/center/table/tbody/tr[8]/td/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[2]/div/select", 11)
                self.__hacer_click(driver, self.Adicionar_Usuarios)

                self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[9]/td/table/tbody/tr/td/table/tbody/tr/td[3]/nobr[2]/input")
                self.__hacer_click(driver, "/html/body/form/center/table/tbody/tr[10]/td/table/tbody/tr/td/table/tbody/tr/td[3]/nobr[1]/input")

                # Click en "Generar reporte"
                    
                driver.find_element(By.ID, "btnGenerar").click()
                des = self.__esperar_descarga('reporteUsuarios.xls')
                if des:
                    print("segundo archivo descargado correctamente")
                else: 
                    print("no se descargo el segundo archivo")
            else: 
                print("no se descargo el primer archivo")
        finally:
                # Cerrar el navegador
                driver.quit()

    def CopiarArchivo(self):
        archivo = f'Adulto_mayor.xlsx'
        path_origen = os.path.join(self.path_,'Estandar','Adulto_mayor.xlsx')
        path_Destino = os.path.join(self.path_,'Resultado',archivo)

        try:
            shutil.copy(path_origen, path_Destino)
            print(f"Archivo copiado a {path_Destino}")
        except FileNotFoundError:
            print("Archivo no encontrado. Verifica la ruta.")
        except Exception as e:
            print(f"Ocurrió un error: {e}")
