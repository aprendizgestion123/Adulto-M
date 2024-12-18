import os
import shutil
import subprocess
import zipfile
import time as t
from os.path import join, exists
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_chrome_version():
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True
        ).decode('utf-8')
        for line in output.split('\n'):
            if 'version' in line:
                version = line.split()[-1]
                # Divide la versión en partes y omite la última parte
                version_parts = version.split('.')
                main_version = '.'.join(version_parts[:-1])
                return main_version
    except subprocess.CalledProcessError as e:
        print(f"Error al obtener la versión de Chrome: {e}")
    return None

#version = get_chrome_version()
#print(version)

def download_chromedriver(base_path, version):
    
    # Ruta de la carpeta donde se guardará el driver
    #base_path = r"C:\Users\aprendiz.serviciosti\Documents\Desarrollos\AnulacionColillas"
    driver_path = join(base_path, "driver")
    chromedriver_path = join(driver_path, "chromedriver-win64", "chromedriver.exe")
    url_driver = "https://googlechromelabs.github.io/chrome-for-testing/"

    # Validar si la versión del driver es la misma que la actual
    with open(join(driver_path, "link_driver.txt"), "r") as file:
        link = file.read()
        link = link.split("/")[4]
        # Extraer 125.0.6422 de la cadena 125.0.6422.1
        link = link.split(".")[0:3]
        link = '.'.join(link)
        
        print(f"Versión del driver: {link}")
        print(f"Versión actual del driver: {version}")
        
        if link == version:
            print("La versión del driver es la misma que la actual.")
            return
        else:
            print("La versión del driver es diferente a la actual.")

    if not exists(chromedriver_path):
        print(f"El archivo ChromeDriver no se encuentra en la ruta especificada: {chromedriver_path}")
        return

    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": driver_path}
    chrome_options.add_experimental_option("prefs", prefs)
    
    service = Service(executable_path=chromedriver_path)

    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get(url_driver)
        t.sleep(3)
        
        stable_link = driver.find_element(By.XPATH, "//a[@href='#stable']")
        stable_link.click()
        t.sleep(3)
        
        wait = WebDriverWait(driver, 5)
        table_element = wait.until(EC.presence_of_element_located((By.XPATH, "//section[@id='stable']//table")))
        table_text = table_element.text

        download_link = None
        for line in table_text.split('\n'):
            if 'chromedriver-win64' in line:
                download_link = line.split()[2]
                break
        
        if download_link:
            
            print(f"Enlace de descarga: {download_link}")
            
            driver.get(download_link)
            print("Descarga iniciada.")
            
            # Guardar el enlace de descarga en un archivo de texto
            with open(join(driver_path, "link_driver.txt"), "w") as file:
                file.write(download_link)
        else:
            print("No se encontró el enlace de descarga para chromedriver-win64.")
        
        t.sleep(10)
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()

    zip_path = join(driver_path, "chromedriver-win64.zip")
    extract_path = join(driver_path, "chromedriver-win64")

    try:
        if exists(zip_path):
            print(f"Archivo zip encontrado: {zip_path}")
            if exists(extract_path):
                print(f"Eliminando la carpeta existente: {extract_path}")
                shutil.rmtree(extract_path)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(driver_path)
            print("Archivo descomprimido.")
            os.remove(zip_path)
        else:
            print(f"No se encontró el archivo zip en la ruta: {zip_path}")
    except Exception as e:
        print(f"Error al descomprimir el archivo: {e}")

# Llamar a la función para probarla
#get_chromedriver_version()