import os
import shutil
import subprocess
import zipfile
import requests
from os.path import join, exists

def get_chrome_version():
    try:
        output = subprocess.check_output(
            r'C:\Windows\System32\reg.exe query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True
        ).decode('utf-8')
        for line in output.split('\n'):
            if 'version' in line:
                version = line.split()[-1]
                return version
    except subprocess.CalledProcessError as e:
        print(f"Error al obtener la versión de Chrome: {e}")
    return None

def get_current_driver_version(driver_path, default_version):
    try:
        os.makedirs(driver_path, exist_ok=True)  # Asegura que el directorio existe
        with open(join(driver_path, "link_driver.txt"), "r") as file:
            link = file.read()
            link = link.split("/")[4]
            return link
    except FileNotFoundError:
        with open(join(driver_path, "link_driver.txt"), "w") as file:
            file.write(default_version)
        return default_version

def download_chromedriver(base_path, version):
    driver_path = join(base_path, "driver")
    #chromedriver_path = join(driver_path, "chromedriver-win64", "chromedriver.exe")
    default_url = "https://storage.googleapis.com/chrome-for-testing-public/125.0.6422.10/win64/chromedriver-win64.zip"
    url_driver = f"https://storage.googleapis.com/chrome-for-testing-public/{version}/win64/chromedriver-win64.zip"

    # Validar si la versión del driver es la misma que la actual
    current_version = get_current_driver_version(driver_path, default_url)
    
    print(f"Versión del driver: {current_version}")
    print(f"Versión de Chrome: {version}")
    
    if current_version == version:
        print("La versión del driver es la misma que la actual.")
        return
    else:
        print("La versión del driver es diferente a la actual.")

    # Descargar el archivo ZIP del driver
    if not exists(driver_path):
        os.makedirs(driver_path)

    response = requests.get(url_driver)
    zip_path = join(driver_path, "chromedriver-win64.zip")
    with open(zip_path, "wb") as file:
        file.write(response.content)
    print("Descarga completa.")
        
    # Extraer el archivo ZIP
    try:
        # Validar si la carpeta del driver existe
        extract_path = join(driver_path, "chromedriver-win64")
        if exists(extract_path):
            shutil.rmtree(extract_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(driver_path)
        print("Archivo descomprimido.")
        os.remove(zip_path)

        # Guardar el enlace de descarga en un archivo de texto
        with open(join(driver_path, "link_driver.txt"), "w") as file:
            file.write(url_driver)
    except Exception as e:
        print(f"Error al extraer el archivo ZIP: {str(e)}")

# Llamar a la función para probarla
#base_path = r"C:\Users\aprendiz.serviciosti\Documents\Desarrollos\Anulacion_Colillas"
#version = get_chrome_version()
'''if version:
    download_chromedriver(base_path, version)
else:
    print("No se pudo obtener la versión de Chrome.")'''