from Datos.cruce_info import GuardarDatos

# Crear una instancia de la clase GuardarDatos
guardador = GuardarDatos()
#guardador.procesar_excel()
# Verificar o crear el archivo de destino

# Cargar, procesar y guardar los datos en el archivo Excel
guardador.cargar_archivos()
guardador.cargar_archivos1()
#guardador.procesar_excel()
""" guardador.procesar_datos()
guardador.guardar_en_excel('DATA')
guardador.cerrar_excel() """

print("Datos guardados correctamente.")
