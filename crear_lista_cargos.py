import csv
import os

path_a_utilizar = os.path.dirname(os.path.abspath(__file__))
nombre_archivo = 'cargos_cometidos.csv'
nombre_archivo_salida = 'cargos_unicos_con_categoria.csv'

file_path = os.path.join(path_a_utilizar, "clausulas_csv", nombre_archivo)
file_path_salida = os.path.join(path_a_utilizar, "clausulas_csv", nombre_archivo_salida)

# Imprimir la ruta del archivo para verificar
print(f"Ruta del archivo: {file_path}")

def extraer_cargos(file_path):
    cargos = set()  # Usamos un conjunto para evitar duplicados
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter='|')
        for row in reader:
            if len(row) > 1:  # Asegurarse de que la fila tenga al menos dos columnas
                # Crear una tupla con los valores de la primera y segunda columna
                cargo = (row[0].strip(), row[2].strip())
                cargos.add(cargo)  # Agregar la tupla al conjunto
    return list(cargos)  # Convertir el conjunto a una lista

# Extraer los cargos
lista_cargos_unicos = extraer_cargos(file_path)

# Imprimir la lista de cargos
print(lista_cargos_unicos)

# Guardar los cargos en una columna de un archivo CSV
with open(file_path_salida, mode='w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile, delimiter='|')
    writer.writerow(['Servicio de salud', 'Cargo'])  # Escribir el encabezado de las columnas
    for cargo in lista_cargos_unicos:
        writer.writerow(cargo)  # Escribir cada tupla en una nueva fila

# Imprimir la ruta del archivo de salida para verificar
print(f"Archivo de salida: {file_path_salida}")