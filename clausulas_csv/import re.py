import re

# Leer el contenido del archivo contratos.csv
with open('contratos.csv', 'r', encoding='utf-8') as file:
    content = file.read()

# Buscar todos los nombres dentro de corchetes "[]"
nombres = re.findall(r'\[([^\]]+)\]', content)

# Imprimir los nombres encontrados
for nombre in nombres:
    print(nombre)