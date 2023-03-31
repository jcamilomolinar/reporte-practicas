#Libreria para manipular el sistema operativo (No es posible correrlo a traves de Colab, debe ser en local)
import subprocess
import os
import zipfile

text = '''Este codigo fue creado con el fin de generar las calificaciones de las practicas de la 
asignatura simulacion de sistemas mas facilmente:

Asegurese de que el archivo de python ScriptNotas.py y el archivo ejecutarScript.bat 
se encuentren dentro de una carpeta, dentro de esta carpeta debe crear una carpeta llamada practicas 
donde colocara los reportes generados por la plataforma (en formato zip o rar)
la estructura que debe seguir es la siguiente

    nombreEjemplo (carpeta inicial, puede ponerle el nombre que desee)
        -> practicas    (carpeta donde deben ir los reportes)
        -> ejecutarScript.bat
        -> ScriptNotas.py

Codificado por: Juan Camilo Molina Roncancio - Est. Ingenieria de Sistemas e Informatica
'''

print(text)
input("Presione enter para generar el reporte. ")

# Verificar si los paquetes están instalados
try:
    import pandas as pd
    import openpyxl
    import rarfile
except ImportError:
    # Si alguno no está instalado, instalarlo con pip
    subprocess.check_call(['pip', 'install', 'pandas'])
    subprocess.check_call(['pip', 'install', 'openpyxl'])
    subprocess.check_call(['pip', 'install', 'rarfile'])
    import pandas as pd
    import openpyxl
    import rarfile

# Carpeta raiz
root_folder = "practicas"

#Determinar si los archivos estan en formato ZIP o RAR
ext = None
for folder_name in os.listdir(root_folder):
    if folder_name.split(".")[-1] == "zip": ext = "zip"
    else: ext = "rar"
    break

# Combinar la ruta del directorio y el nombre de la carpeta para obtener la ruta completa de la carpeta
folder_desc = "descomprimidos"
destiny = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_desc)

# Crear la carpeta si no existe
if not os.path.exists(destiny): os.mkdir(destiny)

#Iterar a través de las archivos comprimidos para descomprimirlos
if ext == "zip":
    for folder_name in os.listdir(root_folder):
        with zipfile.ZipFile("practicas/"+folder_name, 'r') as zip_ref:
            zip_ref.extractall(destiny)
elif ext == "rar":
    for folder_name in os.listdir(root_folder):
        archivo_rar = rarfile.RarFile("practicas/"+folder_name)
        archivo_rar.extractall(path=destiny)
        archivo_rar.close()

#Cantidad de practicas a calificar y Cantidad de practicas para alcanzar el 5, a partir de este valor se hace la regla de 3
cantidad_notaMax = int(input("\nIngrese la cantidad de practicas para que la nota del estudiante sea 5, a partir de este valor se hace la regla de 3: "))
names = []

#Iterar a través de las subcarpetas, obtenemos nombre e id de entrega
for folder_name in os.listdir(folder_desc):
    folder_path = os.path.join(folder_desc, folder_name)
    if os.path.isdir(folder_path):
        sep = folder_name.split("_")
        names.append(sep[0:2])

#Se eliminan los elementos con el mismo id de entrega
unique_names = []
for name in names: 
    if name not in unique_names: unique_names.append(name)

#Creamos un dataframe con nombre, cantidad de practicas enviadas y nota final, se exporta a un excel
df = pd.DataFrame(unique_names, columns=["Nombre", "practicas enviadas"])
result = df.groupby("Nombre").nunique()
result["Nota"] = result["practicas enviadas"].apply(lambda x: int(x)*5/cantidad_notaMax if int(x) < cantidad_notaMax else 5)
result.to_excel('notas_practicas1.xlsx', 'Notas', index=True)

#El excel se genera en la misma carpeta donde se este trabajando, para que el codigo funcione correctamente
print("El reporte de notas se ha generado correctamente!")