import pandas as pd
import sqlite3

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\orquestador")
#print("Directorio actual:", os.getcwd())

# Leer el archivo 'parametros.csv'
df = pd.read_csv('parametros.csv', header=None)

# Convertir el DataFrame en un diccionario params
params = dict(zip(df[0], df[1]))

# Leer el archivo con el nombre definido en base de params en un DataFrame con todas las columnas como texto
dfbase = pd.read_excel(params.get('base') + '.xlsx', dtype=str)

# Guardar el DataFrame en formato Parquet
#dfbase.to_parquet('archivo.parquet')

# Conectar a la base de datos SQLite (o crearla si no existe)
conn = sqlite3.connect('base.db')

# Guardar el DataFrame en la base de datos SQLite
dfbase.to_sql('tabla', conn, if_exists='replace', index=False)

# Cerrar la conexión
conn.close()
