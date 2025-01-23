import pandas as pd
import sqlite3

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\orquestador")
#print("Directorio actual:", os.getcwd())

# Leer el archivo 'parametros.csv'
print('Cargnando parametros.csv')
df = pd.read_csv('parametros.csv')

# Convertir el DataFrame en un diccionario params
params = dict(zip(df['parametro'], df['valor']))

# Leer el archivo con el nombre definido en base de params en un DataFrame con todas las columnas como texto
print('Cargnando archivo base en formato Excel')
df = pd.read_excel(params.get('base') + '.xlsx', dtype=str)
# Agregar las columnas de seguimiento 'usuario', 'fecha_inicio' y 'fecha_fin' con valores vacíos
df[['usuario', 'robot_inicio', 'robot_fin']] = '', '', ''
#df.to_parquet('archivo.parquet')

# Conectar a la base de datos SQLite (o crearla si no existe)
print('Creando base.db')
conn = sqlite3.connect('base.db')

# Guardar el DataFrame en la base de datos SQLite
df.to_sql('base', conn, if_exists='replace', index=False)

# Cerrar la conexión
conn.close()
print('Proceso de creación finalizado')
