import pandas as pd
import sqlite3

# Leer el archivo CSV de parámetros
dfparams = pd.read_csv('parametros.csv')

# Leer el archivo CSV de parámetros y guardarlo en un diccionario
parametros_dict = dict(zip(dfparams.iloc[:, 0], dfparams.iloc[:, 1]))

# Leer el archivo XLSX
df = pd.read_excel('archivo.xlsx')

# Guardar el DataFrame en formato Parquet
#df.to_parquet('archivo.parquet')

# Conectar a la base de datos SQLite (o crearla si no existe)
conn = sqlite3.connect('base.db')

# Guardar el DataFrame en la base de datos SQLite
df.to_sql('tabla', conn, if_exists='replace', index=False)

# Cerrar la conexión
conn.close()

