import pandas as pd
import sqlite3
from datetime import datetime
import subprocess
import sys

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\orquestador")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
sys.stdout = open('log.txt', 'w')
sys.stderr = sys.stdout
# Leer el archivo 'parametros.csv'
print('Cargnando parametros.csv')
df = pd.read_csv('parametros.csv')

# Convertir el DataFrame en un diccionario params
params = dict(zip(df['parametro'], df['valor']))

# Establecer variables de configuración
robot = params.get('robot')
usuario = params.get('usuario')
base = params.get('base')
llave = params.get('llave')

# Conectar a la base de datos SQLite
conn = sqlite3.connect('base.db')

# Leer el primer registro con la columna 'usuario' igual a ''
query = "SELECT * FROM base WHERE usuario = '' LIMIT 1"
df_result = pd.read_sql_query(query, conn)

# Verificar si hay registros que cumplan la condición
if df_result.empty:
    print('no hay más registros')
    conn.close()
    exit()
else:
    # Imprimir el valor de la primera columna
    print('registro:',df_result[llave][0])

    # Guardar el registro como resultado.csv en la carpeta resultado
    df_result.to_csv(robot + '/' + base + '.csv', index=False)

    # Actualizar el valor de la columna 'usuario' a 'usuario' y 'robot_inicio' con la fecha y hora actual
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    update_query = """
    UPDATE base
    SET usuario = 'usuario', robot_inicio = ?
    WHERE usuario = '' AND {} = ?
    """.format(llave)
    conn.execute(update_query, (current_time, df_result[llave][0]))
    conn.commit()

# Cerrar la conexión
conn.close()


# Ejecutar comprobador.exe y esperar a que termine
print('Ejecutando comprobador.exe')
result = subprocess.run(['comprobador/comprobador.exe'], 
                        cwd='comprobador', # Cambiar el directorio de trabajo
                        capture_output=True,
                        text=True)

# Imprimir la salida del comando
print('Salida de comprobador.exe:', result.stdout)
print('Errores de comprobador.exe:', result.stderr)
