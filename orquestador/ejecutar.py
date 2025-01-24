import pandas as pd
import sqlite3
from datetime import datetime
import subprocess
import sys

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\orquestador")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
class Logger:
    def __init__(self, filename):
        self.terminal = sys.stdout
        self.log = open(filename, 'a')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()

sys.stdout = Logger('log.txt')
sys.stderr = sys.stdout

# Leer el archivo 'parametros.csv'
print('\nCargando parametros.csv')
df = pd.read_csv('parametros.csv')

# Convertir el DataFrame en un diccionario params
params = dict(zip(df['parametro'], df['valor']))

# Establecer variables de configuración
repositorio = params.get('repositorio')
robot = params.get('robot')
usuario = params.get('usuario')
base = params.get('base')
llave = params.get('llave')
continuar = True
respositorio_base = repositorio + 'orquestador\\base.db'
robot_base = robot + '\\' + base + '.csv'
robot_exe = robot + '\\' + robot + '.exe'

while continuar:
    # Conectar a la base de datos SQLite
    conn = sqlite3.connect(respositorio_base)

    # Leer el primer registro con la columna 'usuario' igual a ''
    query = "SELECT * FROM base WHERE usuario = '' LIMIT 1"
    df_result = pd.read_sql_query(query, conn)

    # Verificar si hay registros que cumplan la condición
    if df_result.empty:
        print('no hay más registros')
        conn.close()
        continuar = False
    else:
        # Imprimir el valor de la primera columna
        print('registro:',df_result[llave][0],datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'--------------------------------------------------')

        # Guardar el registro como resultado.csv en la carpeta del robot
        df_result.to_csv(robot_base, index=False)

        # Actualizar el valor de la columna 'usuario' a 'usuario' y 'robot_inicio' con la fecha y hora actual
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        update_query = """
        UPDATE base
        SET usuario = '{}', robot_inicio = ?
        WHERE usuario = '' AND {} = ?
        """.format(usuario,llave)
        conn.execute(update_query, (current_time, df_result[llave][0]))
        conn.commit()

        # Cerrar la conexión
        conn.close()

        # Ejecutar robot y esperar a que termine
        print('Ejecutando comprobador.exe')
        result = subprocess.run([robot_exe], 
                                cwd='comprobador', # Cambiar el directorio de trabajo
                                capture_output=True,
                                text=True)

        # Imprimir la salida del comando
        print('Salida de comprobador.exe:', result.stdout)
        print('Errores de comprobador.exe:', result.stderr)

        # Verificar si no hubo errores
        if result.stderr == '':
            # Conectar a la base de datos SQLite
            conn = sqlite3.connect(respositorio_base)

            # Actualizar el valor de la columna 'robot_fin' con la fecha y hora actual
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            update_query = """
            UPDATE base
            SET robot_fin = ?
            WHERE {} = ?
            """.format(llave)
            conn.execute(update_query, (current_time, df_result[llave][0]))
            conn.commit()

            # Cerrar la conexión
            conn.close()