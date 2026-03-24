"""
Automatizar SECOP II descarga
Descarga archivos de Google Drive según listada en el archivo "Descargas.csv"
"""

# %% Cargar datos
# Cargar librerías
from pandas import read_csv, read_excel
from datetime import datetime
from funciones import redirigir_log, mensaje
import requests
import zipfile

# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()

# Establecer variables de configuración
# Leer el archivo 'parametrosDescarga.csv'
print('\nCargando parametrosDescarga.csv')
df = read_csv('parametrosDescarga.csv')

# Convertir el DataFrame en un diccionario params
params = dict(zip(df['parametro'], df['valor']))

# Establecer variables de configuración
variables = {
    'repositorio': params.get('repositorio'),
    'usuario': params.get('usuario'),
    'base': params.get('base')
}
variables['robot'] = 'descarga_v1'

# Cargar base de datos de contratación "Descarga.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

# %% Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=1
    variables['contrato'] = str(dfbase.loc[i, 'NumDocumento'])
    id_funcionario_zip = str(dfbase.loc[i, 'NumDocumento']) + '.zip'
    url_drive_id = dfbase.loc[i, 'ENLACE'].split('=')[-1]
    direct_download_url = f'https://drive.google.com/uc?export=download&id={url_drive_id}'
    nombre_pdf = dfbase.loc[i, 'NOMBRE DE ARCHIVO'] + '.pdf'
    proceso = 'Descarga ' + nombre_pdf

    # Paso 1: Descargar archivo
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 1: ',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 1: '+proceso+' - '+horainicio)
    try:
        # Descargar el PDF en memoria
        respuesta = requests.get(direct_download_url)
        respuesta.raise_for_status() # Lanza error si la descarga falla
        
        # Manejo del ZIP (Modo 'a' para agregar o crear si no existe)
        with zipfile.ZipFile(variables['repositorio'] +'\\documentos\\' + id_funcionario_zip, mode='a', compression=zipfile.ZIP_DEFLATED) as zf:
            # Escribimos el contenido descargado en el zip con el nombre deseado
            zf.writestr(nombre_pdf, respuesta.content)
            
        horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print('Terminada ', proceso, '-', horafin, '--------------------------------------------------')
        mensaje(variables, 'Terminada '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
        mensaje(variables, proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])
        
    except Exception as e:
        print(f"  ✗ Error al procesar {nombre_pdf}: {e}")

# %%
