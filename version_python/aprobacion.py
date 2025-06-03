"""
Automatizar SECOP II aprobación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
import pandas as pd
from datetime import datetime
import sys
from funciones import iniciar, cerrar, mensaje, esperar

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
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
variables = {
    'repositorio': params.get('repositorio'),
    'robot': 'aprobacion_v1',
    'espera': int(params.get('espera')),
    'usuario': params.get('usuario'),
    'user': params.get('user'),
    'password': params.get('password'),
    'base': params.get('base'),
    'contrato': ''
    # 'llave': params.get('llave'),
    # 'continuar': True,
    # 'respositorio_base': repositorio + 'orquestador\\base.db',
    # 'robot_base': robot + '\\' + base + '.csv',
    # 'robot_exe': robot + '\\' + robot + '.exe'
}

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = pd.read_excel(variables['base'], dtype=str)

# Mostrar el valor de la columna 'NUMERO' para el primer registro
#primer_registro_numero = base.loc[0, 'NUMERO']
#print('Valor de la columna NUMERO para el primer registro:', primer_registro_numero)

# Iniciar robot
print('Iniciar robot', variables['robot'])
r.init(visual_automation = True, turbo_mode=False)
#r.timeout(10)

# Iniciar sesion
iniciar(r, variables)

# Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=0
    proceso = 'CPS-' + dfbase.loc[i, 'NUMERO DE CONTRATO'] + '-' + dfbase.loc[i, 'VIGENCIA']

    # Cargar página principal
    r.url('https://community.secop.gov.co/')

    # Paso 0: Acceder al contrato
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Acceder al contrato --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Acceder al contrato --- proceso '+proceso+' - '+horainicio)
    if not esperar(r, variables, '//*[@value="Procesos"]', 'Paso 0: Menú desplegable Procesos'): continue
    r.click('//*[@value="Procesos"]') # Menú Procesos
    if not esperar(r, variables, '//*[@id="lnkSubItem6"]', 'Submenú Procesos de la Entidad Estatal'): continue
    r.click('//*[@id="lnkSubItem6"]') # Submenú Procesos de la Entidad Estatal
    if not esperar(r, variables, 'txtSimpleSearchInput', 'Paso 0: Campo Búsqueda avanzada'): continue
    r.click('lnkAdvancedSearchLink') # Campo Búsqueda avanzada
    #r.wait(30)
    if not esperar(r, variables, '//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]', 'Paso 0: Menú desplegable Mis procesos'): continue
    r.click('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]') # Menú desplegable Mis procesos
    #r.vision('type(Key.UP)') # Subir una opción a Todos
    #r.vision('type(Key.ENTER)') # Seleccionar Todos
    if not esperar(r, variables, '//*[@id="selFilteringStatesSel_child"]/ul/li[1]', 'Paso 0: Seleccionar Todos'): continue
    r.click('//*[@id="selFilteringStatesSel_child"]/ul/li[1]') # Seleccionar Todos
    #r.wait(20)
    r.wait(2)
    r.type('txtReferenceTextbox', '[clear]' + proceso + '[enter]') # Campo Referencia
    #r.wait(1)
    r.type('//*[@id="dtmbCreateDateFromBox_txt"]', '[clear]01/01/2023') # Campo Fecha de creación desde
    #r.wait(1)
    r.click('btnSearchButton') # Botón Buscar
    if not esperar(r, variables, '//*[@title="' + proceso + '"]', 'Paso 0: Titulo proceso'): continue
    r.click('//*[@title="' + proceso + '"]') # Seleccionar el proceso
    if not esperar(r, variables, 'incBuyerDossierDetaillnkBuyerDossierDetailLink', 'Paso 0: Boton Detalle'): continue
    r.click('lnkProcurementContractViewLink_0') # Referencia

    # Paso 1: 8 Modificaciones del Contrato
    print('Paso 1:  8 Modificaciones del Contrato --- proceso', proceso)
    if not esperar(r, variables, '//*[@id="lnk_stpmStepManager9"]', 'Paso 1: Menú 8 Modificacione del Contrato'): continue
    r.click('//*[@id="lnk_stpmStepManager9"]') # Menú 8 Modificacione del Contrato
    if not esperar(r, variables, 'lnkEditLink_0', 'Paso 1: Enlace Editar'): continue
    r.click('lnkEditLink_0') # Enlace Editar

    # Paso 2: 1 Modificación del Contrato
    print('Paso 2: 1 Modificación del Contrato --- proceso', proceso)
    if not esperar(r, variables, 'IncTaskApproval_btnApproveButton', 'Paso 2: Boton Aprobar'): continue
    r.click('IncTaskApproval_btnApproveButton') # Boton Aprobar
    r.wait(5)

    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Aprobación del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Aprobación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Aprobación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])


# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()