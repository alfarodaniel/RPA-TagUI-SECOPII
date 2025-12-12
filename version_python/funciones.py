"""
Automatizar SECOP II modificación
Funciones
"""
# Cargar librerías
from pandas import read_csv
from datetime import datetime
import sys
import re


# Funcion para redirigir la salida estándar y de error a un archivo log.txt
def redirigir_log():
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


# Funcion para leer el archivo 'parametros.csv' y establecer las variables de configuración
def parametros():
    # Leer el archivo 'parametros.csv'
    print('\nCargando parametros.csv')
    df = read_csv('parametros.csv')

    # Convertir el DataFrame en un diccionario params
    params = dict(zip(df['parametro'], df['valor']))

    # Establecer variables de configuración
    variables = {
        'repositorio': params.get('repositorio'),
        #'robot': 'aprobacion_v1',
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

    return variables


# Funcion mensaje para agregar una fila con el mensaje en los archivos historico.csv
def mensaje(variables, mensaje, repositorio=''):
    #for archivo in [f"{repositorio}historico.csv", seguimiento.csv]:
    with open(f"{repositorio}historico.csv", 'a') as file:
        file.write(variables['robot'] + ',' + 
                variables['usuario'] + ',' + 
                variables['contrato'] + ',' + 
                datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ',' + 
                mensaje + '\n')


# Función para validar si existe el registro en la página
def esperar(r, variables, objeto, descripcion="", boton="", frame="", stepCircle="", popup=""):
    for i in range(1, variables['espera']):
        # Verificar si depende de un boton
        if boton != "":
            r.click(boton)
        # Verificar si se encuentra en un frame
        if frame != "":
            r.frame(frame)
        # Verificar si se encuentra en un stepCircle
        if stepCircle != "":
            r.click(stepCircle)
        # Verificar si se encuentra en un popup
        if popup != "":
            r.popup(popup)
        # Verifica si se encontró objeto
        if r.present(objeto):
            #print('si')
            return True  # Retorna True si el objeto fue encontrado
        
        print(i, 'esperando', descripcion)
        # Esperar 1 segundo
        r.wait(1)

    print('No se encontró', descripcion)
    # Agregar una fila con la falla en los archivos historico.csv y seguimiento.csv
    mensaje(variables, 'No se encuentra ' + descripcion)

    return False  # Retorna False si el objeto no fue encontrado


# Función para iniciar sesión en SECOP II
def iniciar(r, variables):
    # Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
    r.ask('Iniciar')

    # Arbrir INICIAR SESION
    r.url('https://community.secop.gov.co/')

    # Iniciar sesion
    esperar(r, variables, '//*[@id="btnLoginButton"]', 'Botón Iniciar Sesión')
    r.type('//*[@id="txtUserName"]', '[clear]' + variables['user']) # Celda 'Nombre de usuario'
    r.type('//*[@id="txtPassword"]', '[clear]' + variables['password']) # Celda 'Contraseña'
    r.click('//*[@id="btnLoginButton"]') # Botón 'Iniciar Sesión'
    
    # Validar si aparace mensaje informativo
    r.wait(5)
    if r.present('btnAcknowledgeGen'):
        r.click('btnAcknowledgeGen')


# Función para cerrar sesión en SECOP II
def cerrar(r):
    # Cerrar sesión
    print('Cerrar sesion')
    r.click('//*[@id="userImage"]/img') # Imagen usuario
    r.wait(1)
    r.click('//*[@id="logOut"]') # Opción Salir


# Función para acceder al contrato en SECOP II
def acceder_contrato(r, proceso, variables, contratos=1):
    if not esperar(r, variables, '//*[@value="Procesos"]', 'Paso 0: Menú desplegable Procesos'): return False
    r.click('//*[@value="Procesos"]') # Menú Procesos
    if not esperar(r, variables, '//*[@id="lnkSubItem6"]', 'Paso 0: Submenú Procesos de la Entidad Estatal'): return False
    r.click('//*[@id="lnkSubItem6"]') # Submenú Procesos de la Entidad Estatal
    if not esperar(r, variables, 'txtSimpleSearchInput', 'Paso 0: Campo Búsqueda avanzada'): return False
    r.click('lnkAdvancedSearchLink') # Campo Búsqueda avanzada
    #r.wait(30)
    if not esperar(r, variables, '//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]', 'Paso 0: Menú desplegable Mis procesos'): return False
    r.click('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]') # Menú desplegable Mis procesos
    #r.vision('type(Key.UP)') # Subir una opción a Todos
    #r.vision('type(Key.ENTER)') # Seleccionar Todos
    if not esperar(r, variables, '//*[@id="selFilteringStatesSel_child"]/ul/li[1]', 'Paso 0: Seleccionar Todos'): return False
    r.click('//*[@id="selFilteringStatesSel_child"]/ul/li[1]') # Seleccionar Todos
    #r.wait(20)
    r.wait(2)
    r.type('txtReferenceTextbox', '[clear]' + proceso + '[enter]') # Campo Referencia
    #r.wait(1)
    r.type('//*[@id="dtmbCreateDateFromBox_txt"]', '[clear]01/01/2023') # Campo Fecha de creación desde
    #r.wait(1)
    r.click('btnSearchButton') # Botón Buscar
    if not esperar(r, variables, '//*[@title="' + proceso + '"]', 'Paso 0: Titulo proceso'): return False
    r.hover('//*[@title="' + proceso + '"]') # Seleccionar el proceso
    r.click('//*[@title="' + proceso + '"]') # Seleccionar el proceso
    if not esperar(r, variables, 'incBuyerDossierDetaillnkBuyerDossierDetailLink', 'Paso 0: Boton Detalle'): return False
    if contratos > 0:
        r.click('lnkProcurementContractViewLink_0') # Referencia
    else:
        r.click('//*[@id="incBuyerDossierDetaillnkRequestReference"]') # Enlace proceso
    return True


# Función para anexar documento en SECOP II
def anexar_documento(r, variables, documento, i):
    # Popup ANEXAR DOCUMENTO
    r.popup('DocumentAlternateUpload')
    #esperar('divAddFilesButton', 'Boton Buscar documento', popup='DocumentAlternateUpload')
    if not esperar(r, variables, 'divAddFilesButton', 'Boton Buscar documento', popup='DocumentAlternateUpload'): return False
    r.click('divAddFilesButton') # Boton Buscar documento
    r.wait(5)
    rutaarchivo = re.sub(r'\\+', r'\\', f'{variables["repositorio"]}documentos\\{documento}')
    r.vision(f'type("{rutaarchivo}")') # Ruta del documento
    r.vision('type(Key.ENTER)')
    r.wait(5)
    r.click('btnUploadFilesButtonBottom') # Botón Anexar
    if not esperar(r, variables, '//*[@id="tblFilesTable"]//*[@processed="success"]', 'Progreso DOCUMENTO ANEXO'): return False
    r.click('btnCancelBottomButtom') # Botón Cerrar
    r.popup(None) # Cierra el contexto del popup
    r.wait(5)
    return True
