"""
Automatizar SECOP II modificación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
import pandas as pd
import re
from datetime import datetime
import sys

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
repositorio = params.get('repositorio')
robot = 'modificación_v1'
espera = int(params.get('espera'))
usuario = params.get('usuario')
user = params.get('user')
password = params.get('password')
base = params.get('base')
contrato = ''
#llave = params.get('llave')
#continuar = True
#respositorio_base = repositorio + 'orquestador\\base.db'
#robot_base = robot + '\\' + base + '.csv'
#robot_exe = robot + '\\' + robot + '.exe'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = pd.read_excel(base, dtype=str)

# Mostrar el valor de la columna 'NUMERO' para el primer registro
#primer_registro_numero = base.loc[0, 'NUMERO']
#print('Valor de la columna NUMERO para el primer registro:', primer_registro_numero)


# Funcion mensaje para agregar una fila con el mensaje en los archivos historico.csv
def mensaje(mensaje, repositorio=''):
    #for archivo in [f"{repositorio}historico.csv", seguimiento.csv]:
    with open(f"{repositorio}historico.csv", 'a') as file:
        file.write(robot + ',' + 
                usuario + ',' + 
                contrato + ',' + 
                datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ',' + 
                mensaje + '\n')


# Función para validar si existe el registro en la página
def esperar(objeto, descripcion="", boton="", frame="", stepCircle="", popup=""):
    for i in range(1, espera):
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
    mensaje('No se encuentra' + descripcion)

    return False  # Retorna False si el objeto no fue encontrado


# Iniciar robot
print('Iniciar robot', robot)
r.init(visual_automation = True, turbo_mode=False)
#r.timeout(10)

# Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
r.ask('Iniciar')

# Arbrir INICIAR SESION
r.url('https://community.secop.gov.co/')

# Iniciar sesion
esperar('//*[@id="btnLoginButton"]', 'Botón Iniciar Sesión')
r.type('//*[@id="txtUserName"]', '[clear]' + user) # Celda 'Nombre de usuario'
r.type('//*[@id="txtPassword"]', '[clear]' + password) # Celda 'Contraseña'
r.click('//*[@id="btnLoginButton"]') # Botón 'Iniciar Sesión'
#esperar('??????????????????????????????????????????????', 'Nombre de usuario')
#r.click('//*[@id="btnButton1"]')

# Validar si aparace mensaje informativo
r.wait(5)
if r.present('btnAcknowledgeGen'):
    r.click('btnAcknowledgeGen')

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
    mensaje('Paso 0: Acceder al contrato --- proceso '+proceso+' - '+horainicio)
    if not esperar('//*[@value="Procesos"]', 'Menú desplegable Procesos'): continue
    r.click('//*[@value="Procesos"]') # Menú Procesos
    if not esperar('//*[@id="lnkSubItem6"]', 'Submenú Procesos de la Entidad Estatal'): continue
    r.click('//*[@id="lnkSubItem6"]') # Submenú Procesos de la Entidad Estatal
    if not esperar('txtSimpleSearchInput', 'Campo Búsqueda avanzada'): continue
    r.click('lnkAdvancedSearchLink') # Campo Búsqueda avanzada
    #r.wait(30)
    if not esperar('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]', 'Menú desplegable Mis procesos'): continue
    r.click('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]') # Menú desplegable Mis procesos
    #r.vision('type(Key.UP)') # Subir una opción a Todos
    #r.vision('type(Key.ENTER)') # Seleccionar Todos
    if not esperar('//*[@id="selFilteringStatesSel_child"]/ul/li[1]', 'Seleccionar Todos'): continue
    r.click('//*[@id="selFilteringStatesSel_child"]/ul/li[1]') # Seleccionar Todos
    #r.wait(20)
    r.wait(2)
    r.type('txtReferenceTextbox', '[clear]' + proceso + '[enter]') # Campo Referencia
    #r.wait(1)
    r.type('//*[@id="dtmbCreateDateFromBox_txt"]', '[clear]01/01/2023') # Campo Fecha de creación desde
    #r.wait(1)
    r.click('btnSearchButton') # Botón Buscar
    if not esperar('//*[@title="' + proceso + '"]', 'Titulo proceso'): continue
    r.click('//*[@title="' + proceso + '"]') # Seleccionar el proceso
    if not esperar('incBuyerDossierDetaillnkBuyerDossierDetailLink', 'Boton Detalle'): continue
    r.click('lnkProcurementContractViewLink_0') # Referencia

    # Paso 1: 8 Modificaciones del Contrato
    print('Paso 1:  8 Modificaciones del Contrato --- proceso', proceso)
    if not esperar('//*[@id="lnk_stpmStepManager9"]', 'Menú 8 Modificacione del Contrato'): continue
    r.click('//*[@id="lnk_stpmStepManager9"]') # Menú 8 Modificacione del Contrato
    if not esperar('//*[@id="btnMakeModification"]', 'Botón Modificar'): continue
    r.click('//*[@id="btnMakeModification"]') # Botón Modificar
    
    # Paso 2: 1 Modificación del Contrato
    print('Paso 2: 1 Modificación del Contrato --- proceso', proceso)
    if not esperar('lnkModifyContractGeneralLink', 'Enlace Modificar el contrato'): continue
    r.click('lnkModifyContractGeneralLink') # Enlace Modificar el contrato

    # Frame TIPO DE MODIFICACION
    #if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
    r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
    if not esperar('btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
    #r.click('body')
    r.click('//*[@id="chkBypassWorkflowCheck"]')
    r.wait(1)
    #r.type('//*[@id="chkBypassWorkflowCheck"]', '') # Check box ¿Requiere reconocimiento del proveedor?
    #r.vision('type(Key.SPACE)')
    #r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ?????? 
    #r.wait(2)
    #r.type('btnConfirmGen', '') # Botón Confirmar
    #r.vision('type(Key.SPACE)')
    r.click('btnConfirmGen')
    r.wait(5)
    r.frame()

    # Paso 3: 2 Información general
    if dfbase.loc[i, 'TIPO_MODIFICACION'] == 'ADICION - PRORROGA':
        print('Paso 3: 2 Información general --- proceso', proceso)
        r.click('stepCircle_2') # Información general
        if not esperar('//*[@id="stepCircleSelected_2"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_2'): continue
        r.type('//*[@id="dtmbContractEnd_txt"]', '[clear]' + dfbase.loc[i, 'PRORROGA_FECHA'] + ' 23:59') # Fecha de terminación del contrato

    # Paso 4: 3 Condiciones
    if dfbase.loc[i, 'TIPO_MODIFICACION'] == 'ADICION - PRORROGA':
        print('Paso 4: 3 Condiciones --- proceso', proceso)
        r.click('stepCircle_3') # Información general
        if not esperar('//*[@id="stepCircleSelected_3"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_3'): continue
        r.type('//*[@id="dtmbContractRenewalDateGen_txt"]', '[clear]' + dfbase.loc[i, 'PRORROGA_FECHA'] + ' 23:59') # Fecha de notificación de prorrogación

    # Paso 5: 4 Bienes y Servicios
    print('Paso 5: 4 Bienes y Servicios --- proceso', proceso)
    r.click('stepCircle_4') # Bienes y Servicios
    if not esperar('//*[@id="stepCircleSelected_4"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_4'): continue
    r.click('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[2]/td[4]') # Símbolo + en Incluya el precio como lo indique la Entidad Estatal
    if not esperar('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', 'Campo Precio unitario'): continue
    r.dclick('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input') # Campo Precio unitario
    r.type('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR TOTAL']) # Campo Precio unitario

    # Paso 6: 7 Informacion presupuestal
    print('Paso 6: 7 Información presupuestal --- proceso', proceso)
    r.click('stepCircle_7') # Informacion presupuestal
    if not esperar('//*[@id="stepCircleSelected_7"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_7'): continue
    if not esperar('cbxOwnResourcesAGRIValue', 'Campo Recursos Propios'): continue
    r.dclick('cbxOwnResourcesAGRIValue') # Campo Recursos Propios
    r.type('cbxOwnResourcesAGRIValue', dfbase.loc[i, 'VALOR TOTAL']) # Campo Recursos Propios
    if not esperar('//*[@id="SIIFModal_iframe"]', 'Frame Información presupuestal',boton='btnAddCode'): continue
    #r.click('btnAddCode') # Botón Agregar de CDP/Vigencias Futuras

    # Frame Información presupuestal
    r.frame('SIIFModal_iframe')
    #esperar('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe')
    if not esperar('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe'): continue
    r.type('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Yes') # Radio button CDP
    r.vision('type(Key.SPACE)') # Radio button CDP
    r.type('txtSIIFIntegrationItemTextbox', '[clear]') # Campo Código 
    r.vision(f'type("{dfbase.loc[i, 'CDP_1']}")') # Campo Código
    #r.type('cbxSIIFIntegrationItemBalanceTextbox', dfbase.loc[i, 'VALOR TOTAL']) # Campo Saldo
    r.type('cbxSIIFIntegrationItemUsedValueTextbox', dfbase.loc[i, 'ADICION_VALOR']) # Campo Valor a utilizar
    r.type('txtSIIFIntegrationItemPCICodebox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código unidad ejecutora
    r.type('btnSIIFIntegrationItemButton', 'yes') # Botón Crear
    #r.type('btnSIIFIntegrationItemCancelButton', 'yes') # Botón Cancelar
    r.vision('type(Key.SPACE)') # Botón Crear
    r.frame()
    
    # Paso 7: 1 Modificación del Contrato
    print('Paso 7: 1 Modificación del Contrato --- proceso', proceso)
    r.click('stepCircle_1') # Menù 1 Modificación del Contrato
    if not esperar('//*[@id="stepCircleSelected_1"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_1'): continue
    r.wait(2)
    r.click('cmAttachmentsOptions') # Lista Anexar documentos
    r.wait(2)
    r.click('linkUploadNew') # Item Anexar nuevo documento
    #r.vision('type(Key.TAB)')
    #r.vision('type(Key.ENTER)')
    #r.wait(5)
    
    # Popup ANEXAR DOCUMENTO
    r.popup('DocumentAlternateUpload')
    #esperar('divAddFilesButton', 'Boton Buscar documento', popup='DocumentAlternateUpload')
    if not esperar('divAddFilesButton', 'Boton Buscar documento', popup='DocumentAlternateUpload'): continue
    r.click('divAddFilesButton') # Boton Buscar documento
    r.wait(5)
    rutaarchivo = re.sub(r'\\+', r'\\', f'{repositorio}documentos\\{dfbase.loc[i, "NOMBRE_DOCUMENTO"]}.zip')
    r.vision(f'type("{rutaarchivo}")') # Ruta del documento
    r.vision('type(Key.ENTER)')
    r.wait(5)
    r.click('btnUploadFilesButtonBottom') # Botón Anexar
    if not esperar('//*[@id="tblFilesTable"]//*[@processed="success"]', 'Progreso DOCUMENTO ANEXO'): continue
    r.click('btnCancelBottomButtom') # Botón Cerrar
    r.popup(None) # Cierra el contexto del popup
    r.wait(5)
    
    print('--- Finalizar Modificacion --- proceso', proceso)
    r.click('body')
    if not esperar('txaModificationPurpose', 'Campo Justificación de la modificación'): continue
    r.type('txaModificationPurpose', 'Modificación') # Campo Justificación de la modificación
    r.click('btnOption_tbContractToolbar_Finish') # Finalizar Modificacion
    if not esperar('chkCheckBoxAgreeTerms', 'Check box Acepto el valor del contrato'): continue
    #r.type('chkCheckBoxAgreeTerms', 'yes') # Check box Acepto el valor del contrato
    r.click('chkCheckBoxAgreeTerms') # Check box Acepto el valor del contrato
    r.wait(2)
    #r.vision('type(Key.SPACE)')
    #r.wait(5)
    #r.type('btnContractTotalValueValidationConfirmDialogModal', '') # Boton Confirmar
    r.click('btnContractTotalValueValidationConfirmDialogModal') # Boton Confirmar
    #r.click('btnContractTotalValueValidationCancelDialogModal') # Boton Cancelar
    #r.vision('type(Key.SPACE)')

    # Frame Confirmar
    r.frame('StartApprovalSupportModal_iframe')
    if not esperar('//*[@id="btnConfirmGen"]', 'Boton Confirmar',frame='StartApprovalSupportModal_iframe'): continue
    r.type('//*[@id="btnConfirmGen"]', 'Yes') # Boton Confirmar
    #r.wait(2)
    #r.type('//*[@id="btnCancelGen"]', 'Yes') # Boton Cancelar
    r.vision('type(Key.SPACE)') # Botón Confirmar
    r.frame()
    r.wait(5)
    #r.wait(20)
    
    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Modificación del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje('Terminada Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje('Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, repositorio)


# Cerrar sesión
print('Cerrar sesion')
r.click('//*[@id="userImage"]/img') # Imagen usuario
r.wait(1)
r.click('//*[@id="logOut"]') # Opción Salir

# Cerrar robot
print('Cerrar robot', robot)
r.close()