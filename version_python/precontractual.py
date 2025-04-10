"""
Automatizar SECOP II precontractual
Cargue de información pecontracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
import pandas as pd
from datetime import datetime
import sys

import os
os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
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
robot = 'precontractual_v1'
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

# Funcion mensaje para agregar una fila con el mensaje en los archivos historico.csv y seguimiento.csv
def mensaje(mensaje):
    for archivo in [f"{repositorio}historico.csv", "seguimiento.csv"]:
        with open(archivo, 'a') as file:
            file.write(robot + ',' + 
                    usuario + ',' + 
                    contrato + ',' + 
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ',' + 
                    mensaje + '\n')

# Función para validar si existe el registro en la página
def esperar(objeto, descripcion="", frame="", stepCircle=""):
    for i in range(1, espera):
        # Verificar si se encuentra en un frame
        if frame != "":
            r.frame(frame)
        # Verificar si se encuentra en un stepCircle
        if stepCircle != "":
            r.click(stepCircle)
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
if r.present('btnAcknowledgeGen'):
    r.click('btnAcknowledgeGen')

# Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Cargar página principal
    r.url('https://community.secop.gov.co/')

    # Paso 0: Crear proceso
    contrato = dfbase.loc[i, 'NUMERO DE CONTRATO']
    print('Paso 0: Crear proceso -',contrato,'-',datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'--------------------------------------------------')
    # Agregar una fila con el inicio en los archivos historico.csv y seguimiento.csv
    mensaje('Inicia')
    # Menú Procesos
    if not esperar('//*[@value="Procesos"]', 'Botón Procesos'): continue
    r.click('//*[@value="Procesos"]') # Menú Procesos
    if not esperar('//*[@id="lnkSubItem9"]', 'Subnemú Tipos de procesos'): continue
    r.click('//*[@id="lnkSubItem9"]') # Submenú Tipos de procesos

    # En la página de Tipos de procesos
    if not esperar('//*[@id="btnCreateProcedureButton12"]', 'Botón Crear Contratación régimen especial'): continue
    r.click('//*[@id="btnCreateProcedureButton12"]') # Botón Crear Contratación régimen especial

    # Frame CREAR PROCESO
    if not esperar('CreateProcedure_iframe', 'Frame CREAR PROCESO'): continue
    r.frame('CreateProcedure_iframe')
    if not esperar('txtProcedureReference', 'Campo Número del proceso',frame='CreateProcedure_iframe'): continue
    r.type('txtProcedureReference', '[clear]' 
           + dfbase.loc[i, 'TIPOLOGIA'] + '-' 
           + contrato + '-' 
           + dfbase.loc[i, 'VIGENCIA']) # Número de proceso
    r.type('txtProcedureName', '[clear]' 
           + 'PRESTAR SERVICIOS PROFESIONALES Y APOYO A LA GESTION') # Nombre
    r.type('txtBusinessOperationText', '[clear]' 
           + 'DIRECCIÓN DE CONTRATACIÓN') # Unidad de contratación
    r.vision('type(" - COMPRAS")')
    r.wait(2)
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.type('btnSaveCurrentDossierTop', '[enter]') # Botón Confirmar
    r.frame()

    # Paso 1: Información general
    print('Paso 1: Información general')
    if not esperar('txaDossierDescription', 'Campo Descripción'): continue
    r.type('txaDossierDescription', '[clear]' + dfbase.loc[i, 'PERFIL (PROFESION)']) # Campo Descripción
    r.click('divCategorizationRow_incDossierCategorizationUnspscMain_0_Lookup_LookupText') # Código UNSPSC
    r.vision('type("85101600")')
    r.wait(2)
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.type('btnAddAcquisitionButton', '[enter]') # Adquisición del PAA

    # Frame BUSCAR POR ADQUISICIONES PLANEADAS
    if not esperar('wndSearchPlannedAcquisitions_iframe', 'Frame CREAR PROCESO'): continue
    r.frame('wndSearchPlannedAcquisitions_iframe')
    if not esperar('txtSearchAcquisitionTXT', 'Campo Número del proceso',frame='CreateProcedure_iframe'): continue
    r.type('rdbgType_1', '') # ?????
    r.vision('type(Key.SPACE)')
    r.wait(5)
    r.vision('type(Key.TAB)')
    r.vision('type(Key.ENTER)')
    r.wait(5)
    if dfbase.loc[i, 'NOMBRE RUBRO'] == "Honorarios":
        r.type('chkGridAcqCheckBox_3', '') # ??????
    elif dfbase.loc[i, 'NOMBRE RUBRO'] == "Remuneracion Servicios Tecnicos":
        r.type('chkGridAcqCheckBox_4', '') # ??????
    else:
        r.type('grdPlannedAcquisitionsGrid_Paginator_goToPage_Next', '') # ??????
        r.vision('type(Key.SPACE)')
        r.wait(2)
        if dfbase.loc[i, 'NOMBRE RUBRO'] == "Contratacion Servicios Asistenciales Generales":
            r.type('chkGridAcqCheckBox_6', '') # ??????
        elif dfbase.loc[i, 'NOMBRE RUBRO'] == "Contratacion Servicios Asistenciales PIC":
            r.type('chkGridAcqCheckBox_5', '') # ??????
    r.vision('type(Key.SPACE)')
    r.wait(2)
    r.type('btnConfirmAcquisitionsSelection', '') # ??????
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.frame()

    # Información del contrato ??????
    if not esperar('rdbgOnlyPublicityOptions_1', '??????'): continue
    r.click('rdbgOnlyPublicityOptions_1') # ??????
    r.wait(5)
    r.select('selTypeOfContractSelect', 'ServicesProvisioning') # ??????
    if not esperar('selJustificationTypeOfContractSelected', '??????'): continue
    r.select('selJustificationTypeOfContractSelected', 'ApplicableRule') # ??????
    # Verificar la Duración estimada del contrato
    if dfbase.loc[i, 'DIAS'] == '0':
        r.type('nbxDurationGen', dfbase.loc[i, 'MESES'])  # Ingresar meses si los días son 0
        r.select('selDurationTypeP2Gen', 2)  # Seleccionar tipo de duración en meses
    else:
        r.type('nbxDurationGen', dfbase.loc[i, 'DIAS'])  # Campo Duración estimada del contrato
    r.wait(5)

    # Guardar información general
    r.wait(5)
    r.click('btnSaveProcedureTop') # Botón Guardar
    r.wait(5) # Esperar que aparezca el mensaje verde Proceso guardado con éxito
    r.click('btnApproveDossier') # Botón Continuar
    r.wait(5)
    if not esperar('//*[@id="btnNoPAAPublishedCurrentYearConfirmDialogModal"]', '??????'): continue
    r.click('//*[@id="btnNoPAAPublishedCurrentYearConfirmDialogModal"]') # ??????
    r.wait(5)

    # Paso 2: Configuración
    print('Paso 2: Configuración')
    # Cronograma
    if not esperar('//*[@id="stepDiv_2"][@class="LeftMenuButtonOn Black"]', '??????'): continue
    r.click('rdbgComplyWithMinimalPurchaseValue_1') # ??????
    r.wait(5)
    r.click('rdbgProcessAssociatedWithSentenceT302Value_1') # ??????
    r.wait(5)
    r.type('dtmbContractSignatureDate_txt', dfbase.loc[i, 'FECHA_INICIO'] + ' 00:00') # Campo Fecha de firma del contrato
    r.type('dtmbStartDateExecutionOfContract_txt', dfbase.loc[i, 'FECHA_INICIO'] + ' 00:00') # Campo Fecha de inicio de ejecución del contrato
    r.type('dtmbExecutionOfContractTerm_txt', dfbase.loc[i, 'FECHA_TERMINACION'] + ' 23:59') # Campo Plazo de ejecución del contrato
    r.click('body')

    # Configuración financiera
    r.click('rdbgWarrantiesField_0')
    r.wait(5)
    if dfbase.loc[i, 'PORCENTAJE_CONTRATO'] == "10":
        r.click('//*[@id="rdbgWarrantiesByLotsGroupsStagesField_1"]')
        r.wait(5)
        r.click('chkComplianceContractCB')
        r.wait(5)
        r.click('//*[@id="tdComplianceContractPercentageRBCell_divComplianceContractPercentageRBDiv_rdbComplianceContractPercentageRB"]')
        r.wait(5)
        r.type('//*[@id="nbxComplianceContractPercentageField"]', '[clear]' + dfbase.loc[i, 'PORCENTAJE_CONTRATO'])
        r.wait(2)
        r.type('dtmbComplianceContractStartDateBox_txt', '[clear]' + dfbase.loc[i, 'FECHA_DESDE_POLIZA'] + ' 00:00')
        r.type('dtmbComplianceContractEndDateBox_txt', '[clear]' + dfbase.loc[i, 'FECHA_HASTA_POLIZA'] + ' 23:59')
        r.click('rdbgCivilLiabilityField_1')
    elif dfbase.loc[i, 'PORCENTAJE_CONTRATO'] == "0":
        r.click('//*[@id="rdbgWarrantiesByLotsGroupsStagesField_1"]')
        r.wait(5)
        r.click('//*[@id="rdbgComplianceField_1"]')
        r.wait(3)
        r.click('//*[@id="rdbgCivilLiabilityField_0"]')
        r.wait(5)
        r.type('cbxCivilLiabilityValueField', dfbase.loc[i, 'PORCENTAJE_CONTRATO'])
    r.wait(5)

    # Precios
    r.type('cbxBasePrice', dfbase.loc[i, 'VALOR CONTRATO']) # Campo Valor estimado
    r.click('body')
    r.wait(5)

    # Información presupuestal
    r.click('//*[@id="rdbgFrameworkAgreementValue_1"]') # Radio button Implementación del Acuerdo de Paz
    r.wait(5)
    r.type('rdbgFrameworkAgreementValue_1', '') # Radio button Implementación del Acuerdo de Paz No
    r.vision('type(Key.ENTER)')
    r.vision('type(Key.SPACE)')

    r.click('selBudgetSourceSelect') # Lista Destinación del gasto
    r.vision('type(key.DOWN)')
    r.vision('type(key.ENTER)')

    r.click('body')
    r.wait(2)

    if r.present('cbxOwnResourcesAGRIValue'):
        r.click('rdbgBudgetOriginGNBCheckValueP2Gen_1')
        r.wait(2)
        r.click('rdbgBudgetOriginGSPCheckValueP2Gen_1')
        r.wait(2)
        r.click('rdbgBudgetOriginGRSCheckValueP2Gen_1')
        r.wait(2)
        r.click('rdbgBudgetOriginOwnResourcesAGRICheckValueP2Gen_1')
        r.wait(2)
        r.click('rdbgBudgetOriginCreditResourcesCheckValueP2Gen_1')
        r.wait(2)
        r.click('rdbgBudgetOriginOwnResourcesCheckValueP2Gen_0')
        r.wait(2)
        r.type('cbxBudgetOriginOwnResourcesValue', '[clear]')
        r.wait(2)
        r.click('cbxBudgetOriginOwnResourcesValue')
        r.vision('type(Key.DELETE)')
        #r.vision_step(f'valor_contrato = "{dfbase.loc[i, 'VALOR CONTRATO']}"')
        r.vision(f'type({dfbase.loc[i, 'VALOR CONTRATO']})')

    r.wait(5)
    r.click('body')
    r.wait(3)
    r.type('btnAddCode', '[enter]')
    r.vision('type(Key.SPACE)')

    # Frame Información presupuestal
    if not esperar('SIIFModal_iframe', 'Frame Información presupuestal'): continue
    r.frame('SIIFModal_iframe')
    if not esperar('rdbgOptionsToSelectRadioButton_0', 'Radio button CDP',frame='SIIFModal_iframe'): continue
    r.type('rdbgOptionsToSelectRadioButton_0', '') # Radio button CDP
    r.vision('type(Key.SPACE)')
    r.wait(5)
    r.type('txtSIIFIntegrationItemTextbox', dfbase.loc[i, 'N° DEL CDP']) # Campo Código
    r.type('cbxSIIFIntegrationItemBalanceTextbox', dfbase.loc[i, 'VALOR DEL CDP']) # Campo Saldo
    r.type('cbxSIIFIntegrationItemUsedValueTextbox', dfbase.loc[i, 'VALOR CONTRATO']) # Campo Saldo a comprometer
    r.type('txtSIIFIntegrationItemPCICodebox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código unidad ejecutora
    r.type('btnSIIFIntegrationItemButton', '') # Botón Crear
    r.vision('type(Key.ENTER)')
    r.frame()
    r.wait(5)

    r.click('btnSaveProcedureTop') # Botón Guardar
    r.wait(5)

    # Paso 3: Cuestionario
    print('Paso 3: Cuestionario')
    r.click('stepCircle_3') # stepCircle Cuestionario
    if not esperar('//*[@id="stepDiv_3"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_3'): continue
    r.wait(5)
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[3]/td[4]') # ??????
    r.wait(5)
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input[1]') # Campo Código UNSPSC
    r.vision('type("85101600")')
    r.wait(2)
    r.vision('type(key.DOWN)')
    r.vision('type(key.ENTER)')
    r.wait(2)
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[3]/input', dfbase.loc[i, 'PERFIL (PROFESION)']) # Campo Descripción
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[4]/input', '1') # Campo Cantidad
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[6]/input', dfbase.loc[i, 'VALOR CONTRATO'])
    r.click('btnSaveProcedureTop') # Botón Guardar
    r.wait(5)

    # Paso 4: Documentos del proceso
    print('Paso 4: Documentos del proceso')
    r.click('stepCircle_4') # stepCircle Documentos del proceso
    if not esperar('//*[@id="stepDiv_4"][@class="LeftMenuButtonOn Black"]', 'stepCircle Documentos del proceso en negro', stepCircle='stepCircle_4'): continue
    r.wait(5)
    r.type('incContractDocumentstxaExternalCommentsGen', 'Se anexa requerimiento segun necesidad de la institucion')
    r.wait(10)
    r.click('incContractDocumentsbtnUploadDocumentGen') # Botón Anexar documento
    r.wait(15)

    # Ventana emergente ANEXAR DOCUMENTO
    r.popup('OnDocumentsUploaded')
    r.click('divAddFilesButton') # Botón Buscar documento
    #r.vision_step(f'req = "{req}"')
    r.wait(10)
    r.vision(f'type("{repositorio}\\documentos\\{dfbase.loc[i, 'REQUERIMIENTO']}.pdf")') # Ruta del documento
    r.wait(10)
    r.vision('type(Key.ENTER)')
    r.wait(10)
    r.click('btnUploadFilesButtonBottom') # Botón subir ??????
    if not esperar('//*[@id="tblFilesTable"]//*[@processed="success"]', '?????? progreso'): continue
    r.click('btnCancelBottomButtom') # Botón cerrar ??????
    r.popup(None)  # Cierra el contexto del popup

    r.wait(10)
    r.click('body')
    r.wait(10)
    r.click('btnSaveProcedureTop') # Botón Guardar documentos anexos
    r.wait(15)
    r.click('btnOption_trRowToolbarTop_tdCell1_tbToolBar_Finish') # Botón Ir a publicar
    r.wait(5)

    # Frame PUBLICAR PROCESO
    r.frame('StartApprovalSupportModal_iframe')
    if not esperar('btnConfirmGen', 'Botón Confirmar',frame='StartApprovalSupportModal_iframe'): continue
    r.type('btnConfirmGen', '') # Botón Confirmar
    r.vision('type(Key.ENTER)')
    r.frame()

    # Agregar una fila con el final en los archivos historico.csv y seguimiento.csv
    mensaje('Finaliza')
    
# Cerrar robot
print('Cerrar robot', robot)
r.close()