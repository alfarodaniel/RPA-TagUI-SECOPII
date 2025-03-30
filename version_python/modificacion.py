"""
Automatizar SECOP II modificación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
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


# Función para terminar la edición del contrato
def terminar_edicion_contrato(i, proceso):
    # Terminar Edición del Contrato
    # Paso 2: Información general
    print('Paso 2: Información general')
    print('--- 2 Informacion General --- proceso', proceso)
    r.click('stepCircle_2') # Información general
    if not esperar('//*[@id="stepCircleSelected_2"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_2'): return False
    r.type('txtContractReference1Gen', '[clear]' + proceso) # ??????
    r.type('dtmbContractEnd_txt', '[clear]' + dfbase.loc[i, 'FECHA FINAL'] + ' 23:59') # Campo Fecha de terminación del contrato
    r.wait(3)

    # Paso 3: Condiciones
    print('Paso 3: Condiciones')
    print('--- 3 Condiciones --- proceso', proceso)
    r.click('stepCircle_3') # Condiciones
    if not esperar('//*[@id="stepCircleSelected_3"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_3'): return False
    r.click('//*[@id="rdbgComplyWithMinimalPurchaseValue_1"]')
    r.wait(5)
    if r.present('//*[@id="rdbgProcessAssociatedWithSentenceT302Value_1"]'):
        r.click('//*[@id="rdbgProcessAssociatedWithSentenceT302Value_1"]')
        r.wait(5)
    r.select('selIncoterm', 'NXTWY.DLVY.2') # ??????
    if r.present('dtmbDeliverEndDateGen_txt'):
        r.type('dtmbDeliverEndDateGen_txt', '[clear]' + dfbase.loc[i, 'FECHA FINAL'] + ' 23:59') # Campo Fecha final de ejecución (estimada)
    r.click('body')
    r.wait(2)
    r.click('rdbgRenewableContract_0') # ??????
    if r.present('dtmbContractRenewalDateGen_txt'):
        r.type('dtmbContractRenewalDateGen_txt', '[clear]' + dfbase.loc[i, 'FECHA FINAL'] + ' 23:59') # Campo Fecha de notificación de prorrogación
    
    # Paso 4: Bienes y Servicios
    print('Paso 3: Bienes y Servicios')
    print('--- 3 Bienes y Servicios --- proceso', proceso)
    r.click('stepCircle_4') # Bienes y Servicios
    if not esperar('//*[@id="stepCircleSelected_4"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_4'): return False
    r.click('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[2]/td[4]') # ??????
    r.wait(2)
    r.dclick('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input')
    r.type('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR TOTAL']) # Campo Precio unitario

    # Paso 7: Informacion presupuestal
    print('Paso 7: Informacion presupuestal')
    print('--- 7 Informacion presupuestal --- proceso', proceso)
    r.click('stepCircle_7') # Informacion presupuestal
    if not esperar('//*[@id="stepCircleSelected_7"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_7'): return False
    if r.present('rdbgBudgetOriginGNBCheckValueP2Gen_1'):
        r.click('rdbgBudgetOriginGNBCheckValueP2Gen_1') # ??????
        r.wait(2)
        r.click('rdbgBudgetOriginGSPCheckValueP2Gen_1') # ??????
        r.wait(2)
        r.click('rdbgBudgetOriginGRSCheckValueP2Gen_1') # ??????
        r.wait(2)
        r.click('rdbgBudgetOriginOwnResourcesAGRICheckValueP2Gen_1') # ??????
        r.wait(2)
        r.click('rdbgBudgetOriginCreditResourcesCheckValueP2Gen_1') # ??????
        r.wait(2)
        r.click('rdbgBudgetOriginOwnResourcesCheckValueP2Gen_0') # ??????
        r.wait(2)
    r.wait(5)
    if r.present('cbxOwnResourcesAGRIValue'):
        r.type('cbxOwnResourcesAGRIValue', '[clear]0') # ??????
        r.click('body')
        r.type('cbxBudgetOriginCreditResourcesValue', '[clear]0') # ??????
        r.click('body')
        r.wait(2)
        r.dclick('cbxBudgetOriginOwnResourcesValue')
        r.type('cbxBudgetOriginOwnResourcesValue', dfbase.loc[i, 'VALOR TOTAL']) # ??????
    r.wait(5)
    if dfbase.loc[i, 'CLASE'] == "adicion":
        print('--- Tiene Adicion --- proceso', proceso)
        r.click('body')
        r.wait(5)
        r.type('btnAddCode', '[enter]') # Botón Agregar
        r.wait(2)
        r.vision('type(Key.SPACE)')

        # Frame Información presupuestal
        if not esperar('SIIFModal_iframe', 'Frame Información presupuestal'): return False
        r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
        if not esperar('rdbgOptionsToSelectRadioButton_0', 'Radio button CDP',frame='SIIFModal_iframe'): return False
        r.type('rdbgOptionsToSelectRadioButton_0', '') # Radio button CDP
        r.vision('type(Key.SPACE)')
        r.wait(5)
        r.type('txtSIIFIntegrationItemTextbox', dfbase.loc[i, 'CDP_1']) # Campo Código
        r.type('cbxSIIFIntegrationItemUsedValueTextbox', dfbase.loc[i, 'ADICION_VALOR']) # Campo Saldo a comprometer
        r.type('txtSIIFIntegrationItemPCICodebox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código unidad ejecutora
        r.type('btnSIIFIntegrationItemButton', '') # Botón Crear
        r.vision('type(Key.SPACE)')
        r.wait(5)
        if r.present('//*[@id="msgMessagesPanel"]/tbody/tr'):
            r.type('//*[@id="btnSIIFIntegrationItemCancelButton"]', '') # ??????
            r.vision('type(Key.SPACE)')
        r.frame()
        
        r.click('body')
        r.wait(5)             

    # Paso 1: Modificacion del Contrato
    print('Paso 1: Modificacion del Contrato')
    print('--- 1 Modificacion del Contrato --- proceso', proceso)
    r.click('stepCircle_1') # Modificacion del Contrato
    if not esperar('//*[@id="stepCircleSelected_1"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_7'): return False
    r.click('cmAttachmentsOptions') # Lista Anexar documentos
    r.wait(2)
    r.vision('type(Key.TAB)')
    r.vision('type(Key.ENTER)')
    r.wait(5)
    
    # Popup ANEXAR DOCUMENTO
    r.popup('DocumentAlternateUpload')
    r.wait(5)
    r.click('divAddFilesButton') # Boton Buscar documento
    r.wait(5)
    r.vision(f'type("{repositorio}\\documentos\\OTROSI_{dfbase.loc[i, 'NUMERO']}_CPS_{dfbase.loc[i, 'NUMERO DE CONTRATO']}_{dfbase.loc[i, 'VIGENCIA']}.pdf")') # Ruta del documento
    r.vision('type(Key.ENTER)')
    r.wait(5)
    r.click('btnUploadFilesButtonBottom') # ??????
    if not esperar('//*[@id="tblFilesTable"]//*[@processed="success"]', '??????'): return False
    r.click('btnCancelBottomButtom') # ??????
    r.popup(None) # Cierra el contexto del popup

    r.wait(5)
    if dfbase.loc[i, 'CLASE'] == "adicion":
        print('--- Tiene Adicion --- proceso', proceso)
        r.wait(10)
        r.click('cmAttachmentsOptions') # ??????
        r.wait(2)
        r.vision('type(Key.TAB)')
        r.vision('type(Key.ENTER)')
        r.wait(10)

        # Popup ANEXAR DOCUMENTO
        r.popup('DocumentAlternateUpload')
        r.wait(5)
        r.click('divAddFilesButton') # Boton Buscar documento
        r.wait(5)
        r.vision(f'type("{repositorio}\\disponibilidad\\OTROSI_{dfbase.loc[i, 'CDP_1']}_{dfbase.loc[i, 'VIGENCIA']}.pdf")') # Ruta del documento
        r.vision('type(Key.ENTER)')
        r.wait(5)
        r.click('btnUploadFilesButtonBottom') # ??????
        if not esperar('//*[@id="tblFilesTable"]//*[@processed="success"]', '??????'): return False
        r.click('btnCancelBottomButtom') # ??????
        r.popup(None) # Cierra el contexto del popup
   
    r.wait(5)
    if dfbase.loc[i, 'OFICIO'] == "SI":
        print('--- Tiene Adicion --- proceso', proceso)
        r.wait(10)
        r.click('cmAttachmentsOptions') # ??????
        r.wait(2)
        r.vision('type(Key.TAB)')
        r.vision('type(Key.ENTER)')
        r.wait(10)

        # Popup ANEXAR DOCUMENTO
        r.popup('DocumentAlternateUpload')
        r.wait(5)
        r.click('divAddFilesButton') # Boton Buscar documento
        r.wait(5)
        r.vision(f'type("{repositorio}\\anexo\\{dfbase.loc[i, 'AREA']}.pdf")') # Ruta del documento
        r.vision('type(Key.ENTER)')
        r.wait(5)
        r.click('btnUploadFilesButtonBottom') # ??????
        if not esperar('//*[@id="tblFilesTable"]//*[@processed="success"]', '??????'): return False
        r.click('btnCancelBottomButtom') # ??????
        r.popup(None) # Cierra el contexto del popup
        
    r.wait(5)
    subir = r.read('//*[@id="spnDocumentNameValue_0"]') # ??????
    r.write(proceso + ',' + subir, 'subir.csv')
    r.wait(2)
    r.type('txaModificationPurpose', 'modificacion') # Campo Justificación de la modificación
    r.wait(2)

    print('--- Finalizar Modificacion ---')
    r.click('btnOption_tbContractToolbar_Finish') # ??????
    r.wait(5)
    r.type('chkCheckBoxAgreeTerms', '') # Check box Acepto el valor del contrato
    r.wait(2)
    r.vision('type(Key.SPACE)')
    r.wait(5)
    r.type('btnContractTotalValueValidationConfirmDialogModal', '') # Boton Confirmar
    r.wait(2)
    r.vision('type(Key.SPACE)')
    r.wait(5)
    if r.present('//*[@id="btnConfirmSIIFCodeRegistrationWarningButton"]'):
        r.type('btnConfirmSIIFCodeRegistrationWarningButton', '') # ??????
        r.wait(5)
        r.vision('type(Key.SPACE)')

    # Frame FLUJOS DE APROBACIÓN
    if not esperar('StartApprovalSupportModal_iframe', 'Frame FLUJOS DE APROBACIÓN'): return False
    r.frame('StartApprovalSupportModal_iframe')
    if not esperar('btnConfirmGen', 'Boton Confirmar',frame='StartApprovalSupportModal_iframe'): return False
    r.type('btnConfirmGen', '') # Boton Confirmar
    r.vision('type(Key.ENTER)')
    r.frame()

    # Esperar cargue del elemento esperado en la página
    if not esperar('IncTaskApproval_spnThisDocumentIsWaitingForAWorkfGen', '??????'): return False
    if r.present('IncTaskApproval_btnApproveButton'):
        r.click('IncTaskApproval_btnApproveButton') # ??????
        # Esperar cargue del elemento esperado en la página
        if not esperar('IncTaskApproval_spnThisDocumentIsWaitingForAWorkfGen', '??????'): return False
    subio = r.read('//*[@id="spnDocumentNameValue_0"]') # ??????
    r.write(proceso + ',' + subio, 'subio.csv')

    return True


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
    # Variables
    #i=0
    proceso = 'CPS-' + dfbase.loc[i, 'NUMERO DE CONTRATO'] + '-' + dfbase.loc[i, 'VIGENCIA']

    # Cargar página principal
    r.url('https://community.secop.gov.co/')

    # Paso 0: Acceder al contrato
    print('Paso 0: Acceder al contrato --- proceso',proceso,'-',datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'--------------------------------------------------')
    if not esperar('//*[@value="Procesos"]', 'Menú desplegable Procesos'): continue
    r.click('//*[@value="Procesos"]') # Menú Procesos
    r.click('//*[@id="lnkSubItem6"]') # Submenú Procesos de la Entidad Estatal
    if not esperar('txtSimpleSearchInput', 'Campo Búsqueda avanzada'): continue
    r.click('lnkAdvancedSearchLink') # Campo Búsqueda avanzada
    #r.wait(30)
    if not esperar('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]', 'Menú desplegable Mis procesos'): continue
    r.click('//*[@id="selFilteringStatesSel_msdd"]//*[@class="ddArrow arrowoff"]') # Menú desplegable Mis procesos
    #r.vision('type(Key.UP)') # Subir una opción a Todos
    #r.vision('type(Key.ENTER)') # Seleccionar Todos
    r.click('//*[@id="selFilteringStatesSel_child"]/ul/li[1]') # Seleccionar Todos
    #r.wait(20)
    r.wait(2)
    r.type('txtReferenceTextbox', '[clear]' + proceso + '[enter]') # Campo Referencia
    r.wait(2)
    r.type('//*[@id="dtmbCreateDateFromBox_txt"]', '[clear]01/01/2023') # Campo Fecha de creación desde
    r.wait(2)
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
    r.wait(2)
    #r.type('//*[@id="chkBypassWorkflowCheck"]', '') # Check box ¿Requiere reconocimiento del proveedor?
    #r.vision('type(Key.SPACE)')
    #r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ?????? 
    r.wait(2)
    #r.type('btnConfirmGen', '') # Botón Confirmar
    #r.vision('type(Key.SPACE)')
    r.click('btnConfirmGen')
    r.wait(5)
    r.frame()

    # Paso 3: 4 Bienes y Servicios
    print('Paso 3: 4 Bienes y Servicios --- proceso', proceso)
    r.click('stepCircle_4') # Bienes y Servicios
    if not esperar('//*[@id="stepCircleSelected_4"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_4'): return False
    r.click('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[2]/td[4]') # Símbolo + en Incluya el precio como lo indique la Entidad Estatal
    if not esperar('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', 'Campo Precio unitario'): return False
    r.dclick('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input') # Campo Precio unitario
    r.type('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR TOTAL']) # Campo Precio unitario

    # Paso 4: 7 Informacion presupuestal
    print('Paso 4: 7 Información presupuestal --- proceso', proceso)
    r.click('stepCircle_7') # Informacion presupuestal
    if not esperar('//*[@id="stepCircleSelected_7"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_7'): return False
    if not esperar('cbxOwnResourcesAGRIValue', 'Campo Recursos Propios'): return False
    r.dclick('cbxOwnResourcesAGRIValue') # Campo Recursos Propios
    r.type('cbxOwnResourcesAGRIValue', dfbase.loc[i, 'VALOR TOTAL']) # Campo Recursos Propios
    #r.wait(2)
    r.click('btnAddCode') # Botón Agregar de CDP/Vigencias Futuras

    # Frame Información presupuestal
    r.frame('SIIFModal_iframe')
    esperar('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe')
    #if not esperar('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe'): return False
    r.type('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Yes') # Radio button CDP
    r.vision('type(Key.SPACE)') 
    r.type('txtSIIFIntegrationItemTextbox', '[clear]') # Campo Código
    r.vision(f'type("{dfbase.loc[i, 'CDP_1']}")') 
    r.type('cbxSIIFIntegrationItemBalanceTextbox', dfbase.loc[i, 'VALOR TOTAL']) # Campo Código
    r.type('cbxSIIFIntegrationItemUsedValueTextbox', dfbase.loc[i, 'ADICION_VALOR']) # Campo Saldo a comprometer445445
    r.type('txtSIIFIntegrationItemPCICodebox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código unidad ejecutora
    r.type('btnSIIFIntegrationItemButton', 'yes') # Botón Crear
    #r.type('btnSIIFIntegrationItemCancelButton', 'yes') # Botón Cancelar
    r.vision('type(Key.SPACE)')
    r.frame()
    
















    general = r.read('//*[@id="spnContractState"]') # ??????
    fecha_general = r.read('//*[@id="dtmbContractEnd_txt"]') # ??????
    r.write(proceso + ',' + general + ',' + fecha_general, 'general.csv')

    # Verificar si esta inhabilitado, de lo contrario continúa
    if r.present('//*[@id="stepDiv_8"]/div[3]'): # ??????
        if r.present('//*[@id="spnContractState"]'): # ??????
            sin_modificacion = r.read('//*[@id="spnContractState"]') # ??????
            r.write(proceso + ',' + sin_modificacion, 'sin_modificacion.csv')
    else:
        r.click('stepCircle_8') # ??????

        # Esperar cargue del elemento esperado en la página
        if not esperar('//*[@id="stepCircleSelected_8"][@class="MainColor4 circle22 Black stepOn"]', '????'): continue
        print('--- Evalua --- proceso', proceso)
        if r.present('//*[@id="spnModificationStatusValue_0"]'): # ??????
            modificacion = r.read('//*[@id="spnModificationStatusValue_0"]') # ??????
            fecha = r.read('//*[@id="dtmbModificationDateValue_0_txt"]') # ??????
            r.write(proceso + ',' + modificacion + ',' + fecha, 'modificacion.csv')
        else:
            r.write(proceso + ',ninguno', 'modificacion.csv')
        
        print('--- 1 Informacion General --- proceso', proceso)
        if r.present('//*[@id="IncTaskApproval_incTreeView_1TaskTreeGroup1Task1CellApproveDate"]/span/font'): # ??????
            r.click('btnMakeModification') # ??????
            r.wait(5)
            if r.present('lnkModifyContractGeneralLink'): # ??????
                print('--- Ingresa --- proceso', proceso)
                r.click('lnkModifyContractGeneralLink') # ??????

                # Frame TIPO DE MODIFICACION
                if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
                r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
                if not esperar('btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
                r.click('body')
                r.wait(2)
                r.type('chkBypassWorkflowCheck', '') # Check box ¿Requiere reconocimiento del proveedor?
                r.vision('type(Key.SPACE)')
                r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ??????
                r.wait(2)
                r.type('btnConfirmGen', '') # Botón Confirmar
                r.vision('type(Key.SPACE)')
                r.wait(5)
                r.frame()

                # Terminar Edición del Contrato
                terminar_edicion_contrato(i, proceso)
                continue

            print('--- Modificacion ---')
            if r.present('//*[@id="msgMessagesPanel"]/tbody/tr'):
                # Paso 8: ??????
                print('Paso 8: ????????')
                print('--- 8 ?????? --- proceso', proceso)
                r.click('stepCircle_8') # Modificacion del Contrato
                if not esperar('//*[@id="stepCircleSelected_8"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_8'): continue
                print('--- Evalua En Edicion---')
                if r.present('En edici'): # ??????
                    r.wait(2)
                    modificacion = r.read('//*[@id="spnModificationStatusValue_0"]') # ??????
                    r.write(proceso + ',' + modificacion, 'modificacion.csv')
                    r.click('//*[@id="lnkEditLink_0"]')
                    r.wait(25)
                    if r.present('//*[@id="stepDiv_2"]/div[3]'): # ??????
                        r.click('lnkModifyContractGeneralLink') # ??????
                        
                        # Frame TIPO DE MODIFICACION
                        if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
                        r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
                        if not esperar('btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
                        r.click('body')
                        r.wait(2)
                        r.type('chkBypassWorkflowCheck', '') # Check box ¿Requiere reconocimiento del proveedor?
                        r.vision('type(Key.SPACE)')
                        r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ??????
                        r.wait(2)
                        r.type('btnConfirmGen', '') # Botón Confirmar
                        r.vision('type(Key.SPACE)')
                        r.wait(5)
                        r.frame()
                        
                    # Terminar Edición del Contrato
                    terminar_edicion_contrato(i, proceso)
                    continue

                r.wait(10)
                print('--- Evalua Aceptado por el proveedor ---')
                if r.present('Aceptado por'):
                    r.click('//*[@id="lnkEditLink_0"]')
                    if not esperar('//*[@id="lblRowModifyContractGeneralMsg"]', '??????'): continue
                    r.click('//*[@id="btnFinishModification"]') # ??????
                    r.wait(10)
                    if r.present('//*[@id="btnConfirmInvoicesWarningButton"]'):
                        r.click('//*[@id="btnConfirmInvoicesWarningButton"]') # ??????
                        r.wait(10)
                    completo_pbl = r.read('//*[@id="dtmbContractEnd_txt"]') # ??????
                    r.write(proceso + ',' + completo_pbl, 'completo_pbl.csv')
                    r.wait(10)
                    if not esperar('//*[@id="btnMakeModification"]', '??????'): continue
                    if r.present('//*[@id="trRowToolbarTop_tdCell1_tbToolBar_lnkBack"]'):
                        r.click('//*[@id="trRowToolbarTop_tdCell1_tbToolBar_lnkBack"]') # ??????
                        if not esperar('//*[@id="btnMakeModification"]', '??????'): continue

                    # Paso 8: ??????
                    print('Paso 8: ????????')
                    print('--- 8 ?????? --- proceso', proceso)
                    r.click('stepCircle_8') # Modificacion del Contrato
                    if not esperar('//*[@id="stepCircleSelected_8"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_8'): continue
                    if r.present('//*[@id="IncTaskApproval_incTreeView_1TaskTreeGroup1Task1CellApproveDate"]/span/font'):
                         r.click('btnMakeModification') # ??????
                         r.wait(5)
                         if r.present('lnkModifyContractGeneralLink'):
                            print('--- Ingresa --- proceso', proceso)
                            r.click('lnkModifyContractGeneralLink') # ??????

                            # Frame TIPO DE MODIFICACION
                            if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
                            r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
                            if not esperar('btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
                            r.click('body')
                            r.wait(2)
                            r.type('chkBypassWorkflowCheck', '') # Check box ¿Requiere reconocimiento del proveedor?
                            r.vision('type(Key.SPACE)')
                            r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ??????
                            r.wait(2)
                            r.type('btnConfirmGen', '') # Botón Confirmar
                            r.vision('type(Key.SPACE)')
                            r.wait(5)
                            r.frame()

                            # Terminar Edición del Contrato
                            terminar_edicion_contrato(i, proceso)
                            continue
                
                print('--- Evalua En Aprobada---')
                if r.present('Aprobada'):
                    r.click('//*[@id="lnkEditLink_0"]')
                    if not esperar('//*[@id="lblRowModifyContractGeneralMsg"]', '??????'): continue
                    r.click('//*[@id="btnFinishModification"]')
                    r.wait(15)
                    if r.present('//*[@id="btnConfirmInvoicesWarningButton"]'):
                        r.type('//*[@id="btnConfirmInvoicesWarningButton"]', '')
                        r.vision('type(Key.SPACE)')
                        r.wait(3)
                    if not esperar('//*[@id="btnMakeModification"]', '??????'): continue
                    if r.present('//*[@id="trRowToolbarTop_tdCell1_tbToolBar_lnkBack"]'):
                        r.click('//*[@id="trRowToolbarTop_tdCell1_tbToolBar_lnkBack"]')
                        if not esperar('//*[@id="btnMakeModification"]', '??????'): continue
                    
                    # Paso 8: ??????
                    print('Paso 8: ????????')
                    print('--- 8 ?????? --- proceso', proceso)
                    r.click('stepCircle_8') # Modificacion del Contrato
                    if not esperar('//*[@id="stepCircleSelected_8"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_8'): continue
                    r.wait(5)
                    print('--- Evalua --- proceso', proceso)
                    if r.present('//*[@id="spnModificationStatusValue_0"]'):
                        modificacion = r.read('//*[@id="spnModificationStatusValue_0"]') # ??????
                        fecha = r.read('//*[@id="dtmbModificationDateValue_0_txt"]') # ??????
                        r.write(proceso + ',' + modificacion + ',' + fecha, 'modificacion.csv')
                    else:
                        r.write(proceso + ',ninguno', 'modificacion.csv')
                    
                    if r.present('//*[@id="IncTaskApproval_incTreeView_1TaskTreeGroup1Task1CellApproveDate"]/span/font'):
                        r.click('btnMakeModification')
                        r.wait(5)
                        if r.present('lnkModifyContractGeneralLink'):
                            print('--- Ingresa --- proceso', proceso)
                            r.click('lnkModifyContractGeneralLink')
                        
                            # Frame TIPO DE MODIFICACION
                            if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
                            r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
                            if not esperar('btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
                            r.click('body')
                            r.wait(2)
                            r.type('chkBypassWorkflowCheck', '') # Check box ¿Requiere reconocimiento del proveedor?
                            r.vision('type(Key.SPACE)')
                            r.click('/html/body/div[2]/div/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/input[1]') # ??????
                            r.wait(2)
                            r.type('btnConfirmGen', '') # Botón Confirmar
                            r.vision('type(Key.SPACE)')
                            r.wait(5)
                            r.frame()

                            # Terminar Edición del Contrato
                            terminar_edicion_contrato(i, proceso)
                            continue


# Cerrar robot
print('Cerrar robot', robot)
r.close()