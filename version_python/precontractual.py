"""
Automatizar SECOP II aprobación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.xlsx"
"""
# %% Cargar datos
# Cargar librerías
import rpa as r
from pandas import read_excel
from datetime import datetime
from funciones import redirigir_log, parametros, iniciar, cerrar, mensaje, esperar, acceder_contrato, anexar_documento

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()

# Establecer variables de configuración
variables = parametros()
variables['robot'] = 'precontractual_v3'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

# Iniciar robot
print('Iniciar robot', variables['robot'])
r.init(visual_automation = True, turbo_mode=False)
#r.timeout(10)

# Iniciar sesion
iniciar(r, variables)

# %% Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=0
    proceso = 'CPS-' + dfbase.loc[i, 'NUMERO DE CONTRATO'] + '-' + dfbase.loc[i, 'VIGENCIA']

    # Cargar página principal
    r.url('https://community.secop.gov.co/')


    # Paso 0: Crear proceso
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Crear proceso --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Crear proceso --- proceso '+proceso+' - '+horainicio)
    # Menú Procesos
    if not esperar(r, variables, '//*[@value="Procesos"]', 'Botón Procesos'): continue
    r.click('//*[@value="Procesos"]') # Menú Procesos
    if not esperar(r, variables, '//*[@id="lnkSubItem9"]', 'Subnemú Tipos de procesos'): continue
    r.click('//*[@id="lnkSubItem9"]') # Submenú Tipos de procesos
    # En la página de Tipos de procesos
    if not esperar(r, variables, '//*[@id="btnCreateProcedureButton12"]', 'Botón Crear Contratación régimen especial'): continue
    r.click('//*[@id="btnCreateProcedureButton12"]') # Botón Crear Contratación régimen especial

    # Frame CREAR PROCESO
    if not esperar(r, variables, 'CreateProcedure_iframe', 'Frame CREAR PROCESO'): continue
    r.frame('CreateProcedure_iframe')
    if not esperar(r, variables, 'txtProcedureReference', 'Campo Número del proceso',frame='CreateProcedure_iframe'): continue
    #r.wait(3)
    r.type('txtProcedureReference', '[clear]' + proceso) # Número de proceso
    r.type('txtProcedureName', '[clear]PRESTACION DE SERVICIOS PROFESIONALES Y APOYO A LA GESTION') # Nombre
    r.type('txtBusinessOperationText', '[clear]DIRECCIÓN DE CONTRATACIÓN') # Unidad de contratación
    r.vision('type(" - COMPRAS")')
    r.wait(2)
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.type('btnSaveCurrentDossierTop', '[enter]') # Botón Confirmar
    r.wait(2)
    if r.exist('Ya existe un proceso con el mismo número.'):
        print('El proceso ya existe, se inicia modificaión')
        r.click('btnCloseGen') # Botón Cancelar
        r.frame()
        r.wait(2)
        if not acceder_contrato(r, proceso, variables, contratos=0): continue
    else:
        r.frame()


    # Paso 1: 1 Información general
    print('Paso 1: 1 Información general', proceso)
    if not esperar(r, variables, '//*[@id="stepDiv_1"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_1'): continue
    if not esperar(r, variables, 'txaDossierDescription', 'Campo Descripción'): continue
    r.type('txaDossierDescription', '[clear]' + dfbase.loc[i, 'PERFIL (PROFESION)']) # Campo Descripción
    # Clasificación del bien o servicio
    #r.wait(3)
    r.click('//*[@id="divCategorizationRow_incDossierCategorizationUnspscMain_0_Lookup_LookupText"]') # Campo Código UNSPSC
    r.type('//*[@id="divCategorizationRow_incDossierCategorizationUnspscMain_0_Lookup_LookupText"]', '[clear]') # Campo Código UNSPSC
    r.vision('type("85101600")') # Código UNSPSC
    r.wait(2)
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.click('btnAddAcquisitionButton') # Botón Agregar de Adquisiciones planeadas
    
    # Frame BUSCAR POR ADQUISICIONES PLANEADAS
    if not esperar(r, variables, 'wndSearchPlannedAcquisitions_iframe', 'Frame CREAR PROCESO'): continue
    r.frame('wndSearchPlannedAcquisitions_iframe')
    if not esperar(r, variables, 'txtSearchAcquisitionTXT', 'Campo Descripción',frame='wndSearchPlannedAcquisitions_iframe'): continue
    r.wait(3)
    r.type('txtSearchAcquisitionTXT', '[clear]' + dfbase.loc[i, 'NOMBRE RUBRO']) # Campo Descripción
    r.type('rdbgType_1', 'Yes') # Radio button Todos
    r.vision('type(Key.SPACE)') # Radio button Todos
    r.type('btnSearchAcquisition', 'Yes') # Botón Buscar
    r.vision('type(Key.SPACE)') # Botón Buscar
    r.wait(3)
    r.type('chkGridAcqCheckBox_0', 'Yes') # Checkbox Adquisisiciones planeadas
    r.vision('type(Key.SPACE)') # Checkbox Adquisisiciones planeadas
    r.type('chkGridAcqCheckBox_0', '') # Checkbox Adquisisiciones planeadas
    r.wait(2)
    r.type('btnConfirmAcquisitionsSelection', 'Yes') # Botón Buscar
    r.vision('type(Key.SPACE)') # Botón Buscar
    r.wait(5)
    r.frame()

    # Publicidad Proceso
    if not esperar(r, variables, 'rdbgOnlyPublicityOptions_1', 'Radio button Uso del módulo de forma publicitaria'): continue
    r.click('rdbgOnlyPublicityOptions_1') # Radio button No
    r.wait(5)
    # Información del contrato
    r.select('selTypeOfContractSelect', 'ServicesProvisioning') # Lista desplegable Tipo
    if not esperar(r, variables, 'selJustificationTypeOfContractSelected', 'Lista desplegable Justificación de la modalidad de contratación'): continue
    r.select('selJustificationTypeOfContractSelected', 'ApplicableRule') # Lista desplegable Justificación de la modalidad de contratación
    # Verificar la Duración estimada del contrato
    r.wait(2)
    if dfbase.loc[i, 'DIAS'] == '0':
        r.type('nbxDurationGen', dfbase.loc[i, 'MESES'])  # Ingresar meses si los días son 0
        r.select('selDurationTypeP2Gen', 2)  # Seleccionar tipo de duración en meses
    else:
        r.type('nbxDurationGen', '[clear]' + dfbase.loc[i, 'DIAS'])  # Campo Duración del contrato
    r.wait(2)
    r.type('dtmbContractEndDateGen_txt', '[clear]' + dfbase.loc[i, 'FECHA_TERMINACION'] + ' 23:59') # Campo Descripción
    r.click('btnApproveDossier') # Botón Continuar
    r.wait(2)
    

    # Paso 2: 2 Configuración
    print('Paso 2: 2 Configuración', proceso)
    # Decreto 248 de 2021
    if not esperar(r, variables, '//*[@id="stepDiv_2"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_2'): continue
    r.click('rdbgComplyWithMinimalPurchaseValue_1') # Radio button No
    r.wait(2)
    # Sentencia T-302 de 2017
    r.click('rdbgProcessAssociatedWithSentenceT302Value_1') # Radio button No
    r.wait(2)
    # Cronograma
    r.type('dtmbContractSignatureDate_txt', '[clear]' + dfbase.loc[i, 'FECHA_INICIO'] + ' 23:59') # Campo Fecha de firma del contrato
    r.type('dtmbStartDateExecutionOfContract_txt', '[clear]' + dfbase.loc[i, 'FECHA_INICIO'] + ' 23:59') # Campo Fecha de inicio de ejecución del contrato
    r.type('dtmbExecutionOfContractTerm_txt', '[clear]' + dfbase.loc[i, 'FECHA_TERMINACION'] + ' 23:59') # Campo Plazo de ejecución del contrato
    r.click('body')
    # Configuración financiera
    if dfbase.loc[i, 'SOLICITUD_DE_GARANTIAS'] == "Si":
        r.click('rdbgWarrantiesField_0') # Solicitud de garantías Radio button Sí
        if not esperar(r, variables, 'rdbgWarrantiesByLotsGroupsStagesField_1', 'Garantias por lotes, grupos o etapas Radio button No'): continue
        r.click('rdbgWarrantiesByLotsGroupsStagesField_1') # Garantias por lotes, grupos o etapas Radio button No
        r.click('rdbgComplianceField_1') # Cumplimiento Radio button No
        r.click('rdbgCivilLiabilityField_0') # Responsabilidad civil extra contractual Radio button Sí
        if not esperar(r, variables, 'tdCivilLiabilityMinWagesRBCell_rdbCivilLiabilityMinWagesRB', 'Radio button No. de SMMLV'): continue
        r.click('tdCivilLiabilityMinWagesRBCell_rdbCivilLiabilityMinWagesRB') # Responsabilidad civil extra contractual Radio button Sí
        r.wait(2)
        r.type('nbxCivilLiabilityMinWagesField', '[clear]100') # Campo No. de SMMLV
    """
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
    """
    # Precios
    r.type('cbxBasePrice', '[clear]') # Campo Valor estimado
    r.wait(2)
    r.type('cbxBasePrice', '[clear]' + dfbase.loc[i, 'VALOR CONTRATO']) # Campo Valor estimado
    r.click('body')
    r.wait(5)
    # Información presupuestal
    r.click('rdbgFrameworkAgreementValue_1') # Implementación del Acuerdo de Paz Radio button No
    r.wait(5)
    r.select('selExpenseTypeSelect', '0') # Lista desplegable Destinación del gasto
    r.wait(2)
    r.click('rdbgBudgetOriginGNBCheckValueP2Gen_1') # Presupuesto General PGN Radio button No
    r.wait(2)
    r.click('rdbgBudgetOriginGSPCheckValueP2Gen_1') # Sistema General SGP Radio button No
    r.wait(2)
    r.click('rdbgBudgetOriginGRSCheckValueP2Gen_1') # Sistema General SGR Radio button No
    r.wait(2)
    r.click('rdbgBudgetOriginOwnResourcesAGRICheckValueP2Gen_0') # Recursos Propios Radio button Si
    r.wait(2)
    r.click('rdbgBudgetOriginCreditResourcesCheckValueP2Gen_1') # Recursos de Crédito Radio button No
    r.wait(2)
    r.click('rdbgBudgetOriginOwnResourcesCheckValueP2Gen_1') # Otros Recursos Radio button No
    r.wait(2)
    r.click('cbxOwnResourcesAGRIValue')
    r.vision('type(Key.DELETE)')
    r.vision(f'type("{dfbase.loc[i, 'VALOR CONTRATO']}")')
    r.click('body')
    r.wait(2)
    r.click('btnAddCode')
    
    # Frame Información presupuestal
    if not esperar(r, variables, 'SIIFModal_iframe', 'Frame Información presupuestal'): continue
    r.frame('SIIFModal_iframe')
    if not esperar(r, variables, 'rdbgOptionsToSelectRadioButton_0', 'Radio button CDP',frame='SIIFModal_iframe'): continue
    #r.wait(3)
    r.type('rdbgOptionsToSelectRadioButton_0', 'Yes') # Radio button CDP
    r.vision('type(Key.SPACE)')
    if not esperar(r, variables, 'txtSIIFIntegrationItemTextbox', 'Campo Código'): continue
    r.wait(3)
    r.type('txtSIIFIntegrationItemTextbox', 'Yes') # Campo Código
    r.vision(f'type("{dfbase.loc[i, 'N° DEL CDP']}")') # Campo Código
    r.type('cbxSIIFIntegrationItemBalanceTextbox', dfbase.loc[i, 'VALOR DEL CDP']) # Campo Saldo
    r.type('cbxSIIFIntegrationItemUsedValueTextbox', dfbase.loc[i, 'VALOR CONTRATO']) # Campo Saldo a comprometer
    r.type('txtSIIFIntegrationItemPCICodebox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código unidad ejecutora
    r.type('btnSIIFIntegrationItemButton', 'Yes') # Botón Crear
    r.vision('type(Key.ENTER)')
    r.wait(3)
    if r.exist('Este código ya existe para este tipo de información presupuestal'):
        print('El código ya existe, se omite')
        r.type('btnSIIFIntegrationItemCancelButton', 'Yes') # Botón Cancelar
        r.vision('type(Key.ENTER)')
    r.frame()
    r.wait(3)


    # Paso 3: 3 Cuestionario
    print('Paso 3: 3 Cuestionario', proceso)
    r.click('stepCircle_3') # stepCircle Cuestionario
    if not esperar(r, variables, '//*[@id="stepDiv_3"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_3'): continue
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[3]/td[4]') # Lista de precios de la oferta 1+
    r.wait(5)
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input[1]') # Campo Código UNSPSC
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input[1]','[clear]') # Código UNSPSC
    r.vision('type("85101600")') # Código UNSPSC
    r.wait(2)
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.wait(2)
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[3]/input', '[clear]' + dfbase.loc[i, 'PERFIL (PROFESION)']) # Campo Descripción
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[4]/input', '[clear]') # Campo Cantidad
    r.click('body')
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[4]/input') # Campo Cantidad
    r.vision('type("1")')
    r.click('body')
    r.type('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[6]/input', '[clear]')
    r.click('body')
    r.click('//*[@id="incQuestionnairefltDataSheet"]/table/tbody/tr[7]/td[5]/table/tbody/tr/td[2]/table/tbody/tr[1]/td[6]/input') # Campo Precio unitario estimado
    r.vision(f'type("{dfbase.loc[i, "VALOR CONTRATO"]}")')
    r.click('body')
    r.click('btnSaveProcedureTop') # Botón Guardar
    r.wait(5)

    """
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
    """


    # Paso 4: Publicar
    print('Paso 4: Publicar', proceso)
    r.click('btnOption_trRowToolbarTop_tdCell1_tbToolBar_Finish') # Botón Ir a publicar
    r.wait(5)
    if not esperar(r, variables, '//input[@id="btnPublishRequest" and @title="Publicar"]', 'Botón Publicar'): continue
    #r.click('btnPublishRequest') # Botón Publicar
    r.wait(3)
    r.type('btnPublishRequest', 'Yes') # Botón Publicar
    r.vision('type(Key.ENTER)')
    r.wait(5)
    r.vision('type(Key.ENTER)')
    if not esperar(r, variables, 'stpBuyerDossierInfoAnchor', 'Label Información general'): continue
    r.click('trRowToolbarTop_tdCell1_tbToolBar_lnkBack') # Botón Volver
    if not esperar(r, variables, 'btnFinishRequest', 'Botón Finalizar'): continue
    r.wait(5)
    r.type('btnFinishRequest', 'Yes') # Botón Finalizar
    r.vision('type(Key.ENTER)')
    #r.wait(2)
    if not esperar(r, variables, 'btnFinishRequestConfirmDialogModal', 'Botón Confirmar'): continue
    r.click('btnFinishRequestConfirmDialogModal') # Botón Confirmar
    r.wait(5)
    if not esperar(r, variables, '//input[@id="btnCreateContractButton" and @title="Crear"]', 'Botón Crear contrato'): continue
    r.click('btnCreateContractButton') # Botón Crear contrato
    
    """
    # Frame PUBLICAR PROCESO
    r.frame('StartApprovalSupportModal_iframe')
    if not esperar('btnConfirmGen', 'Botón Confirmar',frame='StartApprovalSupportModal_iframe'): continue
    r.type('btnConfirmGen', '') # Botón Confirmar
    r.vision('type(Key.ENTER)')
    r.frame()
    """


    # Paso 5: 1 Información general
    print('Paso 5: 1 Información general', proceso)
    # Identificacióndel contrato
    if not esperar(r, variables, 'txtContractReference1Gen', 'Campo Número del contrato'): continue
    r.type('txtContractReference1Gen', '[clear]' + proceso) # Campor Número de contrato
    r.type('txaContractDescription1Gen', '[clear]PRESTAR SERVICIOS PROFESIONALES Y DE APOYO A LA GESTIÓN COMO ' + dfbase.loc[i, 'PERFIL (PROFESION)']) # Campo Objeto del contrato
    r.click('rdbgLiquidationValue_1') # Liquidación Radio button No
    r.click('rdbgEnvironmentObligationValue_1') # Obligaciones Ambientales Radio button No
    r.click('rdbgPostConsumptionObligationValue_1') # Obligaciones por consumo Radio button No
    r.click('rdbgReversionValue_1') # Reversión Radio button No
    # Información del Proveedor contratista
    r.click('btnSelectAwardedSupplier') # Botón Seleccionar
    
    # Frame Buscar entidad para seleccionar
    if not esperar(r, variables, 'SelectAwardedSupplier_iframe', 'Frame Buscar entidad para seleccionar'): continue
    r.frame('SelectAwardedSupplier_iframe')
    if not esperar(r, variables, 'txtAllWords2Search', 'Campo Buscar',frame='SelectAwardedSupplier_iframe'): continue
    r.wait(3)
    r.type('txtAllWords2Search', dfbase.loc[i, 'DOCUMENTO DE IDENTIFICACION']) # Campo Buscar
    r.type('btnSearchCompanies', 'Yes') # Botón Buscar
    r.vision('type(Key.ENTER)')
    if not esperar(r, variables, 'grdSingleCompaniesMatchingResults_rdbSelection_0', 'Radio button'): continue
    r.wait(3)
    r.type('grdSingleCompaniesMatchingResults_rdbSelection_0', 'Yes') # Radio button
    r.vision('type(Key.SPACE)')
    r.type('btnSelectGenCompany', 'Yes') # Botón Agregar
    r.vision('type(Key.SPACE)')
    r.frame()
    r.wait(3)


    # Paso 6: 2 Condiciones
    print('Paso 6: 2 Condiciones', proceso)
    r.click('stepCircle_2') # stepCircle Condiciones
    if not esperar(r, variables, '//*[@id="stepDiv_2"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_2'): continue
    # Condiciones ejecución y entrega
    r.select('selIncoterm', 'A definir') # Lista desplegable Condiciones de entrega
    r.wait(3)
    r.click('rdbgRenewableContract_0') # El contrato puede ser prorrogado button Si
    r.type('dtmbContractRenewalDateGen_txt', '[clear]' + dfbase.loc[i, 'FECHA_TERMINACION'] + ' 23:59') # Campo Plazo de ejecución del contrato
    r.click('body')
    # Condiciones de facturación y pago
    r.select('selPaymentMethod', 'Abono en cuenta') # Lista desplegable Forma de pago
    r.select('selPaymentTerm', 'A definir') # Lista desplegable Plazo de pago de la factura
    # Anexos del contrato
    r.click('btnUploadDocumentGen') # Botón Anexar documentos
    r.wait(3)
    anexar_documento(r, variables, dfbase.loc[i, "NOMBRE_DOCUMENTO_ANEXO"]) # Popup ANEXAR DOCUMENTO
    
    # Municipio de ejecución del contrato
    r.click('btnAddLocationGenPC') # Botón Agregar ubicación

    # Frame BUSCAR UBICACIÓN
    if not esperar(r, variables, 'LocationSelectView_iframe', 'Frame BUSCAR UBICACIÓN'): continue
    r.frame('LocationSelectView_iframe')
    if not esperar(r, variables, 'txtAddressSearchText', 'Campo Buscar',frame='LocationSelectView_iframe'): continue
    r.wait(3)
    r.type('txtAddressSearchText', 'Calle 66 # 15 - 41') # Campo Dirección
    r.type('txtZipCodeSearchText', '111221') # Campo Código postal
    r.type('btnSearchGen', 'Yes') # Botón Buscar
    r.vision('type(Key.ENTER)')
    if not esperar(r, variables, 'grdLocations_rdbSelection_0', 'Radio button'): continue
    r.wait(3)
    r.type('grdLocations_rdbSelection_0', 'Yes') # Radio button
    r.vision('type(Key.SPACE)')
    r.type('tbToolBarPlaceHolder_btnOKGen', 'Yes') # Botón Aceptar
    r.vision('type(Key.SPACE)')
    r.frame()
    r.wait(3)


    # Paso 7: 3 Bienes y servicios
    print('Paso 7: 3 Bienes y servicios', proceso)
    r.click('stepCircle_3') # stepCircle Bienes y servicios
    if not esperar(r, variables, '//*[@id="stepDiv_3"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_3'): continue
    r.click('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td[4]') # Bienes y servicios 1+
    r.wait(5)
    r.type('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', '[clear]') # Campo Precio unitario
    #r.type('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR CONTRATO']) # Campo Precio unitario
    r.click('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input') # Campo Precio unitario
    r.vision('type(Key.DELETE)')
    r.vision(f'type("{dfbase.loc[i, 'VALOR CONTRATO']}")') # Campo Precio unitario
    

    # Paso 8: 5 Documentos del contrato
    print('Paso 8: 5 Documentos del contrato', proceso)
    r.click('stepCircle_5') # stepCircle Documentos del contrato
    if not esperar(r, variables, '//*[@id="stepDiv_5"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_5'): continue
    # Anexos del contrato
    r.click('btnUploadContractDocument') # Botón Anexar documentos
    r.wait(3)
    anexar_documento(r, variables, dfbase.loc[i, "NOMBRE_DOCUMENTO"]) # Popup ANEXAR DOCUMENTO
    
    
    # Paso 9: Confirmar
    print('Paso 9: Confirmar', proceso)
    r.click('btnOption_tbContractToolbar_Finish') # Botón Confirmar
    r.wait(5)
    r.click('chkCheckBoxAgreeTerms') # Checkbox Acepto el valor del contrato
    r.wait(2)
    r.click('btnContractTotalValueValidationConfirmDialogModal') # Botón Confirmar
    
    # Frame Confirmar
    if not esperar(r, variables, 'StartApprovalSupportModal_iframe', 'Frame Confirmar'): continue
    r.frame('StartApprovalSupportModal_iframe')
    if not esperar(r, variables, 'btnConfirmGen', 'Boton Confirmar',frame='StartApprovalSupportModal_iframe'): continue
    r.wait(3)
    r.type('btnConfirmGen', 'Yes') # Botón Confirmar
    r.vision('type(Key.SPACE)')
    r.frame()
    r.wait(3)

    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Pre contractual del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Pre contractual del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Pre contractual del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])

# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()