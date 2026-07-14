"""
Automatizar SECOP II modificación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
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
variables['robot'] = 'modificación_v2'

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


    # Paso 0: Acceder al contrato
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Acceder al contrato --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Acceder al contrato --- proceso '+proceso+' - '+horainicio)
    if not acceder_contrato(r, proceso, variables): continue
    r.wait(10)


    # Paso 1: 8 Modificaciones del Contrato
    print('Paso 1:  8 Modificaciones del Contrato --- proceso', proceso)
    if not esperar(r, variables, '//*[@id="lnk_stpmStepManager9"]', 'Menú 8 Modificacione del Contrato'): continue
    r.click('//*[@id="lnk_stpmStepManager9"]') # Menú 8 Modificacione del Contrato
    if not esperar(r, variables, '//*[@id="btnMakeModification"]', 'Botón Modificar'): continue
    # Validar si el contrato está en edición
    if r.read('//*[@id="spnModificationStatusValue_0"]') == 'InEdition':
        r.click('//*[@id="lnkEditLink_0"]') # Enlace Editar
    else:
        r.click('//*[@id="btnMakeModification"]') # Botón Modificar
    

    # Paso 2: 1 Modificación del Contrato
    print('Paso 2: 1 Modificación del Contrato --- proceso', proceso)
    if not esperar(r, variables, 'lnkModifyContractGeneralLink', 'Enlace Modificar el contrato'): continue
    r.click('lnkModifyContractGeneralLink') # Enlace Modificar el contrato

    # Frame TIPO DE MODIFICACION
    #if not esperar('ProcurementContractModificationConfirmCreateTypeModal_iframe', 'Frame TIPO DE MODIFICACION'): continue
    r.frame('ProcurementContractModificationConfirmCreateTypeModal_iframe')
    if not esperar(r, variables, 'btnConfirmGen', 'Campo Número del proceso',frame='ProcurementContractModificationConfirmCreateTypeModal_iframe'): continue
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
        if not esperar(r, variables, '//*[@id="stepCircleSelected_2"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_2'): continue
        r.type('//*[@id="dtmbContractEnd_txt"]', '[clear]' + dfbase.loc[i, 'PRORROGA_FECHA'] + ' 23:59') # Fecha de terminación del contrato


    # Paso 4: 3 Condiciones
    if dfbase.loc[i, 'TIPO_MODIFICACION'] == 'ADICION - PRORROGA':
        print('Paso 4: 3 Condiciones --- proceso', proceso)
        r.click('stepCircle_3') # Información general
        if not esperar(r, variables, '//*[@id="stepCircleSelected_3"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_3'): continue
        r.type('//*[@id="dtmbContractRenewalDateGen_txt"]', '[clear]' + dfbase.loc[i, 'PRORROGA_FECHA'] + ' 23:59') # Fecha de notificación de prorrogación


    # Paso 5: 4 Bienes y Servicios
    print('Paso 5: 4 Bienes y Servicios --- proceso', proceso)
    r.click('stepCircle_4') # Bienes y Servicios
    if not esperar(r, variables, '//*[@id="stepCircleSelected_4"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_4'): continue
    r.click('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[2]/td[4]') # Símbolo + en Incluya el precio como lo indique la Entidad Estatal
    if not esperar(r, variables, '//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', 'Campo Precio unitario'): continue
    r.dclick('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input') # Campo Precio unitario
    r.type('//*[@id="incCatalogueItemsfltDataSheet"]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR TOTAL']) # Campo Precio unitario


    # Paso 6: 7 Informacion presupuestal
    print('Paso 6: 7 Información presupuestal --- proceso', proceso)
    r.click('stepCircle_7') # Informacion presupuestal
    if not esperar(r, variables, '//*[@id="stepCircleSelected_7"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro',stepCircle='stepCircle_7'): continue
    if not esperar(r, variables, 'cbxOwnResourcesAGRIValue', 'Campo Recursos Propios'): continue
    r.dclick('cbxOwnResourcesAGRIValue') # Campo Recursos Propios
    r.type('cbxOwnResourcesAGRIValue', dfbase.loc[i, 'VALOR TOTAL']) # Campo Recursos Propios
    if not esperar(r, variables, '//*[@id="SIIFModal_iframe"]', 'Frame Información presupuestal',boton='btnAddCode'): continue
    #r.click('btnAddCode') # Botón Agregar de CDP/Vigencias Futuras

    # Frame Información presupuestal
    r.frame('SIIFModal_iframe')
    #esperar('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe')
    if not esperar(r, variables, '//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Radio button CDP',frame='SIIFModal_iframe'): continue
    #r.wait(3)
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
    if not esperar(r, variables, '//*[@id="stepCircleSelected_1"][@class="MainColor4 circle22 Black stepOn"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_1'): continue
    r.wait(2)
    r.click('cmAttachmentsOptions') # Lista Anexar documentos
    r.wait(2)
    r.click('linkUploadNew') # Item Anexar nuevo documento
    r.wait(3)
    anexar_documento(r, variables, dfbase.loc[i, "NOMBRE_DOCUMENTO"]) # Popup ANEXAR DOCUMENTO
    
    print('--- Finalizar Modificacion --- proceso', proceso)
    r.click('body')
    if not esperar(r, variables, 'txaModificationPurpose', 'Campo Justificación de la modificación'): continue
    r.type('txaModificationPurpose', 'Modificación') # Campo Justificación de la modificación
    r.click('btnOption_tbContractToolbar_Finish') # Finalizar Modificacion
    if not esperar(r, variables, 'chkCheckBoxAgreeTerms', 'Check box Acepto el valor del contrato'): continue
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
    if not esperar(r, variables, '//*[@id="btnConfirmGen"]', 'Boton Confirmar',frame='StartApprovalSupportModal_iframe'): continue
    r.type('//*[@id="btnConfirmGen"]', 'Yes') # Boton Confirmar
    #r.wait(2)
    #r.type('//*[@id="btnCancelGen"]', 'Yes') # Boton Cancelar
    r.vision('type(Key.SPACE)') # Botón Confirmar
    r.frame()
    r.wait(5)
    #r.wait(20)
    
    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Modificación del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])

# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()