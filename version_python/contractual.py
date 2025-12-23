"""
Automatizar SECOP II contractual
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
variables['robot'] = 'contractual_v1'

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


    # Paso 1: 1 Información general
    print('Paso 1: 1 Información general', proceso)
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


    # Paso 2: 2 Condiciones
    print('Paso 2: 2 Condiciones', proceso)
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
    anexar_documento(r, variables, dfbase.loc[i, "NOMBRE_DOCUMENTO_ANEXO"], i) # Popup ANEXAR DOCUMENTO
    
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


    # Paso 3: 3 Bienes y servicios
    print('Paso 3: 3 Bienes y servicios', proceso)
    r.click('stepCircle_3') # stepCircle Bienes y servicios
    if not esperar(r, variables, '//*[@id="stepDiv_3"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_3'): continue
    r.click('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[2]/td[4]') # Bienes y servicios 1+
    r.wait(5)
    r.type('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', '[clear]') # Campo Precio unitario
    #r.type('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input', dfbase.loc[i, 'VALOR CONTRATO']) # Campo Precio unitario
    r.click('//*[@id="incCatalogueItemstblDataSheetContent"]/tbody/tr[2]/td/div[2]/table/tbody/tr[5]/td[5]/table/tbody/tr/td/table/tbody/tr[1]/td[7]/input') # Campo Precio unitario
    r.vision('type(Key.DELETE)')
    r.vision(f'type("{dfbase.loc[i, 'VALOR CONTRATO']}")') # Campo Precio unitario
    

    # Paso 4: 5 Documentos del contrato
    print('Paso 4: 5 Documentos del contrato', proceso)
    r.click('stepCircle_5') # stepCircle Documentos del contrato
    if not esperar(r, variables, '//*[@id="stepDiv_5"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_5'): continue
    # Anexos del contrato
    r.click('btnUploadContractDocument') # Botón Anexar documentos
    r.wait(3)
    anexar_documento(r, variables, dfbase.loc[i, "NOMBRE_DOCUMENTO"], i) # Popup ANEXAR DOCUMENTO
    
    
    # Paso 5: Confirmar
    print('Paso 5: Confirmar', proceso)
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
    print('Terminada Contractual del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Contractual del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Contractual del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])

# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()