"""
Descrgar Historias DGH
Descargar de DGH las Historias listadas en el archivo "Base_descargas_HC.xlsx"
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
variables['robot'] = 'historias_v1'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

# Iniciar robot
print('Iniciar robot', variables['robot'])
r.init(visual_automation = True, turbo_mode=False)
#r.timeout(10)

# Iniciar sesion
# Función para iniciar sesión en DGH
# Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
r.ask('Iniciar')

# Arbrir INICIAR SESION
r.url('https://clinico.subrednorte.gov.co:50434/')

# Iniciar sesion
esperar(r, variables, '//app-seguridad-inicio//fa-icon', 'Botón Ingresar')
r.type('//div[1]/dx-text-box//input', '[clear]' + variables['user']) # Celda 'Usuario'
r.type('//div[2]/dx-text-box//input', '[clear]' + variables['password']) # Celda 'Contraseña'
r.click('//app-seguridad-inicio//fa-icon') # Botón 'Ingresar'

# Información asistencial
esperar(r, variables, '//app-seguridad-asistencial//fa-icon', 'Botón Continuar')
r.click('//dx-radio-group/div/div[1]') # Radio Button 'Consulta externa'
r.type('//dx-text-box//input', '[clear]23') # Celda 'Buscar Centro de Atención'
r.click('body')
r.click('//app-seguridad-asistencial//fa-icon') # Botón 'Continuar'

# Acceso Consultar Histórico de Paciente
esperar(r, variables, 'dx-select-box', 'Lista desplegable Seleccione un módulo')
r.click('dx-select-box')
r.click('HISTORIAS CLÍNICAS')
r.click('Consultar Histórico de Paciente')

    
# %% Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=0
    proceso = dfbase.loc[i, 'Var5 Tipo de Identificación del usuario'] + dfbase.loc[i, 'Var6 Número de Identificación del usuario']

    # Deshacer
    r.click('Deshacer (Ctrl + F3)')
    r.wait(2)


    # Paso 0: Buscar la historia
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Buscar la historia  --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Buscar la historia --- proceso '+proceso+' - '+horainicio)
    #if not acceder_contrato(r, proceso, variables): continue
    r.type('//dx-text-box//input', '[clear]' + dfbase.loc[i, 'Var6 Número de Identificación del usuario']) # Campo 'Buscar'
    r.click('body')
    r.wait(2)
    # Validar si se encontró la historia
    if r.read('dghc-sh-info-paciente') == '- \xa0 \xa0 \xa0 \xa0 -':
        print('No se encontró', proceso)
        # Agregar una fila con la falla en los archivos historico.csv y seguimiento.csv
        mensaje(variables, 'No se encuentra ' + proceso)
        continue

    # Validar que existan registros
    if r.read('//div[7]//tr[1]/td[4]') == '':
        print('No se encontraron registros', proceso)
        # Agregar una fila con la falla en los archivos historico.csv y seguimiento.csv
        mensaje(variables, 'No se encuentran registros ' + proceso)
        continue

    # Filtrar fecha
    r.click('//td[3]/div/div[1]') # Celda 'Buscar Fecha'
    r.click('//div[4]//li[7]') # Lista desplegable 'Entre'
    r.type('//div[4]/div/div[1]/div/div//div[1]/input', '[clear]07/01/2025') # Campo Inicio
    r.type('//div[4]/div/div[2]/div/div//div[1]/input', '[clear]06/30/2026') # Campo Fin
    r.click('body')

    # Validar que existan registros
    if r.read('//div[7]//tr[1]/td[4]') == '':
        # Filtrar por fecha anterior
        r.click('//td[3]/div[1]/div[2]') # Celda 'Buscar Fecha Entre'
        r.type('//div[4]/div/div[1]/div/div//div[1]/input', '[clear]07/01/2024') # Campo Inicio
        r.click('body')
    
        # Validar que existan registros
        if r.read('//div[7]//tr[1]/td[4]') == '':
            print('No se encontraron registros', proceso)
            # Agregar una fila con la falla en los archivos historico.csv y seguimiento.csv
            mensaje(variables, 'No se encuentran registros ' + proceso)
            continue


    # Paso 1: Imprimir Folios
    print('Paso 1:  Imprimir Folios --- proceso', proceso)
    r.click('//div[6]//dx-button') # Botón 'Imprimir Folios'
    
    if not esperar(r, variables, '//div[1]/div/dx-button/div/span', 'Botón Todos'): continue
    r.click('//div[1]/div/dx-button/div/span') # Botón 'Todos
    r.click('//div[2]/div/div/div/div/div/div/div/div/div/div/div[1]//span') # Botón 'Generar Reporte'
    
    r.click('//li//dxa-image-template') # Lista desplegable 'Exportar a'
    r.click('//div[@aria-label="PDF"]') # Opción 'PDF'

//div[@title="Buscar" and @aria-expanded="false"]
//div[@title='Grupo Siguiente' and contains(@class,'dx-state-disabled')]


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