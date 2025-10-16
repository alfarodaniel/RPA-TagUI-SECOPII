"""
Automatizar SECOP II aprobación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
from pandas import read_excel
from datetime import datetime
from funciones import redirigir_log, parametros, iniciar, cerrar, mensaje, esperar, acceder_contrato

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()

# Establecer variables de configuración
variables = parametros()
variables['robot'] = 'aprobacion_v4'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

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
    if not acceder_contrato(r, proceso, variables): continue

    r.wait(10)
    # Validar tipo de aprobación
    if r.present('IncTaskApproval_btnApproveButton'):
        # Paso 1: Aprobación por Proceso pendiente de aprobación/apertura
        print('Paso 1: Aprobación por Proceso pendiente de aprobación/apertura --- proceso', proceso)
        r.click('IncTaskApproval_btnApproveButton') # Boton Aprobar
        r.wait(10)
        # Validar si aparece el botón Aceptar nuevamente
        if r.present('IncTaskApproval_btnApproveButton'):
            r.click('IncTaskApproval_btnApproveButton') # Boton Aprobar
            r.wait(5)
        # Validar si aparece el botón Enviar al proveedor
        if r.present('btnSendToSupplier'):
            r.click('btnSendToSupplier') # Enviar al proveedor
            r.wait(5)
    elif r.present('btnOption_tbContractToolbar_FinishAfterAcknowledge'):
        # Paso 1: Aprobación por Enviar para aprobación
        print('Paso 1: Aprobación por Enviar para aprobación --- proceso', proceso)
        r.click('btnOption_tbContractToolbar_FinishAfterAcknowledge') # Boton Enviar para aprobación

        # Frame Confirmar
        r.frame('StartApprovalSupportModal_iframe')
        if not esperar(r, variables, 'btnConfirmGen', 'Botón confirmar',frame='StartApprovalSupportModal_iframe'): continue
        r.click('btnConfirmGen')
        r.wait(5)
        r.frame()

        r.click('IncTaskApproval_btnApproveButton') # Boton Aprobar
        r.wait(5)
        
    else:
        print('Aprobación por Modificaciones del Contrato')
        # Paso 1: 8 Modificaciones del Contrato
        print('Paso 1: 8 Modificaciones del Contrato --- proceso', proceso)
        if not esperar(r, variables, '//*[@id="lnk_stpmStepManager9"]', 'Paso 1: Menú 8 Modificacione del Contrato'): continue
        r.click('//*[@id="lnk_stpmStepManager9"]') # Menú 8 Modificacione del Contrato
        if not esperar(r, variables, 'lnkEditLink_0', 'Paso 1: Enlace Editar'): continue
        r.click('lnkEditLink_0') # Enlace Editar

        # Paso 2: 1 Modificación del Contrato
        print('Paso 2: 1 Modificación del Contrato --- proceso', proceso)
        if not esperar(r, variables, 'IncTaskApproval_btnApproveButton', 'Paso 2: Boton Aprobar'): continue
        r.click('IncTaskApproval_btnApproveButton') # Boton Aprobar
        r.wait(5)
        esperar(r, variables, '//*[@id="stepCircleSelected_1"]', 'Paso 2: Menú 1 Modificación del contrato') # Menú 1 Modificación del contrato
        # Verificar si se encuentra el botón de Publicar modificación
        if r.present("//input[@id='btnFinishModification' and @value='Publicar modificación']"):
            print('Paso 2: Publicar modificación --- proceso', proceso)
            r.click('//*[@id="btnFinishModification"]')
            r.wait(5)
            # Validar si aparace mensaje informativo
            if not r.present("//div[@id='mdboxInvoicesWarningDialog' and contains(@style,'display: none')]"):
                r.click('btnConfirmInvoicesWarningButton')
                r.wait(5)
            esperar(r, variables, '//*[@id="stepCircleSelected_1"]', 'Paso 2: Menú 1 Modificación del contrato') # Menú 1 Modificación del contrato

    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Aprobación del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Aprobación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Aprobación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])


# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()