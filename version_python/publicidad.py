"""
Automatizar SECOP II ejecución
Anexar documento de publicidad en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

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
variables['robot'] = 'publicidad_v1'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

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
    if not acceder_contrato(r, proceso, variables, contratos=0): continue
    
    # Paso 1: Crear modificación
    print('Paso 1: Crear modificación', proceso)
    if not esperar(r, variables, 'ctxm_cmViewOptionsContextMenu', 'Paso 1: Lista Opciones'): continue
    r.click('ctxm_cmViewOptionsContextMenu') # Lista Opciones
    r.wait(2)   
    r.click('//*[@id="createAmendmentLink"]') # Opción Crear Modificación
    
    # Paso 2: 1 Modificación
    print('Paso 2: 1 Modificación', proceso)
    if not esperar(r, variables, '//*[@id="txaAmendmentPurpose1Gen"]', 'Paso 2: Campo Justificación de la modificación'): continue
    r.type('//*[@id="txaAmendmentPurpose1Gen"]', '[clear]Modificación') # Campo Justificación de la modificación

    # Paso 3: 4 Documentos del Proceso
    print('Paso 3: 4 Documentos del Proceso', proceso)
    r.click('stepCircle_4') # stepCircle Documentos del Proceso
    if not esperar(r, variables, '//*[@id="stepDiv_4"][@class="LeftMenuButtonOn Black"]', 'stepCircle Configuracion en negro', stepCircle='stepCircle_4'): continue
    #r.wait(5)
    r.click('incContractDocumentsbtnUploadDocumentGen') # Botón Anexar documento
    r.wait(5)
  
    # Popup ANEXAR DOCUMENTO
    if not anexar_documento(r, variables, dfbase, i): continue
    r.wait(5)

    r.click('btnOption_trRowToolbarTop_tdCell1_tbToolBar_Finish') # Botón Realizar modificación
    if not esperar(r, variables, '//*[@id="btnApplyAmendment"]', 'Paso 3: Botón Aplicar la modificación'): continue
    r.click('//*[@id="btnApplyAmendment"]') # Botón Aplicar la modificación
    r.wait(5)

    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Publicidad del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Publicidad del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Publicidad del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])


# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()