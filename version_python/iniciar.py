"""
Automatizar SECOP II aprobación
Modificación de información contracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
from pandas import read_excel
from datetime import datetime
from funciones import redirigir_log, parametros, iniciar, cerrar, mensaje, esperar, acceder_contrato
from datetime import datetime, timedelta

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()

# Establecer variables de configuración
variables = parametros()
variables['robot'] = 'iniciar_v1'

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
    if not acceder_contrato(r, proceso, variables): continue

    # Paso 1: 1 Información general
    print('Paso 1:  1 Información general --- proceso', proceso)
    if not esperar(r, variables, '//*[@id="stepCircleSelected_1"][@class="MainColor4 circle22 Black stepOn"]', 'Paso 1: 1 Información general', stepCircle='stepCircle_1'): continue
    fecha_str = r.read('dtmbBuyerApprovalDateTimeBox_txt') # Capturar valor Aprobador - Fecha de probación
    fecha_part = fecha_str.split(' (')[0] # Extraer la parte principal de la fecha (antes del paréntesis)
    fecha_original = datetime.strptime(fecha_part, '%d/%m/%Y %I:%M:%S %p') # Convertir a objeto datetime
    fecha_modificada = fecha_original + timedelta(minutes=1) # Agregar 1 minuto
    fecha_inicio = fecha_modificada.strftime('%d/%m/%Y %H:%M') # Formatear en formato de 24 horas
    r.type('dtmbContractStart_txt', '[clear]' + fecha_inicio) # Campo Fecha de inicio de contrato 
    
    # Paso 2: 6 Información presupuestal
    print('Paso 2: 6 Información presupuestal --- proceso', proceso)
    if not esperar(r, variables, '//*[@id="stepCircleSelected_6"][@class="MainColor4 circle22 Black stepOn"]', 'Paso 2: 6 Información presupuestal', stepCircle='stepCircle_6'): continue
    recursos_propios = r.read('cbxOwnResourcesAGRIValue') # Capturar valor Recursos Propios (Alcaldías y Gobernaciones)
    r.click('btnCommitmentAddCode') # Botón Agregar de Compromiso presupuestal de gastos

    # Frame Código del compromiso
    r.frame('SIIFModal_iframe')
    if not esperar(r, variables, '//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Paso 2: Radio button CDP',frame='SIIFModal_iframe'): continue
    r.type('//*[@id="rdbgOptionsToSelectRadioButton_0"]', 'Yes') # Radio button CDP
    r.vision('type(Key.SPACE)') # Radio button CDP
    r.type('txtSIIFCommitmentIntegrationItemTextbox', dfbase.loc[i, 'CODIGO RUBRO']) # Campo Código del compromiso
    r.type('cbxSIIFCommitmentIntegrationItemBalanceTextbox', recursos_propios) # Campo Valor actual compromiso
    r.type('//*[@id="selRelatedBudgetValue"]', 'Yes') # Menú desplegable Código del presupuesto relacionado
    r.vision('type(Key.ENTER)') # Menú desplegable Código del presupuesto relacionado
    r.vision('type(Key.DOWN)') # Bajar una opción
    r.vision('type(Key.ENTER)') # Seleccionar
    r.type('//*[@id="btnSIIFCommitmentIntegrationItemButton"]', 'yes') # Botón Guardar
    r.vision('type(Key.SPACE)') # Botón Guardar
    r.frame()

    # Iniciar ejecución
    r.click('tbToolBarPlaceHolder_btnStartExecution') # Botón Iniciar ejecución
    esperar(r, variables, '//*[@id="stepCircleSelected_1"][@class="MainColor4 circle22 Black stepOn"]', 'Paso 2: Botón Iniciar ejecución')

    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Iniciar Ejecución del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Iniciar Ejecución del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Iniciar Ejecución del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])


# Cerrar sesión
cerrar(r)

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()