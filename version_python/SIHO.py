"""
Automatizar registro del Sistema de Información Hospitalario - SIHO
Modificación de información en SIHO de la información listada en el archivo "SIHO.xlsx"
"""

# Cargar librerías
import rpa as r
from pandas import read_csv, read_excel, to_datetime
from datetime import datetime
from funciones import redirigir_log, mensaje, esperar

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
#print("Directorio actual:", os.getcwd())


# Función para Cargar página principal
def inicio(variables):
    # Cargar página principal
    r.wait(2)
    r.url(variables['pagina'])
    esperar(r, variables, '_ctl0_iGrabarFtr', 'Botón Grabar') # Botón Grabar


# Función para buscar un contrato
def buscar(variables, proceso):
    inicio(variables) # Cargar página principal
     
    # Buscar el contrato
    if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_dgContratacion"]//td[normalize-space(.)="'+ proceso +'"]', 'Paso 0: Buscar número contrato'): return False # Validar si existe el contrato
    r.click('//table[@id="_ctl0_ContentPlaceHolder1_dgContratacion"]//td[normalize-space(.)="'+ proceso +'"]/..//a') # Enlace del contrato
    esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_tabForm" and @style="DISPLAY: block;"]', 'Cuadro Código único de registro') # Cuadro Código único de registro
    
    return True  # Retorna True si el contrato fue encontrado


# Función para validar si Resposable del diligenciamiento está vacio
def validar_responsable():
    if r.read('_ctl0_ContentPlaceHolder1_tbRcedula') == '':
        r.type('_ctl0_ContentPlaceHolder1_tbRcedula', '[clear]74150723') # Campo Cédula
        r.type('_ctl0_ContentPlaceHolder1_tbRnombre', '[clear]Jose Mauricio Robayo Pulido') # Campo Nombre
        r.select('//*[@id="_ctl0_ContentPlaceHolder1_ddRcarg_codigo"]', 'Técnico Administrativo') # Selección Cargo
        r.type('_ctl0_ContentPlaceHolder1_tbRtelefono', '[clear]30346379870') # Campo Teléfono


# Función para validar si el campo tiene decimales
def validar_decimales(campo):
    valor = r.read(campo).strip()
    if valor.endswith(',00'):  # Termina en ',00'
        r.type(campo, '[clear]' + valor[:-3])  # Eliminar ',00'


# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()


# Establecer variables de configuración
# Leer el archivo 'parametros.csv'
print('\nCargando parametrosSIHO.csv')
df = read_csv('parametrosSIHO.csv')

# Convertir el DataFrame en un diccionario params
params = dict(zip(df['parametro'], df['valor']))

# Establecer variables de configuración
variables = {
    'repositorio': params.get('repositorio'),
    'espera': int(params.get('espera')),
    'usuario': params.get('usuario'),
    'trimestre': params.get('trimestre'),
    'user': params.get('user'),
    'password': params.get('password'),
    'base': params.get('base'),
    'contrato': ''
}
variables['robot'] = 'SIHO_v1'
variables['pagina'] = 'https://prestadores.minsalud.gov.co/SIHO/formularios/contratacion.aspx?pageTitle=Contrataci%f3n%20desde%202020&pageHlp=/SIHO/ayudas/formularios/contrataciones.pdf&periodo=TRIMESTRAL'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)
dfbase['fecha_inicio_contrato'] = to_datetime(dfbase['fecha_inicio_contrato']).dt.strftime('%Y/%m/%d') # Formatear fecha
dfbase['fecha_final_contrato'] = to_datetime(dfbase['fecha_final_contrato']).dt.strftime('%Y/%m/%d') # Formatear fecha

# Iniciar robot
print('Iniciar robot', variables['robot'])
#r.init(visual_automation = True, turbo_mode=False)
r.init(visual_automation = False, turbo_mode=True)
#r.timeout(10)

# Iniciar sesion
# Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
#r.ask('Iniciar')

# Abrir INICIAR SESION
r.url('https://prestadores.minsalud.gov.co/siho/')

# Iniciar sesion
# Frame Area de trabajo
r.frame('areawork')
esperar(r, variables, '//*[@id="btnIngresar"]', 'Iniciar sesión',frame='areawork') # Botón 'Iniciar Sesión'
r.type('//*[@id="tbid_usuario"]', '[clear]' + variables['user']) # Celda 'Usuario'
r.type('//*[@id="tbcontrasena"]', '[clear]' + variables['password']) # Celda 'Contraseña'
r.click('//*[@id="btnIngresar"]') # Botón 'Iniciar Sesión'
r.frame()
r.wait(2)

# Abrir PARÁMETROS DE CAPTURA DE FORMULARIOS
r.url('https://prestadores.minsalud.gov.co/siho/formularios/contratacionperiodo.aspx?pageTitle=Contrataci%F3n%20desde%202020&pageHlp=/SIHO/ayudas/formularios/contrataciones.pdf')

# Selección Periodo
esperar(r, variables, '//*[@id="ddmes"]', 'Selección Periodo') # Selección Periodo
r.select('//*[@id="ddmes"]', variables['trimestre']) # Selección Trimestre
r.click('//*[@id="ibAceptar"]') # Botón Aceptar


# Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=0
    proceso = dfbase.loc[i, 'numero_contrato']

    # Paso 0: Acceder al contrato
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Acceder al contrato --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Acceder al contrato --- proceso '+proceso+' - '+horainicio)
        
    # Paso 1: Modificar información del contrato
    print('Paso 1: Modificar información del contrato --- proceso', proceso)
    
    if dfbase.loc[i, 'observación'] in ['Actualizar solo fecha', 'Actualizar solo valor', 'Actualizar valor y fecha']:
        if not buscar(variables, proceso): continue  # Buscar el contrato
        if dfbase.loc[i, 'observación'] == 'Actualizar solo fecha':
            print('Paso 1: Actualizar solo fecha')
        else:
            print('Paso 1:', dfbase.loc[i, 'observación'])
            r.type('_ctl0_ContentPlaceHolder1_tbvalor_total_contrato', '[clear]' + dfbase.loc[i, 'valor_total_contrato']) # Campo Valor Total del Contrato
            validar_responsable() # Validar si Resposable del diligenciamiento está vacio
            r.click('_ctl0_ibGrabarFtr') # Botón Grabar
            if not buscar(variables, proceso): continue # Buscar el contrato
    
        r.click('_ctl0_ContentPlaceHolder1_btnpersonas') # Enlace "Ir a Registro de Personas para este Contrato"
        if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_dgContratacionpersonas"]//td[normalize-space(.)="'+ proceso +'"]'): continue # Tabla Registro de Personas para este Contrato
        r.dom('document.evaluate(\'(//table[@id="_ctl0_ContentPlaceHolder1_dgContratacionpersonas"]//td[normalize-space(.)="' + proceso + '"]/..//a)[1]\', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()') # Celda "Nombres Completos"
        if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_tabForm" and @style="DISPLAY: block;"]', 'Cuadro Registro de Personas para este Contrato'): continue # Cuadro Registro de Personas para este Contrato
        
        if dfbase.loc[i, 'observación'] in ['Actualizar solo fecha', 'Actualizar valor y fecha']:
            r.type('_ctl0_ContentPlaceHolder1_tbfecha_final_contrato', '[clear]' + dfbase.loc[i, 'fecha_final_contrato']) # Campo Fecha terminación contrato
        
        if dfbase.loc[i, 'observación'] in ['Actualizar solo valor', 'Actualizar valor y fecha']:
            r.type('_ctl0_ContentPlaceHolder1_tbvalor_total_contrato', '[clear]' + dfbase.loc[i, 'valor_total_contrato']) # Campo Valor total contrato
        
        validar_decimales('_ctl0_ContentPlaceHolder1_tbvalor_total_contrato') # Validar decimales Campo Valor total contrato
        validar_decimales('_ctl0_ContentPlaceHolder1_tbvalor_honorarios_mes') # Validar decimales Campo Valor honorarios mes
        validar_responsable() # Validar si Resposable del diligenciamiento está vacio
        r.click('_ctl0_ibGrabarFtr') # Botón Grabar
    
    if dfbase.loc[i, 'observación'] == 'Eliminar':
        if not buscar(variables, proceso): continue  # Buscar el contrato
        print('Paso 1: Eliminar')
        r.click('_ctl0_ContentPlaceHolder1_btnpersonas') # Enlace "Ir a Registro de Personas para este Contrato"
        if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_dgContratacionpersonas"]//td[normalize-space(.)="'+ proceso +'"]'): continue # Tabla Registro de Personas para este Contrato
        r.dom('document.evaluate(\'(//table[@id="_ctl0_ContentPlaceHolder1_dgContratacionpersonas"]//td[normalize-space(.)="' + proceso + '"]/..//a)[1]\', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.click()') # Celda "Nombres Completos"
        if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_tabForm" and @style="DISPLAY: block;"]', 'Cuadro Registro de Personas para este Contrato'): continue # Cuadro Registro de Personas para este Contrato
        r.click('_ctl0_ibEliminarFtr') # Botón Eliminar Persona
        if not buscar(variables, proceso): continue # Buscar el contrato
        r.click('_ctl0_ibEliminarFtr') # Botón Eliminar Contrato
    
    if dfbase.loc[i, 'observación'] == 'Nuevo contrato':
        inicio(variables) # Cargar página principal
        print('Paso 1: Nuevo contrato')
        r.click('_ctl0_ibNuevoHdr') # Botón Nuevo
        if not esperar(r, variables, '_ctl0_ibGrabarFtr', 'Botón Grabar'): continue # Botón Grabar
        r.type('_ctl0_ContentPlaceHolder1_tbnumero_contrato', '[clear]' + dfbase.loc[i, 'numero_contrato']) # Campo Número de Contrato
        r.select('//*[@id="_ctl0_ContentPlaceHolder1_ddcodigo_tipo_persona"]', dfbase.loc[i, 'tipo_persona']) # Selección Tipo persona
        r.wait(4)
        r.type('_ctl0_ContentPlaceHolder1_tbnumero_personas_contratadas', '1') # Campo Número de PErsonas Contratadas
        r.type('_ctl0_ContentPlaceHolder1_tbvalor_total_contrato', '[clear]' + dfbase.loc[i, 'valor_total_contrato']) # Campo Valor Total del Contrato
        r.click('_ctl0_ibGrabarFtr') # Botón Grabar
        if not buscar(variables, proceso): continue # Buscar el contrato
        r.click('_ctl0_ContentPlaceHolder1_btnpersonas') # Enlace "Ir a Registro de Personas para este Contrato"
        if not esperar(r, variables, '_ctl0_ContentPlaceHolder1_dgContratacionpersonas'): continue # Tabla Registro de Personas para este Contrato
        r.click('_ctl0_ibNuevoHdr') # Botón Grabar
        if not esperar(r, variables, '//table[@id="_ctl0_ContentPlaceHolder1_tabForm" and @style="DISPLAY: block;"]', 'Cuadro Registro de Personas para este Contrato'): continue # Cuadro Registro de Personas para este Contrato
        r.type('_ctl0_ContentPlaceHolder1_tbnombre_apellidos', '[clear]' + dfbase.loc[i, 'nombres_apellidos']) # Campo Nombres y apellidos
        r.type('_ctl0_ContentPlaceHolder1_tbnumero_cedula', '[clear]' + dfbase.loc[i, 'numero_cedula']) # Campo Número de cédula
        r.select('//*[@id="_ctl0_ContentPlaceHolder1_ddcodigo_actividad"]', dfbase.loc[i, 'actividad']) # Selección Actividad a Desarrollar
        r.wait(1)
        r.select('//*[@id="_ctl0_ContentPlaceHolder1_ddcodigo_perfil"]', dfbase.loc[i, 'perfil']) # Selección Perfil de la Actividad
        r.wait(1)
        r.type('_ctl0_ContentPlaceHolder1_tbvalor_total_contrato', '[clear]' + dfbase.loc[i, 'valor_total_contrato']) # Campo Valor total contrato
        r.select('//*[@id="_ctl0_ContentPlaceHolder1_ddcodigo_frecuencia_pago"]', 'Mensual') # Selección Frecuencia de pago
        r.type('_ctl0_ContentPlaceHolder1_tbvalor_honorarios_mes', '[clear]' + dfbase.loc[i, 'valor_honorarios_mes']) # Campo Valor honorarios mes
        r.type('_ctl0_ContentPlaceHolder1_tbfecha_inicio_contrato', '[clear]' + dfbase.loc[i, 'fecha_inicio_contrato']) # Campo Fecha terminación contrato
        r.type('_ctl0_ContentPlaceHolder1_tbfecha_final_contrato', '[clear]' + dfbase.loc[i, 'fecha_final_contrato']) # Campo Fecha terminación contrato
        r.click('_ctl0_ibGrabarFtr') # Botón Grabar
    
    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Modificación del Contrato --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Modificación del Contrato --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])

          
# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()