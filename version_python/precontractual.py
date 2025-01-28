"""
Automatizar SECOP II precontractual
Cargue de información pecontracual en SECOP II de la información listada en el archivo "Base_de_datos_Contratacion.csv"
"""

# Cargar librerías
import rpa as r
import pandas as pd

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python")
#print("Directorio actual:", os.getcwd())

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
base = pd.read_excel('base_de_datos_Contratacion.xlsx', dtype=str)

# Mostrar el valor de la columna 'NUMERO' para el primer registro
#primer_registro_numero = base.loc[0, 'NUMERO']
#print('Valor de la columna NUMERO para el primer registro:', primer_registro_numero)

# Función para validar si existe el registro en la página
def esperar(objeto, descripcion="", frame=""):
    for i in range(1, 10):
        # Verificar si se encuentra en un frame
        if frame != "":
            r.frame(frame)
        # Verifica si se encontró objeto
        if r.present(objeto):
            #print('si')
            break        # Verifica si se encontró
        
        print(i, 'esperando', descripcion)
        # Esperar 1 segundo
        r.wait(1)

# Iniciar robot
print('Iniciar robot')
r.init(visual_automation = True, turbo_mode=False)
#r.timeout(10)

# Arbrir INICIAR SESION
r.url('https://community.secop.gov.co/')

# Cargar credenciales
Credenciales = r.load('credenciales.txt').splitlines()

# Iniciar sesion
esperar('//*[@id="btnLoginButton"]', 'Botón Iniciar Sesión')
# Celda 'Nombre de usuario'
r.type('//*[@id="txtUserName"]', '[clear]' + Credenciales[0])
# Celda 'Contraseña'
r.type('//*[@id="txtPassword"]', '[clear]' + Credenciales[1])
# Botón 'Iniciar Sesión'
r.click('//*[@id="btnLoginButton"]')
#esperar('??????????????????????????????????????????????', 'Nombre de usuario')
#r.click('//*[@id="btnButton1"]')

# Validar si aparace mensaje informativo
if r.present('btnAcknowledgeGen'):
    r.click('btnAcknowledgeGen')

# Recorrer la base de datos
for i in range(0, len(base)):
    i = 0

    # Cargar página principal
    r.url('https://community.secop.gov.co/')

    # Crear proceso
    # Menú Procesos
    esperar('//*[@value="Procesos"]', 'Botón Procesos')
    r.click('//*[@value="Procesos"]') # Menú Procesos
    esperar('//*[@id="lnkSubItem9"]', 'Subnemú Tipos de procesos')
    r.click('//*[@id="lnkSubItem9"]') # Submenú Tipos de procesos

    # En la página de Tipos de procesos
    esperar('//*[@id="btnCreateProcedureButton12"]', 'Botón Crear Contratación régimen especial')
    r.click('//*[@id="btnCreateProcedureButton12"]') # Botón Crear Contratación régimen especial

    # En el frame CREAR PROCESO
    esperar('CreateProcedure_iframe', 'Frame CREAR PROCESO')
    r.frame('CreateProcedure_iframe')
    esperar('txtProcedureReference', 'Campo Número del proceso','CreateProcedure_iframe')
    r.type('txtProcedureReference', '[clear]' 
           + base.loc[i, 'TIPOLOGIA'] + '-' 
           + base.loc[i, 'NUMERO DE CONTRATO'] + '-' 
           + base.loc[i, 'VIGENCIA']) # Número de proceso
    r.type('txtProcedureName', '[clear]' 
           + 'PRESTAR SERVICIOS PROFESIONALES Y APOYO A LA GESTION') # Nombre
    r.type('txtBusinessOperationText', '[clear]' 
           + 'DIRECCIÓN DE CONTRATACIÓN') # Unidad de contratación
    r.vision('type(" - COMPRAS")')
    r.vision('type(Key.DOWN)')
    r.vision('type(Key.ENTER)')
    r.type('btnSaveCurrentDossierTop', '[enter]') # Botón Confirmar
    r.frame()



"""
# Acceder a Busqueda
esperar('//*[@id="SessionBarWidget"]/div[5]/div[1]/div[1]/div/div')
r.click('//*[@id="SessionBarWidget"]/div[5]/div[1]/div[1]/div/div')
esperar('//*[@id="lnkSubItem4"]')
r.click('//*[@id="lnkSubItem4"]')

# Prueba de busqueda
esperar('//*[@id="btnSearchButton"]') # Botón Buscar
r.type('//*[@id="txtProcedureDataAdvancedSearch"]', '[clear]' + 'SA-01-2024')
r.type('//*[@id="txtRequestReference"]', '[clear]' + 'SA-01-2024')
r.type('//*[@id="txtRequestName"]', '[clear]' + 'MANTENIMIENTO Y ADECUACION CUBIERTA')
r.click('//*[@id="btnSearchButton"]')
r.hover('VORTAL')
r.wait(3)

# Cerrar sesion
r.click('//*[@id="userImage"]/img')
r.click('//*[@id="logOut"]/span')



# Acceder a subasta
# //*[@id="mainGrid"]/tbody/tr/td[10]/div/div/input
esperar('//*[@id="mainGrid"]/tbody/tr/td/div/div/input', 'Campo próximo lance') # Campo Próximo lance
r.type('//*[@id="mainGrid"]/tbody/tr/td/div/div/input"]', '[clear]' + '0') # valor
r.click('//*[@id="tableBids"]/div[1]/button') # Botón 'Presentar Lance'
r.click('Presentar')

"""

# Cerrar robot
print('Cerrar robot')
r.close()