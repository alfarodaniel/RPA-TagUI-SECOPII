"""
Automatizar SECOP II modificación
Funciones
"""
# Cargar librerías
from datetime import datetime

# Funcion mensaje para agregar una fila con el mensaje en los archivos historico.csv
def mensaje(variables, mensaje, repositorio=''):
    #for archivo in [f"{repositorio}historico.csv", seguimiento.csv]:
    with open(f"{repositorio}historico.csv", 'a') as file:
        file.write(variables['robot'] + ',' + 
                variables['usuario'] + ',' + 
                variables['contrato'] + ',' + 
                datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ',' + 
                mensaje + '\n')


# Función para validar si existe el registro en la página
def esperar(r, variables, objeto, descripcion="", boton="", frame="", stepCircle="", popup=""):
    for i in range(1, variables['espera']):
        # Verificar si depende de un boton
        if boton != "":
            r.click(boton)
        # Verificar si se encuentra en un frame
        if frame != "":
            r.frame(frame)
        # Verificar si se encuentra en un stepCircle
        if stepCircle != "":
            r.click(stepCircle)
        # Verificar si se encuentra en un popup
        if popup != "":
            r.popup(popup)
        # Verifica si se encontró objeto
        if r.present(objeto):
            #print('si')
            return True  # Retorna True si el objeto fue encontrado
        
        print(i, 'esperando', descripcion)
        # Esperar 1 segundo
        r.wait(1)

    print('No se encontró', descripcion)
    # Agregar una fila con la falla en los archivos historico.csv y seguimiento.csv
    mensaje(variables, 'No se encuentra ' + descripcion)

    return False  # Retorna False si el objeto no fue encontrado


# Función para iniciar sesión en SECOP II
def iniciar(r, variables):
    # Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
    r.ask('Iniciar')

    # Arbrir INICIAR SESION
    r.url('https://community.secop.gov.co/')

    # Iniciar sesion
    esperar(r, variables, '//*[@id="btnLoginButton"]', 'Botón Iniciar Sesión')
    r.type('//*[@id="txtUserName"]', '[clear]' + variables['user']) # Celda 'Nombre de usuario'
    r.type('//*[@id="txtPassword"]', '[clear]' + variables['password']) # Celda 'Contraseña'
    r.click('//*[@id="btnLoginButton"]') # Botón 'Iniciar Sesión'
    
    # Validar si aparace mensaje informativo
    r.wait(5)
    if r.present('btnAcknowledgeGen'):
        r.click('btnAcknowledgeGen')


# Función para cerrar sesión en SECOP II
def cerrar(r):
    # Cerrar sesión
    print('Cerrar sesion')
    r.click('//*[@id="userImage"]/img') # Imagen usuario
    r.wait(1)
    r.click('//*[@id="logOut"]') # Opción Salir