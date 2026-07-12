"""
Descargar Laboratorios
Descargar Laboratorios las Historias listadas en el archivo "Base_descargas_HC.xlsx"
"""

# %% Cargar datos
# Cargar librerías
import rpa as r
from pandas import read_excel
from datetime import datetime
from funciones import redirigir_log, parametros, mensaje, esperar
import shutil
import os
import glob
import sys
from PyPDF2 import PdfMerger
#pip install PyPDF2
import subprocess

#import os
#os.chdir("C:\\Mis Documentos\\trabajos\\contratacion\\robots\\RPA-TagUI-SECOPII\\version_python\\")
#print("Directorio actual:", os.getcwd())

# Redirigir la salida estándar y de error a un archivo log.txt
redirigir_log()

# Establecer variables de configuración
variables = parametros()
variables['robot'] = 'laboratorio_v1'

# Cargar base de datos de contratación "base_de_datos_Contratacion.xlsx" en solo texto
dfbase = read_excel(variables['base'], dtype=str)

# Iniciar robot
print('Iniciar robot', variables['robot'])
r.init(visual_automation = False, turbo_mode=False, headless_mode=True)

# Carpeta donde está el ejecutable (o el .py)
if getattr(sys, 'frozen', False):
    CARPETA_BASE = os.path.dirname(sys.executable)
else:
    CARPETA_BASE = os.path.dirname(os.path.abspath(__file__))

# Iniciar sesion
# Función para iniciar sesión en DGH
# Muestra un mensaje ara garantizar que se de click dentro de la ventana para que funcione r.vision
#r.ask('Iniciar')

# Arbrir INICIAR SESION
r.url('https://infolab.subrednorte.gov.co/REDNORTE/login.aspx')

# Iniciar sesion
esperar(r, variables, '//*[@id="SECTION4"]/input', 'Botón Iniciar Sesión')
r.type('//*[@id="vUSUARIOLOGIN"]', '[clear]' + variables['user']) # Celda 'Usuario'
r.type('//*[@id="vUSUARIOPASSWORD"]', '[clear]' + variables['password']) # Celda 'Contraseña'
r.select('//*[@id="vUSUARIOLABORATORIO"]', 'BUENAVISTA') # Lista despelgable 'Seleccionar laboratorio'
r.click('//*[@id="SECTION4"]/input') # Botón 'Iniciar Sesión'


def _buscar_ghostscript():
    """
    Retorna la ruta del ejecutable gswin64c.exe o gswin32c.exe.
    Lanza FileNotFoundError si no lo encuentra.
    """

    # Buscar primero en el PATH
    for exe in ("gswin64c.exe", "gswin32c.exe"):
        ruta = shutil.which(exe)
        if ruta:
            return ruta

    # Buscar en Program Files
    carpetas = [
        r"C:\Program Files\gs",
        r"C:\Program Files (x86)\gs"
    ]

    ejecutables = []

    for carpeta in carpetas:
        if os.path.exists(carpeta):
            ejecutables.extend(glob.glob(os.path.join(carpeta, "*", "bin", "gswin64c.exe")))
            ejecutables.extend(glob.glob(os.path.join(carpeta, "*", "bin", "gswin32c.exe")))

    if ejecutables:
        # Toma la versión más reciente
        ejecutables.sort(reverse=True)
        return ejecutables[0]

    raise FileNotFoundError(
        "Ghostscript no está instalado o no fue encontrado."
    )


def comprimir_pdf(pdf,
                   calidad="ebook",
                   reemplazar=True):
    """
    Comprime un PDF utilizando Ghostscript.

    calidad:
        screen   -> máxima compresión
        ebook    -> recomendada
        printer  -> alta calidad
        prepress -> muy alta calidad
        default  -> por defecto

    reemplazar=True
        Reemplaza el PDF original.
    """

    if not os.path.exists(pdf):
        raise FileNotFoundError(pdf)

    gs = _buscar_ghostscript()

    carpeta = os.path.dirname(pdf)
    
    temporal = os.path.join(carpeta, "__tmp_comp.pdf")

    comando = [
        gs,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{calidad}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        "-dJPEGQ=10",
        "-dDetectDuplicateImages=true",
        "-dPrinted=false",
        "-dOptimize=true",
        "-dColorImageDownsampleType=/Bicubic",
        "-dColorImageResolution=76",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dGrayImageResolution=76",
        "-dMonoImageDownsampleType=/Bicubic",
        "-dMonoImageResolution=76",
        "-dDownsampleColorImages=true",
        "-dDownsampleGrayImages=true",
        "-dDownsampleMonoImages=true",
        "-dAutoFilterColorImages=false",
        "-dAutoFilterGrayImages=false",
        "-dColorImageFilter=/DCTEncode",      # JPEG para color
        "-dGrayImageFilter=/DCTEncode",       # JPEG para grises
        "-dMonoImageFilter=/CCITTFaxEncode",  # Fax para B/N
        "-dCompressFonts=true",
        "-dSubsetFonts=true",
        "-dEmbedAllFonts=false",
        f"-sOutputFile={temporal}",
        pdf
    ]

    subprocess.run(comando, check=True)

    if reemplazar:
        os.remove(pdf)
        os.rename(temporal, pdf)
        return pdf
    else:
        salida = pdf.replace(".pdf", "_comprimido.pdf")
        os.rename(temporal, salida)
        return salida

    
# %% Recorrer la base de datos
for i in range(0, len(dfbase)):
    # Variables
    #i=0
    proceso = dfbase.loc[i, 'Var5 Tipo de Identificación del usuario'] + dfbase.loc[i, 'Var6 Número de Identificación del usuario']

    # Iniciar consulta
    r.url('https://infolab.subrednorte.gov.co/REDNORTE/consultaexternaresultados.aspx')
    r.wait(5)
    esperar(r, variables, '//*[@id="vFLTPACIENTEEXPEDIENTE"]', 'Campo Historia')
    

    # Paso 0: Buscar la historia
    horainicio = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Paso 0: Buscar la historia  --- proceso',proceso,'-',horainicio,'--------------------------------------------------')
    mensaje(variables, 'Paso 0: Buscar la historia --- proceso '+proceso+' - '+horainicio)
    #if not acceder_contrato(r, proceso, variables): continue
    r.type('//*[@id="vFLTPACIENTEEXPEDIENTE"]', '[clear]' + dfbase.loc[i, 'Var6 Número de Identificación del usuario']) # Campo 'Historia'
    r.type('//*[@id="vFROMFECHA"]', '[clear]01/07/2025') # Campo Inicio
    r.type('//*[@id="vTOFECHA"]', '[clear]30/06/2026') # Campo Fin
    r.click('Buscar - F4') # Botón 'Buscar'
    r.wait(5)
    # Vaidar si hay historias sino agrandar periodo
    if not r.present('//*[@id="vIMGATENCION_0001"]'):
        r.type('//*[@id="vFROMFECHA"]', '[clear]01/07/2024') # Campo Inicio
        r.click('Buscar - F4') # Botón 'Buscar'
        r.wait(5)
    # Vaidar si hay historias sino pasar a la siguiente
    if not r.present('//*[@id="vIMGATENCION_0001"]'):
        dfbase.loc[i, 'Estado'] = 'Sin Historia'
        print('Sin Historia --- proceso '+proceso)
        mensaje(variables, 'Sin Historia')
        # Actualizar archivo
        dfbase.loc[i, 'Estado'] = 'Sin Historia'
        dfbase.to_excel(variables['base'], index=False)
        continue

        
    # Paso 1: Descargar Folios
    # OJO configurar Chrome para descargar PDF en vez de abrir en el navegador con 
    #  chrome://settings/content/pdfDocuments
    print('Paso 1: Descargar Folios --- proceso', proceso)
    
    # Crear carpeta
    repositorio = str(variables['repositorio'].replace('\\\\', '\\'))
    carpeta_proceso = os.path.join(repositorio, "historias")
    carpeta_proceso = os.path.join(carpeta_proceso, proceso)
    os.makedirs(carpeta_proceso, exist_ok=True)

    # Descarga cada folio
    for n in range(1, 9999):
        #n=1
        # Abre la historia n
        if r.present(f'//*[@id="vIMGATENCION_{n:04d}"]'):
            # Valida que esté habilitado el enlace
            if not r.present(f'//*[@id="vIMGATENCION_{n:04d}" and @disabled=""]'):
                r.click(f'//*[@id="vIMGATENCION_{n:04d}"]') # Enlace Descargar Historia

                # Tomar el PDF generado (InformeResultados*.pdf)
                archivo_pdf = None
                # Esperar máximo 30 segundos a que aparezca el archivo en descargas
                for _ in range(30):
                    r.wait(1)
                    pdfs = glob.glob(os.path.join(CARPETA_BASE, "InformeResultados*.pdf"))
                    # Ignorar archivos aún descargándose
                    pdfs = [x for x in pdfs if not x.endswith(".crdownload")]
                    if pdfs:
                        # Selecciona el más reciente por si hay varios
                        r.wait(1)   
                        archivo_pdf = max(pdfs, key=os.path.getctime)
                        break
                
                if archivo_pdf is None:
                    print(f"No se descargó el PDF {n}")
                    continue

                # Nombre destino: proceso-n.pdf
                nombre_nuevo = f"{proceso}_LAB-{n}.pdf"
                ruta_destino = os.path.join(carpeta_proceso, nombre_nuevo)
                # Mover a la subcarpeta del proceso
                shutil.move(str(archivo_pdf), str(ruta_destino))
        else:
            break
            
    # Paso 2: Consolidar todos los PDFs de la carpeta del proceso en uno solo
    print('Paso 2: Consolidar todos los PDFs --- proceso', proceso)
    pdfs_individuales = sorted(glob.glob(os.path.join(carpeta_proceso, "*.pdf")))

    if pdfs_individuales:
        merger = PdfMerger()
        for pdf in pdfs_individuales:
            # Validar si el PDF está vacío
            if os.path.getsize(pdf) == 0:
                print(f"PDF vacío omitido: {pdf}")
                continue
            merger.append(str(pdf))
        # Guardar el PDF consolidado
        ruta_consolidado = os.path.join(carpeta_proceso, f"{proceso}_LAB.pdf")
        merger.write(str(ruta_consolidado))
        merger.close()
        # Eliminar los PDFs individuales
        for pdf in pdfs_individuales:
            if os.path.basename(pdf) != f"{proceso}.pdf":
                os.remove(pdf)
        # Comprimir pdf
        comprimir_pdf(ruta_consolidado, calidad="screen", reemplazar=True)
    else:
        print(f"No se encontraron PDFs individuales para consolidar en: {carpeta_proceso}")


    horafin = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print('Terminada Descarga de Laboratorios ' + str(n-1) + ' --- proceso', proceso, '-', horafin, '--------------------------------------------------')
    mensaje(variables, 'Terminada Descarga de Laboratorios ' + str(n-1) + ' --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin)
    mensaje(variables, 'Descarga de Laboratorios ' + str(n-1) + ' --- proceso '+proceso+' - inicio: '+horainicio+' - fin: '+horafin, variables['repositorio'])
    # Actualizar archivo
    dfbase.loc[i, 'Estado'] = 'Descargadas ' + str(n)
    dfbase.to_excel(variables['base'], index=False)


# Cerrar sesion
r.click('//*[@id="CUSTOMMENU_MPAGE"]/ul/li/a/span') # Menú Usuario
r.click('//*[@id="salir"]') # Botón Salir

# Cerrar robot
print('Cerrar robot', variables['robot'])
r.close()