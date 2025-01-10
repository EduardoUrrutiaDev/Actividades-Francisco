import os
import time
import comtypes.client
import fitz
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from pptx import Presentation
from PIL import Image
from io import BytesIO

# Ruta de descarga
DOWNLOAD_FOLDER = "C:/Users/eduardo.urrutia/Downloads/"

# Configurar opciones del navegador
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_FOLDER,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# Ruta del chromedriver
service = Service('C:/Users/eduardo.urrutia/Documents/Chrome Driver/chromedriver-win64/chromedriver.exe')

chrome_options.add_experimental_option("detach", True)

chrome_options.add_argument("--disable-blink-features=AutomationControlled")




# Inicializar WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    """
})


driver.maximize_window()

# Navegar al sitio web
driver.get("https://app.powerbi.com/groups/me/reports/d58a7dff-8a48-4b7a-ae23-7da4009c7cbf/ReportSection128671da9f6b6ceb5749?ctid=f31661d6-6630-4f8a-a619-abd1d6192f40&experience=power-bi")

# Digitación de correo
username = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id='email']")))
username.clear()
username.send_keys("eduardo.urrutia@lindcorp.pe")

# Botón de enviar
button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id='submitBtn']"))).click()

# Digitación de contraseña
password = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='passwd']")))
password.clear()
password.send_keys("Lindcorp2024*")

# Botón de iniciar sesión
buttonSignIn = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

# Mantener sesión iniciada
button2 = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

# Clic en exportar
buttonExportar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "exportMenuBtn"))).click()

time.sleep(3) 

# Hacer clic en PPT
buttonPpt = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@tabindex='0' and @role='menuitem' and contains(@class, 'mat-menu-item')]"))
).click()

# Botón abrir en PowerPoint
buttonOpen = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "okButton"))).click()

# Esperar a que se descargue el archivo
print("Esperando que se complete la descarga...")
def esperar_descarga(carpeta, tiempo_max=1000, extension=".pptx"):
    tiempo_inicio = time.time()
    archivos_previos = {}  # Guardar tamaños previos de los archivos para verificar cambios

    while time.time() - tiempo_inicio < tiempo_max:
        archivos = os.listdir(carpeta)
        archivos_descargados = [archivo for archivo in archivos if archivo.endswith(extension)]

        for archivo in archivos_descargados:
            ruta_archivo = os.path.join(carpeta, archivo)
            try:
                tamano_actual = os.path.getsize(ruta_archivo)

                # Si es la primera vez que vemos el archivo, guardamos su tamaño
                if archivo not in archivos_previos:
                    archivos_previos[archivo] = tamano_actual

                # Si el tamaño no ha cambiado durante 5 segundos, asumimos que la descarga ha terminado
                elif tamano_actual == archivos_previos[archivo]:
                    print(f"Descarga completada para: {archivo}")
                    return ruta_archivo
                else:
                    archivos_previos[archivo] = tamano_actual  # Actualizamos el tamaño del archivo

            except FileNotFoundError:
                continue  # En caso de que el archivo haya sido eliminado

        time.sleep(1)  # Esperar 1 segundo antes de comprobar nuevamente

    raise Exception("La descarga tomó demasiado tiempo o falló.")

# Llamar a la función
try:
    ruta_archivo_pptx = esperar_descarga("C:/Users/eduardo.urrutia/Downloads/")
    print(f"Archivo descargado: {ruta_archivo_pptx}")
except Exception as e:
    print(f"Error: {e}")

esperar_descarga(DOWNLOAD_FOLDER)

# Identificar el archivo descargado
archivo_descargado = [archivo for archivo in os.listdir(DOWNLOAD_FOLDER) if archivo.endswith('.pptx')][0]
ruta_archivo_pptx = os.path.join(DOWNLOAD_FOLDER, archivo_descargado)
print(f"Archivo descargado: {ruta_archivo_pptx}")

def convertir_a_raw_string(ruta):
    return f'r"{ruta.replace("/", "\\")}"'

ruta_convertida = convertir_a_raw_string(ruta_archivo_pptx)
print(ruta_convertida)


def convertir_pptx_a_pdf_con_datos_dinamicos(ruta_pptx, ruta_pdf):
    """
    Convierte un archivo PowerPoint (.pptx) a un archivo PDF, asegurando que los datos dinámicos se actualicen.
    
    :param ruta_pptx: Ruta del archivo PowerPoint (.pptx) a convertir.
    :param ruta_pdf: Ruta donde se guardará el archivo PDF resultante.
    """
    # Asegúrate de que la ruta del archivo de entrada sea válida
    if not os.path.exists(ruta_pptx):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_pptx}")
    
    # Crear la carpeta de salida si no existe
    directorio_pdf = os.path.dirname(ruta_pdf)
    if directorio_pdf and not os.path.exists(directorio_pdf):
        os.makedirs(directorio_pdf)
    
    # Iniciar PowerPoint
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  # Hacer PowerPoint visible para depuración

    try:
        # Abrir la presentación
        presentation = powerpoint.Presentations.Open(ruta_pptx)

        # Forzar la actualización de enlaces dinámicos
        print("Actualizando enlaces dinámicos en la presentación...")
        presentation.UpdateLinks()

        # Recorrer las diapositivas para asegurar la actualización de gráficos y tablas
        for i, slide in enumerate(presentation.Slides):
            print(f"Actualizando diapositiva {i + 1}...")
            for shape in slide.Shapes:
                if shape.HasChart:  # Si el objeto es un gráfico
                    print(f"  - Actualizando gráfico en la diapositiva {i + 1}")
                    shape.Chart.Refresh()  # Refrescar el gráfico
                    shape.Chart.Application.Update()  # Forzar actualización completa
                    time.sleep(2)  # Breve espera para la actualización

                if shape.HasTable:  # Si el objeto es una tabla
                    print(f"  - Actualizando tabla en la diapositiva {i + 1}")
                    time.sleep(2)  # Breve espera para forzar actualización

        # Esperar un poco más para asegurarse de que todos los datos dinámicos se actualicen
        print("Esperando para garantizar la actualización completa de los datos dinámicos...")
        time.sleep(60)

        # Exportar la presentación como PDF
        print(f"Convirtiendo {ruta_pptx} a PDF...")
        presentation.SaveAs(ruta_pdf, 32)  # 32 es el formato para PDF en PowerPoint
        
        print(f"Archivo PDF guardado en: {ruta_pdf}")
        
        # Cerrar la presentación
        presentation.Close()
    finally:
        # Cerrar PowerPoint
        powerpoint.Quit()

# Ruta del archivo PowerPoint y salida como PDF
ruta_pptx = r"C:\Users\eduardo.urrutia\Downloads\Microsoft-Power-BI-Storytelling.pptx"
ruta_pdf = r"C:\Users\eduardo.urrutia\Downloads\Microsoft-Power-BI-Storytelling.pdf"

# Llamar a la función para convertir a PDF
convertir_pptx_a_pdf_con_datos_dinamicos(ruta_pptx, ruta_pdf)

def convertir_pdf_a_imagenes(pdf_path, output_folder, resolution=300):
    """
    Convierte cada página de un archivo PDF a imágenes de alta resolución.
    
    :param pdf_path: Ruta del archivo PDF de entrada.
    :param output_folder: Carpeta donde se guardarán las imágenes.
    :param resolution: Resolución en DPI (píxeles por pulgada) para las imágenes.
    """
    # Asegúrate de que la ruta del archivo PDF sea válida
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No se encontró el archivo PDF: {pdf_path}")

    # Crear la carpeta de salida si no existe
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Abrir el archivo PDF
    pdf_document = fitz.open(pdf_path)
    print(f"Convirtiendo {pdf_path} a imágenes...")

    for page_num in range(len(pdf_document)):
        # Obtener la página
        page = pdf_document[page_num]
        
        # Ajustar la resolución
        matrix = fitz.Matrix(resolution / 72, resolution / 72)
        
        # Renderizar la página como imagen
        pix = page.get_pixmap(matrix=matrix, alpha=False)

        # Guardar la imagen
        image_path = os.path.join(output_folder, f"pagina_{page_num + 1}.png")
        pix.save(image_path)

        print(f"Página {page_num + 1} guardada como imagen en: {image_path}")

    # Cerrar el archivo PDF
    pdf_document.close()
    print("Conversión completada.")

# Parámetros
pdf_path = r"C:\Users\eduardo.urrutia\Downloads\Microsoft-Power-BI-Storytelling.pdf"
output_folder = r"C:\Users\eduardo.urrutia\Downloads\Imagenes_PDF"
resolution = 300  # DPI para alta resolución

# Llamar a la función para convertir
convertir_pdf_a_imagenes(pdf_path, output_folder, resolution)

#Ingresar a cuenta de gmail

driver.get("https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fmail.google.com%2Fmail%2Fu%2F0%2F&emr=1&followup=https%3A%2F%2Fmail.google.com%2Fmail%2Fu%2F0%2F&osid=1&passive=1209600&service=mail&ifkv=AVdkyDkfVZNhZJBoLImXrFTIjPXoEAiwJYCfqKr1mmOnVh1f-v_DUuS8JFSnaUrLzWou1neGIhjuYA&ddm=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin")

# Digitación de correo
username = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id='identifierId']")))
username.clear()
username.send_keys("eduardo.urrutia@lindcorp.pe")

#Siguiente
button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Siguiente']]"))
).click()

#Redactar password
password_field = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.NAME, "Passwd"))
)
password_field.clear()
password_field.send_keys("Benjamin250624")

#Ingresar
button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Siguiente']]"))
)
button.click()

buttonCerrar = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Cerrar']"))
)
buttonCerrar.click()

# Esperar hasta que el elemento sea clickable y hacer clic
buttonRedactar = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Redactar')]"))
)
buttonRedactar.click()

#Digitar correo
destinatario = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "input[role='combobox'][aria-label='Destinatarios']"))
)
destinatario.clear()
destinatario.send_keys("urrutiaveduardo@gmail.com")

#Agregar asunto
asunto = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.NAME, "subjectbox"))
)
asunto.clear()
asunto.send_keys("Este es el asunto del correo")


# time.sleep(15)
elemento = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.XPATH, "//div[@id=':wn']"))
)
elemento.click()

#click Enviar correo
boton_enviar = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.XPATH, "//div[@aria-label='Enviar']"))
)
boton_enviar.click()  # Hacer clic en el botón "Enviar"
# Cerrar el navegador
driver.quit()
print("Proceso completado.")