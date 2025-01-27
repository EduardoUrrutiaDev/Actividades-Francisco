from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import time
import comtypes.client
import fitz  # PyMuPDF
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

from PIL import Image
from io import BytesIO

from src.constant import USER_CC, USER_TO

print("init process")
# Ruta de descarga
DOWNLOAD_FOLDER = r"C:\Users\eduardo.urrutia\Downloads"


# Configurar opciones del navegador
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_FOLDER,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "plugins.always_open_pdf_externally": True
})

print("selected chromedriver")
# Ruta del chromedriver
service = Service('C:/Users/eduardo.urrutia/Documents/Chrome Driver/chromedriver-win64/chromedriver.exe')

chrome_options.add_experimental_option("detach", True)

chrome_options.add_argument("--disable-blink-features=AutomationControlled")

chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-extensions")


print("init webdriver")
# Inicializar WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

print("add script to evaluate on new document")
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    """
})

print("maximize window")
driver.maximize_window()

print("navigate to site")
# Navegar al sitio web
driver.get("https://app.powerbi.com/groups/me/reports/6d3ff701-1a25-410d-895b-d82dfd133e74/e46735aca7547a529d08?ctid=f31661d6-6630-4f8a-a619-abd1d6192f40&experience=power-bi")

# Digitación de correo
username = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id='email']")))
username.clear()
username.send_keys("eduardo.urrutia@lindcorp.pe")

# Botón de enviar
button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[id='submitBtn']"))).click()

# Digitación de contraseña
password = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='passwd']")))
password.clear()
password.send_keys("Lindcorp2024*")

#Botón de iniciar sesión
buttonSignIn = WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

#Mantener sesión iniciada
button2 = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()

# Clic en exportar
buttonExportar = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "exportMenuBtn"))).click()

time.sleep(5) 

# Hacer clic en Pdf
buttonPpt = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'mat-menu-item')]//span[text()='PDF']"))
).click()

# Click en exportar 
buttonOpen = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "okButton"))).click()

# Esperar a que se descargue el archivo
downloaded_file = None
MAX_WAIT_TIME = 300  # 5 minutos
poll_interval = 5  # Intervalo en segundos para verificar el archivo

start_time = time.time()

while True:
    print("Esperando a que se complete la descarga...")
    # Lista los archivos en la carpeta de descargas
    downloaded_files = os.listdir(DOWNLOAD_FOLDER)

    for file_name in downloaded_files:
        print(file_name)
        if file_name.lower().endswith(".pdf"):  # Insensible a mayúsculas
            downloaded_file = os.path.join(DOWNLOAD_FOLDER, file_name)
            break

    if downloaded_file:  # Si se encuentra, rompe el bucle principal
        break

    # Verifica si ha pasado el tiempo máximo de espera
    if time.time() - start_time > MAX_WAIT_TIME:
        print(f"Tiempo máximo de espera alcanzado. El archivo no se descargó.")
        break

    # Espera antes de volver a comprobar
    time.sleep(poll_interval)

if not downloaded_file:
    raise Exception("El archivo no se descargó en el tiempo establecido.")

print(f"Archivo descargado: {downloaded_file}")


#Configuracion
EMAIL_SENDER = "multifuncional@lindcorp.pe"
EMAIL_RECEIVER = "francisco.esparza@lindcorp.pe"
EMAIL_ALIAS = "reportes.diario@lindcorp.pe"
EMAIL_PASSWORD = "Lind#T4mb#23"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
OUTPUT_DIR = os.path.join(DOWNLOAD_FOLDER, "images")
PDF_FILE = downloaded_file

#Obtener fecha actual
fecha_actual = datetime.now().strftime("%y%m%d")

# Crear carpeta de salida dentro de la carpeta de descargas si no existe
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Abrir PDF y dividir cada página en imágenes
def pdf_to_images(pdf_file, output_dir):
    doc = fitz.open(pdf_file)
    image_paths = []
    for page_num in range(len(doc)):
        if page_num in [4]:  # Omitir página 5 (índice 4)
            continue
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        image_path = os.path.join(output_dir, f"pagina_{page_num + 1}.png")
        pix.save(image_path)
        image_paths.append(image_path)
    doc.close()
    return image_paths

# Enviar un correo con cada imagen como adjunto
def send_email_with_images(image_paths, sender, password):
    msg = MIMEMultipart()
    msg['From'] = f"{EMAIL_ALIAS} <{sender}>"
    msg['To'] = ', '.join(USER_TO)
    msg['Subject'] = f"Reporte diario de ventas - {fecha_actual}"
    msg['Bcc'] = ", ".join(USER_CC)
    
    body = "Buenos días. Se adjunta el reporte de venta diario. Saludos"
    msg.attach(MIMEText(body, 'plain'))
    all_recipients = USER_TO + USER_CC
    # Adjuntar cada imagen
    for image_path in image_paths:
        with open(image_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(image_path)}")
            msg.attach(part)

    
    try:
        # Enviar correo
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, all_recipients, msg.as_string())

         # Confirmación del envío
        print("Imágenes creadas y correo enviado exitosamente con cada imagen como adjunto.")

    except smtplib.SMTPException as e:
        # Manejo de errores en el envío
        print(f"Error al enviar el correo: {e}")

# Ejecución de las funciones
image_paths = pdf_to_images(PDF_FILE, OUTPUT_DIR)  # Divide el PDF en imágenes
send_email_with_images(image_paths, EMAIL_SENDER, EMAIL_PASSWORD)  # Envía el correo con las imágenes adjuntas

