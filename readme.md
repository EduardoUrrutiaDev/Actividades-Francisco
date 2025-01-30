**Automatización de Reportes Mensuales de Backus**

**Descripción y contexto**

Este proyecto es una automatización que realiza las siguientes tareas:

1. Obtiene la fecha actual.
1. Calcula el rango de fechas del mes anterior.
1. Ejecuta un procedimiento almacenado en SQL Server con el rango de fechas calculado.
1. Convierte la data obtenida en un archivo Excel.
1. Envía el archivo Excel por correo mediante Gmail.

**Guía de usuario**

Ejecutar el script principal:

	python main.py

Esto realizará todo el proceso automáticamente.

**Guía de instalación**

**Requisitos**

- Python 3.12
- Instalar las siguientes librerias : 
  - pandas (para manejar datos y generar el archivo Excel)
  - pyodbc y sqlalchemy (para trabajar con base de datos)
  - openpyxl (para manipular archivos Excel)
  - smtplib y email (para enviar correos con Gmail)
  - decouple (para cargar variables de entorno desde un archivo  .env)

- Conexión a SQL Server con los permisos adecuados.
- Cuenta de Gmail configurada para el envío de correos.



**Estructura del Proyecto**

📂 ReporteBackus 

│── 📂 repository

│   └── db_connection.py        # Conexion con la base de datos

│── 📂 services 

│   └── data_extractor.py  # Extraer data del procedimiento almacenado

│   └── email_sender.py  # Enviar el archivo excel por correo

│   └── excel_generator.py  # Generar el excel con la data

│── 📂 utils  

│   └── date_range.py  # Obtener rango de fechas del mes anterior

│   └── date_today.py  # Obtener fecha actual

│── .env                 # Archivo de variables de entorno

│── main.py              # Script principal

**Consideraciones**

- Validar que el procedimiento almacenado en SQL Server devuelve los datos esperados.
- Configurar correctamente los permisos en el servidor de base de datos y en la cuenta de correo.


**Autor/es**

- Eduardo Urrutia

[Python Version]: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAOxAAADsQBlSsOGwAAABpJREFUWIXtwQEBAAAAgiD/r25IQAEAAADvBhAgAAFHAaCIAAAAAElFTkSuQmCC
