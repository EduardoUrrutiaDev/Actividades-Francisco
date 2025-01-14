import http.client
import datetime
import json
from urllib.parse import urlencode

# Obtener la fecha y hora actuales y calcular el timestamp de hace 10 minutos
end_time = datetime.datetime.utcnow()
start_time = end_time - datetime.timedelta(minutes=10)

print(f"Hora inicio: {start_time}")
print(f"Hora fin:    {end_time}")

# Convertir a formato ISO 8601
start_time_str = start_time.isoformat() + 'Z'
end_time_str = end_time.isoformat() + 'Z'

# Crear los parámetros de la consulta
params = {
    'f_creationDate': f'creationDate:[{start_time_str} TO {end_time_str}]',
    'per_page': 15  # Ajusta según la cantidad de órdenes que desees
}

# Codificar los parámetros
query_string = urlencode(params)

# Establecer la conexión
conn = http.client.HTTPSConnection("aruma.vtexcommercestable.com.br")

# Definir los headers
headers = {
    'Accept': "application/json",
    'Content-Type': "application/json",
    'X-VTEX-API-AppKey': "",
    'X-VTEX-API-AppToken': ""
}

# Realizar la solicitud GET con los parámetros
url = f"/api/oms/pvt/orders?{query_string}"
conn.request("GET", url, headers=headers)

# Obtener y procesar la respuesta
res = conn.getresponse()
data = res.read()

# Convertir la respuesta a JSON
response_json = json.loads(data.decode("utf-8"))

# Lista para almacenar todos los orderId
order_ids = []

# Verificar si la clave "list" existe en la respuesta
if "list" in response_json:  
    for order in response_json["list"]:
        order_id = order.get("orderId", None)
        if order_id:  # Asegurarse de que el orderId existe
            order_ids.append(order_id)
else:
    print("No orders found or error in response.")

# Imprimir los orderIds almacenados
print("-" * 50)  # Separador entre órdenes
if not order_ids:
    print("No hay órdenes")      
else:
    print(order_ids)
print("-" * 50)  # Separador entre órdenes





# Función para obtener Order ID, DNI y la promoción aplicada
def obtener_datos_ordenes(order_ids):
    # Establecer la conexión con la plataforma VTEX
    conn = http.client.HTTPSConnection("aruma.vtexcommercestable.com.br")
    
    # Definir los headers para la autenticación en la API
    headers = {
        'Accept': "application/json",
        'Content-Type': "application/json",
        'X-VTEX-API-AppKey': "",
        'X-VTEX-API-AppToken': ""
    }
    
    resultados = []  # Lista para almacenar los objetos con los datos

    for order_id in order_ids:
        # Realizar la solicitud GET con el orderId específico
        conn.request("GET", f"/api/oms/pvt/orders/{order_id}", headers=headers)
        res = conn.getresponse()
        data = res.read()
        response_json = json.loads(data.decode("utf-8"))

        # Extraer el DNI del cliente
        client_profile_data = response_json.get("clientProfileData", {})
        dni_cliente = client_profile_data.get("document", "No disponible")

        # Extraer la promoción aplicada
        rates_and_benefits_data = response_json.get("ratesAndBenefitsData", {})
        promotions = rates_and_benefits_data.get("rateAndBenefitsIdentifiers", [])
        promocion_aplicada = promotions[0]["name"] if promotions else "No hay promociones aplicadas"

        

        # Crear el objeto con los datos
        if dni_cliente and promocion_aplicada != "No hay promociones aplicadas":
            # Crear el objeto con los datos
            resultado = {
                "orderId": order_id,
                "document": dni_cliente,
                "name": promocion_aplicada
            }
            resultados.append(resultado)


    return resultados

# Ejemplo de uso
resultados = obtener_datos_ordenes(order_ids)

# Imprimir los resultados como lista de objetos

if  resultados:
    print(resultados)     
else:
    print("[]")
