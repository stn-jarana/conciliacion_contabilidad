import os
import requests

url = 'https://api.decolecta.com/v1/tipo-cambio/sunat'

API_KEY = os.getenv('api_key')

headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {API_KEY}',
}

params = {
    "date": '2024-06-01',
}

response = requests.get(url, headers=headers, params=params)

print("Status:", response.status_code)
print(response.text)
