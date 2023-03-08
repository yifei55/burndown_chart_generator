import requests

r = requests.get('https://reqbin.com/echo/get/json',
                 headers={'Accept': 'application/json'})

print(f"Response: {r.json()}")