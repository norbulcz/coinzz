from datetime import datetime
from requests import Request, Session
import json
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects

url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest'
# url = 'https://api.coindesk.com/v1/bpi/currentprice.json'
parameters = {
    'start':'1',
    'limit':'5000',
    'convert':'USD'
}
headers = {
    'Accepts': 'application/json',
    'X-CMC_PRO_API_KEY': 'fcf7d05c-d732-4268-8ea2-a5db562f0521'
}

session = Session()
session.headers.update(headers)



try:
    response = session.get(url, params=parameters)
    # response = session.get(url)
    
    data = json.loads(response.text)
    with open('files/data-'+datetime.now().strftime('%d-%m-%Y')+'.json', 'w') as file:
        json.dump(data, file)
except (ConnectionError, Timeout, TooManyRedirects) as e:
  print(e)
