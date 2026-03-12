import sys
sys.path.insert(0, '.')
from sectors_config import N8N_BASE_URL, N8N_API_KEY, SECTORS
print("URL:", N8N_BASE_URL)
print("KEY:", N8N_API_KEY[:20])
import requests, urllib3
urllib3.disable_warnings()
url = N8N_BASE_URL + "/api/v1/data-tables/" + SECTORS["monetary"]["datatable_id"] + "/rows"
print("Запрос к:", url)
r = requests.get(url, headers={"X-N8N-API-KEY": N8N_API_KEY}, timeout=10, verify=False)
print("Статус:", r.status_code)
print("Ответ:", r.text[:200])
