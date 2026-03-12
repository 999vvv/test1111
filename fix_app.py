import re

with open("app.py", "r", encoding="utf-8") as f:
    content = f.read()

# Исправляем fetch_datatable - добавляем verify=False и urllib3
old = '''def fetch_datatable(datatable_id: str) -> list:
    """Получает все строки DataTable через n8n REST API."""
    url = f"{N8N_BASE_URL}/api/v1/data-tables/{datatable_id}/rows"
    headers = {"X-N8N-API-KEY": N8N_API_KEY}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("data", [])'''

new = '''def fetch_datatable(datatable_id: str) -> list:
    """Получает все строки DataTable через n8n REST API."""
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    url = N8N_BASE_URL + "/api/v1/data-tables/" + datatable_id + "/rows"
    headers = {"X-N8N-API-KEY": N8N_API_KEY}
    resp = requests.get(url, headers=headers, timeout=30, verify=False)
    resp.raise_for_status()
    return resp.json().get("data", [])'''

if old in content:
    content = content.replace(old, new)
    print("fetch_datatable - исправлено")
else:
    print("fetch_datatable - не найдено точное совпадение, ищем другой вариант...")
    # Универсальная замена через regex
    pattern = r'def fetch_datatable\(datatable_id.*?return resp\.json\(\)\.get\("data", \[\]\)'
    replacement = '''def fetch_datatable(datatable_id: str) -> list:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    url = N8N_BASE_URL + "/api/v1/data-tables/" + datatable_id + "/rows"
    headers = {"X-N8N-API-KEY": N8N_API_KEY}
    resp = requests.get(url, headers=headers, timeout=30, verify=False)
    resp.raise_for_status()
    return resp.json().get("data", [])'''
    content, n = re.subn(pattern, replacement, content, flags=re.DOTALL)
    print(f"regex замена: {n} вхождений")

# Исправляем webhook запрос - добавляем verify=False
content = content.replace(
    'resp = requests.post(webhook_url, json={"action": "run_all"}, timeout=10)',
    'resp = requests.post(webhook_url, json={"action": "run_all"}, timeout=10, verify=False)'
)

with open("app.py", "w", encoding="utf-8") as f:
    f.write(content)

print("app.py обновлён")
