import requests
import json

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)...",
    "Accept-Language": "es-AR,es;q=0.9",
})

resp = session.get("https://boletinoficial.cba.gov.ar/", timeout=30)
print(resp.status_code)

# Guardar resultado para descargarlo como artifact
with open("resultados.json", "w") as f:
    json.dump({"status": resp.status_code, "html": resp.text[:2000]}, f)
