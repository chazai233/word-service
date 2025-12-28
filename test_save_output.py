import requests
import json
import base64
import os

url = 'http://localhost:8001/generate-from-template'

# Dummy data similar to what Dify sends
chinese_data = json.dumps([
    {"seq": 1, "location": "右岸道路", "content": "R10路段回填沙砾石", "quantity": "20m", "shift": "白天"}
])

english_data = json.dumps({
    "translated_data": [
        {"seq": 1, "location_en": "Right Bank Roads", "content_en": "Backfill of R10 Road with sand and gravel", "quantity_en": "20m", "remarks_en": "Day"}
    ]
})

payload = {
    "chinese_data": chinese_data,
    "english_data": english_data,
    "cn_template_base64": None,
    "en_template_base64": None
}

print(f"Sending request to {url}...")
try:
    r = requests.post(url, json=payload)
    data = r.json()
    if data.get('success'):
        en_base64 = data.get('en_document_base64')
        if en_base64:
            with open("test_output_en.docx", "wb") as f:
                f.write(base64.b64decode(en_base64))
            print("Saved test_output_en.docx")
        else:
            print("No en_document_base64 in response")
        
        cn_base64 = data.get('cn_document_base64')
        if cn_base64:
            with open("test_output_cn.docx", "wb") as f:
                f.write(base64.b64decode(cn_base64))
            print("Saved test_output_cn.docx")
    else:
        print(f"Request failed: {data}")
except Exception as e:
    print(f"Error: {e}")
