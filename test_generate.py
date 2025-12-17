import requests
import json
import base64

# 读取测试数据
data = json.load(open('test_data.json', 'r', encoding='utf-8'))

# 读取模板文件
with open(r'd:\Projects\Dify\[CN]北本水电站施工日报.docx', 'rb') as f:
    cn_b64 = base64.b64encode(f.read()).decode()
with open(r'd:\Projects\Dify\[EN]Pak Beng daily construction report.docx', 'rb') as f:
    en_b64 = base64.b64encode(f.read()).decode()

# 请求
r = requests.post('http://localhost:8000/generate-from-template', json={
    'chinese_data': data['chinese_data'],
    'english_data': data['english_data'],
    'cn_template_base64': cn_b64,
    'en_template_base64': en_b64
})
result = r.json()

if result.get('success'):
    with open('output_cn.docx', 'wb') as f:
        f.write(base64.b64decode(result['cn_document_base64']))
    with open('output_en.docx', 'wb') as f:
        f.write(base64.b64decode(result['en_document_base64']))
    
    weather = result.get('weather_info', {})
    print('生成成功!')
    print(f"日期: {weather.get('date')}")
    print(f"天气: {weather.get('weather')}")
    print(f"温度: {weather.get('temp')}")
else:
    print('错误:', result)
