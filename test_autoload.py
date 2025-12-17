import requests
import json
import base64

# 测试不传模板时自动加载
r = requests.post('http://localhost:8000/generate-from-template', json={
    'chinese_data': '[{"seq":1,"location":"右岸道路","content":"测试内容","quantity":"100m","shift":""}]',
    'english_data': '{"translated_data":[{"seq":1,"location_en":"Right Bank Roads","content_en":"Test content","quantity_en":"100m","remarks_en":""}]}'
})
result = r.json()

if result.get('success'):
    print('自动加载模板成功!')
    print(f"天气信息: {result.get('weather_info')}")
    
    # 保存测试文件
    with open('test_cn.docx', 'wb') as f:
        f.write(base64.b64decode(result['cn_document_base64']))
    with open('test_en.docx', 'wb') as f:
        f.write(base64.b64decode(result['en_document_base64']))
    print('文档已保存: test_cn.docx, test_en.docx')
else:
    print('错误:', result)
