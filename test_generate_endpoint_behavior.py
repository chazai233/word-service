import base64
import io
import unittest
from unittest.mock import patch

from docx import Document
from fastapi.testclient import TestClient

import main


def make_template_base64(rows: int = 6, cols: int = 5) -> str:
    doc = Document()
    doc.add_table(rows=rows, cols=cols)
    buf = io.BytesIO()
    doc.save(buf)
    return base64.b64encode(buf.getvalue()).decode()


def doc_from_base64(document_base64: str) -> Document:
    return Document(io.BytesIO(base64.b64decode(document_base64)))


class GenerateEndpointBehaviorTests(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(main.app)

    def test_generate_updates_cn_en_water_level(self):
        cn_template = make_template_base64(rows=6, cols=5)
        en_template = make_template_base64(rows=7, cols=5)
        with patch("main.fetch_feishu_water_level", return_value="123.45"):
            response = self.client.post(
                "/generate-from-template",
                json={
                    "chinese_data": '[{"seq":1,"location":"右岸道路","content":"测试内容","quantity":"100m"}]',
                    "english_data": '{"translated_data":[{"seq":1,"location_en":"Right Bank Roads","content_en":"Test content","quantity_en":"100m","remarks_en":""}]}',
                    "cn_template_base64": cn_template,
                    "en_template_base64": en_template,
                    "update_date_weather": True,
                    "feishu_app_token": "app_token",
                    "feishu_table_id": "table_id",
                },
            )

        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertTrue(body["success"])
        self.assertEqual(body["weather_info"]["water_level"], "123.45")
        self.assertEqual(body["weather_info"]["water_level_status"], "ok")
        self.assertNotIn("warnings", body)

        cn_doc = doc_from_base64(body["cn_document_base64"])
        en_doc = doc_from_base64(body["en_document_base64"])
        self.assertEqual(cn_doc.tables[0].rows[0].cells[3].text, "123.45")
        self.assertEqual(en_doc.tables[0].rows[0].cells[3].text, "123.45")

    def test_generate_returns_warnings_when_feishu_fails(self):
        cn_template = make_template_base64(rows=6, cols=5)
        en_template = make_template_base64(rows=7, cols=5)
        with patch("main.fetch_feishu_water_level", side_effect=RuntimeError("boom")):
            response = self.client.post(
                "/generate-from-template",
                json={
                    "chinese_data": '[{"seq":1,"location":"右岸道路","content":"测试内容","quantity":"100m"}]',
                    "english_data": '{"translated_data":[{"seq":1,"location_en":"Right Bank Roads","content_en":"Test content","quantity_en":"100m","remarks_en":""}]}',
                    "cn_template_base64": cn_template,
                    "en_template_base64": en_template,
                    "update_date_weather": True,
                    "feishu_app_token": "app_token",
                    "feishu_table_id": "table_id",
                },
            )

        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertTrue(body["success"])
        self.assertIn("warnings", body)
        self.assertEqual(len(body["warnings"]), 2)
        self.assertEqual(body["weather_info"]["water_level"], "")
        self.assertEqual(body["weather_info"]["water_level_status"], "failed")
        self.assertIn("cn_document_base64", body)
        self.assertIn("en_document_base64", body)

    def test_generate_requires_both_templates(self):
        response = self.client.post(
            "/generate-from-template",
            json={
                "chinese_data": '[{"seq":1,"location":"右岸道路","content":"测试内容","quantity":"100m"}]',
                "english_data": '{"translated_data":[{"seq":1,"location_en":"Right Bank Roads","content_en":"Test content","quantity_en":"100m","remarks_en":""}]}',
            },
        )

        self.assertEqual(response.status_code, 422)

    def test_generate_rejects_shared_templates(self):
        template = make_template_base64(rows=6, cols=5)
        response = self.client.post(
            "/generate-from-template",
            json={
                "chinese_data": '[{"seq":1,"location":"右岸道路","content":"测试内容","quantity":"100m"}]',
                "english_data": '{"translated_data":[{"seq":1,"location_en":"Right Bank Roads","content_en":"Test content","quantity_en":"100m","remarks_en":""}]}',
                "cn_template_base64": template,
                "en_template_base64": template,
            },
        )

        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertFalse(body["success"])
        self.assertIn("must be different templates", body["message"])


if __name__ == "__main__":
    unittest.main()
