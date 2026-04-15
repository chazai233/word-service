# Word Service

基于 FastAPI 的 Word 文档生成服务，支持：

- 中文/英文日报一键生成（`/generate-from-template`）
- 文本按规则格式化写入模板单元格（标题、缩进、局部加粗）
- 表头日期/天气/气温/水位更新（支持 Feishu Bitable 水位读取）
- 附录表格工程量回填

## 快速启动

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 本地运行

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

### 3. 查看接口文档

- Swagger: `http://127.0.0.1:8000/docs`
- 健康检查: `GET /health`

## 核心接口

### `POST /generate-from-template`

用于同时生成中文和英文 Word 文档，是推荐主入口。

请求体核心字段：

- `chinese_data`: 中文数据（`list` / `dict` / JSON 字符串）
- `english_data`: 英文数据（`list` / `dict` / JSON 字符串）
- `cn_template_base64`, `en_template_base64`: 可选模板（base64）
- `update_date_weather`: 是否在生成时更新表头信息
- `cn_table_index/cn_row_index/cn_col_index`: 中文写入目标
- `en_table_index/en_row_index/en_col_index`: 英文写入目标
- Feishu 参数（可选）：
- `feishu_token`, `feishu_app_token`, `feishu_table_id`, `feishu_view_id`
- `feishu_water_level_field`, `feishu_date_field`, `feishu_date_value`
- `feishu_app_id`, `feishu_app_secret`

返回字段：

- `success`: 是否成功
- `cn_document_base64`, `en_document_base64`: 生成的文档
- `weather_info`: 表头更新结果（含 `water_level`、`water_level_status`）
- `warnings`: 可选。若 Feishu 水位获取失败，仍返回文档并附告警

### `POST /fill-template`

把多行文本写入指定模板单元格，按规则处理：

- 大标题：不缩进，整行加粗（如 `1、右岸道路`）
- 统计项：首段缩进，关键词前缀加粗（如 `人员投入：`）
- 子标题/正文：首段缩进

### `POST /update-date-weather`

更新已有文档中的日期/天气/气温/水位。

### `POST /update-personnel-stats`

在文末追加人员统计文本。

### `POST /update-appendix-tables`

按 `row_name` 匹配附录表行，写入今日/累计工程量。

## 部署说明（Render）

### 推荐配置

- Build Command: `pip install -r requirements.txt`
- Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
- 分支：`main`

### 部署后检查

1. 打开 `/docs` 确认存在 `POST /generate-from-template`
2. 再让 Dify 调用接口

## Dify 调用注意事项

1. URL 必须指向：`/generate-from-template`
2. `Content-Type` 使用 `application/json`
3. 若出现 `404 Not Found`，通常是线上服务仍是旧版本或未重部署
4. 建议先在 `/docs` 里手动试调成功，再接入工作流

## 测试

```bash
python -m unittest -q
```
