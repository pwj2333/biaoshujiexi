# 标书解析与登记 Demo

这是一个独立 demo，支持：

- 上传文本型 `PDF` / `DOCX` 标书
- 调用用户自填的 OpenAI 兼容接口解析标书
- 生成摘取结果、登记预览、AI 对话
- 导出 3 类文件：
  - 摘取 Excel
  - 登记 Excel
  - 解析总览 Markdown
- 将历史记录保存到本地目录 `data/projects`

## 启动

直接双击 `一键启动.bat`。

如果要手动启动：

```powershell
python -m pip install -r requirements.txt
python app.py
```

浏览器打开 `http://127.0.0.1:8008`

## 数据目录

- `data/config.json`：AI 配置
- `data/projects/*.json`：历史记录
- `templates/`：Excel 模板

## 当前范围

- 只支持可提取文本的 PDF 和 DOCX
- 不做 OCR
- 不接公司内部运力、港口、客户系统
- 历史记录使用本地 JSON 文件，不用数据库
