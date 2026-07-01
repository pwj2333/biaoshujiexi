# 标书解析与登记 Demo

## 项目简介

这是一个面向招投标场景的内部演示 Demo，用于辅助处理外部标书文件。

当前版本支持：

- 上传文本型 `PDF` / `DOC` / `DOCX` / `XLS` / `XLSX` 标书
- 调用用户自填的 OpenAI 兼容接口解析标书
- 提取标书核心字段并生成结构化结果
- 生成登记预览与 AI 对话分析
- 导出 3 类文件：
  - 摘取结果 Excel
  - 招标登记 Excel
  - 解析总览 Markdown
- 将项目结果保存为本地历史记录，便于后续查询和复用

## 核心能力

- 标书解析：
  提取项目名称、招标编号、投标截止时间、保证金、资格要求、船舶要求、评标标准、报价要求等关键信息。

- 原文依据保留：
  导出的摘取结果文件包含两页。
  第一页为摘取结果。
  第二页为原文依据，且尽量保留完整原文条款块，不使用省略号代替。

- 招标登记导出：
  使用项目目录下的 `模板文件/招标登记.xlsx` 作为唯一登记模板。
  登记年份由用户手动填写；若模板中不存在该年份 sheet，程序会复制模板 sheet 自动生成。

- 多标的 / 多标段识别：
  如果同一份文件中包含多个标的、多个合同包、多个标段、多个航线或多个独立报价单元，系统会尽量拆分并统计数量。

- 历史记录管理：
  当前解析结果可保存到本地 JSON 文件中，后续可从左侧历史记录菜单查询并重新打开。

## 技术方案

- 后端：`FastAPI`
- 前端：原生 `HTML + CSS + JavaScript`
- 文件解析：
  - `pypdf` 解析文本型 PDF
  - `zip/xml` 方式解析 DOCX 正文
  - `openpyxl` 解析 XLSX
  - Windows 本机 Office COM 兜底解析 DOC / XLS
- Excel 导出：`openpyxl`
- AI 调用：OpenAI 兼容 `chat/completions`

## 目录结构

```text
biaoshujiexi/
├─ app.py                    # 后端入口
├─ requirements.txt          # Python 依赖
├─ README.md                 # 项目说明
├─ test_smoke.py             # 最小烟雾测试
├─ 一键启动.bat               # Windows 一键启动脚本
├─ static/                   # 前端页面资源
│  ├─ index.html
│  ├─ app.js
│  └─ styles.css
├─ 模板文件/                  # 打包与正式运行时使用的模板目录
│  ├─ extraction_template.xlsx
│  └─ 招标登记.xlsx
├─ templates/                # 早期模板备份目录
├─ data/                     # 本地运行数据
│  ├─ config.json            # AI 配置
│  ├─ projects/              # 历史记录
│  └─ tmp/                   # 临时上传文件
└─ 招标文件/                  # 本地测试用标书样本（默认不提交）
```

## 模板文件说明

程序当前强制只认项目目录下的 `模板文件` 目录。

必须存在：

- `模板文件/extraction_template.xlsx`
- `模板文件/招标登记.xlsx`

如果缺少任一模板，后端会直接返回明确错误，不再回退到其他目录。

## 启动方式

### 方式一：一键启动

直接双击：

```text
一键启动.bat
```

一键启动脚本会自动执行这些步骤：

1. 检测本机是否已有可用 Python 3
2. 如果没有，优先尝试用 `winget` 自动安装 Python 3.11
3. 如果 `winget` 不可用，再尝试从 Python 官网下载安装包并静默安装
4. 自动创建项目本地虚拟环境 `.venv`
5. 自动安装 `requirements.txt` 里的依赖
6. 自动启动服务并打开浏览器

注意：

- 首次启动如果需要下载安装 Python 或依赖，会比平时慢一些
- 如果网络较慢，依赖安装过程可能需要等待一会儿
- 如果目录里已经带了旧的 `.venv`，脚本会自动检测是否可用；发现是别的电脑拷过来的无效环境时，会自动删除并重建
- 如果启动失败，请查看项目目录下的 `startup.log`

### 方式二：命令行启动

```powershell
python -m pip install -r requirements.txt
python app.py
```

启动后访问：

```text
http://127.0.0.1:8008
```

## 使用流程

1. 在左侧 `AI 配置` 中填写：
   - Base URL
   - API Key
   - 模型名称

2. 上传标书文件：
   - 支持文本型 `PDF`
   - 支持 `DOC` / `DOCX`
   - 支持 `XLS` / `XLSX`

3. 选择登记展示方式：
   - 每个标段一行
   - 整份标书一行

4. 手动填写登记年份

5. 点击 `开始解析`

6. 在右侧查看：
   - 解析总览
   - 摘取结果
   - 登记预览
   - AI 对话

7. 如需保留，点击 `保存当前记录`

8. 如需导出，可分别导出：
   - 摘取 Excel
   - 登记 Excel
   - 解析总览 Excel
   - 我司报价一览表
   - 竞对报价对比
   - 历史台账汇总

## 数据存储说明

本项目使用本地文件存储，不依赖数据库。

- AI 配置：
  `data/config.json`

- 历史记录：
  `data/projects/*.json`

- 上传过程中的临时文件：
  `data/tmp/`

## Docker 部署与配置持久化

Docker 部署时必须持久化 `/app/data`，否则容器重建后 AI 配置、飞书配置、历史项目都会丢失。

推荐使用仓库内的 `docker-compose.yml` 启动：

```bash
docker compose pull
docker compose up -d
```

默认会创建 Docker volume `biaoshujiexi_data` 并挂载到容器 `/app/data`。网页端保存的 AI 配置和飞书配置会写入：

```text
/app/data/config.json
```

如果不用 compose，直接 `docker run` 时也要挂载数据卷：

```bash
docker run -d \
  --name biaoshujiexi \
  --restart unless-stopped \
  -p 8008:8008 \
  -v biaoshujiexi_data:/app/data \
  ghcr.io/pwj2333/biaoshujiexi:latest
```

查看容器内配置是否已保存：

```bash
docker exec -it biaoshujiexi ls -l /app/data
docker exec -it biaoshujiexi cat /app/data/config.json
```

## 打包与交付注意事项

如果要把项目交给别人使用，至少需要带上这些内容：

- `app.py`
- `requirements.txt`
- `static/`
- `模板文件/`
- `一键启动.bat`

如果希望保留历史记录，也需要一并带上：

- `data/projects/`

不建议把这些内容直接打包给别人：

- `data/config.json`
  因为里面可能包含真实 API Key

- `招标文件/`
  因为里面通常是测试或业务原始文件

- `.venv/`
  因为 Windows 虚拟环境不能跨电脑稳定复用，交付时不需要带上

## 当前限制

- 只支持可提取文本的 `PDF` / `DOC` / `DOCX` / `XLS` / `XLSX`
- 暂不支持扫描件 OCR
- 不接公司内部运力、港口、货种、客户评级等正式系统
- AI 返回格式不稳定时，后端会自动尝试最多 3 轮 JSON 修复，但不能保证所有异常输出都能恢复
- 多标的拆分能力依赖模型理解，仍建议人工复核

## 建议的交付方式

如果这是要打包给业务同事使用，推荐保留以下结构：

```text
biaoshujiexi/
├─ app.py
├─ requirements.txt
├─ 一键启动.bat
├─ static/
├─ 模板文件/
└─ data/
```

这样后续只需要维护 `模板文件` 和 `data/projects`，不需要业务同事理解代码。
