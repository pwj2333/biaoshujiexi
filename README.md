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
- 通过飞书与配置好的 AI Agent 对话，录入商机货盘、新造船信息、解析标书、查询台账

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

- 飞书 AI Agent：
  普通聊天由 AI 自然回复；业务消息会由 Agent 调用后端白名单工具完成草稿识别、查询、标书解析和确认保存。写入类动作必须先生成草稿，用户明确回复“确认保存/保存”后才会落库。

- 市场情报：
  支持商机货盘和新造船信息录入、AI 识别、市场分析、台账查询和导出。飞书保存的记录会保留来源、open_id、chat_id、message_id、原始消息和确认时间。

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
│  ├─ market_skill/           # 商机货盘和新造船市场情报
│  ├─ feishu/                 # 飞书会话、事件和 Agent 日志
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

如果服务器上使用固定目录保存数据，也可以这样挂载：

```bash
docker rm -f biaoshujiexi || true
docker pull ghcr.io/pwj2333/biaoshujiexi:latest
docker run -d \
  --name biaoshujiexi \
  --restart unless-stopped \
  -p 8008:8008 \
  -v /opt/biaoshujiexi/data:/app/data \
  ghcr.io/pwj2333/biaoshujiexi:latest
```

查看容器内配置是否已保存：

```bash
docker exec -it biaoshujiexi ls -l /app/data
docker exec -it biaoshujiexi cat /app/data/config.json
```

查看容器日志：

```bash
docker logs --tail=200 biaoshujiexi
docker logs -f biaoshujiexi
```

进入容器检查数据目录：

```bash
docker exec -it biaoshujiexi sh
ls -R /app/data
```

## 飞书 Agent 配置

在网页左侧 `AI 配置` 中保存 Base URL、API Key、模型名称。这些配置会永久保存到 `/app/data/config.json`，容器重建后只要 `/app/data` 挂载不变，就不需要重新填写。

在网页左侧 `飞书机器人` 中保存：

- App ID
- App Secret
- Verification Token
- Encrypt Key（如果飞书事件回调启用了加密）
- 接收模式：HTTP 回调或长连接
- 允许 open_id / 允许 chat_id（可选，留空表示不限制）

HTTP 回调地址为：

```text
https://你的域名或公网地址/api/feishu/events
```

飞书消息行为：

- 普通聊天：由 AI 自然回复，不返回固定菜单。
- 商机货盘：生成草稿，用户确认后保存到市场情报台账。
- 新造船：生成草稿，用户确认后保存到新造船台账。
- 标书文件：上传 PDF/DOC/DOCX/XLS/XLSX 并发送“解析标书”，解析结果会进入当前飞书会话，后续可继续追问，也可确认保存为历史项目。
- 查询台账：Agent 会查询系统已有项目和市场情报，不编造不存在的记录。

## Agent 日志

每次飞书 Agent 对话都会写入日志：

```text
/app/data/feishu/agent_logs/YYYYMMDD/*.json
```

日志包含 open_id、chat_id、message_id、用户原文、AI 决策 JSON、工具调用结果、回复内容和错误原因，并隐藏 API Key、飞书密钥、服务器路径等敏感信息。

管理员登录后可通过接口查看或导出：

```text
GET /api/feishu/agent-logs?limit=50
GET /api/feishu/agent-logs?date=20260703&limit=100
GET /api/export/feishu-agent-logs?date=20260703&limit=500
```

## 常见排障

标书在本地可解析、服务器失败时，优先检查：

```bash
docker logs --tail=300 biaoshujiexi
docker exec -it biaoshujiexi ls -l /app/data
docker exec -it biaoshujiexi cat /app/data/config.json
```

常见原因：

- `/app/data` 没有挂载，导致 AI 配置或飞书配置没有持久化。
- 服务器网络无法访问配置的 AI Base URL。
- 上传的是扫描件 PDF，当前版本不支持 OCR。
- 飞书文件权限或 file_key 下载失败。
- 模板文件缺失或容器镜像不是最新版本。

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
