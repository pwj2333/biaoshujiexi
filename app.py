from __future__ import annotations

import io
import json
import re
import uuid
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import httpx
import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from pydantic import BaseModel, Field
from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / 'static'
TEMPLATE_DIR = BASE_DIR / 'templates'
PACKAGE_TEMPLATE_DIR = BASE_DIR / '模板文件'
DATA_DIR = BASE_DIR / 'data'
TMP_DIR = DATA_DIR / 'tmp'
PROJECTS_DIR = DATA_DIR / 'projects'
CONFIG_PATH = DATA_DIR / 'config.json'
PACKAGE_EXTRACTION_TEMPLATE_PATH = PACKAGE_TEMPLATE_DIR / 'extraction_template.xlsx'
PACKAGE_REGISTER_TEMPLATE_PATH = PACKAGE_TEMPLATE_DIR / '招标登记.xlsx'
SESSION_TTL = timedelta(hours=4)

for folder in (DATA_DIR, TMP_DIR, PROJECTS_DIR):
    folder.mkdir(parents=True, exist_ok=True)

app = FastAPI(title='标书解析与登记 Demo')
app.mount('/static', StaticFiles(directory=STATIC_DIR), name='static')

EXTRACTION_FIELDS = [
    {'key': 'open_time', 'label': '开标时间', 'cell': 'B2'},
    {'key': 'submission_method', 'label': '投标文件递交方式', 'cell': 'B3'},
    {'key': 'deposit_amount', 'label': '投标保证金', 'cell': 'B4'},
    {'key': 'deposit_method', 'label': '保证金缴纳方式', 'cell': 'E4'},
    {'key': 'deposit_notes', 'label': '保证金备注', 'cell': 'G4'},
    {'key': 'performance_bond', 'label': '履约保证金', 'cell': 'B5'},
    {'key': 'agency_fee', 'label': '招标代理服务费', 'cell': 'F5'},
    {'key': 'service_period', 'label': '服务期限', 'cell': 'B6'},
    {'key': 'service_scope_note', 'label': '服务范围要求', 'cell': 'B7'},
    {'key': 'qualification_requirements', 'label': '投标人资格要求', 'cell': 'B8'},
    {'key': 'package_rule', 'label': '合同包/标段规则', 'cell': 'B9'},
    {'key': 'vessel_requirements', 'label': '船舶要求', 'cell': 'B10'},
    {'key': 'evaluation_method', 'label': '评标方法', 'cell': 'B11'},
    {'key': 'technical_business', 'label': '技术/商务要求', 'cell': 'B12'},
    {'key': 'quotation', 'label': '报价要求', 'cell': 'B13'},
]

REGISTER_COLUMNS = [
    {'key': 'receive_date', 'label': '接收日期'},
    {'key': 'project_name', 'label': '项目名称'},
    {'key': 'is_awarded', 'label': '是否中标'},
    {'key': 'bid_no', 'label': '招标编号'},
    {'key': 'tenderer', 'label': '招标人'},
    {'key': 'bid_deadline', 'label': '投标截止日期'},
    {'key': 'deposit_amount', 'label': '保证金金额'},
    {'key': 'contract_note', 'label': '保证金退款备注（合同期限）'},
    {'key': 'deposit_return_status', 'label': '保证金退还情况'},
    {'key': 'payment_method', 'label': '缴纳方式'},
    {'key': 'acquisition_method', 'label': '招标文件获取途径'},
    {'key': 'submission_method', 'label': '投标文件递交方式'},
    {'key': 'remark', 'label': '备注'},
]

DEFAULT_CONFIG = {
    'base_url': '',
    'api_key': '',
    'model': '',
    'temperature': 0.1,
}

STATUS_LABELS = {'ok', '已明确满足', '需人工确认', '疑似风险/否决项'}
ELLIPSIS_MARKERS = ('...', '…', '⋯', '。。。')

PARSER_SYSTEM_PROMPT = """
你是航运招投标标书解析助手，只能输出一个 JSON 对象，不能输出 Markdown 或额外解释。

请根据用户上传的招标文件正文，提取标书核心信息，并结合以下公司流程输出分析：
1. 标书解析：投标截止时间、报价要求、资质文件、技术方案、商务条款、评标标准。
2. 业务匹配：只输出“需人工确认项”，不要擅自判断公司运力或运量一定满足。
3. 安全匹配：只输出“需人工确认项”和明显否决风险提示。
4. 询比价场景提示：
   - 老客户 + 已承运货种 + 已靠泊港口：可直接询比价。
   - 新货种 / 新港口：先做适装适靠性评估。
   - 首次合作新客户：先做客户资质及评级审核。
   - 高风险运区：提示上报航运部经理并组织海务、法务、安全联合评估。
5. 航运部客户经理应于接收当日完成《招标文件接收登记表》登记。
6. 如果文档包含多个合同包或标段，请在 register_rows 中按“每个标段一行”返回；无法明确拆分时返回一行。
7. 不确定的字段留空，不要编造。
8. extraction_fields 中的 source_excerpt 必须是招标文件中的原文连续片段，不得使用省略号，不得改写。
9. source_excerpt 的原则是“宁长勿短、宁全勿漏”：
   - 不要只摘一句标题或半句结论。
   - 对于保证金、履约保证金、招标代理费、合同包规则、资格要求、船舶要求、评标办法、技术商务要求、报价要求，必须尽量摘取完整条款块。
   - 如果原文是多行、多项列表、分合同包金额、特别说明、备注、补充条件，必须把相关原文整体保留到 source_excerpt 中，保留换行和分项，不要压缩成一句短语。
   - 如果某字段的依据分散在相邻几段正文中，应合并摘取这些连续原文，优先保留完整条件、金额、时间、适用范围、例外说明。
   - 尤其像“合同包1-8分别对应不同保证金/履约保证金”“资格条件第1-7条”“评标规则9.1/9.2/9.3”“船舶要求及特别说明”这类内容，必须保留完整原文区块，不能只截最前面一行。
10. extraction_fields.value 可以是提炼后的摘要，但 source_excerpt 必须尽可能接近人工摘取标准，详细保留原文上下文，便于直接粘贴到“原文依据”模板页。
11. 如果一个文件中存在多个标的、多个合同包、多个标段、多个航线或多个独立报价单元，必须进行统计并逐项拆分：
   - register_rows 中每个标的/标段/合同包至少返回一行。
   - analysis.summary 中明确说明总共识别到多少个标的/标段/合同包。
   - package_rule 中尽量摘取原文中关于“可兼投兼中、分包、分标段、分航线、分别报价”的完整说明。

JSON 顶层字段必须完整包含：
{
  "document_summary": {
    "project_name": "",
    "bid_no": "",
    "tenderer": "",
    "bid_deadline": "",
    "open_time": "",
    "submission_method": "",
    "deposit_amount": "",
    "service_period": "",
    "qualification_requirements": "",
    "vessel_requirements": "",
    "technical_business": "",
    "quotation": "",
    "evaluation_method": ""
  },
  "extraction_fields": {
    "open_time": {"value": "", "source_excerpt": ""},
    "submission_method": {"value": "", "source_excerpt": ""},
    "deposit_amount": {"value": "", "source_excerpt": ""},
    "deposit_method": {"value": "", "source_excerpt": ""},
    "deposit_notes": {"value": "", "source_excerpt": ""},
    "performance_bond": {"value": "", "source_excerpt": ""},
    "agency_fee": {"value": "", "source_excerpt": ""},
    "service_period": {"value": "", "source_excerpt": ""},
    "service_scope_note": {"value": "", "source_excerpt": ""},
    "qualification_requirements": {"value": "", "source_excerpt": ""},
    "package_rule": {"value": "", "source_excerpt": ""},
    "vessel_requirements": {"value": "", "source_excerpt": ""},
    "evaluation_method": {"value": "", "source_excerpt": ""},
    "technical_business": {"value": "", "source_excerpt": ""},
    "quotation": {"value": "", "source_excerpt": ""}
  },
  "register_rows": [
    {
      "receive_date": "",
      "project_name": "",
      "is_awarded": "",
      "bid_no": "",
      "tenderer": "",
      "bid_deadline": "",
      "deposit_amount": "",
      "contract_note": "",
      "deposit_return_status": "",
      "payment_method": "",
      "acquisition_method": "",
      "submission_method": "",
      "remark": ""
    }
  ],
  "analysis": {
    "summary": "",
    "qualification_files": [],
    "business_points": [],
    "scoring_points": [],
    "risks": []
  },
  "match_review": [
    {
      "area": "",
      "item": "",
      "status": "已明确满足|需人工确认|疑似风险/否决项",
      "reason": "",
      "action": ""
    }
  ]
}
""".strip()

CHAT_SYSTEM_PROMPT = """
你是标书问答助手，只能基于本次标书解析结果和摘取的原文片段回答。
如果问题没有明确依据，直接回答：
未在本次标书中找到明确依据，请人工复核原文。
回答尽量简洁，优先引用结构化结果。
""".strip()

JSON_REPAIR_SYSTEM_PROMPT = """
你是 JSON 修复助手。
你的唯一任务是把用户给出的内容修复成一个严格有效的 JSON 对象。
要求：
1. 只能输出 JSON 对象本身。
2. 不要输出 Markdown，不要输出解释，不要输出多余文字。
3. 保留原始语义，不要擅自新增不存在的业务结论。
4. 如果原文里有换行、引号、列表，按 JSON 字符串规范正确转义。
5. 如果存在尾逗号、智能引号、格式错误、半截内容，尽最大努力修复为合法 JSON。
""".strip()

SESSIONS: dict[str, dict[str, Any]] = {}


class ConfigPayload(BaseModel):
    base_url: str = ''
    api_key: str = ''
    model: str = ''
    temperature: float = 0.1


class ChatPayload(BaseModel):
    session_id: str
    question: str = Field(min_length=1)


class ExportExtractionPayload(BaseModel):
    result: dict[str, Any] = Field(default_factory=dict)
    extraction_fields: dict[str, dict[str, str]] = Field(default_factory=dict)


class ExportRegisterPayload(BaseModel):
    rows: list[dict[str, Any]]
    sheet_name: str = ''


class ExportOverviewPayload(BaseModel):
    result: dict[str, Any]


class ProjectPayload(BaseModel):
    title: str = ''
    source_file_name: str = ''
    register_mode: str = 'packages'
    sheet_name: str = ''
    result: dict[str, Any] = Field(default_factory=dict)


def cleanup_sessions() -> None:
    now = datetime.now()
    expired = [
        session_id
        for session_id, payload in SESSIONS.items()
        if now - payload['created_at'] > SESSION_TTL
    ]
    for session_id in expired:
        SESSIONS.pop(session_id, None)


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return DEFAULT_CONFIG.copy()
    try:
        raw = json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
    except json.JSONDecodeError:
        return DEFAULT_CONFIG.copy()
    return {**DEFAULT_CONFIG, **raw}


def save_config(payload: dict[str, Any]) -> dict[str, Any]:
    current = load_config()
    api_key = payload.get('api_key', '').strip() or current.get('api_key', '')
    cleaned = {
        'base_url': payload.get('base_url', '').strip(),
        'api_key': api_key,
        'model': payload.get('model', '').strip(),
        'temperature': float(payload.get('temperature', 0.1) or 0.1),
    }
    CONFIG_PATH.write_text(
        json.dumps(cleaned, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    return cleaned


def mask_config(config: dict[str, Any]) -> dict[str, Any]:
    masked = config.copy()
    masked['has_api_key'] = bool(masked.get('api_key'))
    if masked.get('api_key'):
        masked['api_key'] = '*' * 8
    return masked


def require_config() -> dict[str, Any]:
    config = load_config()
    if not config['base_url'] or not config['api_key'] or not config['model']:
        raise HTTPException(
            status_code=400,
            detail='请先在左侧 AI 配置中保存 base URL、API Key 和模型名称。',
        )
    return config


def normalize_base_url(base_url: str) -> str:
    value = base_url.rstrip('/')
    if value.endswith('/chat/completions'):
        return value
    if value.endswith('/v1'):
        return f'{value}/chat/completions'
    return f'{value}/v1/chat/completions'


def parse_json_text(content: str) -> dict[str, Any]:
    def strip_code_fence(text: str) -> str:
        text = text.strip()
        if text.startswith('```'):
            text = re.sub(r'^```(?:json)?\s*', '', text)
            text = re.sub(r'\s*```$', '', text)
        return text.strip()

    def extract_balanced_object(text: str) -> str:
        start = text.find('{')
        if start < 0:
            return ''
        depth = 0
        in_string = False
        escape = False
        for index in range(start, len(text)):
            char = text[index]
            if in_string:
                if escape:
                    escape = False
                elif char == '\\':
                    escape = True
                elif char == '"':
                    in_string = False
                continue
            if char == '"':
                in_string = True
            elif char == '{':
                depth += 1
            elif char == '}':
                depth -= 1
                if depth == 0:
                    return text[start : index + 1]
        return text[start:]

    def sanitize_json_candidate(text: str) -> str:
        text = text.replace('\ufeff', '')
        text = (
            text.replace('“', '"')
            .replace('”', '"')
            .replace('‘', '"')
            .replace('’', '"')
        )
        text = re.sub(r',(\s*[}\]])', r'\1', text)
        buffer: list[str] = []
        in_string = False
        escape = False
        for char in text:
            if in_string:
                if escape:
                    buffer.append(char)
                    escape = False
                    continue
                if char == '\\':
                    buffer.append(char)
                    escape = True
                    continue
                if char == '"':
                    buffer.append(char)
                    in_string = False
                    continue
                if char == '\n':
                    buffer.append('\\n')
                    continue
                if char == '\r':
                    buffer.append('\\r')
                    continue
                if char == '\t':
                    buffer.append('\\t')
                    continue
            else:
                if char == '"':
                    in_string = True
            buffer.append(char)
        return ''.join(buffer).strip()

    text = strip_code_fence(content)
    candidates = [text]
    balanced = extract_balanced_object(text)
    if balanced:
        candidates.append(balanced)
    candidates.extend(sanitize_json_candidate(item) for item in list(candidates))

    seen: set[str] = set()
    last_error: Exception | None = None
    for candidate in candidates:
        candidate = candidate.strip()
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        try:
            return json.loads(candidate)
        except json.JSONDecodeError as exc:
            last_error = exc
            continue

    detail = 'AI 返回的 JSON 无法解析。'
    if last_error:
        detail = f'AI 返回的 JSON 无法解析：{last_error.msg}'
    raise HTTPException(status_code=502, detail=detail)


def call_chat_completion(
    config: dict[str, Any],
    *,
    system_prompt: str,
    user_prompt: str,
    temperature: float | None = None,
) -> str:
    payload = {
        'model': config['model'],
        'temperature': config.get('temperature', 0.1) if temperature is None else temperature,
        'messages': [
            {'role': 'system', 'content': system_prompt},
            {'role': 'user', 'content': user_prompt},
        ],
    }
    headers = {
        'Authorization': f"Bearer {config['api_key']}",
        'Content-Type': 'application/json',
    }
    url = normalize_base_url(config['base_url'])
    timeout = httpx.Timeout(connect=30.0, read=180.0, write=60.0, pool=30.0)
    last_error: Exception | None = None
    response = None
    for attempt in range(2):
        try:
            response = httpx.post(url, json=payload, headers=headers, timeout=timeout)
            break
        except (httpx.ConnectTimeout, httpx.ReadTimeout, httpx.WriteTimeout) as exc:
            last_error = exc
            if attempt == 0:
                continue
            raise HTTPException(
                status_code=504,
                detail=f'AI 接口超时：{type(exc).__name__}: {exc}',
            ) from exc
        except httpx.TimeoutException as exc:
            raise HTTPException(
                status_code=504,
                detail=f'AI 接口超时：{type(exc).__name__}: {exc}',
            ) from exc
        except httpx.HTTPError as exc:
            raise HTTPException(
                status_code=502,
                detail=f'AI 接口调用失败：{type(exc).__name__}: {exc}',
            ) from exc
    if response is None:
        raise HTTPException(
            status_code=504,
            detail=f'AI 接口超时：{type(last_error).__name__}: {last_error}' if last_error else 'AI 接口超时，请稍后重试。',
        )
    if response.status_code == 401:
        raise HTTPException(status_code=401, detail='AI 接口鉴权失败，请检查 API Key。')
    if response.status_code >= 400:
        raise HTTPException(
            status_code=502,
            detail=f"AI 接口返回错误：{response.status_code} {response.text[:200]}",
        )
    try:
        data = response.json()
        return data['choices'][0]['message']['content']
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail='AI 接口响应格式不兼容 chat/completions。',
        ) from exc


def clean_text(text: str) -> str:
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def extract_pdf_text(file_path: Path) -> str:
    reader = PdfReader(str(file_path))
    parts = [(page.extract_text() or '').strip() for page in reader.pages]
    return clean_text('\n\n'.join(part for part in parts if part))


def extract_docx_text(file_path: Path) -> str:
    with ZipFile(file_path) as archive:
        try:
            xml_bytes = archive.read('word/document.xml')
        except KeyError as exc:
            raise HTTPException(status_code=400, detail='DOCX 结构无效，无法读取正文。') from exc
    root = ET.fromstring(xml_bytes)
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    parts: list[str] = []
    for paragraph in root.findall('.//w:p', namespace):
        texts = [node.text for node in paragraph.findall('.//w:t', namespace) if node.text]
        line = ''.join(texts).strip()
        if line:
            parts.append(line)
    return clean_text('\n'.join(parts))


def extract_document_text(file_path: Path, suffix: str) -> str:
    if suffix == '.pdf':
        return extract_pdf_text(file_path)
    if suffix == '.docx':
        # ponytail: 只读 document.xml 正文，不处理页眉页脚和图片；要更完整再接专门的 docx 解析器。
        return extract_docx_text(file_path)
    raise HTTPException(status_code=400, detail='当前只支持 PDF 和 DOCX。')


def compact_text(value: Any) -> str:
    text = str(value or '').replace('\u3000', ' ')
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n+', ' ', text)
    return text.strip(' ;；，,')


def compact_paragraph(value: Any) -> str:
    text = str(value or '').replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def normalize_awarded(value: Any) -> str:
    text = compact_text(value)
    lowered = text.lower()
    if lowered in {'√', '是', 'yes', 'y', 'true', '已中标'}:
        return '√'
    if lowered in {'×', '否', 'no', 'n', 'false', '未中标'}:
        return ''
    return text


def normalize_summary(raw: Any) -> dict[str, str]:
    source = raw if isinstance(raw, dict) else {}
    return {
        'project_name': compact_text(source.get('project_name')),
        'bid_no': compact_text(source.get('bid_no')),
        'tenderer': compact_text(source.get('tenderer')),
        'bid_deadline': compact_text(source.get('bid_deadline')),
        'open_time': compact_text(source.get('open_time')),
        'submission_method': compact_text(source.get('submission_method')),
        'deposit_amount': compact_text(source.get('deposit_amount')),
        'service_period': compact_text(source.get('service_period')),
        'qualification_requirements': compact_paragraph(source.get('qualification_requirements')),
        'vessel_requirements': compact_paragraph(source.get('vessel_requirements')),
        'technical_business': compact_paragraph(source.get('technical_business')),
        'quotation': compact_paragraph(source.get('quotation')),
        'evaluation_method': compact_paragraph(source.get('evaluation_method')),
    }


def normalize_extraction_fields(raw: Any) -> dict[str, dict[str, str]]:
    source = raw if isinstance(raw, dict) else {}
    normalized: dict[str, dict[str, str]] = {}
    for field in EXTRACTION_FIELDS:
        item = source.get(field['key'], {})
        if isinstance(item, dict):
            value = compact_paragraph(item.get('value', ''))
            excerpt = compact_paragraph(item.get('source_excerpt', ''))
        else:
            value = compact_paragraph(item)
            excerpt = ''
        normalized[field['key']] = {'value': value, 'source_excerpt': excerpt}
    return normalized


def refine_register_row(row: dict[str, Any], summary: dict[str, str]) -> dict[str, str]:
    return {
        'receive_date': compact_text(row.get('receive_date') or datetime.now().strftime('%Y-%m-%d')),
        'project_name': compact_text(row.get('project_name') or summary['project_name']),
        'is_awarded': normalize_awarded(row.get('is_awarded')),
        'bid_no': compact_text(row.get('bid_no') or summary['bid_no']),
        'tenderer': compact_text(row.get('tenderer') or summary['tenderer']),
        'bid_deadline': compact_text(row.get('bid_deadline') or summary['bid_deadline']),
        'deposit_amount': compact_text(row.get('deposit_amount') or summary['deposit_amount']),
        'contract_note': compact_text(row.get('contract_note') or summary['service_period']),
        'deposit_return_status': compact_text(row.get('deposit_return_status')),
        'payment_method': compact_text(row.get('payment_method')),
        'acquisition_method': compact_text(row.get('acquisition_method')),
        'submission_method': compact_text(row.get('submission_method') or summary['submission_method']),
        'remark': compact_text(row.get('remark')),
    }


def normalize_register_rows(raw: Any, summary: dict[str, str]) -> list[dict[str, str]]:
    rows = raw if isinstance(raw, list) and raw else [{}]
    normalized: list[dict[str, str]] = []
    for row in rows:
        normalized.append(refine_register_row(row if isinstance(row, dict) else {}, summary))
    return normalized


def normalize_analysis(raw: Any) -> dict[str, Any]:
    source = raw if isinstance(raw, dict) else {}
    return {
        'summary': compact_paragraph(source.get('summary')),
        'qualification_files': [compact_text(item) for item in source.get('qualification_files', []) if compact_text(item)],
        'business_points': [compact_text(item) for item in source.get('business_points', []) if compact_text(item)],
        'scoring_points': [compact_text(item) for item in source.get('scoring_points', []) if compact_text(item)],
        'risks': [compact_text(item) for item in source.get('risks', []) if compact_text(item)],
    }


def normalize_match_review(raw: Any) -> list[dict[str, str]]:
    rows = raw if isinstance(raw, list) else []
    normalized: list[dict[str, str]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        status = compact_text(row.get('status') or '需人工确认')
        if status not in STATUS_LABELS:
            status = '需人工确认'
        normalized.append(
            {
                'area': compact_text(row.get('area')),
                'item': compact_text(row.get('item')),
                'status': status,
                'reason': compact_paragraph(row.get('reason')),
                'action': compact_paragraph(row.get('action')),
            }
        )
    return normalized


def merge_register_rows(rows: list[dict[str, Any]]) -> dict[str, Any]:
    if not rows:
        return {column['key']: '' for column in REGISTER_COLUMNS}
    first = rows[0].copy()
    if len(rows) == 1:
        return first
    bid_numbers = [row.get('bid_no', '') for row in rows if row.get('bid_no')]
    deposits = [row.get('deposit_amount', '') for row in rows if row.get('deposit_amount')]
    if bid_numbers:
        first['bid_no'] = ' / '.join(dict.fromkeys(bid_numbers))
    if deposits:
        first['deposit_amount'] = ' / '.join(dict.fromkeys(deposits))
    package_note = f'识别到 {len(rows)} 个标段/合同包'
    remark = first.get('remark', '')
    first['remark'] = package_note if not remark else f'{remark}；{package_note}'
    return first


def normalize_result_payload(data: dict[str, Any]) -> dict[str, Any]:
    summary = normalize_summary(data.get('document_summary'))
    register_rows = normalize_register_rows(data.get('register_rows'), summary)
    document_text_full = compact_paragraph(data.get('document_text_full') or data.get('document_text_excerpt') or '')
    package_count = len(register_rows)
    return {
        'document_summary': summary,
        'extraction_fields': normalize_extraction_fields(data.get('extraction_fields')),
        'register_rows': register_rows,
        'document_register_row': merge_register_rows(register_rows),
        'package_count': package_count,
        'has_multiple_packages': package_count > 1,
        'analysis': normalize_analysis(data.get('analysis')),
        'match_review': normalize_match_review(data.get('match_review')),
        'document_text_full': document_text_full,
        'document_text_excerpt': document_text_full[:12000],
    }


def build_parser_prompt(document_text: str) -> str:
    return (
        '以下是招标文件正文，请按约定 JSON 结构解析。\n'
        '注意：本次抽取要求接近人工摘取表标准，source_excerpt 要尽可能详细完整，尤其不要遗漏多标段金额、分项条件、特别说明、备注、适用范围。\n\n'
        f'{document_text[:50000]}'
    )


def build_json_repair_prompt(raw_content: str, error_detail: str) -> str:
    return (
        '下面是上一次模型输出的内容，但它不是有效 JSON。\n'
        f'解析错误：{error_detail}\n\n'
        '请修复它，并只返回严格有效的 JSON 对象：\n\n'
        f'{raw_content[:60000]}'
    )


def build_chat_prompt(session: dict[str, Any], question: str) -> str:
    return json.dumps(
        {
            'document_summary': session['result']['document_summary'],
            'analysis': session['result']['analysis'],
            'match_review': session['result']['match_review'],
            'register_rows': session['result']['register_rows'],
            'document_text_excerpt': session['result']['document_text_excerpt'],
            'question': question,
        },
        ensure_ascii=False,
    )


def create_session_for_result(result: dict[str, Any]) -> str:
    cleanup_sessions()
    session_id = uuid.uuid4().hex
    SESSIONS[session_id] = {'created_at': datetime.now(), 'result': result}
    return session_id


def parse_ai_document(document_text: str, config: dict[str, Any]) -> dict[str, Any]:
    content = call_chat_completion(
        config,
        system_prompt=PARSER_SYSTEM_PROMPT,
        user_prompt=build_parser_prompt(document_text),
        temperature=min(float(config.get('temperature', 0.1)), 0.3),
    )
    raw = None
    current_content = content
    last_error_detail = 'AI 返回的 JSON 无法解析。'
    for _ in range(4):
        try:
            raw = parse_json_text(current_content)
            break
        except HTTPException as exc:
            last_error_detail = str(exc.detail)
            repair_prompt = build_json_repair_prompt(current_content, last_error_detail)
            current_content = call_chat_completion(
                config,
                system_prompt=JSON_REPAIR_SYSTEM_PROMPT,
                user_prompt=repair_prompt,
                temperature=0,
            )
    if raw is None:
        raise HTTPException(status_code=502, detail=f'AI 连续 3 轮 JSON 修复仍失败：{last_error_detail}')
    result = normalize_result_payload(raw)
    result['document_text_full'] = document_text
    result['document_text_excerpt'] = document_text[:12000]
    return result


def parse_excel_date(value: str) -> Any:
    text = compact_text(value)
    if not text:
        return ''
    for fmt in (
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%Y.%m.%d',
        '%Y年%m月%d日',
        '%Y-%m-%d %H:%M',
        '%Y/%m/%d %H:%M',
        '%Y年%m月%d日 %H:%M',
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return text


def parse_excel_amount(value: str) -> Any:
    text = compact_text(value)
    if not text:
        return ''
    plain = text.replace(',', '').replace('，', '')
    if re.fullmatch(r'-?\d+(?:\.\d+)?', plain):
        return float(plain) if '.' in plain else int(plain)
    match = re.fullmatch(r'(-?\d+(?:\.\d+)?)\s*万(?:元)?', plain)
    if match:
        number = float(match.group(1)) * 10000
        return int(number) if number.is_integer() else number
    match = re.fullmatch(r'(-?\d+(?:\.\d+)?)\s*元', plain)
    if match:
        number = float(match.group(1))
        return int(number) if number.is_integer() else number
    return text


def find_register_start_row(sheet: Any) -> int:
    row = 3
    while sheet[f'A{row}'].value not in (None, ''):
        row += 1
    return row


def get_register_template_path() -> Path:
    if not PACKAGE_REGISTER_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f'缺少登记模板文件：{PACKAGE_REGISTER_TEMPLATE_PATH.name}，请放到 模板文件 目录。',
        )
    return PACKAGE_REGISTER_TEMPLATE_PATH


def get_extraction_template_path() -> Path:
    if not PACKAGE_EXTRACTION_TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f'缺少摘取模板文件：{PACKAGE_EXTRACTION_TEMPLATE_PATH.name}，请放到 模板文件 目录。',
        )
    return PACKAGE_EXTRACTION_TEMPLATE_PATH


def ensure_register_sheet(workbook: Any, sheet_name: str) -> Any:
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    source_sheet = workbook[workbook.sheetnames[0]]
    cloned = workbook.copy_worksheet(source_sheet)
    cloned.title = sheet_name
    return cloned


def apply_register_row_style(sheet: Any, row_index: int) -> None:
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for column in range(1, 15):
        cell = sheet.cell(row_index, column)
        cell.border = border
        cell.alignment = center
    sheet[f'B{row_index}'].number_format = 'yyyy/m/d'
    sheet[f'G{row_index}'].number_format = 'yyyy/m/d h:mm'
    if isinstance(sheet[f'H{row_index}'].value, (int, float)):
        sheet[f'H{row_index}'].number_format = '#,##0.00'


def project_path(project_id: str) -> Path:
    if not re.fullmatch(r'[a-f0-9]{32}', project_id):
        raise HTTPException(status_code=400, detail='项目 ID 无效。')
    return PROJECTS_DIR / f'{project_id}.json'


def project_meta(project: dict[str, Any]) -> dict[str, Any]:
    summary = project.get('result', {}).get('document_summary', {})
    return {
        'project_id': project['project_id'],
        'title': project.get('title', ''),
        'source_file_name': project.get('source_file_name', ''),
        'register_mode': project.get('register_mode', 'packages'),
        'sheet_name': project.get('sheet_name', ''),
        'created_at': project.get('created_at', ''),
        'updated_at': project.get('updated_at', ''),
        'project_name': summary.get('project_name', ''),
        'bid_no': summary.get('bid_no', ''),
    }


def read_project_file(path: Path) -> dict[str, Any] | None:
    try:
        raw = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None
    project_id = str(raw.get('project_id') or path.stem)
    return {
        'project_id': project_id,
        'title': compact_text(raw.get('title')),
        'source_file_name': compact_text(raw.get('source_file_name')),
        'register_mode': compact_text(raw.get('register_mode') or 'packages') or 'packages',
        'sheet_name': compact_text(raw.get('sheet_name')),
        'created_at': compact_text(raw.get('created_at')),
        'updated_at': compact_text(raw.get('updated_at')),
        'result': normalize_result_payload(raw.get('result') or {}),
    }


def load_project(project_id: str) -> dict[str, Any]:
    path = project_path(project_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail='项目不存在。')
    project = read_project_file(path)
    if not project:
        raise HTTPException(status_code=500, detail='项目文件损坏，无法读取。')
    return project


def save_project(project_id: str, payload: ProjectPayload, *, created_at: str | None = None) -> dict[str, Any]:
    now = datetime.now().isoformat(timespec='seconds')
    result = normalize_result_payload(payload.result)
    title = compact_text(payload.title) or result['document_summary']['project_name'] or '未命名项目'
    project = {
        'project_id': project_id,
        'title': title,
        'source_file_name': compact_text(payload.source_file_name),
        'register_mode': payload.register_mode if payload.register_mode in {'packages', 'document'} else 'packages',
        'sheet_name': compact_text(payload.sheet_name),
        'created_at': created_at or now,
        'updated_at': now,
        'result': result,
    }
    project_path(project_id).write_text(
        json.dumps(project, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    return project


def build_overview_text(result: dict[str, Any]) -> str:
    summary = result['document_summary']
    analysis = result['analysis']
    rows = result['register_rows']
    review = result['match_review']

    def block(title: str, items: list[str]) -> str:
        lines = [item for item in items if item]
        return f"{title}\n" + ('\n'.join(f'- {item}' for item in lines) if lines else '- 无')

    sections = [
        '# 标书解析总览',
        '',
        '## 核心信息',
        f"项目名称：{summary['project_name'] or '未识别'}",
        f"招标编号：{summary['bid_no'] or '未识别'}",
        f"招标人：{summary['tenderer'] or '未识别'}",
        f"投标截止时间：{summary['bid_deadline'] or '未识别'}",
        f"开标时间：{summary['open_time'] or '未识别'}",
        f"投标文件递交方式：{summary['submission_method'] or '未识别'}",
        f"保证金金额：{summary['deposit_amount'] or '未识别'}",
        f"服务期限：{summary['service_period'] or '未识别'}",
        '',
        '## AI 解析',
        analysis['summary'] or '无',
        '',
        block('## 资格文件清单', analysis['qualification_files']),
        '',
        block('## 报价与商务关注点', analysis['business_points']),
        '',
        block('## 评标得分关注点', analysis['scoring_points']),
        '',
        block('## 风险提示', analysis['risks']),
        '',
        '## 登记预览',
    ]

    if rows:
        for index, row in enumerate(rows, start=1):
            sections.extend(
                [
                    f"{index}. 项目名称：{row.get('project_name') or '未识别'}",
                    f"   招标编号：{row.get('bid_no') or '未识别'}",
                    f"   招标人：{row.get('tenderer') or '未识别'}",
                    f"   投标截止日期：{row.get('bid_deadline') or '未识别'}",
                    f"   保证金金额：{row.get('deposit_amount') or '未识别'}",
                    f"   备注：{row.get('remark') or '-'}",
                ]
            )
    else:
        sections.append('无')

    sections.extend(['', '## 匹配复核'])
    if review:
        for item in review:
            sections.extend(
                [
                    f"- 领域：{item.get('area') or '未分类'}",
                    f"  事项：{item.get('item') or '-'}",
                    f"  状态：{item.get('status') or '需人工确认'}",
                    f"  原因：{item.get('reason') or '-'}",
                    f"  建议动作：{item.get('action') or '-'}",
                ]
            )
    else:
        sections.append('- 无')
    return '\n'.join(sections).strip() + '\n'


def clear_sheet_rows(sheet: Any, start_row: int) -> None:
    if sheet.max_row >= start_row:
        sheet.delete_rows(start_row, sheet.max_row - start_row + 1)


def resolve_source_excerpt(value: str, excerpt: str, document_text_full: str) -> str:
    clean_excerpt = compact_paragraph(excerpt)
    if clean_excerpt and not any(marker in clean_excerpt for marker in ELLIPSIS_MARKERS):
        return clean_excerpt
    if not value or not document_text_full:
        return clean_excerpt
    index = document_text_full.find(value)
    if index < 0:
        match = SequenceMatcher(None, value, document_text_full).find_longest_match(
            0,
            len(value),
            0,
            len(document_text_full),
        )
        # ponytail: 只做一次最长公共片段回退；还不够稳时再上专门的模糊匹配。
        if match.size < 4:
            return clean_excerpt
        index = match.b
    start = document_text_full.rfind('\n', 0, index)
    end = document_text_full.find('\n', index + len(value))
    start = 0 if start < 0 else start + 1
    end = len(document_text_full) if end < 0 else end
    line = compact_paragraph(document_text_full[start:end])
    return line or clean_excerpt


def project_register_rows(rows: list[dict[str, Any]]) -> list[dict[str, str]]:
    return [refine_register_row(row if isinstance(row, dict) else {}, normalize_summary({})) for row in rows]


@app.get('/')
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / 'index.html')


@app.get('/api/template-meta')
def template_meta() -> dict[str, Any]:
    workbook = load_workbook(get_register_template_path(), read_only=True)
    sheet_names = workbook.sheetnames
    workbook.close()
    current_year = str(datetime.now().year)
    return {
        'extraction_fields': EXTRACTION_FIELDS,
        'register_columns': REGISTER_COLUMNS,
        'register_sheets': sheet_names,
        'default_sheet': current_year,
    }


@app.get('/api/config')
def get_config() -> dict[str, Any]:
    return mask_config(load_config())


@app.post('/api/config')
def update_config(payload: ConfigPayload) -> dict[str, Any]:
    return mask_config(save_config(payload.model_dump()))


@app.post('/api/config/test')
def test_config(payload: ConfigPayload) -> dict[str, Any]:
    config = payload.model_dump()
    if not config['base_url'] or not config['api_key'] or not config['model']:
        raise HTTPException(status_code=400, detail='请先填写 URL、API Key 和模型名称。')
    content = call_chat_completion(
        config,
        system_prompt='你是连通性测试助手，只返回 OK。',
        user_prompt='回复 OK',
        temperature=0,
    )
    return {'ok': 'OK' in content.upper(), 'message': content.strip()}


@app.get('/api/projects')
def list_projects(q: str = '') -> dict[str, Any]:
    keyword = q.strip().lower()
    items: list[dict[str, Any]] = []
    for path in PROJECTS_DIR.glob('*.json'):
        project = read_project_file(path)
        if not project:
            continue
        meta = project_meta(project)
        text = ' '.join(
            [
                meta.get('title', ''),
                meta.get('project_name', ''),
                meta.get('bid_no', ''),
                meta.get('source_file_name', ''),
            ]
        ).lower()
        if keyword and keyword not in text:
            continue
        items.append(meta)
    items.sort(key=lambda item: item.get('updated_at', ''), reverse=True)
    return {'items': items}


@app.post('/api/projects')
def create_project(payload: ProjectPayload) -> dict[str, Any]:
    project = save_project(uuid.uuid4().hex, payload)
    return {**project_meta(project), 'project': project}


@app.get('/api/projects/{project_id}')
def get_project(project_id: str) -> dict[str, Any]:
    project = load_project(project_id)
    session_id = create_session_for_result(project['result'])
    return {**project, 'session_id': session_id}


@app.put('/api/projects/{project_id}')
def update_project(project_id: str, payload: ProjectPayload) -> dict[str, Any]:
    existing = load_project(project_id)
    project = save_project(project_id, payload, created_at=existing.get('created_at'))
    return {**project_meta(project), 'project': project}


@app.delete('/api/projects/{project_id}')
def delete_project(project_id: str) -> dict[str, Any]:
    path = project_path(project_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail='项目不存在。')
    path.unlink()
    return {'ok': True}


@app.post('/api/parse')
async def parse_bid(file: UploadFile = File(...)) -> dict[str, Any]:
    cleanup_sessions()
    suffix = Path(file.filename or '').suffix.lower()
    if suffix not in {'.pdf', '.docx'}:
        raise HTTPException(status_code=400, detail='当前只支持文本型 PDF 或 DOCX。')
    temp_path = TMP_DIR / f'{uuid.uuid4().hex}{suffix}'
    temp_path.write_bytes(await file.read())
    try:
        document_text = extract_document_text(temp_path, suffix)
    finally:
        temp_path.unlink(missing_ok=True)
    if len(document_text.strip()) < 80:
        raise HTTPException(
            status_code=400,
            detail='当前版本仅支持可提取文本的 PDF/DOCX，扫描件或空白文件暂不支持。',
        )
    result = parse_ai_document(document_text, require_config())
    session_id = create_session_for_result(result)
    return {'session_id': session_id, 'source_file_name': file.filename or '', **result}


@app.post('/api/chat')
def chat(payload: ChatPayload) -> dict[str, Any]:
    cleanup_sessions()
    session = SESSIONS.get(payload.session_id)
    if not session:
        raise HTTPException(status_code=404, detail='当前会话已失效，请重新解析或重新打开记录。')
    answer = call_chat_completion(
        require_config(),
        system_prompt=CHAT_SYSTEM_PROMPT,
        user_prompt=build_chat_prompt(session, payload.question),
        temperature=0.2,
    )
    return {'answer': answer.strip()}


@app.post('/api/export/extraction')
def export_extraction(payload: ExportExtractionPayload) -> StreamingResponse:
    raw_result = payload.result or {'extraction_fields': payload.extraction_fields}
    result = normalize_result_payload(raw_result)
    workbook = load_workbook(get_extraction_template_path())
    value_sheet = workbook[workbook.sheetnames[0]]
    value_sheet.title = '摘取结果'
    for field in EXTRACTION_FIELDS:
        item = result['extraction_fields'].get(field['key'], {})
        value_sheet[field['cell']] = compact_paragraph(item.get('value', ''))

    evidence_sheet = workbook.copy_worksheet(value_sheet)
    evidence_sheet.title = '原文依据'
    for field in EXTRACTION_FIELDS:
        item = result['extraction_fields'].get(field['key'], {})
        evidence_sheet[field['cell']] = resolve_source_excerpt(
            compact_paragraph(item.get('value', '')),
            compact_paragraph(item.get('source_excerpt', '')),
            result.get('document_text_full', ''),
        )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    filename = f"bid_extraction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/api/export/register')
def export_register(payload: ExportRegisterPayload) -> StreamingResponse:
    workbook = load_workbook(get_register_template_path())
    sheet_name = compact_text(payload.sheet_name) or str(datetime.now().year)
    sheet = ensure_register_sheet(workbook, sheet_name)
    row_index = find_register_start_row(sheet)
    previous_no = sheet[f'A{row_index - 1}'].value if row_index > 3 else 0
    try:
        serial = int(previous_no or 0)
    except (TypeError, ValueError):
        serial = row_index - 3

    for offset, row in enumerate(payload.rows, start=1):
        current_row = row_index + offset - 1
        serial += 1
        clean_row = refine_register_row(row if isinstance(row, dict) else {}, normalize_summary({}))
        sheet[f'A{current_row}'] = serial
        sheet[f'B{current_row}'] = parse_excel_date(clean_row.get('receive_date', ''))
        sheet[f'C{current_row}'] = clean_row.get('project_name', '')
        sheet[f'D{current_row}'] = clean_row.get('is_awarded', '')
        sheet[f'E{current_row}'] = clean_row.get('bid_no', '')
        sheet[f'F{current_row}'] = clean_row.get('tenderer', '')
        sheet[f'G{current_row}'] = parse_excel_date(clean_row.get('bid_deadline', ''))
        sheet[f'H{current_row}'] = parse_excel_amount(clean_row.get('deposit_amount', ''))
        sheet[f'I{current_row}'] = clean_row.get('contract_note', '')
        sheet[f'J{current_row}'] = clean_row.get('deposit_return_status', '')
        sheet[f'K{current_row}'] = clean_row.get('payment_method', '')
        sheet[f'L{current_row}'] = clean_row.get('acquisition_method', '')
        sheet[f'M{current_row}'] = clean_row.get('submission_method', '')
        sheet[f'N{current_row}'] = clean_row.get('remark', '')
        apply_register_row_style(sheet, current_row)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    filename = f"bid_register_{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/api/export/overview')
def export_overview(payload: ExportOverviewPayload) -> StreamingResponse:
    result = normalize_result_payload(payload.result)
    output = io.BytesIO(build_overview_text(result).encode('utf-8'))
    filename = f"bid_overview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
    return StreamingResponse(
        output,
        media_type='text/markdown; charset=utf-8',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.exception_handler(HTTPException)
async def http_error_handler(_: Any, exc: HTTPException) -> JSONResponse:
    return JSONResponse(status_code=exc.status_code, content={'detail': exc.detail})


def _self_check() -> None:
    sample = parse_json_text('```json {"ok": true} ```')
    assert sample['ok'] is True
    merged = merge_register_rows(
        [
            {'project_name': '项目A', 'bid_no': '01', 'deposit_amount': '1000', 'remark': ''},
            {'project_name': '项目A', 'bid_no': '02', 'deposit_amount': '2000', 'remark': ''},
        ]
    )
    assert '02' in merged['bid_no']
    assert '2 个标段' in merged['remark']
    assert resolve_source_excerpt('电子化股份有限公司', '...', '福建省东南电化股份有限公司') == '福建省东南电化股份有限公司'
    project = save_project(
        '0' * 32,
        ProjectPayload(
            title='测试项目',
            result={
                'document_summary': {'project_name': '测试项目'},
                'register_rows': [{'project_name': '测试项目'}],
                'document_text_full': '测试项目原文',
            },
        ),
    )
    assert project['title'] == '测试项目'
    assert project_path('0' * 32).exists()
    project_path('0' * 32).unlink(missing_ok=True)


if __name__ == '__main__':
    _self_check()
    uvicorn.run('app:app', host='127.0.0.1', port=8008, reload=False, app_dir=str(BASE_DIR))
