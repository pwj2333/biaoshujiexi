from __future__ import annotations

import io
import json
import hashlib
import hmac
import re
import subprocess
import uuid
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

import httpx
import uvicorn
from fastapi import Cookie, Depends, FastAPI, File, Header, HTTPException, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
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
USERS_PATH = DATA_DIR / 'users.json'
PACKAGE_EXTRACTION_TEMPLATE_PATH = PACKAGE_TEMPLATE_DIR / 'extraction_template.xlsx'
PACKAGE_REGISTER_TEMPLATE_PATH = PACKAGE_TEMPLATE_DIR / '招标登记.xlsx'
SESSION_TTL = timedelta(hours=4)
AUTH_TTL = timedelta(hours=12)
PASSWORD_SALT = 'ruico-bid-parser'
DEFAULT_ADMIN_USERNAME = 'ruico'
DEFAULT_ADMIN_PASSWORD = 'Ruico668@'
AUTH_COOKIE_NAME = 'bid_parser_token'

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
BID_STATUS_OPTIONS = ('待跟进', '准备投标', '已投标', '放弃', '未投标')
AWARD_STATUS_OPTIONS = ('未知', '待定', '已中标', '未中标')
TIMELINE_TYPE_OPTIONS = ('parse', 'quote', 'award', 'note')

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
AUTH_SESSIONS: dict[str, dict[str, Any]] = {}


def hash_password(password: str) -> str:
    # ponytail: 先用 salted sha256 落地单机 demo；要上公网时再换成 bcrypt/argon2。
    return hashlib.sha256(f'{PASSWORD_SALT}:{password}'.encode('utf-8')).hexdigest()


def default_admin_user() -> dict[str, Any]:
    return {
        'username': DEFAULT_ADMIN_USERNAME,
        'password_hash': hash_password(DEFAULT_ADMIN_PASSWORD),
        'display_name': '系统管理员',
        'role': 'admin',
        'created_at': datetime.now().isoformat(timespec='seconds'),
    }


def load_users() -> dict[str, Any]:
    if not USERS_PATH.exists():
        payload = {'users': [default_admin_user()]}
        USERS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
        return payload
    try:
        payload = json.loads(USERS_PATH.read_text(encoding='utf-8'))
    except json.JSONDecodeError:
        payload = {'users': [default_admin_user()]}
        USERS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
        return payload
    users = payload.get('users') if isinstance(payload, dict) else None
    if not isinstance(users, list):
        payload = {'users': [default_admin_user()]}
        USERS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
        return payload
    if not any((item or {}).get('username') == DEFAULT_ADMIN_USERNAME for item in users if isinstance(item, dict)):
        users.insert(0, default_admin_user())
        USERS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return payload


def save_users(payload: dict[str, Any]) -> dict[str, Any]:
    USERS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return payload


def find_user(username: str) -> dict[str, Any] | None:
    username = compact_text(username).lower()
    for user in load_users().get('users', []):
        if compact_text((user or {}).get('username', '')).lower() == username:
            return user
    return None


def sanitize_user(user: dict[str, Any]) -> dict[str, Any]:
    return {
        'username': compact_text(user.get('username', '')),
        'display_name': compact_text(user.get('display_name', '')),
        'role': compact_text(user.get('role', 'user')) or 'user',
        'created_at': compact_text(user.get('created_at', '')),
    }


def cleanup_auth_sessions() -> None:
    now = datetime.now()
    expired = [
        token
        for token, payload in AUTH_SESSIONS.items()
        if now - payload['created_at'] > AUTH_TTL
    ]
    for token in expired:
        AUTH_SESSIONS.pop(token, None)


def create_auth_session(user: dict[str, Any]) -> dict[str, Any]:
    cleanup_auth_sessions()
    token = uuid.uuid4().hex
    payload = {
        'token': token,
        'user': sanitize_user(user),
        'created_at': datetime.now(),
    }
    AUTH_SESSIONS[token] = payload
    return {'token': token, 'user': payload['user']}


def extract_auth_token(authorization: str | None = None, auth_cookie: str | None = None) -> str:
    if authorization and authorization.startswith('Bearer '):
        return authorization.split(' ', 1)[1].strip()
    return compact_text(auth_cookie)


def require_auth(
    authorization: str | None = Header(default=None),
    auth_cookie: str | None = Cookie(default=None, alias=AUTH_COOKIE_NAME),
) -> dict[str, Any]:
    cleanup_auth_sessions()
    token = extract_auth_token(authorization, auth_cookie)
    if not token:
        raise HTTPException(status_code=401, detail='请先登录。')
    session = AUTH_SESSIONS.get(token)
    if not session:
        raise HTTPException(status_code=401, detail='登录已失效，请重新登录。')
    return session['user']


def require_admin(
    authorization: str | None = Header(default=None),
    auth_cookie: str | None = Cookie(default=None, alias=AUTH_COOKIE_NAME),
) -> dict[str, Any]:
    user = require_auth(authorization, auth_cookie)
    if user.get('role') != 'admin':
        raise HTTPException(status_code=403, detail='只有管理员可以执行该操作。')
    return user


def resolve_auth_user(authorization: str | None = None, auth_cookie: str | None = None) -> dict[str, Any] | None:
    cleanup_auth_sessions()
    token = extract_auth_token(authorization, auth_cookie)
    if not token:
        return None
    session = AUTH_SESSIONS.get(token)
    if not session:
        return None
    return session['user']


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
    title: str = ''
    result: dict[str, Any]
    follow_up: dict[str, Any] = Field(default_factory=dict)
    our_quotes: list[dict[str, Any]] = Field(default_factory=list)
    competitor_quotes: list[dict[str, Any]] = Field(default_factory=list)
    timeline: list[dict[str, Any]] = Field(default_factory=list)


class ExportQuotePayload(BaseModel):
    title: str = ''
    follow_up: dict[str, Any] = Field(default_factory=dict)
    rows: list[dict[str, Any]] = Field(default_factory=list)


class ExportLedgerPayload(BaseModel):
    q: str = ''
    year: str = ''
    award_status: str = ''
    bid_status: str = ''


class ProjectPayload(BaseModel):
    title: str = ''
    source_file_name: str = ''
    register_mode: str = 'packages'
    sheet_name: str = ''
    result: dict[str, Any] = Field(default_factory=dict)
    follow_up: dict[str, Any] = Field(default_factory=dict)
    our_quotes: list[dict[str, Any]] = Field(default_factory=list)
    competitor_quotes: list[dict[str, Any]] = Field(default_factory=list)
    timeline: list[dict[str, Any]] = Field(default_factory=list)


class LoginPayload(BaseModel):
    username: str = Field(min_length=1)
    password: str = Field(min_length=1)


class CreateUserPayload(BaseModel):
    username: str = Field(min_length=1)
    password: str = Field(min_length=1)
    display_name: str = ''
    role: str = 'user'


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
    timeout = httpx.Timeout(connect=30.0, read=300.0, write=60.0, pool=30.0)
    response = None
    try:
        response = httpx.post(url, json=payload, headers=headers, timeout=timeout)
    except httpx.ConnectTimeout as exc:
        # ponytail: 仅对连接阶段重试一次；读超时不重试，避免模型已成功执行却重复扣费。
        try:
            response = httpx.post(url, json=payload, headers=headers, timeout=timeout)
        except httpx.ConnectTimeout as retry_exc:
            raise HTTPException(
                status_code=504,
                detail=f'AI 接口连接超时：{type(retry_exc).__name__}: {retry_exc}',
            ) from retry_exc
    except (httpx.ReadTimeout, httpx.WriteTimeout) as exc:
        raise HTTPException(
            status_code=504,
            detail=(
                f'AI 接口响应超时：{type(exc).__name__}: {exc}。'
                '请求可能已到达模型侧并正在处理，请避免立即重复提交，可稍后重试或缩短解析内容。'
            ),
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
            detail='AI 接口超时，请稍后重试。',
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
    try:
        with ZipFile(file_path) as archive:
            try:
                xml_bytes = archive.read('word/document.xml')
            except KeyError as exc:
                raise HTTPException(status_code=400, detail='DOCX 结构无效，无法读取正文。') from exc
    except BadZipFile as exc:
        raise HTTPException(
            status_code=400,
            detail='上传的 DOCX 文件格式无效，当前文件不是可解析的 Office 文档，请重新导出为标准 DOCX 后再上传。',
        ) from exc
    root = ET.fromstring(xml_bytes)
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    parts: list[str] = []
    for paragraph in root.findall('.//w:p', namespace):
        texts = [node.text for node in paragraph.findall('.//w:t', namespace) if node.text]
        line = ''.join(texts).strip()
        if line:
            parts.append(line)
    return clean_text('\n'.join(parts))


def extract_xlsx_text(file_path: Path) -> str:
    try:
        workbook = load_workbook(file_path, data_only=True)
    except (BadZipFile, InvalidFileException, OSError) as exc:
        raise HTTPException(
            status_code=400,
            detail='上传的 Excel 文件格式无效，当前文件不是可解析的标准 Excel 文档，请重新保存为 xlsx 后再上传。',
        ) from exc
    parts: list[str] = []
    for sheet in workbook.worksheets:
        sheet_lines: list[str] = []
        for row in sheet.iter_rows(values_only=True):
            cells = [compact_text(cell) for cell in row if compact_text(cell)]
            if cells:
                sheet_lines.append(' | '.join(cells))
        if sheet_lines:
            parts.append(f'工作表：{sheet.title}')
            parts.extend(sheet_lines)
    workbook.close()
    return clean_text('\n'.join(parts))


def extract_with_office_com(file_path: Path, kind: str) -> str:
    escaped = str(file_path).replace("'", "''")
    if kind == 'word':
        script = rf"""
$ErrorActionPreference = 'Stop'
$path = '{escaped}'
$word = $null
$doc = $null
try {{
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($path, $false, $true)
  $text = $doc.Content.Text
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output $text
}} finally {{
  if ($doc) {{ $doc.Close() }}
  if ($word) {{ $word.Quit() }}
}}
""".strip()
    else:
        script = rf"""
$ErrorActionPreference = 'Stop'
$path = '{escaped}'
$excel = $null
$workbook = $null
try {{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($path, 0, $true)
  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($sheet in $workbook.Worksheets) {{
    $lines.Add("工作表：$($sheet.Name)")
    $range = $sheet.UsedRange
    $rowCount = $range.Rows.Count
    $colCount = $range.Columns.Count
    for ($r = 1; $r -le $rowCount; $r++) {{
      $cells = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $colCount; $c++) {{
        $value = $range.Item($r, $c).Text
        if ($value) {{ $cells.Add($value.Trim()) }}
      }}
      if ($cells.Count -gt 0) {{ $lines.Add(($cells -join ' | ')) }}
    }}
  }}
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  Write-Output ($lines -join [Environment]::NewLine)
}} finally {{
  if ($workbook) {{ $workbook.Close($false) }}
  if ($excel) {{ $excel.Quit() }}
}}
""".strip()
    try:
        completed = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script],
            capture_output=True,
            text=True,
            encoding='utf-8',
            timeout=120,
            check=False,
        )
    except Exception as exc:
        raise HTTPException(
            status_code=400,
            detail='当前环境无法调用 Office 解析该文件，请优先上传 PDF、DOCX 或 XLSX。',
        ) from exc
    if completed.returncode != 0:
        raise HTTPException(
            status_code=400,
            detail='当前环境无法解析该 Office 文件。若本机未安装 Word/Excel，请先另存为 DOCX/XLSX/PDF 后再上传。',
        )
    return clean_text(completed.stdout)


def extract_doc_text(file_path: Path) -> str:
    return extract_with_office_com(file_path, 'word')


def extract_xls_text(file_path: Path) -> str:
    return extract_with_office_com(file_path, 'excel')


def extract_document_text(file_path: Path, suffix: str) -> str:
    if suffix == '.pdf':
        return extract_pdf_text(file_path)
    if suffix == '.docx':
        # ponytail: 只读 document.xml 正文，不处理页眉页脚和图片；要更完整再接专门的 docx 解析器。
        return extract_docx_text(file_path)
    if suffix == '.doc':
        # ponytail: 旧版 doc 直接走本机 Word COM；无 Office 时引导另存为 docx/pdf。
        return extract_doc_text(file_path)
    if suffix == '.xlsx':
        return extract_xlsx_text(file_path)
    if suffix == '.xls':
        # ponytail: 旧版 xls 直接走本机 Excel COM；无 Office 时引导另存为 xlsx/pdf。
        return extract_xls_text(file_path)
    raise HTTPException(status_code=400, detail='当前只支持 PDF、DOC、DOCX、XLS、XLSX。')


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


def normalize_follow_up(raw: Any, *, sheet_name: str = '') -> dict[str, str]:
    source = raw if isinstance(raw, dict) else {}
    bid_status = compact_text(source.get('bid_status') or '待跟进')
    award_status = compact_text(source.get('award_status') or '未知')
    if bid_status not in BID_STATUS_OPTIONS:
        bid_status = '待跟进'
    if award_status not in AWARD_STATUS_OPTIONS:
        award_status = '未知'
    return {
        'bid_status': bid_status,
        'award_status': award_status,
        'award_date': compact_text(source.get('award_date')),
        'award_company': compact_text(source.get('award_company')),
        'our_award_amount': compact_text(source.get('our_award_amount')),
        'competitor_award_amount': compact_text(source.get('competitor_award_amount')),
        'tracking_note': compact_paragraph(source.get('tracking_note')),
        'information_source': compact_text(source.get('information_source')),
        'register_year': compact_text(source.get('register_year') or sheet_name),
    }


def normalize_quote_row(raw: Any, *, default_company: str = '') -> dict[str, Any]:
    source = raw if isinstance(raw, dict) else {}
    try:
        round_no = max(1, int(source.get('round_no') or 1))
    except (TypeError, ValueError):
        round_no = 1
    return {
        'id': compact_text(source.get('id') or uuid.uuid4().hex[:12]),
        'package_name': compact_text(source.get('package_name')),
        'round_no': round_no,
        'quote_date': compact_text(source.get('quote_date')),
        'quote_company': compact_text(source.get('quote_company') or default_company),
        'currency': compact_text(source.get('currency') or 'CNY'),
        'tax_mode': compact_text(source.get('tax_mode')),
        'unit_price': compact_text(source.get('unit_price')),
        'total_price': compact_text(source.get('total_price')),
        'ranking': compact_text(source.get('ranking')),
        'is_submitted': bool(source.get('is_submitted')),
        'is_awarded': bool(source.get('is_awarded')),
        'source': compact_text(source.get('source')),
        'remark': compact_paragraph(source.get('remark')),
    }


def normalize_quote_rows(raw: Any, *, default_company: str = '') -> list[dict[str, Any]]:
    rows = raw if isinstance(raw, list) else []
    return [normalize_quote_row(row, default_company=default_company) for row in rows]


def normalize_timeline_rows(raw: Any) -> list[dict[str, str]]:
    rows = raw if isinstance(raw, list) else []
    normalized: list[dict[str, str]] = []
    for row in rows:
        source = row if isinstance(row, dict) else {}
        item_type = compact_text(source.get('type') or 'note')
        if item_type not in TIMELINE_TYPE_OPTIONS:
            item_type = 'note'
        note = compact_paragraph(source.get('note'))
        if not note:
            continue
        normalized.append(
            {
                'id': compact_text(source.get('id') or uuid.uuid4().hex[:12]),
                'date': compact_text(source.get('date') or datetime.now().strftime('%Y-%m-%d')),
                'type': item_type,
                'note': note,
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


def build_auto_timeline(result: dict[str, Any], source_file_name: str = '') -> list[dict[str, str]]:
    note = '完成标书解析并生成摘取结果'
    if source_file_name:
        note = f'完成标书解析：{source_file_name}'
    return [
        {
            'id': uuid.uuid4().hex[:12],
            'date': datetime.now().strftime('%Y-%m-%d'),
            'type': 'parse',
            'note': note,
        }
    ]


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
    for repair_round in range(3):
        try:
            raw = parse_json_text(current_content)
            break
        except HTTPException as exc:
            last_error_detail = str(exc.detail)
            if repair_round == 2:
                break
            repair_prompt = build_json_repair_prompt(current_content, last_error_detail)
            try:
                current_content = call_chat_completion(
                    config,
                    system_prompt=JSON_REPAIR_SYSTEM_PROMPT,
                    user_prompt=repair_prompt,
                    temperature=0,
                )
            except HTTPException as repair_exc:
                raise HTTPException(
                    status_code=repair_exc.status_code,
                    detail=f'AI 首轮已返回，但第 {repair_round + 1} 轮 JSON 修复失败：{repair_exc.detail}',
                ) from repair_exc
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
    follow_up = project.get('follow_up', {})
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
        'tenderer': summary.get('tenderer', ''),
        'register_year': follow_up.get('register_year', '') or project.get('sheet_name', ''),
        'bid_status': follow_up.get('bid_status', '待跟进'),
        'award_status': follow_up.get('award_status', '未知'),
        'our_quote_count': len(project.get('our_quotes', [])),
        'competitor_quote_count': len(project.get('competitor_quotes', [])),
    }


def read_project_file(path: Path) -> dict[str, Any] | None:
    try:
        raw = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None
    project_id = str(raw.get('project_id') or path.stem)
    sheet_name = compact_text(raw.get('sheet_name'))
    result = normalize_result_payload(raw.get('result') or {})
    follow_up = normalize_follow_up(raw.get('follow_up'), sheet_name=sheet_name)
    our_quotes = normalize_quote_rows(raw.get('our_quotes'), default_company='我司')
    competitor_quotes = normalize_quote_rows(raw.get('competitor_quotes'))
    timeline = normalize_timeline_rows(raw.get('timeline'))
    if not timeline and result.get('document_summary', {}).get('project_name'):
        timeline = build_auto_timeline(result, compact_text(raw.get('source_file_name')))
    return {
        'project_id': project_id,
        'title': compact_text(raw.get('title')),
        'source_file_name': compact_text(raw.get('source_file_name')),
        'register_mode': compact_text(raw.get('register_mode') or 'packages') or 'packages',
        'sheet_name': sheet_name,
        'created_at': compact_text(raw.get('created_at')),
        'updated_at': compact_text(raw.get('updated_at')),
        'result': result,
        'follow_up': follow_up,
        'our_quotes': our_quotes,
        'competitor_quotes': competitor_quotes,
        'timeline': timeline,
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
    sheet_name = compact_text(payload.sheet_name)
    follow_up = normalize_follow_up(payload.follow_up, sheet_name=sheet_name)
    our_quotes = normalize_quote_rows(payload.our_quotes, default_company='我司')
    competitor_quotes = normalize_quote_rows(payload.competitor_quotes)
    timeline = normalize_timeline_rows(payload.timeline)
    if not timeline and result.get('document_summary', {}).get('project_name'):
        timeline = build_auto_timeline(result, compact_text(payload.source_file_name))
    project = {
        'project_id': project_id,
        'title': title,
        'source_file_name': compact_text(payload.source_file_name),
        'register_mode': payload.register_mode if payload.register_mode in {'packages', 'document'} else 'packages',
        'sheet_name': sheet_name,
        'created_at': created_at or now,
        'updated_at': now,
        'result': result,
        'follow_up': follow_up,
        'our_quotes': our_quotes,
        'competitor_quotes': competitor_quotes,
        'timeline': timeline,
    }
    project_path(project_id).write_text(
        json.dumps(project, ensure_ascii=False, indent=2),
        encoding='utf-8',
    )
    return project


def build_overview_text(
    title: str,
    result: dict[str, Any],
    *,
    follow_up: dict[str, Any] | None = None,
    our_quotes: list[dict[str, Any]] | None = None,
    competitor_quotes: list[dict[str, Any]] | None = None,
    timeline: list[dict[str, Any]] | None = None,
) -> str:
    summary = result['document_summary']
    analysis = result['analysis']
    rows = result['register_rows']
    review = result['match_review']
    follow_up = normalize_follow_up(follow_up or {})
    our_quotes = normalize_quote_rows(our_quotes or [], default_company='我司')
    competitor_quotes = normalize_quote_rows(competitor_quotes or [])
    timeline = normalize_timeline_rows(timeline or [])

    def block(title: str, items: list[str]) -> str:
        lines = [item for item in items if item]
        return f"{title}\n" + ('\n'.join(f'- {item}' for item in lines) if lines else '- 无')

    sections = [
        '# 标书解析总览',
        '',
        f"归档标题：{title or summary['project_name'] or '未命名项目'}",
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
        '## 跟踪结果',
        f"投标状态：{follow_up['bid_status'] or '待跟进'}",
        f"中标状态：{follow_up['award_status'] or '未知'}",
        f"中标日期：{follow_up['award_date'] or '-'}",
        f"中标单位：{follow_up['award_company'] or '-'}",
        f"我司中标价：{follow_up['our_award_amount'] or '-'}",
        f"竞对中标价：{follow_up['competitor_award_amount'] or '-'}",
        f"信息来源：{follow_up['information_source'] or '-'}",
        f"跟进备注：{follow_up['tracking_note'] or '-'}",
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

    sections.extend(['', '## 我司报价一览'])
    if our_quotes:
        for index, row in enumerate(our_quotes, start=1):
            sections.extend(
                [
                    f"{index}. 标段：{row.get('package_name') or '-'}",
                    f"   轮次：第 {row.get('round_no') or 1} 轮",
                    f"   报价日期：{row.get('quote_date') or '-'}",
                    f"   币种：{row.get('currency') or 'CNY'}",
                    f"   单价：{row.get('unit_price') or '-'}",
                    f"   总价：{row.get('total_price') or '-'}",
                    f"   是否已投：{'是' if row.get('is_submitted') else '否'}",
                    f"   是否中标：{'是' if row.get('is_awarded') else '否'}",
                    f"   备注：{row.get('remark') or '-'}",
                ]
            )
    else:
        sections.append('- 无')

    sections.extend(['', '## 竞对报价一览'])
    if competitor_quotes:
        for index, row in enumerate(competitor_quotes, start=1):
            sections.extend(
                [
                    f"{index}. 公司：{row.get('quote_company') or '-'}",
                    f"   标段：{row.get('package_name') or '-'}",
                    f"   报价日期：{row.get('quote_date') or '-'}",
                    f"   币种：{row.get('currency') or 'CNY'}",
                    f"   单价：{row.get('unit_price') or '-'}",
                    f"   总价：{row.get('total_price') or '-'}",
                    f"   排名：{row.get('ranking') or '-'}",
                    f"   是否中标：{'是' if row.get('is_awarded') else '否'}",
                    f"   来源：{row.get('source') or '-'}",
                    f"   备注：{row.get('remark') or '-'}",
                ]
            )
    else:
        sections.append('- 无')

    sections.extend(['', '## 跟进时间线'])
    if timeline:
        for item in timeline:
            sections.append(f"- [{item.get('date') or '-'}] {item.get('type') or 'note'}：{item.get('note') or '-'}")
    else:
        sections.append('- 无')

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


def autosize_sheet(sheet: Any, widths: dict[int, int]) -> None:
    for index, width in widths.items():
        sheet.column_dimensions[chr(64 + index)].width = width


def apply_plain_table_style(sheet: Any, row_count: int, col_count: int) -> None:
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_font = Font(bold=True)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row_index in range(1, row_count + 1):
        for col_index in range(1, col_count + 1):
            cell = sheet.cell(row_index, col_index)
            cell.border = border
            cell.alignment = center if row_index <= 2 or col_index <= 2 else Alignment(vertical='top', wrap_text=True)
            if row_index in {1, 2, 4}:
                cell.font = header_font


def build_quote_workbook(
    *,
    title: str,
    follow_up: dict[str, Any],
    rows: list[dict[str, Any]],
    is_competitor: bool,
) -> BytesIO:
    follow_up = normalize_follow_up(follow_up or {})
    normalized_rows = normalize_quote_rows(rows or [], default_company='' if is_competitor else '我司')
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '报价一览'
    sheet['A1'] = '竞对报价对比' if is_competitor else '我司报价一览表'
    sheet['A2'] = '项目名称'
    sheet['B2'] = title or '未命名项目'
    sheet['C2'] = '投标状态'
    sheet['D2'] = follow_up.get('bid_status', '待跟进')
    sheet['E2'] = '中标状态'
    sheet['F2'] = follow_up.get('award_status', '未知')
    sheet['G2'] = '我司中标价' if not is_competitor else '竞对中标价'
    sheet['H2'] = follow_up.get('our_award_amount' if not is_competitor else 'competitor_award_amount', '')

    headers = (
        ['序号', '标段/合同包', '轮次', '报价日期', '报价主体', '币种', '税率/口径', '单价', '总价', '已投递', '是否中标', '备注']
        if not is_competitor
        else ['序号', '竞对公司', '标段/合同包', '报价日期', '币种', '单价', '总价', '排名', '是否中标', '来源', '备注']
    )
    start_row = 4
    for index, header in enumerate(headers, start=1):
        sheet.cell(start_row, index, header)

    for offset, row in enumerate(normalized_rows, start=1):
        row_index = start_row + offset
        if not is_competitor:
            values = [
                offset,
                row.get('package_name', ''),
                row.get('round_no', 1),
                row.get('quote_date', ''),
                row.get('quote_company', '我司'),
                row.get('currency', 'CNY'),
                row.get('tax_mode', ''),
                row.get('unit_price', ''),
                row.get('total_price', ''),
                '是' if row.get('is_submitted') else '否',
                '是' if row.get('is_awarded') else '否',
                row.get('remark', ''),
            ]
        else:
            values = [
                offset,
                row.get('quote_company', ''),
                row.get('package_name', ''),
                row.get('quote_date', ''),
                row.get('currency', 'CNY'),
                row.get('unit_price', ''),
                row.get('total_price', ''),
                row.get('ranking', ''),
                '是' if row.get('is_awarded') else '否',
                row.get('source', ''),
                row.get('remark', ''),
            ]
        for col_index, value in enumerate(values, start=1):
            sheet.cell(row_index, col_index, value)

    final_row = max(start_row + len(normalized_rows), start_row)
    apply_plain_table_style(sheet, final_row, len(headers))
    autosize_sheet(
        sheet,
        {
            1: 8,
            2: 24,
            3: 18,
            4: 16,
            5: 18,
            6: 12,
            7: 14,
            8: 14,
            9: 14,
            10: 14,
            11: 26,
            12: 24,
        },
    )
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def collect_project_items(
    *,
    q: str = '',
    year: str = '',
    award_status: str = '',
    bid_status: str = '',
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    keyword = q.strip().lower()
    year = compact_text(year)
    award_status = compact_text(award_status)
    bid_status = compact_text(bid_status)
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
                meta.get('tenderer', ''),
                meta.get('source_file_name', ''),
            ]
        ).lower()
        if keyword and keyword not in text:
            continue
        if year and meta.get('register_year', '') != year:
            continue
        if award_status and meta.get('award_status', '') != award_status:
            continue
        if bid_status and meta.get('bid_status', '') != bid_status:
            continue
        items.append(meta)
    items.sort(key=lambda item: item.get('updated_at', ''), reverse=True)
    stats = {
        'total_projects': len(items),
        'awarded_projects': sum(1 for item in items if item.get('award_status') == '已中标'),
        'submitted_projects': sum(1 for item in items if item.get('bid_status') == '已投标'),
        'pending_projects': sum(1 for item in items if item.get('bid_status') == '待跟进'),
        'our_quote_rows': sum(int(item.get('our_quote_count') or 0) for item in items),
        'competitor_quote_rows': sum(int(item.get('competitor_quote_count') or 0) for item in items),
    }
    return items, stats


def build_ledger_workbook(
    *,
    q: str = '',
    year: str = '',
    award_status: str = '',
    bid_status: str = '',
) -> BytesIO:
    items, stats = collect_project_items(q=q, year=year, award_status=award_status, bid_status=bid_status)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '历史台账'
    sheet['A1'] = '历史台账汇总'
    sheet['A2'] = '筛选关键词'
    sheet['B2'] = compact_text(q) or '全部'
    sheet['C2'] = '登记年份'
    sheet['D2'] = compact_text(year) or '全部'
    sheet['E2'] = '投标状态'
    sheet['F2'] = compact_text(bid_status) or '全部'
    sheet['G2'] = '中标状态'
    sheet['H2'] = compact_text(award_status) or '全部'

    stat_labels = [
        ('项目总数', stats['total_projects']),
        ('已中标', stats['awarded_projects']),
        ('已投标', stats['submitted_projects']),
        ('待跟进', stats['pending_projects']),
        ('我司报价条数', stats['our_quote_rows']),
        ('竞对报价条数', stats['competitor_quote_rows']),
    ]
    for index, (label, value) in enumerate(stat_labels, start=1):
        column = (index - 1) * 2 + 1
        sheet.cell(3, column, label)
        sheet.cell(3, column + 1, value)

    headers = ['序号', '项目标题', '项目名称', '招标编号', '招标人', '登记年份', '投标状态', '中标状态', '我司报价条数', '竞对报价条数', '最后更新时间']
    for col_index, header in enumerate(headers, start=1):
        sheet.cell(5, col_index, header)

    for row_index, item in enumerate(items, start=6):
        values = [
            row_index - 5,
            item.get('title', ''),
            item.get('project_name', ''),
            item.get('bid_no', ''),
            item.get('tenderer', ''),
            item.get('register_year', ''),
            item.get('bid_status', ''),
            item.get('award_status', ''),
            item.get('our_quote_count', 0),
            item.get('competitor_quote_count', 0),
            item.get('updated_at', ''),
        ]
        for col_index, value in enumerate(values, start=1):
            sheet.cell(row_index, col_index, value)

    final_row = max(5, 5 + len(items))
    apply_plain_table_style(sheet, final_row, len(headers))
    autosize_sheet(sheet, {1: 8, 2: 26, 3: 30, 4: 22, 5: 24, 6: 12, 7: 12, 8: 12, 9: 12, 10: 14, 11: 20})
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def write_overview_block(sheet: Any, row_index: int, title: str, lines: list[str]) -> int:
    sheet.cell(row_index, 1, title)
    sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=4)
    row_index += 1
    if not lines:
        lines = ['无']
    for line in lines:
        sheet.cell(row_index, 1, line)
        sheet.merge_cells(start_row=row_index, start_column=1, end_row=row_index, end_column=4)
        row_index += 1
    return row_index + 1


def build_overview_workbook(
    title: str,
    result: dict[str, Any],
    *,
    follow_up: dict[str, Any] | None = None,
    our_quotes: list[dict[str, Any]] | None = None,
    competitor_quotes: list[dict[str, Any]] | None = None,
    timeline: list[dict[str, Any]] | None = None,
) -> BytesIO:
    summary = result['document_summary']
    analysis = result['analysis']
    rows = result['register_rows']
    review = result['match_review']
    follow_up = normalize_follow_up(follow_up or {})
    our_quotes = normalize_quote_rows(our_quotes or [], default_company='我司')
    competitor_quotes = normalize_quote_rows(competitor_quotes or [])
    timeline = normalize_timeline_rows(timeline or [])

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '解析总览'
    sheet['A1'] = '标书解析总览'
    sheet.merge_cells('A1:D1')
    sheet['A2'] = '归档标题'
    sheet['B2'] = title or summary['project_name'] or '未命名项目'
    sheet['C2'] = '生成时间'
    sheet['D2'] = datetime.now().strftime('%Y-%m-%d %H:%M')

    core_pairs = [
        ('项目名称', summary['project_name'] or '未识别'),
        ('招标编号', summary['bid_no'] or '未识别'),
        ('招标人', summary['tenderer'] or '未识别'),
        ('投标截止时间', summary['bid_deadline'] or '未识别'),
        ('开标时间', summary['open_time'] or '未识别'),
        ('投标文件递交方式', summary['submission_method'] or '未识别'),
        ('保证金金额', summary['deposit_amount'] or '未识别'),
        ('服务期限', summary['service_period'] or '未识别'),
        ('投标状态', follow_up['bid_status'] or '待跟进'),
        ('中标状态', follow_up['award_status'] or '未知'),
        ('中标日期', follow_up['award_date'] or '-'),
        ('中标单位', follow_up['award_company'] or '-'),
        ('我司中标价', follow_up['our_award_amount'] or '-'),
        ('竞对中标价', follow_up['competitor_award_amount'] or '-'),
        ('信息来源', follow_up['information_source'] or '-'),
        ('跟进备注', follow_up['tracking_note'] or '-'),
    ]
    sheet['A4'] = '核心信息'
    sheet.merge_cells('A4:D4')
    row_index = 5
    for label, value in core_pairs:
        sheet.cell(row_index, 1, label)
        sheet.cell(row_index, 2, value)
        sheet.merge_cells(start_row=row_index, start_column=2, end_row=row_index, end_column=4)
        row_index += 1

    row_index += 1
    row_index = write_overview_block(sheet, row_index, 'AI 解析结论', [analysis['summary'] or '无'])
    row_index = write_overview_block(sheet, row_index, '资格文件清单', analysis['qualification_files'])
    row_index = write_overview_block(sheet, row_index, '报价与商务关注点', analysis['business_points'])
    row_index = write_overview_block(sheet, row_index, '评标得分关注点', analysis['scoring_points'])
    row_index = write_overview_block(sheet, row_index, '风险提示', analysis['risks'])

    register_lines = []
    for index, row in enumerate(rows, start=1):
        register_lines.append(
            f"{index}. 项目名称：{row.get('project_name') or '未识别'} | 招标编号：{row.get('bid_no') or '未识别'} | 招标人：{row.get('tenderer') or '未识别'} | 投标截止日期：{row.get('bid_deadline') or '未识别'} | 保证金金额：{row.get('deposit_amount') or '未识别'} | 备注：{row.get('remark') or '-'}"
        )
    row_index = write_overview_block(sheet, row_index, '登记预览', register_lines)

    our_lines = []
    for index, row in enumerate(our_quotes, start=1):
        our_lines.append(
            f"{index}. 标段：{row.get('package_name') or '-'} | 轮次：第 {row.get('round_no') or 1} 轮 | 报价日期：{row.get('quote_date') or '-'} | 币种：{row.get('currency') or 'CNY'} | 单价：{row.get('unit_price') or '-'} | 总价：{row.get('total_price') or '-'} | 已投：{'是' if row.get('is_submitted') else '否'} | 中标：{'是' if row.get('is_awarded') else '否'} | 备注：{row.get('remark') or '-'}"
        )
    row_index = write_overview_block(sheet, row_index, '我司报价一览', our_lines)

    competitor_lines = []
    for index, row in enumerate(competitor_quotes, start=1):
        competitor_lines.append(
            f"{index}. 公司：{row.get('quote_company') or '-'} | 标段：{row.get('package_name') or '-'} | 报价日期：{row.get('quote_date') or '-'} | 币种：{row.get('currency') or 'CNY'} | 单价：{row.get('unit_price') or '-'} | 总价：{row.get('total_price') or '-'} | 排名：{row.get('ranking') or '-'} | 中标：{'是' if row.get('is_awarded') else '否'} | 来源：{row.get('source') or '-'} | 备注：{row.get('remark') or '-'}"
        )
    row_index = write_overview_block(sheet, row_index, '竞对报价一览', competitor_lines)

    timeline_lines = [f"[{item.get('date') or '-'}] {item.get('type') or 'note'}：{item.get('note') or '-'}" for item in timeline]
    row_index = write_overview_block(sheet, row_index, '跟进时间线', timeline_lines)

    review_lines = []
    for item in review:
        review_lines.append(
            f"领域：{item.get('area') or '未分类'} | 事项：{item.get('item') or '-'} | 状态：{item.get('status') or '需人工确认'} | 原因：{item.get('reason') or '-'} | 建议动作：{item.get('action') or '-'}"
        )
    write_overview_block(sheet, row_index, '匹配复核', review_lines)

    apply_plain_table_style(sheet, sheet.max_row, 4)
    sheet.row_dimensions[1].height = 28
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 34
    sheet.column_dimensions['C'].width = 18
    sheet.column_dimensions['D'].width = 34
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


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
def root() -> RedirectResponse:
    return RedirectResponse(url='/login', status_code=302)


@app.get('/app')
def index(
    authorization: str | None = Header(default=None),
    auth_cookie: str | None = Cookie(default=None, alias=AUTH_COOKIE_NAME),
):
    if not resolve_auth_user(authorization, auth_cookie):
        return RedirectResponse(url='/login', status_code=302)
    return FileResponse(STATIC_DIR / 'index.html')


@app.get('/login')
def login_page() -> FileResponse:
    return FileResponse(STATIC_DIR / 'login.html')


@app.get('/api/template-meta')
def template_meta(_: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
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


@app.post('/api/auth/login')
def login(payload: LoginPayload, response: Response) -> dict[str, Any]:
    user = find_user(payload.username)
    if not user:
        raise HTTPException(status_code=401, detail='账号或密码错误。')
    password_hash = hash_password(payload.password)
    if not hmac.compare_digest(user.get('password_hash', ''), password_hash):
        raise HTTPException(status_code=401, detail='账号或密码错误。')
    session = create_auth_session(user)
    response.set_cookie(
        key=AUTH_COOKIE_NAME,
        value=session['token'],
        httponly=True,
        samesite='lax',
        max_age=int(AUTH_TTL.total_seconds()),
        path='/',
    )
    return session


@app.post('/api/auth/logout')
def logout(
    response: Response,
    authorization: str | None = Header(default=None),
    auth_cookie: str | None = Cookie(default=None, alias=AUTH_COOKIE_NAME),
) -> dict[str, Any]:
    token = extract_auth_token(authorization, auth_cookie)
    if token:
        AUTH_SESSIONS.pop(token, None)
    response.delete_cookie(AUTH_COOKIE_NAME, path='/')
    return {'ok': True}


@app.get('/api/auth/me')
def me(user: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    return {'user': user}


@app.get('/api/users')
def list_users(_: dict[str, Any] = Depends(require_admin)) -> dict[str, Any]:
    users = [sanitize_user(user) for user in load_users().get('users', []) if isinstance(user, dict)]
    return {'items': users}


@app.post('/api/users')
def create_user(payload: CreateUserPayload, _: dict[str, Any] = Depends(require_admin)) -> dict[str, Any]:
    username = compact_text(payload.username)
    if not username:
        raise HTTPException(status_code=400, detail='用户名不能为空。')
    if find_user(username):
        raise HTTPException(status_code=400, detail='该用户名已存在。')
    role = compact_text(payload.role).lower() or 'user'
    if role not in {'admin', 'user'}:
        raise HTTPException(status_code=400, detail='角色只能是 admin 或 user。')
    users_payload = load_users()
    users = users_payload.get('users', [])
    users.append(
        {
            'username': username,
            'password_hash': hash_password(payload.password),
            'display_name': compact_text(payload.display_name) or username,
            'role': role,
            'created_at': datetime.now().isoformat(timespec='seconds'),
        }
    )
    save_users(users_payload)
    return {'user': sanitize_user(users[-1])}


@app.get('/api/config')
def get_config(_: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    return mask_config(load_config())


@app.post('/api/config')
def update_config(payload: ConfigPayload, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    return mask_config(save_config(payload.model_dump()))


@app.post('/api/config/test')
def test_config(payload: ConfigPayload, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
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
def list_projects(
    q: str = '',
    year: str = '',
    award_status: str = '',
    bid_status: str = '',
    _: dict[str, Any] = Depends(require_auth),
) -> dict[str, Any]:
    items, stats = collect_project_items(q=q, year=year, award_status=award_status, bid_status=bid_status)
    return {'items': items, 'stats': stats}


@app.post('/api/projects')
def create_project(payload: ProjectPayload, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    project = save_project(uuid.uuid4().hex, payload)
    return {**project_meta(project), 'project': project}


@app.get('/api/projects/{project_id}')
def get_project(project_id: str, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    project = load_project(project_id)
    session_id = create_session_for_result(project['result'])
    return {**project, 'session_id': session_id}


@app.put('/api/projects/{project_id}')
def update_project(project_id: str, payload: ProjectPayload, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    existing = load_project(project_id)
    project = save_project(project_id, payload, created_at=existing.get('created_at'))
    return {**project_meta(project), 'project': project}


@app.delete('/api/projects/{project_id}')
def delete_project(project_id: str, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    path = project_path(project_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail='项目不存在。')
    path.unlink()
    return {'ok': True}


@app.post('/api/parse')
async def parse_bid(file: UploadFile = File(...), _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
    cleanup_sessions()
    suffix = Path(file.filename or '').suffix.lower()
    if suffix not in {'.pdf', '.doc', '.docx', '.xls', '.xlsx'}:
        raise HTTPException(status_code=400, detail='当前只支持 PDF、DOC、DOCX、XLS、XLSX。')
    temp_path = TMP_DIR / f'{uuid.uuid4().hex}{suffix}'
    temp_path.write_bytes(await file.read())
    try:
        document_text = extract_document_text(temp_path, suffix)
    finally:
        temp_path.unlink(missing_ok=True)
    if len(document_text.strip()) < 80:
        raise HTTPException(
            status_code=400,
            detail='当前版本仅支持可提取文本的 PDF、Word、Excel 文件，扫描件、空白文件或只有图片的内容暂不支持。',
        )
    result = parse_ai_document(document_text, require_config())
    session_id = create_session_for_result(result)
    return {'session_id': session_id, 'source_file_name': file.filename or '', **result}


@app.post('/api/chat')
def chat(payload: ChatPayload, _: dict[str, Any] = Depends(require_auth)) -> dict[str, Any]:
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
def export_extraction(payload: ExportExtractionPayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
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
def export_register(payload: ExportRegisterPayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
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
def export_overview(payload: ExportOverviewPayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
    result = normalize_result_payload(payload.result)
    output = build_overview_workbook(
        payload.title,
        result,
        follow_up=payload.follow_up,
        our_quotes=payload.our_quotes,
        competitor_quotes=payload.competitor_quotes,
        timeline=payload.timeline,
    )
    filename = f"bid_overview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/api/export/our-quotes')
def export_our_quotes(payload: ExportQuotePayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
    output = build_quote_workbook(
        title=compact_text(payload.title),
        follow_up=payload.follow_up,
        rows=payload.rows,
        is_competitor=False,
    )
    filename = f"our_quotes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/api/export/competitor-quotes')
def export_competitor_quotes(payload: ExportQuotePayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
    output = build_quote_workbook(
        title=compact_text(payload.title),
        follow_up=payload.follow_up,
        rows=payload.rows,
        is_competitor=True,
    )
    filename = f"competitor_quotes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.post('/api/export/ledger')
def export_ledger(payload: ExportLedgerPayload, _: dict[str, Any] = Depends(require_auth)) -> StreamingResponse:
    output = build_ledger_workbook(
        q=payload.q,
        year=payload.year,
        award_status=payload.award_status,
        bid_status=payload.bid_status,
    )
    year = compact_text(payload.year) or 'all'
    filename = f"history_ledger_{year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{filename}"'},
    )


@app.exception_handler(HTTPException)
async def http_error_handler(_: Any, exc: HTTPException) -> JSONResponse:
    return JSONResponse(status_code=exc.status_code, content={'detail': exc.detail})


def _self_check() -> None:
    users_payload = load_users()
    assert any(user.get('username') == DEFAULT_ADMIN_USERNAME for user in users_payload.get('users', []))
    admin = find_user(DEFAULT_ADMIN_USERNAME)
    assert admin is not None
    assert hmac.compare_digest(admin.get('password_hash', ''), hash_password(DEFAULT_ADMIN_PASSWORD))
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
