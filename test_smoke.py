import tempfile
from io import BytesIO
from pathlib import Path

import httpx
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app import ProjectPayload, app, call_chat_completion, extract_document_text, load_users, merge_register_rows, parse_json_text, project_path, resolve_source_excerpt, save_project, save_users


def cleanup(project_id: str) -> None:
    project_path(project_id).unlink(missing_ok=True)


def cleanup_user(username: str) -> None:
    payload = load_users()
    payload['users'] = [user for user in payload.get('users', []) if user.get('username') != username]
    save_users(payload)


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
assert resolve_source_excerpt('股份有限公司', '...', '福建省东南电化股份有限公司') == '福建省东南电化股份有限公司'

with tempfile.TemporaryDirectory() as temp_dir:
    xlsx_path = Path(temp_dir) / 'sample.xlsx'
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '招标信息'
    sheet['A1'] = '项目名称'
    sheet['B1'] = '测试招标项目'
    sheet['A2'] = '投标截止'
    sheet['B2'] = '2026-03-04 09:00'
    workbook.save(xlsx_path)
    workbook.close()
    extracted = extract_document_text(xlsx_path, '.xlsx')
    assert '工作表：招标信息' in extracted
    assert '项目名称 | 测试招标项目' in extracted


class DummyReadTimeoutClient:
    calls = 0

    @staticmethod
    def post(*args, **kwargs):
        DummyReadTimeoutClient.calls += 1
        raise httpx.ReadTimeout('simulated read timeout')


original_httpx_post = httpx.post
httpx.post = DummyReadTimeoutClient.post
try:
    try:
        call_chat_completion(
            {'base_url': 'https://example.com/v1', 'api_key': 'sk-test', 'model': 'demo', 'temperature': 0.1},
            system_prompt='test',
            user_prompt='test',
        )
        raise AssertionError('expected timeout')
    except Exception as exc:
        assert '响应超时' in str(exc.detail)
        assert DummyReadTimeoutClient.calls == 1
finally:
    httpx.post = original_httpx_post

project_id = 'f' * 32
cleanup(project_id)
save_project(
    project_id,
    ProjectPayload(
        title='测试项目',
        result={
            'document_summary': {'project_name': '测试项目'},
            'register_rows': [{'project_name': '测试项目'}],
            'document_text_full': '测试项目原文',
        },
    ),
)
assert project_path(project_id).exists()
cleanup(project_id)

client = TestClient(app)
cleanup_user('demo_user')

root_resp = client.get('/', follow_redirects=False)
assert root_resp.status_code in (302, 307)
assert root_resp.headers['location'] == '/login'

app_redirect_resp = client.get('/app', follow_redirects=False)
assert app_redirect_resp.status_code in (302, 307)
assert app_redirect_resp.headers['location'] == '/login'

login_resp = client.post('/api/auth/login', json={'username': 'ruico', 'password': 'Ruico668@'})
assert login_resp.status_code == 200
token = login_resp.json()['token']
auth_headers = {'Authorization': f'Bearer {token}'}
assert 'bid_parser_token=' in (login_resp.headers.get('set-cookie') or '')

app_resp = client.get('/app', headers=auth_headers)
assert app_resp.status_code == 200

project_payload = {
    'title': '接口测试项目',
    'source_file_name': 'demo.pdf',
    'register_mode': 'packages',
    'sheet_name': '2026',
    'follow_up': {
        'bid_status': '已投标',
        'award_status': '已中标',
        'award_date': '2026-03-10',
        'award_company': '我司',
        'our_award_amount': '1280000',
        'competitor_award_amount': '1310000',
        'information_source': '业务补录',
        'tracking_note': '完成中标补录',
        'register_year': '2026',
    },
    'our_quotes': [
        {
            'package_name': '合同包1',
            'round_no': 1,
            'quote_date': '2026-03-01',
            'quote_company': '我司',
            'currency': 'CNY',
            'tax_mode': '含税',
            'unit_price': '650',
            'total_price': '1280000',
            'is_submitted': True,
            'is_awarded': True,
            'remark': '最终中标价',
        }
    ],
    'competitor_quotes': [
        {
            'quote_company': '竞对A',
            'package_name': '合同包1',
            'quote_date': '2026-03-01',
            'currency': 'CNY',
            'unit_price': '665',
            'total_price': '1310000',
            'ranking': '2',
            'is_awarded': False,
            'source': '中标公示',
            'remark': '次低价',
        }
    ],
    'timeline': [
        {'date': '2026-01-27', 'type': 'parse', 'note': '完成标书解析'},
        {'date': '2026-03-10', 'type': 'award', 'note': '完成中标结果补录'},
    ],
    'result': {
        'document_summary': {
            'project_name': '福建省东南电化股份有限公司2026年液碱海路运输业务',
            'bid_no': 'WHTY/B-QT-2026-EB008-01',
            'tenderer': '福建省东南电化股份有限公司',
            'bid_deadline': '2026/3/4 9:00',
            'open_time': '2026/3/4 9:00',
            'submission_method': '投标平台递交',
            'deposit_amount': '10万',
            'service_period': '2026年3月-2027年3月',
            'qualification_requirements': '资质齐全',
            'vessel_requirements': '适装适靠',
            'technical_business': '按招标文件执行',
            'quotation': '固定单价',
            'evaluation_method': '综合评分法',
        },
        'extraction_fields': {
            'open_time': {'value': '2026/3/4 9:00', 'source_excerpt': '开标时间：2026/3/4 9:00'},
            'submission_method': {'value': '投标平台递交', 'source_excerpt': '投标文件递交方式：投标平台递交'},
        },
        'register_rows': [
            {
                'receive_date': '2026/1/27',
                'project_name': '福建省东南电化股份有限公司2026年液碱海路运输业务',
                'is_awarded': '√',
                'bid_no': 'WHTY/B-QT-2026-EB008-01',
                'tenderer': '福建省东南电化股份有限公司',
                'bid_deadline': '2026/3/4 9:00',
                'deposit_amount': '10万',
                'contract_note': '2026年3月-2027年3月',
                'deposit_return_status': '已退还10万元（履约保证金）',
                'payment_method': '对公转账',
                'acquisition_method': '万华电子招标投标交易网',
                'submission_method': '投标平台递交',
                'remark': '',
            }
        ],
        'analysis': {
            'summary': '测试摘要',
            'qualification_files': ['营业执照'],
            'business_points': ['注意报价口径'],
            'scoring_points': ['关注服务方案'],
            'risks': ['需人工确认船期'],
        },
        'match_review': [
            {
                'area': '业务匹配',
                'item': '运力核对',
                'status': '需人工确认',
                'reason': '未接内部运力系统',
                'action': '人工确认可用运力',
            }
        ],
        'document_text_excerpt': '示例原文',
        'document_text_full': '投标文件递交方式：投标平台递交\n开标时间：2026/3/4 9:00\n福建省东南电化股份有限公司',
    },
}

create_resp = client.post('/api/projects', json=project_payload, headers=auth_headers)
assert create_resp.status_code == 200
created = create_resp.json()
created_id = created['project_id']

list_resp = client.get('/api/projects?q=接口测试', headers=auth_headers)
assert list_resp.status_code == 200
assert any(item['project_id'] == created_id for item in list_resp.json()['items'])
assert list_resp.json()['stats']['awarded_projects'] >= 1

get_resp = client.get(f'/api/projects/{created_id}', headers=auth_headers)
assert get_resp.status_code == 200
project = get_resp.json()
result = project['result']

overview_resp = client.post(
    '/api/export/overview',
    json={
        'title': project['title'],
        'result': result,
        'follow_up': project['follow_up'],
        'our_quotes': project['our_quotes'],
        'competitor_quotes': project['competitor_quotes'],
        'timeline': project['timeline'],
    },
    headers=auth_headers,
)
assert overview_resp.status_code == 200
wb_overview = load_workbook(BytesIO(overview_resp.content))
assert wb_overview['解析总览']['A1'].value == '标书解析总览'
assert wb_overview['解析总览']['B2'].value == '接口测试项目'
assert wb_overview['解析总览']['A4'].value == '核心信息'

extraction_resp = client.post('/api/export/extraction', json={'result': result}, headers=auth_headers)
assert extraction_resp.status_code == 200
wb = load_workbook(BytesIO(extraction_resp.content))
assert wb.sheetnames[0] == '摘取结果'
assert '原文依据' in wb.sheetnames
assert wb['摘取结果']['B2'].value == '2026/3/4 9:00'
assert wb['原文依据']['B2'].value == '开标时间：2026/3/4 9:00'

legacy_extraction_resp = client.post('/api/export/extraction', json={'extraction_fields': result['extraction_fields']}, headers=auth_headers)
assert legacy_extraction_resp.status_code == 200

register_resp = client.post('/api/export/register', json={'rows': result['register_rows'], 'sheet_name': '2028'}, headers=auth_headers)
assert register_resp.status_code == 200
wb2 = load_workbook(BytesIO(register_resp.content))
assert '2028' in wb2.sheetnames
sheet = wb2['2028']
assert sheet['C3'].value
assert sheet['B3'].number_format == 'yyyy/m/d'
assert sheet['G3'].number_format == 'yyyy/m/d h:mm'
assert sheet['H3'].value == 100000
assert sheet['A3'].border.left.style == 'thin'
assert sheet['N3'].alignment.horizontal == 'center'

our_quotes_resp = client.post(
    '/api/export/our-quotes',
    json={'title': project['title'], 'follow_up': project['follow_up'], 'rows': project['our_quotes']},
    headers=auth_headers,
)
assert our_quotes_resp.status_code == 200
wb3 = load_workbook(BytesIO(our_quotes_resp.content))
assert wb3['报价一览']['A1'].value == '我司报价一览表'
assert wb3['报价一览']['E5'].value == '我司'

competitor_quotes_resp = client.post(
    '/api/export/competitor-quotes',
    json={'title': project['title'], 'follow_up': project['follow_up'], 'rows': project['competitor_quotes']},
    headers=auth_headers,
)
assert competitor_quotes_resp.status_code == 200
wb4 = load_workbook(BytesIO(competitor_quotes_resp.content))
assert wb4['报价一览']['A1'].value == '竞对报价对比'
assert wb4['报价一览']['B5'].value == '竞对A'

ledger_resp = client.post(
    '/api/export/ledger',
    json={'q': '接口测试', 'year': '2026', 'bid_status': '已投标', 'award_status': '已中标'},
    headers=auth_headers,
)
assert ledger_resp.status_code == 200
wb5 = load_workbook(BytesIO(ledger_resp.content))
assert wb5['历史台账']['A1'].value == '历史台账汇总'
assert wb5['历史台账']['D2'].value == '2026'
assert wb5['历史台账']['F2'].value == '已投标'
assert wb5['历史台账']['H2'].value == '已中标'
assert wb5['历史台账']['B6'].value == '接口测试项目'

bad_docx_resp = client.post(
    '/api/parse',
    files={
        'file': (
            'bad.docx',
            b'not-a-real-docx',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )
    },
    headers=auth_headers,
)
assert bad_docx_resp.status_code == 400
assert 'DOCX 文件格式无效' in bad_docx_resp.json()['detail']

users_resp = client.get('/api/users', headers=auth_headers)
assert users_resp.status_code == 200
assert any(item['username'] == 'ruico' for item in users_resp.json()['items'])

create_user_resp = client.post(
    '/api/users',
    json={'username': 'demo_user', 'password': 'Demo123@', 'display_name': '演示用户', 'role': 'user'},
    headers=auth_headers,
)
assert create_user_resp.status_code == 200

delete_resp = client.delete(f'/api/projects/{created_id}', headers=auth_headers)
assert delete_resp.status_code == 200
cleanup(created_id)
cleanup_user('demo_user')
