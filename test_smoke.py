import tempfile
import uuid
import re
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import httpx
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app import ProjectPayload, app, call_chat_completion, extract_document_text, handle_feishu_message, load_config, load_users, market_record_path, merge_register_rows, parse_json_text, project_path, resolve_source_excerpt, save_config, save_project, save_users, ws_message_to_payload


def cleanup(project_id: str) -> None:
    project_path(project_id).unlink(missing_ok=True)


def cleanup_user(username: str) -> None:
    payload = load_users()
    payload['users'] = [user for user in payload.get('users', []) if user.get('username') != username]
    save_users(payload)


def cleanup_market(kind: str, record_id: str) -> None:
    market_record_path(kind, record_id).unlink(missing_ok=True)


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

cargo_a = {
    'board_type': '\u5373\u65f6\u8d27\u76d8',
    'segment': '\u5185\u8d38\u5316',
    'cargo_name': 'smoke-benzene-market-filter',
    'tonnage': '4000\u5428',
    'load_port': '\u5f20\u5bb6\u6e2f',
    'discharge_port': '\u4e1c\u839e',
    'cargo_owner': '\u8fdc\u5927',
    'status': '\u8ddf\u8fdb\u4e2d',
    'raw_text': 'raw text mentions smoke-methane-market-filter only',
}
cargo_b = {**cargo_a, 'cargo_name': 'smoke-methane-market-filter', 'raw_text': 'same as structured'}
cargo_a_resp = client.post('/api/market-skill/cargo', json={'kind': 'cargo', 'record': cargo_a}, headers=auth_headers)
cargo_b_resp = client.post('/api/market-skill/cargo', json={'kind': 'cargo', 'record': cargo_b}, headers=auth_headers)
assert cargo_a_resp.status_code == 200
assert cargo_b_resp.status_code == 200
cargo_a_id = cargo_a_resp.json()['record']['id']
cargo_b_id = cargo_b_resp.json()['record']['id']
try:
    methane_resp = client.get(
        '/api/market-skill/cargo',
        params={'q': 'smoke-methane-market-filter', 'segment': '\u5185\u8d38\u5316', 'status': '\u8ddf\u8fdb\u4e2d'},
        headers=auth_headers,
    )
    benzene_resp = client.get(
        '/api/market-skill/cargo',
        params={'q': 'smoke-benzene-market-filter', 'segment': '\u5185\u8d38\u5316', 'status': '\u8ddf\u8fdb\u4e2d'},
        headers=auth_headers,
    )
    assert methane_resp.status_code == 200
    assert benzene_resp.status_code == 200
    assert [item['id'] for item in methane_resp.json()['items']] == [cargo_b_id]
    assert [item['id'] for item in benzene_resp.json()['items']] == [cargo_a_id]
finally:
    cleanup_market('cargo', cargo_a_id)
    cleanup_market('cargo', cargo_b_id)

cargo_deal = {
    'board_type': '\u5373\u65f6\u8d27\u76d8',
    'segment': '\u5185\u8d38\u5316',
    'cargo_name': 'smoke-trend-xylene',
    'cargo_standard_name': 'smoke-trend-xylene',
    'tonnage': '3000\u5428',
    'load_port': '\u5b81\u6ce2',
    'discharge_port': '\u4e1c\u839e',
    'route': '\u5b81\u6ce2 - \u4e1c\u839e',
    'cargo_date': '2026-06-01',
    'deal_date': '2026-06-15',
    'status': '\u5df2\u6210\u4ea4',
    'deal_price': '128',
    'price_unit': '\u5143/\u5428',
    'currency': 'CNY',
}
cargo_open = {**cargo_deal, 'status': '\u8ddf\u8fdb\u4e2d', 'deal_price': '999', 'cargo_name': 'smoke-trend-open'}
cargo_deal_resp = client.post('/api/market-skill/cargo', json={'kind': 'cargo', 'record': cargo_deal}, headers=auth_headers)
cargo_open_resp = client.post('/api/market-skill/cargo', json={'kind': 'cargo', 'record': cargo_open}, headers=auth_headers)
assert cargo_deal_resp.status_code == 200
assert cargo_open_resp.status_code == 200
cargo_deal_id = cargo_deal_resp.json()['record']['id']
cargo_open_id = cargo_open_resp.json()['record']['id']
try:
    cargo_report_resp = client.post(
        '/api/market-skill/report',
        json={
            'kind': 'cargo',
            'period': 'monthly',
            'start_date': '2026-06-01',
            'end_date': '2026-06-30',
            'filters': {'q': 'smoke-trend'},
        },
        headers=auth_headers,
    )
    assert cargo_report_resp.status_code == 200
    cargo_report = cargo_report_resp.json()
    assert cargo_report['source_stats']['deal_count'] == 1
    assert cargo_report['trend_points'][0]['avg_price'] == 128
    assert cargo_report['detail_rows'][0]['price'] == 128
    cargo_report_export_resp = client.post(
        '/api/export/market-skill-report',
        json={'kind': 'cargo', 'report': cargo_report, 'format': 'xlsx'},
        headers=auth_headers,
    )
    assert cargo_report_export_resp.status_code == 200
    wb6 = load_workbook(BytesIO(cargo_report_export_resp.content))
    assert wb6['分析报告']['A1'].value == '商机市场分析报告'
    cargo_docx_resp = client.post(
        '/api/export/market-skill-report',
        json={'kind': 'cargo', 'report': cargo_report},
        headers=auth_headers,
    )
    assert cargo_docx_resp.status_code == 200
    with ZipFile(BytesIO(cargo_docx_resp.content)) as docx:
        names = set(docx.namelist())
        assert {'[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml'} <= names
        assert '商机市场 AI 分析报告' in docx.read('word/document.xml').decode('utf-8')
finally:
    cleanup_market('cargo', cargo_deal_id)
    cleanup_market('cargo', cargo_open_id)

newbuilding_record = {
    'stage': '\u5df2\u5b8c\u9020\u5e76\u51fa\u5382\u6295\u8fd0',
    'ship_name': 'smoke-newbuilding-1',
    'update_date': '2026-06-02',
    'status_update_date': '2026-06-20',
    'shipyard': '\u6c5f\u5357\u9020\u8239',
    'owner': '\u6d4b\u8bd5\u8239\u4e1c',
    'dwt': '25000DWT',
    'build_status': '\u5df2\u4ea4\u4ed8',
    'delivery_time': '2026\u5e74',
    'actual_delivery_date': '2026-06-18',
    'status_note': '\u51fa\u5382\u6295\u8fd0',
}
newbuilding_resp = client.post('/api/market-skill/newbuilding', json={'kind': 'newbuilding', 'record': newbuilding_record}, headers=auth_headers)
assert newbuilding_resp.status_code == 200
newbuilding_id = newbuilding_resp.json()['record']['id']
try:
    newbuilding_report_resp = client.post(
        '/api/market-skill/report',
        json={
            'kind': 'newbuilding',
            'period': 'weekly',
            'start_date': '2026-06-01',
            'end_date': '2026-06-30',
            'filters': {'q': 'smoke-newbuilding-1'},
        },
        headers=auth_headers,
    )
    assert newbuilding_report_resp.status_code == 200
    newbuilding_report = newbuilding_report_resp.json()
    assert newbuilding_report['source_stats']['delivered'] == 1
    assert newbuilding_report['detail_rows'][0]['ship_name'] == 'smoke-newbuilding-1'
    newbuilding_export_resp = client.post(
        '/api/export/market-skill-report',
        json={'kind': 'newbuilding', 'report': newbuilding_report, 'format': 'xlsx'},
        headers=auth_headers,
    )
    assert newbuilding_export_resp.status_code == 200
    wb7 = load_workbook(BytesIO(newbuilding_export_resp.content))
    assert wb7['分析报告']['A1'].value == '新造船市场分析报告'
finally:
    cleanup_market('newbuilding', newbuilding_id)

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

original_config = load_config()
feishu_cargo_id = ''
try:
    config_resp = client.post(
        '/api/config',
        json={
            **original_config,
            'base_url': '',
            'api_key': '',
            'model': '',
            'feishu_enabled': True,
            'feishu_receive_mode': 'http',
            'feishu_app_id': 'cli_smoke',
            'feishu_app_secret': 'secret-smoke',
            'feishu_verification_token': 'verify-smoke',
            'feishu_encrypt_key': 'encrypt-smoke',
        },
        headers=auth_headers,
    )
    assert config_resp.status_code == 200
    masked_config = config_resp.json()
    assert masked_config['has_feishu_app_secret'] is True
    assert masked_config['feishu_app_secret'] == '********'

    verify_resp = client.post(
        '/api/feishu/events',
        json={'type': 'url_verification', 'token': 'verify-smoke', 'challenge': 'challenge-ok'},
    )
    assert verify_resp.status_code == 200
    assert verify_resp.json()['challenge'] == 'challenge-ok'

    help_reply = handle_feishu_message({'open_id': 'ou_anyone', 'chat_id': 'chat_anywhere', 'message_id': 'mid-help', 'text': '帮助', 'files': []}, load_config())
    assert '我现在能做这些事' in help_reply
    assert '市场情报' in help_reply
    ws_event = type(
        'WsEvent',
        (),
        {
            'header': type('Header', (), {'event_id': 'ws-event', 'event_type': 'im.message.receive_v1', 'token': ''})(),
            'event': type(
                'Event',
                (),
                {
                    'sender': type('Sender', (), {'sender_id': type('SenderId', (), {'open_id': 'ou_ws', 'user_id': '', 'union_id': ''})()})(),
                    'message': type('Message', (), {'message_id': 'mid-ws', 'chat_id': 'chat_ws', 'message_type': 'text', 'content': '{"text":"帮助"}'})(),
                },
            )(),
        },
    )()
    ws_payload = ws_message_to_payload(ws_event)
    assert ws_payload['header']['event_type'] == 'im.message.receive_v1'
    assert ws_payload['event']['message']['content'] == '{"text":"帮助"}'

    message = {
        'open_id': 'ou_smoke',
        'chat_id': 'chat_smoke',
        'message_id': 'mid-smoke',
        'text': '4000吨甲苯，张家港～东莞，要求月内装出，远大的计划',
        'files': [],
    }
    reply = handle_feishu_message(message, load_config())
    assert '确认保存' in reply
    assert '下一步' in reply
    confirm_reply = handle_feishu_message({**message, 'text': '确认保存'}, load_config())
    feishu_cargo_id = re.search(r'[a-f0-9]{32}', confirm_reply).group(0)
    assert '已保存' in confirm_reply
    assert '下一步' in confirm_reply
    assert market_record_path('cargo', feishu_cargo_id).exists()
finally:
    if feishu_cargo_id:
        cleanup_market('cargo', feishu_cargo_id)
    save_config(original_config)

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
