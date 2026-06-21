from io import BytesIO

from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app import ProjectPayload, app, merge_register_rows, parse_json_text, project_path, resolve_source_excerpt, save_project


def cleanup(project_id: str) -> None:
    project_path(project_id).unlink(missing_ok=True)


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

project_payload = {
    'title': '接口测试项目',
    'source_file_name': 'demo.pdf',
    'register_mode': 'packages',
    'sheet_name': '2026',
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

create_resp = client.post('/api/projects', json=project_payload)
assert create_resp.status_code == 200
created = create_resp.json()
created_id = created['project_id']

list_resp = client.get('/api/projects?q=接口测试')
assert list_resp.status_code == 200
assert any(item['project_id'] == created_id for item in list_resp.json()['items'])

get_resp = client.get(f'/api/projects/{created_id}')
assert get_resp.status_code == 200
result = get_resp.json()['result']

overview_resp = client.post('/api/export/overview', json={'result': result})
assert overview_resp.status_code == 200
assert overview_resp.content.startswith(b'# ')

extraction_resp = client.post('/api/export/extraction', json={'result': result})
assert extraction_resp.status_code == 200
wb = load_workbook(BytesIO(extraction_resp.content))
assert wb.sheetnames[0] == '摘取结果'
assert '原文依据' in wb.sheetnames
assert wb['摘取结果']['B2'].value == '2026/3/4 9:00'
assert wb['原文依据']['B2'].value == '开标时间：2026/3/4 9:00'

legacy_extraction_resp = client.post('/api/export/extraction', json={'extraction_fields': result['extraction_fields']})
assert legacy_extraction_resp.status_code == 200

register_resp = client.post('/api/export/register', json={'rows': result['register_rows'], 'sheet_name': '2028'})
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

delete_resp = client.delete(f'/api/projects/{created_id}')
assert delete_resp.status_code == 200
cleanup(created_id)
