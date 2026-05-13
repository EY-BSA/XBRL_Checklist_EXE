"""
taxonomy_xlsx_parser.py 
IxD 편집기 '구조내려받기' xlsx 파일을 파싱하여
checklist_engine이 사용하는 presentation_rows 형태로 변환
"""

import io, re
from typing import Dict, Optional
import pandas as pd


def _is_consol(text: str) -> Optional[bool]:
    for k in ['Consolidated', 'consolidated', '연결']:
        if k in text: return True
    for k in ['Separated', 'Separate', 'separated', '별도', 'Nonconsolidated']:
        if k in text: return False
    return None


def _extract_table_number(role_def: str) -> str:
    """Role Definition '[D210000] ...' → 'D210000' (가이드 7.비고)"""
    m = re.search(r'\[([A-Za-z]{1,3}X?\d{4,})\]', str(role_def))
    return m.group(1) if m else ''


def _extract_role_code(role_def: str, role_uri: str) -> str:
    code = _extract_table_number(role_def)
    if code: return code
    m = re.search(r'/([A-Z]{1,3}X?\d{4,})$', str(role_uri))
    return m.group(1) if m else ''


def _label_role_short(url: str) -> str:
    s = str(url).strip()
    return '' if not s or s.lower() == 'nan' else s.split('/')[-1]


def _safe(v) -> str:
    if v is None: return ''
    s = str(v).strip()
    return '' if s.lower() == 'nan' else s


def _classify_element(gubn_raw: str, name: str) -> str:
    """가이드 2.2 - Element 분류 (Name 끝 4글자 기준)"""
    if 'lineitem' in name.lower(): return 'item'
    if gubn_raw.strip().upper() == 'FOOTNOTES': return 'FOOTNOTES'
    suffix = name[-4:].lower() if len(name) >= 4 else ''
    return {
        'tory': 'Explanatory', 'ract': 'Abstract',
        'axis': 'Axis',        'lock': 'TextBlock',
        'able': 'Table',       'mber': 'Member',
    }.get(suffix, 'item')


def _classify_gubn(gubn_raw: str, name: str) -> str:
    """구분 컬럼 세분화"""
    g = gubn_raw.strip().upper()
    if g == 'TABLE'     or name.endswith('Table'):                   return 'TABLE'
    if g == 'FOOTNOTES' or name.endswith('TextBlock'):               return 'FOOTNOTES'
    if name.endswith('Axis'):                                         return 'Axis'
    if name.endswith('Member'):                                       return 'Member'
    if name.endswith('Domain'):                                       return 'Domain'
    if name.endswith('LineItems') or name.endswith('LineItem'):       return 'LINEITEM'
    if g == 'LINEITEM':                                               return 'LINEITEM'
    if g in ('DOMAIN','MEMBER','AXIS'):                               return g.capitalize()
    return 'LINEITEM'


class TaxonomyXlsxData:
    class _El:
        def __init__(self, lko='', len_='', lr=''):
            self.label_ko=lko; self.label_en=len_; self.label_role=lr; self.abstract=False
    def __init__(self):
        self.company_name=''; self.report_date=''; self.entity_id=''
        self.presentation_rows=[]; self.errors=[]
        self.elements:Dict[str,object]={}
        self.contexts:Dict[str,object]={}
        self.facts=[]; self._fact_elements:set=set()


def parse_taxonomy_xlsx(file_bytes: bytes) -> TaxonomyXlsxData:
    data = TaxonomyXlsxData()
    try:
        xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None,
                            header=None, dtype=str, na_filter=False)
    except Exception as e:
        data.errors.append(f'xlsx 읽기 실패: {e}'); return data

    rows = []
    for sheet_name, df in xls.items():
        if '기본정보' in sheet_name:
            _parse_basic_info(df, data); continue

        role_uri = ''; role_def = ''; header_idx = None
        for i, row in df.iterrows():
            v0 = _safe(row.iloc[0]); v1 = _safe(row.iloc[1]) if len(row)>1 else ''
            if 'Role URI' in v0:        role_uri = v1
            elif 'Role Definition' in v0: role_def = v1
            elif v0 == '구분' and i >= 1: header_idx = i; break

        if header_idx is None: continue
        if not role_uri: role_uri = f'sheet:{sheet_name}'

        code      = _extract_role_code(role_def, role_uri)
        table_num = _extract_table_number(role_def) or code
        parts     = role_def.split('|', 1)
        name_ko   = re.sub(r'^\[[^\]]+\]\s*', '', parts[0]).strip()
        name_en   = parts[1].strip() if len(parts)>1 else ''
        is_c      = _is_consol(role_def or role_uri)

        # 연결/별도: Role Definition 7번째 자리 (가이드 2.2)
        consol_str = '-'
        if role_def and len(role_def) >= 7:
            ch = role_def[6]
            if ch == '0':   consol_str = '연결'
            elif ch == '5': consol_str = '별도'

        current_table_label = ''  # 시트 내 현재 TABLE 행의 Label(KO) 추적
        for i in range(header_idx + 1, len(df)):
            row = df.iloc[i]
            def col(j): return _safe(row.iloc[j]) if len(row)>j else ''
            gubn_raw=col(0); prefix=col(1); name=col(2)
            lbl_ko=col(3); lbl_en=col(4); lbl_role_url=col(5)
            dtype=col(6); balance=col(7); period=col(8)
            decimal_val=col(9); fact_val=col(10)

            if not name: continue
            if gubn_raw == '구분': continue  # 시트 내 반복 헤더 행 건너뜀

            gubn    = _classify_gubn(gubn_raw, name)
            element = _classify_element(gubn_raw, name)
            lbl_role= _label_role_short(lbl_role_url)
            period_n= period.upper()     # 가이드: 'INSTANT' / 'DURATION'
            bal_n   = balance.lower()

            # TABLE 행을 만날 때마다 현재 테이블 이름 갱신 (다중 테이블 지원)
            if gubn == 'TABLE':
                current_table_label = lbl_ko

            # 가이드 기준: 비확장 = '-', 확장 = '확장'
            ext = '확장' if prefix.startswith('entity') else '-'
            client_negate = 'negate' if 'negated' in lbl_role.lower() else '-'
            alias         = '별칭'   if 'terse'   in lbl_role.lower() else '-'
            has_fact      = bool(fact_val) or (decimal_val == '0')

            if name not in data.elements:
                data.elements[name] = TaxonomyXlsxData._El(lbl_ko, lbl_en, lbl_role)

            rows.append({
                'role_uri': role_uri, 'role_code': code,
                'role_name_ko': name_ko, 'role_name_en': name_en,
                'is_consolidated': is_c,
                'Role Definition': role_def,
                'Sheet': sheet_name, '연결/별도': consol_str,
                'Table_Number': table_num, 'TABLE_NUMBER': table_num,
                'depth': 0, 'parent': '', 'parent_label_ko': '', 'parent_gubn': '',
                'Prefix': prefix, 'Name': name,
                'Label(KO)': lbl_ko, 'Label(EN)': lbl_en,
                'Label Role': lbl_role,
                'DataType': dtype, 'Balance': bal_n,
                'Period': period_n,   # 대문자 유지
                'Decimal': decimal_val, 'Fact': fact_val,
                '구분': gubn, 'Element': element,
                '확장여부': ext, 'Client_Negate': client_negate, '별칭여부': alias,
                'PreferredLabel': lbl_role_url,
                'has_fact': has_fact, 'abstract': False,
                'table_label_ko': current_table_label,
            })

    data.presentation_rows = rows
    return data


def _parse_basic_info(df, data: TaxonomyXlsxData):
    for i, row in df.iterrows():
        for j, cell in enumerate(row):
            s = _safe(cell)

            if '법인명' in s or '회사명' in s:
                # "법인명 : 인탑스 주식회사" 형식 (한 셀)
                if ':' in s:
                    v = s.split(':', 1)[1].strip()
                    if v:
                        data.company_name = v
                        continue
                # label | value 인접 셀 형식
                for k in range(j + 1, min(j + 4, len(row))):
                    v = _safe(row.iloc[k])
                    if v:
                        data.company_name = v
                        break

            if not data.report_date and ('문서작성일' in s or '회계기간종료일' in s):
                # "문서작성일 : 2024-12-31" 형식 (한 셀)
                if ':' in s:
                    v = s.split(':', 1)[1].strip()
                    if v:
                        data.report_date = v
                        continue
                # label | value 인접 셀 형식
                for k in range(j + 1, min(j + 4, len(row))):
                    v = _safe(row.iloc[k])
                    if v:
                        data.report_date = v
                        break