"""
XBRL CoE Checklist 1-1 체커 (독립 실행형)
IxD 편집기 '구조내려받기' xlsx → 1-1 Gross 계정 사용 검토 → XBRL_CoE_Checklist_Result.xlsx 저장
"""

import io
import re
import sys
import os
import threading
import tkinter as tk
import zipfile
from tkinter import filedialog, messagebox, ttk

import openpyxl

from taxonomy_xlsx_parser import parse_taxonomy_xlsx
from checklist_engine import _c1_1, CheckResult
from standard_taxonomy import StandardTaxonomy, enrich_axis_domain_check


def resource_path(relative: str) -> str:
    """PyInstaller 번들 실행 시 내장 리소스 경로 반환."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


AXIS_DOMAIN_PATH = resource_path('Axis_Domain_Check.xlsx')
TEMPLATE_PATH    = resource_path(os.path.join('template', 'XBRL_CoE_Checklist_Result.xlsx'))


# ─── 엑셀 출력 ────────────────────────────────────────────────────────────────

def _role_def(iss) -> str:
    rc = iss.role_code; ko = iss.role_name_ko; en = iss.role_name_en
    if rc and ko:
        return f"[{rc}] {ko} | {en}" if en else f"[{rc}] {ko}"
    return ko or rc


def _table_name(iss) -> str:
    return iss.table_label_ko or iss.role_name_ko or iss.role_code


def _find_sheet_xml(xlsx_path: str, sheet_name: str) -> str:
    """템플릿 xlsx에서 시트 이름에 해당하는 worksheet XML 경로 반환."""
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        wb_xml   = z.read('xl/workbook.xml').decode('utf-8')
        rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')

    # <sheet> 태그에서 name 속성으로 r:id 추출 (속성 순서 무관)
    escaped = re.escape(sheet_name)
    m = re.search(rf'<sheet\b[^>]*\bname="{escaped}"[^>]*/>', wb_xml)
    if not m:
        raise KeyError(f'workbook.xml에서 시트 찾기 실패: {sheet_name}')
    rid_m = re.search(r'r:id="([^"]+)"', m.group(0))
    if not rid_m:
        raise KeyError(f'<sheet> 태그에서 r:id 추출 실패: {sheet_name}')
    rid = rid_m.group(1)

    # rels 파일에서 Target 추출 (속성 순서 무관)
    m2 = re.search(rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*/>', rels_xml)
    if not m2:
        raise KeyError(f'workbook.xml.rels에서 rId 찾기 실패: {rid}')
    target_m = re.search(r'Target="worksheets/([^"]+)"', m2.group(0))
    if not target_m:
        raise KeyError(f'<Relationship> 태그에서 Target 추출 실패: {rid}')
    return f'xl/worksheets/{target_m.group(1)}'


def write_result(result: CheckResult, data, output_path: str):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f'템플릿을 찾을 수 없습니다: {TEMPLATE_PATH}')

    sheet_name = result.sheet  # 'Checklist_1-1'

    # Step 1: openpyxl로 데이터 작성 후 메모리 버퍼에 저장
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f'템플릿에 시트 없음: {sheet_name}')

    ws = wb[sheet_name]

    # 14행 이후 기존 데이터 클리어
    if ws.max_row >= 14:
        for row in ws.iter_rows(min_row=14, max_row=ws.max_row):
            for cell in row:
                cell.value = None

    # 이슈 기록 (14행~)
    for ri, iss in enumerate(result.issues, 14):
        ws.cell(ri, 2, _role_def(iss))          # B: Role Definition
        ws.cell(ri, 3, _table_name(iss))         # C: TABLE NAME
        ws.cell(ri, 4, iss.prefix)               # D: Prefix
        ws.cell(ri, 5, iss.element_name)         # E: Name
        ws.cell(ri, 6, iss.label_ko)             # F: Label(KO)
        ws.cell(ri, 7, iss.label_en)             # G: Label(EN)
        ws.cell(ri, 8, iss.label_role)           # H: Label Role
        ws.cell(ri, 9, iss.data_type)            # I: DataType
        ws.cell(ri, 10, iss.period)              # J: Period

    # B11에 건수를 정적 값으로 기록 (openpyxl은 수식 미계산 → LEAD 참조 셀이 0으로 표시되는 문제 방지)
    ws['B11'] = result.issue_count

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Step 2: 템플릿 기반으로 출력 파일 구성
    # openpyxl 출력에서 가져올 파일: 수정된 시트 XML + sharedStrings + styles
    # styles.xml도 openpyxl 버전 사용 → 시트 XML의 스타일 ID와 일치 보장
    # 나머지(LEAD 시트, drawing, media 등)는 템플릿 원본 사용
    target_xml = _find_sheet_xml(TEMPLATE_PATH, sheet_name)
    preserve_from_modified = {target_xml, 'xl/sharedStrings.xml', 'xl/styles.xml'}

    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as tmpl_zip, \
         zipfile.ZipFile(buf, 'r') as mod_zip:
        mod_names = set(mod_zip.namelist())

        # 템플릿 worksheet XML에서 <drawing> 참조 태그 추출
        # openpyxl은 저장 시 이 태그를 제거하므로, 나중에 다시 주입해야 도형이 유지됨
        tmpl_sheet_raw = tmpl_zip.read(target_xml)
        drawing_tags = re.findall(rb'<drawing[^>]*/>', tmpl_sheet_raw)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for info in tmpl_zip.infolist():
                # calcChain 제거: B11을 정적 값으로 교체해 수식 구조가 바뀌었으므로
                # Excel이 파일 열 때 전체 재계산하도록 강제
                if info.filename == 'xl/calcChain.xml':
                    continue
                raw = tmpl_zip.read(info.filename)
                if info.filename in preserve_from_modified and info.filename in mod_names:
                    raw = mod_zip.read(info.filename)
                    # openpyxl이 제거한 <drawing> 참조를 워크시트 XML에 복원
                    if info.filename == target_xml and drawing_tags:
                        inject = b''.join(drawing_tags)
                        raw = raw.replace(b'</worksheet>', inject + b'</worksheet>')
                # [Content_Types].xml 에서 calcChain Override 항목 제거
                if info.filename == '[Content_Types].xml':
                    raw = re.sub(rb'<Override[^>]*calcChain[^>]*/>', b'', raw)
                # workbook.xml 에 fullCalcOnLoad="1" 추가
                # → Excel이 파일 열 때 캐시 값을 무시하고 모든 수식을 재계산
                if info.filename == 'xl/workbook.xml':
                    raw = re.sub(
                        rb'<calcPr([^/]*)/>', rb'<calcPr\1 fullCalcOnLoad="1"/>', raw
                    )
                out_zip.writestr(info, raw)


# ─── 분석 로직 ────────────────────────────────────────────────────────────────

def run_check(input_path: str):
    with open(input_path, 'rb') as f:
        xlsx_bytes = f.read()

    data = parse_taxonomy_xlsx(xlsx_bytes)

    # Axis_Domain_Check 로드 (내장 파일)
    std = StandardTaxonomy()
    if os.path.exists(AXIS_DOMAIN_PATH):
        enrich_axis_domain_check(std, AXIS_DOMAIN_PATH)

    result = _c1_1(data.presentation_rows, data)
    return data, result


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('XBRL CoE Checklist 1-1')
        self.resizable(False, False)
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f'+{(sw-w)//2}+{(sh-h)//2}')

    def _build_ui(self):
        pad = {'padx': 16, 'pady': 8}

        tk.Label(self, text='XBRL CoE Checklist', font=('Helvetica', 15, 'bold')).pack(**pad)
        tk.Label(self, text='1-1  Gross 계정 사용 검토', font=('Helvetica', 11)).pack(pady=(0, 4))
        tk.Label(
            self,
            text='IxD 편집기 [구조내려받기] xlsx 파일을 선택하세요.',
            fg='#555'
        ).pack()

        self._file_var = tk.StringVar(value='선택된 파일 없음')
        tk.Label(self, textvariable=self._file_var, fg='#333', wraplength=420).pack(**pad)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=4)
        tk.Button(btn_frame, text='📂  파일 선택', width=18,
                  command=self._select_file).grid(row=0, column=0, padx=6)
        self._run_btn = tk.Button(btn_frame, text='▶  체크 실행', width=18,
                                  state='disabled', command=self._run)
        self._run_btn.grid(row=0, column=1, padx=6)

        self._progress = ttk.Progressbar(self, mode='indeterminate', length=420)
        self._progress.pack(pady=(8, 2))

        self._status = tk.StringVar(value='')
        tk.Label(self, textvariable=self._status, fg='#333').pack(pady=(0, 12))

        self._input_path = None

    def _select_file(self):
        path = filedialog.askopenfilename(
            title='XBRL 구조내려받기 파일 선택',
            filetypes=[('Excel 파일', '*.xlsx'), ('모든 파일', '*.*')]
        )
        if not path:
            return
        self._input_path = path
        self._file_var.set(os.path.basename(path))
        self._run_btn.config(state='normal')
        self._status.set('')

    def _run(self):
        if not self._input_path:
            return
        self._run_btn.config(state='disabled')
        self._progress.start(12)
        self._status.set('분석 중...')
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            data, result = run_check(self._input_path)
            self.after(0, self._on_done, data, result)
        except Exception as exc:
            self.after(0, self._on_error, str(exc))

    def _on_done(self, data, result):
        self._progress.stop()

        company = data.company_name or 'XBRL'
        rdate   = (data.report_date or '').replace('/', '').replace('-', '')
        parts   = [p for p in [company, rdate] if p]
        default_name = f"XBRL_CoE_Checklist_Result_{'_'.join(parts)}.xlsx"

        issue_cnt = result.issue_count
        msg = (f'이슈 {issue_cnt}건 발견 되었습니다.\n결과 파일을 저장할 위치를 선택하세요.'
               if issue_cnt else
               '이슈 없음 (Pass). 결과 파일을 저장할 위치를 선택하세요.')
        self._status.set(msg)

        output_path = filedialog.asksaveasfilename(
            title='결과 파일 저장',
            defaultextension='.xlsx',
            initialfile=default_name,
            filetypes=[('Excel 파일', '*.xlsx'), ('모든 파일', '*.*')]
        )
        if not output_path:
            self._run_btn.config(state='normal')
            self._status.set('저장 취소됨.')
            return

        try:
            write_result(result, data, output_path)
            self._status.set(f'저장 완료: {os.path.basename(output_path)}')
            messagebox.showinfo(
                '완료',
                f'1-1 체크 완료\n'
                f'이슈 건수: {issue_cnt}건\n\n'
                f'저장 위치:\n{output_path}'
            )
        except Exception as exc:
            self._on_error(str(exc))
        finally:
            self._run_btn.config(state='normal')

    def _on_error(self, msg: str):
        self._progress.stop()
        self._status.set('오류 발생')
        self._run_btn.config(state='normal')
        messagebox.showerror('오류', msg)


if __name__ == '__main__':
    App().mainloop()
