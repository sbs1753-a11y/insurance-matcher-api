import openpyxl
import zipfile
import shutil
import os
import re
from collections import defaultdict


# ─────────────────────────────────────────────
# ZIP 직접 조작 헬퍼 (서식/색상/폰트 완전 보존)
# ─────────────────────────────────────────────

def _col_to_letter(n):
    """컬럼 번호 → 엑셀 열 문자 (1→A, 26→Z, 27→AA)"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell_addr(row, col):
    return _col_to_letter(col) + str(row)


def _escape_xml(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def write_cells_zip(file_path, cell_updates):
    """
    xlsx 파일을 ZIP으로 직접 조작해 셀 값만 수정.
    styles.xml, theme1.xml 등 서식 파일은 일체 건드리지 않아
    색상/폰트/정렬이 완전히 보존된다.

    file_path   : 수정할 xlsx 파일 경로 (in-place)
    cell_updates: [{'row': int, 'col': int, 'value': str}, ...]
    """
    if not cell_updates:
        return

    # ── 1. 워크시트 XML 읽기 ────────────────────────────────────
    with zipfile.ZipFile(file_path, "r") as zf:
        ws_names = sorted(
            [n for n in zf.namelist()
             if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
        )
        if not ws_names:
            raise ValueError("xlsx 내 워크시트 XML 없음")
        ws_path = ws_names[0]
        ws_xml = zf.read(ws_path).decode("utf-8")

    # ── 2. 수정 맵 생성 ──────────────────────────────────────────
    updates = {}
    for item in cell_updates:
        if item.get("value") is not None:
            updates[_cell_addr(item["row"], item["col"])] = str(item["value"])

    # ── 3. 기존 <c> 요소 업데이트 ───────────────────────────────
    processed = set()

    def replace_cell(m):
        addr   = m.group(1)
        attrs  = m.group(2)   # r="..." 뒤의 나머지 속성들
        inner  = m.group(3)   # </c> 앞 내용

        if addr not in updates:
            return m.group(0)

        val = updates[addr]
        processed.add(addr)

        escaped = _escape_xml(val)
        # 기존 <v>/<f> 제거
        inner = re.sub(r"<v>[^<]*</v>", "", inner)
        inner = re.sub(r"<f[^>]*>.*?</f>", "", inner, flags=re.DOTALL)
        # t 속성 → "str" (인라인 문자열)
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        return f'<c r="{addr}"{attrs} t="str">{inner}<v>{escaped}</v></c>'

    # 닫힌 태그 형태: <c r="A1" ...>...</c>  (자기닫힘 제외: [^/>]* 로 / 차단)
    ws_xml = re.sub(
        r'<c r="([A-Z]+\d+)"([^/>]*)>(.*?)</c>',
        replace_cell,
        ws_xml,
        flags=re.DOTALL,
    )

    # 자기 닫힘 형태만: <c r="A1" ... />  (여는 태그 > 는 매칭 안 함)
    def replace_selfclose(m):
        addr  = m.group(1)
        attrs = m.group(2)
        if addr not in updates:
            return m.group(0)
        val = updates[addr]
        processed.add(addr)
        escaped = _escape_xml(val)
        attrs = re.sub(r'\s*t="[^"]*"', "", attrs)
        return f'<c r="{addr}"{attrs} t="str"><v>{escaped}</v></c>'

    ws_xml = re.sub(
        r'<c r="([A-Z]+\d+)"([^/>]*)/>',
        replace_selfclose,
        ws_xml,
    )

    # ── 4. XML에 없던 셀 삽입 ────────────────────────────────────
    missing = {addr: val for addr, val in updates.items() if addr not in processed}
    if missing:
        ws_xml = _insert_new_cells(ws_xml, missing)

    # ── 5. ZIP에서 워크시트 XML만 교체 ──────────────────────────
    tmp = file_path + ".tmp"
    with zipfile.ZipFile(file_path, "r") as zf_in, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for info in zf_in.infolist():
            if info.filename == ws_path:
                zf_out.writestr(info, ws_xml.encode("utf-8"))
            else:
                zf_out.writestr(info, zf_in.read(info.filename))

    os.replace(tmp, file_path)


def _insert_new_cells(ws_xml, cells_to_insert):
    """XML에 존재하지 않는 셀을 적절한 <row>에 삽입"""
    by_row = defaultdict(dict)
    for addr, val in cells_to_insert.items():
        m = re.match(r"([A-Z]+)(\d+)", addr)
        if m:
            by_row[int(m.group(2))][addr] = val

    for row_num in sorted(by_row):
        cells = by_row[row_num]
        row_pat = re.compile(
            r'(<row\b[^>]*\br="' + str(row_num) + r'"[^>]*>)(.*?)(</row>)',
            re.DOTALL,
        )
        row_m = row_pat.search(ws_xml)

        if row_m:
            # 행이 존재 → 셀 삽입 (열 순서 유지)
            inner = row_m.group(2)
            for addr in sorted(cells):
                escaped = _escape_xml(cells[addr])
                inner += f'<c r="{addr}" t="str"><v>{escaped}</v></c>'
            ws_xml = ws_xml[:row_m.start(2)] + inner + ws_xml[row_m.end(2):]
        else:
            # 행 자체가 없음 → 새 행 삽입
            cells_xml = "".join(
                f'<c r="{addr}" t="str"><v>{_escape_xml(val)}</v></c>'
                for addr, val in sorted(cells.items())
            )
            new_row = f'<row r="{row_num}">{cells_xml}</row>'

            # 행 번호 기준으로 적절한 위치에 삽입
            last_m = None
            for m in re.finditer(r"<row\b[^>]*\br=\"(\d+)\"", ws_xml):
                if int(m.group(1)) < row_num:
                    last_m = m
            if last_m:
                end = ws_xml.index("</row>", last_m.start()) + len("</row>")
                ws_xml = ws_xml[:end] + new_row + ws_xml[end:]
            else:
                ws_xml = ws_xml.replace("</sheetData>", new_row + "</sheetData>", 1)

    return ws_xml


# ─────────────────────────────────────────────
# 읽기 함수 (openpyxl - 변경 없음)
# ─────────────────────────────────────────────

def find_row_by_label(excel_path, label, sheet_name=None, search_col=2):
    """특정 텍스트가 있는 행 번호를 자동 탐지"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active
        for row_idx in range(1, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=search_col).value
            if val and label in str(val).strip():
                return row_idx
        return None
    finally:
        wb.close()


def find_insurer_row(excel_path, sheet_name=None, search_col=2):
    """보험사명이 들어갈 행 자동 탐지"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active
        for row_idx in range(1, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=search_col).value
            if val and "회사" in str(val).strip():
                return row_idx + 1
        return None
    finally:
        wb.close()


def find_structure(excel_path, sheet_name=None, search_col=2):
    """엑셀 보장분석표 구조 자동 탐지 (A형 + B형 지원)"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        structure = {
            "insurer_row": None,
            "product_row": None,
            "premium_row": None,
            "reserve_row": None,
            "start_row": None,
            "template_type": "A",
            "coverage_col": 2,
            "first_amount_col": 4,
        }

        # --- 1단계: A형 탐지 (B열=column 2 기준) ---
        for row_idx in range(1, min(ws.max_row + 1, 20)):
            val = ws.cell(row=row_idx, column=search_col).value
            if val is None:
                continue
            val_str = str(val).strip()

            if "회사" in val_str and structure["insurer_row"] is None:
                structure["insurer_row"] = row_idx + 1
            if "상품" in val_str and structure["product_row"] is None:
                structure["product_row"] = row_idx
            if "보험료" in val_str and "총" not in val_str and structure["premium_row"] is None:
                structure["premium_row"] = row_idx
            if "적립금" in val_str and structure["reserve_row"] is None:
                structure["reserve_row"] = row_idx
            if val_str in ["실비질병/상해 종합입원", "실비질병/상해종합입원"] and structure["start_row"] is None:
                structure["start_row"] = row_idx

        if structure["start_row"] is None and structure["premium_row"]:
            structure["start_row"] = structure["premium_row"] + 3

        if structure["insurer_row"] or structure["premium_row"]:
            structure["template_type"] = "A"
            structure["coverage_col"] = 2
            structure["first_amount_col"] = 4
            return structure

        # --- 2단계: B형 탐지 (C열=column 3 기준) ---
        for row_idx in range(1, min(ws.max_row + 1, 20)):
            val_c = ws.cell(row=row_idx, column=3).value
            if val_c is None:
                continue
            val_str = str(val_c).strip()

            if ("보험사" in val_str or "회사" in val_str) and structure["insurer_row"] is None:
                structure["insurer_row"] = row_idx
            if "상품" in val_str and structure["product_row"] is None:
                structure["product_row"] = row_idx
            if "보험료" in val_str and "총" not in val_str and structure["premium_row"] is None:
                structure["premium_row"] = row_idx

        for row_idx in range(1, min(ws.max_row + 1, 20)):
            val_a = ws.cell(row=row_idx, column=1).value
            val_b = ws.cell(row=row_idx, column=2).value
            a_str = str(val_a).strip() if val_a else ""
            b_str = str(val_b).strip() if val_b else ""

            if ("분류" in a_str and "보장" in b_str) or ("분류" in b_str):
                structure["start_row"] = row_idx + 1
                break

        if structure["start_row"] is None and structure["premium_row"]:
            structure["start_row"] = structure["premium_row"] + 2

        if structure["insurer_row"] or structure["premium_row"] or structure["start_row"]:
            structure["template_type"] = "B"
            structure["coverage_col"] = 2
            structure["first_amount_col"] = 5
        else:
            structure["template_type"] = "unknown"
            structure["coverage_col"] = 2
            structure["first_amount_col"] = 4

        return structure
    finally:
        wb.close()


def read_excel_coverages(excel_path, sheet_name=None, coverage_col=2,
                         amount_col=4, start_row=8):
    """Excel 보장분析표에서 특약명 목록 읽기"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        coverages = []
        skip_values = [
            "주계약", "특약", "합계", "총보험료", "보장항목",
            "담보명", "특약명", "보장내용", "분류", ""
        ]
        category_only = {
            "실손", "수술", "암", "뇌", "심장", "입원", "간호",
            "후유", "사망", "상해", "치매", "운전", "배상"
        }

        for row_idx in range(start_row, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=coverage_col).value
            if cell_value is None:
                continue
            name = str(cell_value).strip()
            if name in skip_values or len(name) < 2 or name in category_only:
                continue
            coverages.append({
                "row": row_idx,
                "특약명": name,
                "amount_col": amount_col
            })

        return coverages
    finally:
        wb.close()


# ─────────────────────────────────────────────
# 쓰기 함수 (ZIP 방식 - 서식 완전 보존)
# ─────────────────────────────────────────────

def write_matched_amounts(excel_path, output_path, matched_data, sheet_name=None):
    """매칭 결과를 Excel에 기록 (ZIP 방식 - 색상/폰트 보존)"""
    if excel_path != output_path:
        shutil.copy2(excel_path, output_path)

    cells = []
    for item in matched_data:
        val = item.get("display_amount") or item.get("가입금액")
        if val is not None:
            cells.append({
                "row": item["row"],
                "col": item["amount_col"],
                "value": str(val),
            })

    write_cells_zip(output_path, cells)


def write_insurer_info(excel_path, output_path, insurer_name, insurer_row,
                       product_name, product_row, col, sheet_name=None):
    """보험사명과 상품명을 Excel 특정 셀에 기록 (ZIP 방식 - 색상/폰트 보존)"""
    if excel_path != output_path:
        shutil.copy2(excel_path, output_path)

    write_cells_zip(output_path, [
        {"row": insurer_row, "col": col, "value": str(insurer_name)},
        {"row": product_row, "col": col, "value": str(product_name)},
    ])


def write_premium(excel_path, output_path, premium, premium_row, col, sheet_name=None):
    """보험료를 Excel 특정 셀에 기록 (ZIP 방식 - 색상/폰트 보존)"""
    if excel_path != output_path:
        shutil.copy2(excel_path, output_path)

    write_cells_zip(output_path, [
        {"row": premium_row, "col": col, "value": str(premium)},
    ])
