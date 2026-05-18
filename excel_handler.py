import openpyxl


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
    """엑셀 보장분석표 구조 자동 탐지 (A형 + B형 지원)
    
    A형: B열에 "회사", "상품", "보험료" + 특약명 B열 + 금액 D열~
    B형: C열에 "보험사", "상품명", "보험료" + 특약명 B열 + 금액 E열~
    """
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        structure = {
            "insurer_row": None,
            "product_row": None,
            "premium_row": None,
            "reserve_row": None,
            "start_row": None,
            # 양식 타입 및 열 정보
            "template_type": "A",    # "A" 또는 "B"
            "coverage_col": 2,       # 특약명 열 (둘 다 B열=2)
            "first_amount_col": 4,   # 첫 번째 금액 열 (A형=D/4, B형=E/5)
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

        # A형 탐지 성공: insurer_row 또는 premium_row가 있으면 A형
        if structure["insurer_row"] or structure["premium_row"]:
            structure["template_type"] = "A"
            structure["coverage_col"] = 2
            structure["first_amount_col"] = 4
            return structure

        # --- 2단계: B형 탐지 (C열=column 3 기준) ---
        # B형: C열에 "보험사", "상품명", "월 보험료" 등이 있고, E열~에 데이터
        for row_idx in range(1, min(ws.max_row + 1, 20)):
            val_c = ws.cell(row=row_idx, column=3).value  # C열
            if val_c is None:
                continue
            val_str = str(val_c).strip()

            if ("보험사" in val_str or "회사" in val_str) and structure["insurer_row"] is None:
                structure["insurer_row"] = row_idx
            if "상품" in val_str and structure["product_row"] is None:
                structure["product_row"] = row_idx
            if "보험료" in val_str and "총" not in val_str and structure["premium_row"] is None:
                structure["premium_row"] = row_idx

        # B형 start_row 탐지: A열에 "분류" 또는 B열에 "보장내용"이 있는 헤더 행의 다음 행
        for row_idx in range(1, min(ws.max_row + 1, 20)):
            val_a = ws.cell(row=row_idx, column=1).value
            val_b = ws.cell(row=row_idx, column=2).value
            a_str = str(val_a).strip() if val_a else ""
            b_str = str(val_b).strip() if val_b else ""

            if ("분류" in a_str and "보장" in b_str) or ("분류" in b_str):
                structure["start_row"] = row_idx + 1
                break

        # B형 폴백: premium_row 기준
        if structure["start_row"] is None and structure["premium_row"]:
            structure["start_row"] = structure["premium_row"] + 2

        # B형이면 타입 마킹
        if structure["insurer_row"] or structure["premium_row"] or structure["start_row"]:
            structure["template_type"] = "B"
            structure["coverage_col"] = 2       # B열 = 특약명 (사용자가 B열로 이동함)
            structure["first_amount_col"] = 5   # E열 = 첫 번째 보험사 금액
        else:
            # 완전 탐지 실패 — 기본값 사용
            structure["template_type"] = "unknown"
            structure["coverage_col"] = 2
            structure["first_amount_col"] = 4

        return structure
    finally:
        wb.close()


def read_excel_coverages(excel_path, sheet_name=None, coverage_col=2, amount_col=4, start_row=8):
    """Excel 보장분석표에서 특약명 목록 읽기"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        coverages = []
        skip_values = [
            "주계약", "특약", "합계", "총보험료", "보장항목",
            "담보명", "특약명", "보장내용", "분류", ""
        ]
        # 분류명만 있는 셀 스킵 (B형에서 A열의 분류가 B열에 혼재할 경우 대비)
        category_only = {
            "실손", "수술", "암", "뇌", "심장", "입원", "간호",
            "후유", "사망", "상해", "치매", "운전", "배상"
        }

        for row_idx in range(start_row, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=coverage_col).value

            if cell_value is None:
                continue

            name = str(cell_value).strip()

            if name in skip_values:
                continue

            if len(name) < 2:
                continue

            # 분류명만 단독으로 있는 행 스킵
            if name in category_only:
                continue

            coverages.append({
                "row": row_idx,
                "특약명": name,
                "amount_col": amount_col
            })

        return coverages
    finally:
        wb.close()


def write_matched_amounts(excel_path, output_path, matched_data, sheet_name=None):
    """매칭 결과를 Excel에 기록"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        for item in matched_data:
            ws.cell(
                row=item["row"],
                column=item["amount_col"],
                value=item["가입금액"]
            )

        wb.save(output_path)
    finally:
        wb.close()


def write_insurer_info(excel_path, output_path, insurer_name, insurer_row, product_name, product_row, col, sheet_name=None):
    """보험사명과 상품명을 Excel 특정 셀에 기록"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        ws.cell(row=insurer_row, column=col, value=insurer_name)
        ws.cell(row=product_row, column=col, value=product_name)

        wb.save(output_path)
    finally:
        wb.close()


def write_premium(excel_path, output_path, premium, premium_row, col, sheet_name=None):
    """보험료를 Excel 특정 셀에 기록"""
    wb = openpyxl.load_workbook(excel_path)
    try:
        ws = wb[sheet_name] if sheet_name else wb.active

        ws.cell(row=premium_row, column=col, value=premium)

        wb.save(output_path)
    finally:
        wb.close()
