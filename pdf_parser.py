import pdfplumber
import re


def detect_insurer(pdf_path):
    """PDF에서 보험사 자동 감지"""
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages[:3]:
            page_text = page.extract_text()
            if page_text:
                text += page_text

    keywords_ordered = [
        ("삼성생명", "samsung_life"),
        ("삼성화재", "samsung"),
        ("메리츠화재", "meritz"),
        ("메리츠", "meritz"),
        ("미래에셋생명", "mirae"),
        ("미래에셋", "mirae"),
        ("KB손해", "kb"),
        ("KB손보", "kb"),
        ("KB 플러스", "kb"),
        ("KB플러스", "kb"),
        ("DB손해", "db"),
        ("DB손보", "db"),
        ("ABL", "abl"),
        ("에이비엘", "abl"),
        ("흥국", "heungkuk"),
        ("한화", "hanwha"),
        ("현대해상", "hyundai"),
        ("롯데손해", "lotte"),
        ("NH농협", "nh"),
        ("동양생명", "dongyang"),
        ("교보생명", "kyobo"),
        ("신한라이프", "shinhan"),
    ]

    for keyword, insurer in keywords_ordered:
        if keyword in text:
            return insurer

    if "kbinsure" in text.lower():
        return "kb"

    if "삼성" in text:
        if any(kw in text for kw in ["생명보험", "건강보험", "종신보험", "The간편한", "다모은"]):
            return "samsung_life"
        return "samsung"

    return None


def detect_product_name(pdf_path):
    """PDF에서 상품명 추출"""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:3]:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                line = line.strip()

                kb_match = re.search(r'(KB\s*플러스\s*[^\(\n]+보험)', line)
                if kb_match:
                    name = kb_match.group(1).strip()
                    name = re.sub(r'\(무배당\).*', '', name).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

                kb_match2 = re.search(r'(KB\s*[^\(\n]*보험[^\(\n]*)\(무배당\)', line)
                if kb_match2:
                    name = kb_match2.group(1).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

                mirae_match = re.search(r'(M-케어\s*건강[^\(]*)', line)
                if mirae_match:
                    name = mirae_match.group(1).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

                if re.match(r'^\(무\)|^\(유\)', line):
                    name = line
                    name = re.sub(r'^\(무\)\s*|^\(유\)\s*', '', name)
                    match = re.match(r'^([^(]+)', name)
                    if match:
                        name = match.group(1).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

                samsung_match = re.match(r'^삼성\s+(.+보험)', line)
                if samsung_match:
                    name = samsung_match.group(1).strip()
                    name = re.sub(r'\(\d{4}\).*', '', name).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

                # 메리츠 상품명 감지
                meritz_match = re.search(r'(메리츠\s*[^\(\n]*보험[^\(\n]*)', line)
                if meritz_match:
                    name = meritz_match.group(1).strip()
                    name = re.sub(r'\(무배당\).*', '', name).strip()
                    if len(name) > 30:
                        name = name[:30]
                    return name

    return "상품명 미확인"


def extract_premium(pdf_path):
    """PDF에서 보험료 추출 (원 단위)"""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:7]:
            text = page.extract_text()
            if not text:
                continue
            patterns = [
                r'실납입보험료\s*([\d,]+)\s*원',
                r'1회차보험료\(할인후\)\s*([\d,]+)\s*원',
                r'할인후초회보험료\s*([\d,]+)\s*원',
                r'보장보험료\s*합계\s*([\d,]+)\s*원',
                r'합\s*계\s*보\s*험\s*료\s*([\d,]+)\s*원',
                r'합계보험료\s*([\d,]+)\s*원',
                r'합\s*계\s*([\d,]+)',
                r'보험료\s*[:\s]?\s*([\d,]+)\s*원',
            ]
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    amount_str = match.group(1).replace(',', '')
                    val = int(amount_str)
                    if val >= 1000:
                        return val
    return None


def find_name_col_from_data(table, header_idx):
    """데이터 행을 보고 특약명이 있는 열을 찾는다"""
    for row in table[header_idx + 1:]:
        if not row:
            continue
        for j, cell in enumerate(row):
            if not cell:
                continue
            cell_clean = cell.strip()
            if len(cell_clean) > 10 and any(kw in cell_clean for kw in [
                "갱신형", "통합간편", "특별약관", "보험료납입",
                "일반상해", "질병", "수술", "진단", "입원", "배상",
                "골절", "화상", "사망", "후유장해", "치매", "암",
                "뇌혈관", "심장", "혈전",
                "보장특약", "무배당", "파워수술", "양성신생물",
            ]):
                return j
    return None


# ══════════════════════════════════════════════
# KB손해보험 전용 파서
# ══════════════════════════════════════════════

def extract_coverage_kb(pdf_path):
    """KB손해보험 PDF 파싱 — 가입담보 페이지 집중 파싱"""
    results = []
    full_text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

        for page_num, page in enumerate(pdf.pages[:10]):
            page_text = page.extract_text()
            if not page_text:
                continue

            is_coverage_page = any(kw in page_text for kw in [
                "가입담보", "가입내용", "보장명", "가입금액"
            ])
            if not is_coverage_page:
                continue

            _parse_kb_coverage_page(page_text, results)

    seen = set()
    unique_results = []
    for r in results:
        if r["특약명"] not in seen:
            seen.add(r["특약명"])
            unique_results.append(r)
    results = unique_results

    _enrich_kb_injury_grade_info(results, full_text)

    return results


def _parse_kb_coverage_page(page_text, results):
    """KB 가입담보 페이지 텍스트에서 보장명+가입금액 추출"""
    lines = page_text.split('\n')

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1

        if not line:
            continue

        match = re.match(r'^(\d{1,3})\s+(.+)', line)
        if not match:
            continue

        num_str = match.group(1)
        name_part = match.group(2).strip()

        num_val = int(num_str)
        if num_val > 500:
            continue

        if _is_kb_skip_line(name_part):
            continue

        amount_in_line = _extract_kb_amount_from_text(name_part)
        if amount_in_line:
            clean_name = _clean_kb_amount_from_name(name_part)
            if clean_name and len(clean_name) >= 2:
                results.append({"특약명": clean_name, "가입금액": amount_in_line})
            continue

        amount = None
        for offset in range(0, 5):
            if i + offset >= len(lines):
                break
            next_line = lines[i + offset].strip()

            if not next_line:
                continue

            if re.match(r'^\d{1,3}\s+\S', next_line):
                break

            amount = _extract_kb_amount_from_text(next_line)
            if amount:
                break

        if amount and name_part and len(name_part) >= 2:
            if not _is_kb_skip_line(name_part):
                results.append({"특약명": name_part, "가입금액": amount})


def _extract_kb_amount_from_text(text):
    """텍스트에서 KB 가입금액 패턴 추출 (원 단위 반환)"""
    match = re.search(r'(\d+천\d*백만원|\d+억\d*천?만?원|\d+천만원|\d+백만원|\d+만원)', text)
    if match:
        return _parse_korean_amount(match.group(1))
    return None


def _parse_korean_amount(text):
    """한글 금액을 원 단위 숫자로 변환"""
    text = text.replace(',', '').replace(' ', '')

    match = re.match(r'(\d+)천(\d*)백만원', text)
    if match:
        thousands = int(match.group(1))
        hundreds = int(match.group(2)) if match.group(2) else 0
        return (thousands * 1000 + hundreds * 100) * 10000

    match = re.match(r'(\d+)억(\d*천?만?)원?', text)
    if match:
        billions = int(match.group(1))
        rest = match.group(2)
        rest_val = 0
        if rest:
            rest_match = re.match(r'(\d+)천만', rest)
            if rest_match:
                rest_val = int(rest_match.group(1)) * 10000000
            else:
                rest_match = re.match(r'(\d+)만', rest)
                if rest_match:
                    rest_val = int(rest_match.group(1)) * 10000
        return billions * 100000000 + rest_val

    match = re.match(r'(\d+)천만원', text)
    if match:
        return int(match.group(1)) * 10000000

    match = re.match(r'(\d+)백만원', text)
    if match:
        return int(match.group(1)) * 1000000

    match = re.match(r'(\d+)만원', text)
    if match:
        return int(match.group(1)) * 10000

    return None


def _clean_kb_amount_from_name(name):
    """보장명에서 금액 부분을 제거하고 순수 보장명만 반환"""
    name = re.sub(r'\d+천\d*백만원', '', name)
    name = re.sub(r'\d+억\d*천?만?원?', '', name)
    name = re.sub(r'\d+천만원', '', name)
    name = re.sub(r'\d+백만원', '', name)
    name = re.sub(r'\d+만원', '', name)
    name = re.sub(r'[\d,]+\s*\d+년/\d+년', '', name)
    name = re.sub(r'[\d,]+$', '', name)
    name = name.strip()
    return name


def _is_kb_skip_line(text):
    """KB PDF에서 스킵해야 할 줄인지 판단"""
    text_clean = text.replace(" ", "")
    skip_words = [
        "주의사항", "고객콜센터", "홈페이지", "영업담당자",
        "발급일시", "계약자용", "장기", "제작",
        "납입형태", "계약사항", "피보험자님",
        "보장합계", "기타유의사항", "구분", "내용",
        "공통사항", "갱신시", "예상만기", "할인후",
        "보험료(원)", "납입|보험기간", "보장명",
        "가입금액", "RQ2", "2026", "인천GA",
        "어센틱금융", "www.", "1544",
    ]
    return any(sk in text_clean for sk in skip_words)


def _enrich_kb_injury_grade_info(results, full_text):
    """KB PDF 본문에서 자동차사고부상 등급별 지급액을 추출"""
    pattern_14 = re.search(
        r'(?:12[~\-]14급|12급[~\-]14급)\s*[:：]\s*(\d[\d,]*)\s*만원',
        full_text
    )
    if pattern_14:
        grade_14_amount = int(pattern_14.group(1).replace(',', '')) * 10000
        for r in results:
            if "사고부상" in r["특약명"] and ("4~14" in r["특약명"] or "4~14급" in r["특약명"]):
                r["14급지급액"] = grade_14_amount
                break


# ══════════════════════════════════════════════
# 미래에셋생명 전용 파서
# ══════════════════════════════════════════════

def extract_coverage_mirae(pdf_path):
    """미래에셋생명 PDF 파싱 — 보험계약 개요 페이지에서 추출"""
    results = []

    with pdfplumber.open(pdf_path) as pdf:
        overview_text = ""
        for page in pdf.pages[:7]:
            page_text = page.extract_text()
            if page_text and ("보험종류" in page_text or "보험가입금액" in page_text):
                overview_text += page_text + "\n"

    if not overview_text:
        with pdfplumber.open(pdf_path) as pdf:
            overview_text = ""
            for page in pdf.pages[:7]:
                page_text = page.extract_text()
                if page_text:
                    overview_text += page_text + "\n"

    lines = overview_text.split('\n')
    person_name = _detect_person_name(lines)
    results = _parse_mirae_blocks(lines, person_name)

    if not results:
        results = _parse_mirae_benefit_section(pdf_path)

    main_info = _detect_main_contract_benefit(pdf_path)

    if main_info:
        benefit_name = main_info["benefit_name"]
        found_main = False
        for r in results:
            if "주계약" in r["특약명"]:
                r["특약명"] = benefit_name
                found_main = True
                break
        if not found_main and main_info["amount"]:
            results.append({
                "특약명": benefit_name,
                "가입금액": main_info["amount"]
            })

    return results


def _detect_person_name(lines):
    """PDF에서 피보험자 이름을 자동 감지"""
    for line in lines:
        line = line.strip()
        match = re.search(r'([가-힣]{2,4})\((?:남자|여자)', line)
        if match:
            return match.group(1)
        match = re.search(r'피보험자\s+([가-힣]{2,4})', line)
        if match:
            return match.group(1)
    return None


def _parse_mirae_blocks(lines, person_name):
    """미래에셋 PDF를 블록 단위로 파싱하여 특약명+가입금액 추출"""
    results = []
    if not person_name:
        return results

    name_buffer = ""
    i = 0

    while i < len(lines):
        line = lines[i].strip()
        line_clean = re.sub(r'^#+\s*', '', line).strip()
        i += 1

        if not line_clean:
            continue

        if any(kw in line_clean for kw in [
            "가입안내서", "발행번호", "Page", "FC :", "Tel :",
            "발행일시", "동일한 번호", "페이지로 구성",
        ]):
            continue

        if person_name in line_clean:
            after_person = line_clean.split(person_name)[-1].strip()
            amount_match = re.search(r'^[\s,]*(\d[\d,]*)', after_person)

            amount = None
            if amount_match:
                val = int(amount_match.group(1).replace(',', ''))
                if 1 <= val <= 100000:
                    amount = val * 10000
            else:
                for offset in range(0, 4):
                    if i + offset >= len(lines):
                        break
                    next_line = re.sub(r'^#+\s*', '', lines[i + offset].strip()).strip()
                    num_match = re.match(r'^(\d[\d,]*)$', next_line)
                    if num_match:
                        val = int(num_match.group(1).replace(',', ''))
                        if 1 <= val <= 100000:
                            amount = val * 10000
                            break

            if amount and name_buffer:
                clean_name = _clean_mirae_coverage_name(name_buffer)
                if clean_name and len(clean_name) >= 2:
                    if "납입면제" in clean_name:
                        name_buffer = ""
                        continue
                    if not any(r["특약명"] == clean_name for r in results):
                        results.append({"특약명": clean_name, "가입금액": amount})

            name_buffer = ""
            continue

        if re.match(r'^[\d,.\s원세년월납]+$', line_clean):
            continue
        if line_clean in ["전기납", "월납", "연납", "0"]:
            continue

        if line_clean in [
            "보험종류", "보험가입금액", "보험기간",
            "보험가입금액 (만원)", "보험료(원)",
            "가입 나이", "납입기간", "납입주기",
            "피보험자",
        ]:
            continue

        if "가입금액" in line_clean and "만원" in line_clean:
            continue

        if re.match(r'^최초계약\s+\d+', line_clean):
            continue
        if re.match(r'^갱신계약\s+\d+', line_clean):
            continue
        if re.match(r'^최대\s+\d+', line_clean):
            continue

        coverage_keywords = [
            "특약", "진단", "수술", "치료", "입원", "통원",
            "골절", "깁스", "사망", "장해", "배상", "벌금",
            "주계약", "보장", "보험금",
        ]

        has_keyword = any(kw in line_clean for kw in coverage_keywords)

        if has_keyword:
            name_buffer = line_clean
        elif name_buffer:
            if any(kw in line_clean for kw in ["최초계약", "갱신형)", "형)", "5)", "간편고지"]):
                name_buffer += " " + line_clean
            elif len(line_clean) > 3 and not re.match(r'^\d', line_clean):
                name_buffer += " " + line_clean

    return results


def _detect_main_contract_benefit(pdf_path):
    """미래에셋 PDF 보장내역 섹션에서 주계약의 실제 보장내용과 금액 감지"""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:15]:
            text = page.extract_text()
            if not text:
                continue
            if "주계약 보장내역" not in text:
                continue

            lines = text.split('\n')
            in_main_section = False
            benefit_name = None
            amount = None

            for idx, line in enumerate(lines):
                line_clean = re.sub(r'^#+\s*', '', line.strip()).strip()
                if "주계약 보장내역" in line_clean:
                    in_main_section = True
                    continue
                if "선택특약 보장내역" in line_clean:
                    break
                if not in_main_section:
                    continue

                bracket_match = re.search(r'\[([^\]]+보험금[^\]]*)\]', line_clean)
                if bracket_match:
                    benefit_raw = bracket_match.group(1).strip()
                    benefit_name = _map_main_benefit_name(benefit_raw)

                if benefit_name and amount is None:
                    amount_match = re.search(r'(\d[\d,]*)\s*만원', line_clean)
                    if amount_match:
                        amount = int(amount_match.group(1).replace(',', '')) * 10000

                if benefit_name and amount:
                    return {"benefit_name": benefit_name, "amount": amount}

    return None


def _map_main_benefit_name(benefit_raw):
    """주계약 보장명을 표준화된 특약명으로 매핑"""
    benefit_clean = benefit_raw.replace(" ", "")
    mapping = {
        "재해사망보험금": "주계약(재해사망)",
        "사망보험금": "주계약(일반사망)",
        "재해사망": "주계약(재해사망)",
        "일반사망": "주계약(일반사망)",
        "질병사망보험금": "주계약(질병사망)",
    }
    for key, value in mapping.items():
        if key in benefit_clean:
            return value
    return f"주계약({benefit_raw})"


def _parse_mirae_benefit_section(pdf_path):
    """미래에셋 PDF 보장내역 섹션에서 지급금액 기반 추출 (보완용)"""
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:15]:
            text = page.extract_text()
            if not text:
                continue
            if "보장내역" not in text and "지급사유" not in text:
                continue

            lines = text.split('\n')
            current_name = None

            for line in lines:
                line_clean = re.sub(r'^#+\s*', '', line.strip()).strip()
                if any(kw in line_clean for kw in ["특약", "주계약"]) and "대상" not in line_clean:
                    current_name = _clean_mirae_coverage_name(line_clean)

                amount_match = re.search(r'(\d[\d,]*)\s*만원', line_clean)
                if amount_match and current_name:
                    val = int(amount_match.group(1).replace(',', ''))
                    amount = val * 10000
                    if amount > 0 and len(current_name) >= 2:
                        if "납입면제" not in current_name:
                            if not any(r["특약명"] == current_name for r in results):
                                results.append({"특약명": current_name, "가입금액": amount})
                    current_name = None

    return results


def _clean_mirae_coverage_name(name):
    """미래에셋 특약명 정리"""
    name = name.strip()
    name = re.sub(r'^#+\s*', '', name)
    name = re.split(r'\s*최초계약', name)[0].strip()
    name = re.split(r'\s*최초\s*계약', name)[0].strip()
    name = name.strip()
    return name


# ══════════════════════════════════════════════
# 삼성생명 파서
# ══════════════════════════════════════════════

def extract_coverage_samsung(pdf_path):
    """삼성생명 PDF 파싱 — 텍스트 기반 (주력)"""
    results = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"

    lines = full_text.split('\n')

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1

        if not line:
            continue

        match = re.match(
            r'^(\d{1,4})\s+(.+?)\s+(\d[\d,]*(?:억|천만|백만|만|천)?원)\s+(\d+년갱신)',
            line
        )
        if match:
            name = match.group(2).strip()
            amount_text = match.group(3).strip()
            amount = parse_amount(amount_text)

            if name and amount and len(name) >= 5:
                name_nospace = name.replace(" ", "")
                skip_words = ["합계보험료", "주보험"]
                if not any(sk in name_nospace for sk in skip_words):
                    if not any(r["특약명"] == name for r in results):
                        results.append({"특약명": name, "가입금액": amount})
            continue

        match_name = re.match(r'^(\d{1,4})\s+(.+)', line)
        if match_name:
            num = match_name.group(1)
            name_part = match_name.group(2).strip()

            if any(kw in name_part for kw in [
                "보장특약", "수술", "진단", "사망", "무배당",
                "양성신생물", "파워수술", "질병", "재해", "여성특정"
            ]):
                amount_in_line = re.search(r'(\d[\d,]*(?:억|천만|백만|만|천)?원)', name_part)
                if amount_in_line:
                    amount = parse_amount(amount_in_line.group(1))
                    name = name_part[:amount_in_line.start()].strip()
                    if name and amount and len(name) >= 5:
                        if not any(r["특약명"] == name for r in results):
                            results.append({"특약명": name, "가입금액": amount})
                    continue

                for look_ahead in range(i, min(i + 3, len(lines))):
                    next_line = lines[look_ahead].strip()
                    amount_match = re.match(r'^(\d[\d,]*(?:억|천만|백만|만|천)?원)', next_line)
                    if amount_match:
                        amount = parse_amount(amount_match.group(1))
                        if name_part and amount and len(name_part) >= 5:
                            if not any(r["특약명"] == name_part for r in results):
                                results.append({"특약명": name_part, "가입금액": amount})
                        break
                continue

        if "재해사망" in line and "보험금" in line:
            amount_match = re.search(r'(\d[\d,]*(?:억|천만|백만|만|천)?원)', line)
            if amount_match:
                amount = parse_amount(amount_match.group(1))
                if amount and not any(r["특약명"] == "주보험 재해사망" for r in results):
                    results.append({"특약명": "주보험 재해사망", "가입금액": amount})
            continue

        if line.startswith("재해사망보험금"):
            for look_ahead in range(i, min(i + 3, len(lines))):
                next_line = lines[look_ahead].strip()
                amount_match = re.search(r'(\d[\d,]*(?:억|천만|백만|만|천)?원)', next_line)
                if amount_match:
                    amount = parse_amount(amount_match.group(1))
                    if amount and not any(r["특약명"] == "주보험 재해사망" for r in results):
                        results.append({"특약명": "주보험 재해사망", "가입금액": amount})
                    break
            continue

    if not results:
        results = extract_coverage_samsung_table(pdf_path)

    return results


def extract_coverage_samsung_table(pdf_path):
    """삼성생명 PDF 테이블 기반 파싱 (보완용)"""
    results = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue
            if not any(kw in text for kw in ["계약사항", "보험가입금액", "보장내용"]):
                continue

            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 5,
                "join_tolerance": 5,
            })

            if not tables:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                })

            for table in tables:
                if not table or len(table) < 2:
                    continue

                header_idx = None
                amount_col = None

                for i_row, row in enumerate(table):
                    if i_row > 5:
                        break
                    for j, cell in enumerate(row):
                        if not cell:
                            continue
                        cell_clean = cell.replace(" ", "").replace("\n", "")
                        if cell_clean in ["구분", "구분번호"]:
                            header_idx = i_row
                        if "보험가입금액" in cell_clean or "가입금액" in cell_clean or "지급금액" in cell_clean:
                            amount_col = j
                    if header_idx is not None and amount_col is not None:
                        break

                if header_idx is None or amount_col is None:
                    continue

                name_col = find_name_col_from_data(table, header_idx)
                if name_col is None:
                    continue

                for row in table[header_idx + 1:]:
                    if not row or len(row) <= max(name_col, amount_col):
                        continue

                    raw_name = row[name_col]
                    if not raw_name:
                        for cell in row:
                            if cell and len(cell.strip()) > 10 and any(kw in cell for kw in [
                                "보장특약", "수술", "진단", "사망", "무배당"
                            ]):
                                raw_name = cell
                                break
                    if not raw_name:
                        continue

                    raw_name = raw_name.replace("\n", " ")
                    raw_name = re.sub(r'\s+', ' ', raw_name).strip()

                    if len(raw_name) < 5:
                        continue

                    name_nospace = raw_name.replace(" ", "")
                    skip_words = ["합계보험료", "보험료", "경과년도", "납입기간", "보험기간"]
                    if any(sk in name_nospace for sk in skip_words):
                        continue

                    amount = None
                    if len(row) > amount_col and row[amount_col]:
                        amount = parse_amount(row[amount_col])
                    if amount is None:
                        for cell in row:
                            if cell and cell != raw_name:
                                parsed = parse_amount(cell)
                                if parsed:
                                    amount = parsed
                                    break

                    if raw_name and amount:
                        clean_name = re.sub(r'^\d+\s+', '', raw_name).strip()
                        if clean_name and not any(r["특약명"] == clean_name for r in results):
                            results.append({"특약명": clean_name, "가입금액": amount})

    return results


# ══════════════════════════════════════════════
# 통합 파싱 (PDF 1회 오픈) + 메인 분기 + 범용 파서
# ══════════════════════════════════════════════

def parse_pdf_all_in_one(pdf_path):
    """PDF를 1번만 열어서 보험사/상품명/보험료/특약 모두 추출 (최적화 버전)"""
    # 1단계: PDF를 1번만 열어서 텍스트 추출 + 키워드 있는 페이지만 테이블 추출
    coverage_keywords = [
        '특약', '담보', '가입금액', '보장내용',
        '보장내역', '가입담보', '보장항목'
    ]
    page_texts = []
    page_tables = {}  # {page_index: tables} — 키워드 있는 페이지만

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            page_texts.append(text or "")

            # 키워드가 있는 페이지만 테이블 추출 (비용이 큰 작업)
            if text and any(kw in text for kw in coverage_keywords):
                tables = page.extract_tables({
                    'vertical_strategy': 'lines',
                    'horizontal_strategy': 'lines',
                    'snap_tolerance': 5,
                    'join_tolerance': 5,
                })
                if not tables:
                    tables = page.extract_tables({
                        'vertical_strategy': 'text',
                        'horizontal_strategy': 'text',
                    })
                if tables:
                    page_tables[i] = tables

    # 2단계: 보험사 감지 (텍스트 재사용)
    combined_3 = "\n".join(page_texts[:3])
    insurer_code = _detect_insurer_from_text(combined_3)

    # 3단계: 상품명 추출 (텍스트 재사용)
    product_name = _detect_product_name_from_text(page_texts[:3])

    # 4단계: 보험료 추출 (텍스트 재사용)
    premium = _extract_premium_from_texts(page_texts[:7])

    # 5단계: 특약 추출 (보험사별 분기)
    if insurer_code == "samsung_life":
        coverages = _extract_coverage_samsung_from_texts(page_texts, pdf_path)
    elif insurer_code == "mirae":
        coverages = extract_coverage_mirae(pdf_path)
    elif insurer_code == "kb":
        coverages = _extract_coverage_kb_from_texts(page_texts, pdf_path)
    else:
        # 범용 파서 — 미리 추출한 텍스트+테이블 재사용 (PDF 재오픈 안 함)
        coverages = _extract_coverage_generic_from_cache(page_texts, page_tables, pdf_path)

    insurer_name_map = {
        "meritz": "메리츠화재", "samsung": "삼성화재",
        "samsung_life": "삼성생명", "kb": "KB손해보험",
        "db": "DB손해보험", "mirae": "미래에셋생명",
        "abl": "ABL생명", "heungkuk": "흥국생명",
        "hanwha": "한화생명", "hyundai": "현대해상",
        "lotte": "롯데손해보험", "nh": "NH농협생명",
        "dongyang": "동양생명", "kyobo": "교보생명",
        "shinhan": "신한라이프",
    }

    return {
        "insurer_code": insurer_code,
        "insurer_name": insurer_name_map.get(insurer_code, insurer_code or "알 수 없음"),
        "product_name": product_name,
        "premium": premium,
        "coverages": coverages,
    }


def _detect_insurer_from_text(text):
    """이미 추출된 텍스트에서 보험사 감지"""
    keywords_ordered = [
        ("삼성생명", "samsung_life"),
        ("삼성화재", "samsung"),
        ("메리츠화재", "meritz"),
        ("메리츠", "meritz"),
        ("미래에셋생명", "mirae"),
        ("미래에셋", "mirae"),
        ("KB손해", "kb"), ("KB손보", "kb"),
        ("KB 플러스", "kb"), ("KB플러스", "kb"),
        ("DB손해", "db"), ("DB손보", "db"),
        ("ABL", "abl"), ("에이비엘", "abl"),
        ("흥국", "heungkuk"), ("한화", "hanwha"),
        ("현대해상", "hyundai"), ("롯데손해", "lotte"),
        ("NH농협", "nh"), ("동양생명", "dongyang"),
        ("교보생명", "kyobo"), ("신한라이프", "shinhan"),
    ]
    for keyword, insurer in keywords_ordered:
        if keyword in text:
            return insurer
    if "kbinsure" in text.lower():
        return "kb"
    if "삼성" in text:
        if any(kw in text for kw in ["생명보험", "건강보험", "종신보험", "The간편한", "다모은"]):
            return "samsung_life"
        return "samsung"
    return None


def _detect_product_name_from_text(page_texts):
    """이미 추출된 페이지 텍스트에서 상품명 추출"""
    for text in page_texts:
        if not text:
            continue
        for line in text.split('\n'):
            line = line.strip()

            kb_match = re.search(r'(KB\s*플러스\s*[^\(\n]+보험)', line)
            if kb_match:
                name = kb_match.group(1).strip()
                name = re.sub(r'\(무배당\).*', '', name).strip()
                return name[:30] if len(name) > 30 else name

            kb_match2 = re.search(r'(KB\s*[^\(\n]*보험[^\(\n]*)\(무배당\)', line)
            if kb_match2:
                name = kb_match2.group(1).strip()
                return name[:30] if len(name) > 30 else name

            mirae_match = re.search(r'(M-케어\s*건강[^\(]*)', line)
            if mirae_match:
                name = mirae_match.group(1).strip()
                return name[:30] if len(name) > 30 else name

            if re.match(r'^\(무\)|^\(유\)', line):
                name = re.sub(r'^\(무\)\s*|^\(유\)\s*', '', line)
                match = re.match(r'^([^(]+)', name)
                if match:
                    name = match.group(1).strip()
                return name[:30] if len(name) > 30 else name

            samsung_match = re.match(r'^삼성\s+(.+보험)', line)
            if samsung_match:
                name = samsung_match.group(1).strip()
                name = re.sub(r'\(\d{4}\).*', '', name).strip()
                return name[:30] if len(name) > 30 else name

            meritz_match = re.search(r'(메리츠\s*[^\(\n]*보험[^\(\n]*)', line)
            if meritz_match:
                name = meritz_match.group(1).strip()
                name = re.sub(r'\(무배당\).*', '', name).strip()
                return name[:30] if len(name) > 30 else name

    return "상품명 미확인"


def _extract_premium_from_texts(page_texts):
    """이미 추출된 페이지 텍스트에서 보험료 추출"""
    patterns = [
        r'실납입보험료\s*([\d,]+)\s*원',
        r'1회차보험료\(할인후\)\s*([\d,]+)\s*원',
        r'할인후초회보험료\s*([\d,]+)\s*원',
        r'보장보험료\s*합계\s*([\d,]+)\s*원',
        r'합\s*계\s*보\s*험\s*료\s*([\d,]+)\s*원',
        r'합계보험료\s*([\d,]+)\s*원',
        r'합\s*계\s*([\d,]+)',
        r'보험료\s*[:\s]?\s*([\d,]+)\s*원',
    ]
    for text in page_texts:
        if not text:
            continue
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                val = int(match.group(1).replace(',', ''))
                if val >= 1000:
                    return val
    return None


def _extract_coverage_samsung_from_texts(page_texts, pdf_path):
    """삼성생명 - 미리 추출된 텍스트로 특약 파싱 (페이지 제한)"""
    results = []
    # 5~8페이지에 특약 정보가 집중 (인덱스 4~7)
    relevant_texts = page_texts[4:8] if len(page_texts) > 4 else page_texts
    full_text = "\n".join(relevant_texts)

    lines = full_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line:
            continue

        match = re.match(
            r'^(\d{1,4})\s+(.+?)\s+(\d[\d,]*(?:억|천만|백만|만|천)?원)\s+(\d+년갱신)',
            line
        )
        if match:
            name = match.group(2).strip()
            amount_text = match.group(3).strip()
            amount = parse_amount(amount_text)
            if name and amount and len(name) >= 5:
                name_nospace = name.replace(" ", "")
                if not any(sk in name_nospace for sk in ["합계보험료", "주보험"]):
                    if not any(r["특약명"] == name for r in results):
                        results.append({"특약명": name, "가입금액": amount})
            continue

        match_name = re.match(r'^(\d{1,4})\s+(.+)', line)
        if match_name:
            name_part = match_name.group(2).strip()
            if any(kw in name_part for kw in [
                "보장특약", "수술", "진단", "사망", "무배당",
                "양성신생물", "파워수술", "질병", "재해", "여성특정"
            ]):
                amount_in_line = re.search(r'(\d[\d,]*(?:억|천만|백만|만|천)?원)', name_part)
                if amount_in_line:
                    amount = parse_amount(amount_in_line.group(1))
                    name = name_part[:amount_in_line.start()].strip()
                    if name and amount and len(name) >= 5:
                        if not any(r["특약명"] == name for r in results):
                            results.append({"특약명": name, "가입금액": amount})
                    continue

                for look_ahead in range(i, min(i + 3, len(lines))):
                    next_line = lines[look_ahead].strip()
                    amount_match = re.match(r'^(\d[\d,]*(?:억|천만|백만|만|천)?원)', next_line)
                    if amount_match:
                        amount = parse_amount(amount_match.group(1))
                        if name_part and amount and len(name_part) >= 5:
                            if not any(r["특약명"] == name_part for r in results):
                                results.append({"특약명": name_part, "가입금액": amount})
                        break
                continue

        if "재해사망" in line and "보험금" in line:
            amount_match = re.search(r'(\d[\d,]*(?:억|천만|백만|만|천)?원)', line)
            if amount_match:
                amount = parse_amount(amount_match.group(1))
                if amount and not any(r["특약명"] == "주보험 재해사망" for r in results):
                    results.append({"특약명": "주보험 재해사망", "가입금액": amount})
            continue

    # 텍스트 기반으로 결과 없으면 테이블 파싱 시도 (원본 함수 사용)
    if not results:
        results = extract_coverage_samsung_table(pdf_path)

    return results


def _extract_coverage_kb_from_texts(page_texts, pdf_path):
    """KB손해보험 - 미리 추출된 텍스트로 특약 파싱"""
    results = []
    full_text = "\n".join(page_texts)

    for page_text in page_texts[:10]:
        if not page_text:
            continue
        is_coverage_page = any(kw in page_text for kw in [
            "가입담보", "가입내용", "보장명", "가입금액"
        ])
        if not is_coverage_page:
            continue
        _parse_kb_coverage_page(page_text, results)

    seen = set()
    unique_results = []
    for r in results:
        if r["특약명"] not in seen:
            seen.add(r["특약명"])
            unique_results.append(r)
    results = unique_results

    _enrich_kb_injury_grade_info(results, full_text)
    return results


def extract_coverage_from_pdf(pdf_path):
    """PDF에서 특약명 + 가입금액 추출 (보험사별 분기) — 레거시 호환"""
    result = parse_pdf_all_in_one(pdf_path)
    return result["coverages"]


def _extract_coverage_generic_from_cache(page_texts, page_tables, pdf_path):
    """범용 파서 — 미리 추출된 텍스트+테이블 캐시 사용 (PDF 재오픈 없음)"""
    results = []
    sub_prefix_pattern = re.compile(r'^┗?\s*\d+\s+')

    header_keywords = [
        "가입담보", "가입담보및보장내용", "담보명", "특약명",
        "보장명", "담보내용", "보장항목", "보장내용",
        "급부명", "보장담보", "보험종목", "보장종목"
    ]
    amount_keywords = ["가입금액", "보험가입금액", "보장금액"]

    for page_idx, tables in page_tables.items():
        for table in tables:
            if not table or len(table) < 2:
                continue

            header_idx = None
            amount_col = None

            for idx, row in enumerate(table):
                if idx > 5:
                    break
                for j, cell in enumerate(row):
                    if not cell:
                        continue
                    cell_clean = cell.replace(" ", "").replace("\n", "")
                    if any(kw in cell_clean for kw in header_keywords):
                        header_idx = idx
                    if any(kw in cell_clean for kw in amount_keywords):
                        amount_col = j
                if header_idx is not None:
                    break

            if header_idx is None or amount_col is None:
                continue

            name_col = find_name_col_from_data(table, header_idx)
            if name_col is None:
                continue

            for row in table[header_idx + 1:]:
                if not row or len(row) <= max(name_col, amount_col):
                    continue

                name = row[name_col]
                if not name:
                    continue

                name = name.replace("\n", " ")
                name = re.sub(r'\s+', ' ', name).strip()

                if len(name) < 5:
                    continue

                skip_names = [
                    "주계약", "선택특약", "필수특약", "의무특약",
                    "합계", "총보험료", "보장보험료", "보험료합계",
                    "2회차이후", "1회차보험료", "계약자명",
                    "보장보험료합계", "적립보험료", "할인보험료",
                    "선택계약", "보험료사항",
                    "보험료자동납입", "주의사항"
                ]
                name_nospace = name.replace(" ", "")

                if name_nospace == "기본계약":
                    continue
                if any(sk in name_nospace for sk in skip_names):
                    continue
                if re.match(r'^[\d,.\s%원]+$', name):
                    continue
                if re.match(r'^\d+년\s*\(\d+세\)', name_nospace):
                    continue
                if "경우" in name and "지급" in name:
                    continue
                if len(name) > 100 and ("보험기간" in name or "최초계약" in name):
                    continue

                name = sub_prefix_pattern.sub('', name).strip()
                name = re.sub(r'\(필수\)|\(선택\)', '', name).strip()

                if len(name) < 5:
                    continue

                amount = None
                if len(row) > amount_col and row[amount_col]:
                    amount = parse_amount(row[amount_col])

                if amount is None:
                    for cell in row:
                        if cell and cell != name:
                            parsed = parse_amount(cell)
                            if parsed:
                                amount = parsed
                                break

                if name and amount:
                    if not any(r["특약명"] == name for r in results):
                        results.append({"특약명": name, "가입금액": amount})

    return results


def extract_coverage_generic(pdf_path):
    """범용 PDF 파싱 (메리츠 등)"""
    results = []
    sub_prefix_pattern = re.compile(r'^┗?\s*\d+\s+')

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue

            if not any(kw in text for kw in [
                "특약", "담보", "가입금액", "보장내용",
                "보장내역", "가입담보", "보장항목"
            ]):
                continue

            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 5,
                "join_tolerance": 5,
            })

            if not tables:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                })

            for table in tables:
                if not table or len(table) < 2:
                    continue

                header_idx = None
                amount_col = None

                header_keywords = [
                    "가입담보", "가입담보및보장내용", "담보명", "특약명",
                    "보장명", "담보내용", "보장항목", "보장내용",
                    "급부명", "보장담보", "보험종목", "보장종목"
                ]
                amount_keywords = ["가입금액", "보험가입금액", "보장금액"]

                for idx, row in enumerate(table):
                    if idx > 5:
                        break
                    for j, cell in enumerate(row):
                        if not cell:
                            continue
                        cell_clean = cell.replace(" ", "").replace("\n", "")
                        if any(kw in cell_clean for kw in header_keywords):
                            header_idx = idx
                        if any(kw in cell_clean for kw in amount_keywords):
                            amount_col = j
                    if header_idx is not None:
                        break

                if header_idx is None or amount_col is None:
                    continue

                name_col = find_name_col_from_data(table, header_idx)
                if name_col is None:
                    continue

                for row in table[header_idx + 1:]:
                    if not row or len(row) <= max(name_col, amount_col):
                        continue

                    name = row[name_col]
                    if not name:
                        continue

                    name = name.replace("\n", " ")
                    name = re.sub(r'\s+', ' ', name).strip()

                    if len(name) < 5:
                        continue

                    skip_names = [
                        "주계약", "선택특약", "필수특약", "의무특약",
                        "합계", "총보험료", "보장보험료", "보험료합계",
                        "2회차이후", "1회차보험료", "계약자명",
                        "보장보험료합계", "적립보험료", "할인보험료",
                        "선택계약", "보험료사항",
                        "보험료자동납입", "주의사항"
                    ]
                    name_nospace = name.replace(" ", "")

                    if name_nospace == "기본계약":
                        continue
                    if any(sk in name_nospace for sk in skip_names):
                        continue
                    if re.match(r'^[\d,.\s%원]+$', name):
                        continue
                    if re.match(r'^\d+년\s*\(\d+세\)', name_nospace):
                        continue
                    if "경우" in name and "지급" in name:
                        continue
                    if len(name) > 100 and ("보험기간" in name or "최초계약" in name):
                        continue

                    name = sub_prefix_pattern.sub('', name).strip()
                    name = re.sub(r'\(필수\)|\(선택\)', '', name).strip()

                    if len(name) < 5:
                        continue

                    amount = None
                    if len(row) > amount_col and row[amount_col]:
                        amount = parse_amount(row[amount_col])

                    if amount is None:
                        for cell in row:
                            if cell and cell != name:
                                parsed = parse_amount(cell)
                                if parsed:
                                    amount = parsed
                                    break

                    if name and amount:
                        if not any(r["특약명"] == name for r in results):
                            results.append({"특약명": name, "가입금액": amount})

    return results


def parse_amount(text):
    """금액 텍스트를 숫자(원 단위)로 변환"""
    if not text:
        return None
    text = text.replace(" ", "").replace(",", "")

    if re.search(r'[가-힣]{2,}', text) and "원" not in text and "만" not in text and "억" not in text:
        return None

    match = re.search(r'(\d+)(억원|억)', text)
    if match:
        return int(match.group(1)) * 100000000

    match = re.search(r'(\d+)(천만원)', text)
    if match:
        return int(match.group(1)) * 10000000

    match = re.search(r'(\d+)(백만원|백만)', text)
    if match:
        return int(match.group(1)) * 1000000

    match = re.search(r'(\d+)(만원|만)', text)
    if match:
        return int(match.group(1)) * 10000

    match = re.search(r'(\d+)(천원)', text)
    if match:
        return int(match.group(1)) * 1000

    match = re.search(r'(\d+)(원)?$', text)
    if match:
        num = int(match.group(1))
        if num >= 10000:
            return num

    return None
