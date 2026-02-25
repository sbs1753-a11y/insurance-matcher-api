from rapidfuzz import fuzz, process
import re


def normalize_name(name):
    """매칭 정확도를 높이기 위한 정규화"""
    name = name.strip()
    roman_map = {"Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5"}
    for roman, arabic in roman_map.items():
        name = name.replace(roman, arabic)
    name = re.sub(r'[\s\-\_\.\·]+', '', name)
    name = name.replace('（', '(').replace('）', ')')
    return name


def simplify_pdf_name(name):
    """PDF 특약명에서 핵심 키워드만 추출"""
    name = name.strip()
    roman_map = {"Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5"}
    for roman, arabic in roman_map.items():
        name = name.replace(roman, arabic)

    # 메리츠 정리
    name = re.sub(r'^┗\s*', '', name)
    name = re.sub(r'^\([^)]*갱신\)\s*', '', name)
    name = re.sub(r'갱신형\s*', '', name)
    name = re.sub(r'\(통합간편[^)]*\)', '', name)
    name = re.sub(r'\[기본계약\]', '', name)
    name = re.sub(r'\(연간\d+회한?\)', '', name)
    name = re.sub(r'\(급여[,)]\s*', '(', name)
    name = re.sub(r'^\(\s*', '', name)
    name = re.sub(r'\(\d+%체증형\)', '', name)
    name = re.sub(r'\(1-\d+종\)', '', name)

    # 삼성생명 정리
    name = re.sub(r'\(갱신형,\s*무배당\)', '', name)
    name = re.sub(r'\(갱신형\)', '', name)
    name = re.sub(r'\(무배당\)', '', name)
    name = re.sub(r'보장특약[Ⅰ-Ⅴ1-5]*U?', '', name)
    name = re.sub(r'특약[Ⅰ-Ⅴ1-5]*U?', '', name)
    name = re.sub(r'U$', '', name)
    name = re.sub(r'·', '', name)

    # 미래에셋 정리
    name = re.sub(r'\(간편고지형\(\d+\)[,\s]*갱신형\)', '', name)
    name = re.sub(r'\(간편고지형\(\d+\)\)', '', name)
    name = re.sub(r'최초계약', '', name)
    name = re.sub(r'##\s*', '', name)

    # KB손해 운전자보험 정리
    name = re.sub(r'\(운전자\)', '', name)
    name = re.sub(r'\(기본계약\)', '', name)
    name = re.sub(r'\(비탑승중포함\)', '', name)
    name = re.sub(r'\(경찰조사포함\)', '', name)
    name = re.sub(r'\(스쿨존사고\s*\d+천?만원한도\)', '', name)
    name = re.sub(r'\(중상해보장확대\)', '', name)
    name = re.sub(r'\(중대법규위반[^)]*\)', '', name)
    name = re.sub(r'\(대물\)', '', name)
    name = re.sub(r'\(치아파절포함\)', '', name)
    name = re.sub(r'\(치아파절제외\)', '', name)
    name = re.sub(r'\(1일\d+회한[,\s]*연간\d+회한[,\s]*급여\)', '', name)
    name = re.sub(r'\(1일\d+회한[,\s]*연간\d+회한\)', '', name)
    name = re.sub(r'\(\d+~\d+급\)', '', name)
    name = re.sub(r'\(\d+급\)', '', name)

    # 공통 정리
    name = re.sub(r'\s+', '', name)
    name = re.sub(r'\(\s*\)', '', name)
    name = name.strip()
    return name


def simplify_excel_name(name):
    """Excel 특약명 정규화"""
    name = name.strip()
    roman_map = {"Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5"}
    for roman, arabic in roman_map.items():
        name = name.replace(roman, arabic)
    name = re.sub(r'[\s\-\_\.\·]+', '', name)
    return name


def get_surgery_grade_amounts(pdf_coverages):
    """1종~7종 수술비의 상해/질병별 금액을 모두 수집하고 각 종별 최소값 반환"""
    grade_amounts = {i: [] for i in range(1, 8)}

    for cov in pdf_coverages:
        name = cov["특약명"]
        amount = cov["가입금액"]

        for grade in range(1, 8):
            patterns = [
                f"[상해{grade}종]",
                f"[질병{grade}종]",
                f"〔상해{grade}종〕",
                f"〔질병{grade}종〕",
                f"_{grade}종수술",
                f"_{grade}종 수술",
                f"_{grade}종수술보험금",
                f"_{grade}종 수술보험금",
            ]
            matched = False
            for pattern in patterns:
                if pattern in name:
                    grade_amounts[grade].append(amount)
                    matched = True
                    break
            if matched:
                break

    result = {}
    for grade, amounts in grade_amounts.items():
        if amounts:
            result[grade] = min(amounts)

    return result


def get_aggregated_amounts(pdf_coverages):
    """합산 규칙이 적용되는 특약들의 금액 계산"""
    simplified = {}
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        simplified[key] = cov["가입금액"]

    result = {}

    # 질병수술비
    disease_surgery = 0
    for key, amount in simplified.items():
        if "질병수술비" in key and "130대" not in key and "5대질환" not in key:
            disease_surgery += amount
        elif "질병재해수술" in key:
            disease_surgery += amount
    if disease_surgery > 0:
        result["질병수술비"] = disease_surgery

    # 상해수술비
    for key, amount in simplified.items():
        if key == "상해수술비":
            result["상해수술비"] = amount
            break
        elif "질병재해수술" in key and "상해수술비" not in result:
            result["상해수술비"] = amount

    # 뇌혈관질환 수술비
    brain_surgery = 0
    for key, amount in simplified.items():
        if "뇌혈관질환수술비" in key and "130대" not in key:
            brain_surgery += amount
        elif "130대질병수술비" in key and "뇌혈관질환" in key:
            brain_surgery += amount
    if brain_surgery > 0:
        result["뇌혈수술비"] = brain_surgery
        result["뇌혈관질환수술비"] = brain_surgery

    # 허혈성심장질환수술비
    heart_surgery = 0
    for key, amount in simplified.items():
        if "허혈성심장질환수술비" in key and "130대" not in key:
            heart_surgery += amount
        elif "130대질병수술비" in key and "심장질환" in key:
            heart_surgery += amount
    if heart_surgery > 0:
        result["허혈성심장질환수술비"] = heart_surgery

    # 골절진단 (합산)
    fracture_diag = 0
    for key, amount in simplified.items():
        if "골절" in key and ("진단" in key or "골절진단" in key) and "수술" not in key:
            fracture_diag += amount
    if fracture_diag > 0:
        result["골절진단"] = fracture_diag

    # 골절수술비
    for key, amount in simplified.items():
        if "골절수술비" in key:
            result["골절수술비"] = amount
            break

    # 뇌혈관질환 진단비
    for key, amount in simplified.items():
        if "뇌혈관질환진단" in key and "수술" not in key:
            result["뇌혈관질환진단비"] = amount
            break

    # 허혈성심장질환 진단비
    for key, amount in simplified.items():
        if "허혈성심장질환진단" in key and "수술" not in key:
            result["허혈성심장질환진단비"] = amount
            break

    # 일반상해사망 / 재해사망
    for key, amount in simplified.items():
        if any(kw in key for kw in ["일반상해사망", "재해사망", "주보험재해사망", "주계약(재해사망)"]):
            result["일반상해사망"] = amount
            break

    # 가족일상배상책임
    for key, amount in simplified.items():
        if "가족일상생활중배상책임" in key or "가족일상배상책임" in key:
            result["가족일상배상책임"] = amount
            break

    # 교통사고처리지원금 (중상해보장확대만 — 6주미만 제외)
    for cov in pdf_coverages:
        original_name = cov["특약명"]
        if "교통사고처리보장" in original_name and "중상해보장확대" in original_name:
            result["교통사고처리지원금"] = cov["가입금액"]
            break
    if "교통사고처리지원금" not in result:
        for key, amount in simplified.items():
            if "교통사고처리보장A" in key and "6주미만" not in key and "중대법규위반" not in key:
                result["교통사고처리지원금"] = amount
                break

    # 자동차사고부상 14등급 지급액
    for cov in pdf_coverages:
        if "14급지급액" in cov:
            result["자동차사고부상14등급"] = cov["14급지급액"]
            break
    if "자동차사고부상14등급" not in result:
        for cov in pdf_coverages:
            original_name = cov["특약명"]
            if "사고부상" in original_name and ("4~14" in original_name or "4~14급" in original_name):
                result["자동차사고부상14등급"] = round(cov["가입금액"] / 30)
                break

    return result


# Excel 약칭 → 매칭 규칙
MATCHING_RULES = {
    "보험료": "special_premium",
    "적립금": None,
    "실비질병/상해종합입원": None,
    "실비질병/상해통원치료": None,

    # 수술
    "질병수술비": {"type": "aggregate", "key": "질병수술비"},
    "상해수술비": {"type": "aggregate", "key": "상해수술비"},
    "1종수술": {"type": "surgery_grade", "grade": 1},
    "2종수술": {"type": "surgery_grade", "grade": 2},
    "3종수술": {"type": "surgery_grade", "grade": 3},
    "4종수술": {"type": "surgery_grade", "grade": 4},
    "5종수술": {"type": "surgery_grade", "grade": 5},
    "6종수술": {"type": "surgery_grade", "grade": 6},
    "7종수술": {"type": "surgery_grade", "grade": 7},

    # 암
    "암진단(일반암)": {"type": "direct", "keywords": [
        "암진단비", "일반암진단비", "암(유사암제외)진단비",
        "암(유사암제외)진단", "암(유사암제외)"
    ]},
    "소액/유사암": {"type": "direct", "keywords": [
        "유사암진단비", "소액암진단비", "유사암진단"
    ]},
    "전이암진단비": {"type": "direct", "keywords": ["전이암진단비"]},
    "암주요치료비": {"type": "direct", "keywords": ["암주요치료비"]},
    "항암방사선약물치료비": {"type": "direct", "keywords": [
        "항암방사선약물치료비", "항암약물치료비",
        "항암방사선치료", "항암약물치료"
    ]},
    "표적항암약물치료비": {"type": "direct", "keywords": [
        "표적항암약물치료비", "표적항암약물허가치료"
    ]},
    "양성자방사선치료비": {"type": "direct", "keywords": [
        "양성자방사선치료비", "항암양성자방사선치료"
    ]},
    "세기조절방사선치료비": {"type": "direct", "keywords": [
        "세기조절방사선치료비", "항암세기조절방사선치료"
    ]},
    "면역항암약물치료비": {"type": "direct", "keywords": [
        "면역항암약물치료비", "특정면역항암약물허가치료",
        "면역항암약물허가치료", "특정면역항암"
    ]},
    "카티항암약물치료비": {"type": "direct", "keywords": ["카티항암약물치료비", "CAR-T"]},
    "다빈치로봇수술비": {"type": "direct", "keywords": [
        "다빈치로봇수술비", "다빈치로봇수술",
        "암다빈치로봇수술"
    ]},

    # 뇌
    "뇌혈관질환진단비": {"type": "aggregate", "key": "뇌혈관질환진단비"},
    "뇌혈관질환": {"type": "aggregate", "key": "뇌혈관질환진단비"},
    "뇌혈수술비": {"type": "aggregate", "key": "뇌혈수술비"},
    "뇌혈관질환수술비": {"type": "aggregate", "key": "뇌혈관질환수술비"},
    "뇌졸증진단비": {"type": "direct", "keywords": ["뇌졸중진단비", "뇌졸증진단비"]},
    "뇌졸증": {"type": "direct", "keywords": ["뇌졸중진단비", "뇌졸증진단비"]},
    "뇌출혈진단비": {"type": "direct", "keywords": ["뇌출혈진단비"]},
    "뇌출혈": {"type": "direct", "keywords": ["뇌출혈진단비"]},

    # 심장
    "심혈관질환진단비": {"type": "direct", "keywords": ["심혈관질환진단비"]},
    "심혈관질환": {"type": "direct", "keywords": ["심혈관질환진단비"]},
    "허혈성심장질환진단비": {"type": "aggregate", "key": "허혈성심장질환진단비"},
    "허혈성심장질환": {"type": "aggregate", "key": "허혈성심장질환진단비"},
    "허혈성심장질환수술비": {"type": "aggregate", "key": "허혈성심장질환수술비"},
    "급성심근경색진단비": {"type": "direct", "keywords": ["급성심근경색진단비", "급성심근경색증진단비"]},
    "급성심근경색": {"type": "direct", "keywords": ["급성심근경색진단비", "급성심근경색증진단비"]},

    # 입원/응급
    "질병입원": {"type": "direct", "keywords": ["질병입원"]},
    "상해입원": {"type": "direct", "keywords": ["상해입원"]},
    "응급실내원(응급)": {"type": "direct", "keywords": ["응급실내원(응급)"]},
    "응급실내원(비응급)": {"type": "direct", "keywords": ["응급실내원(비응급)"]},

    # 사망
    "일반사망": {"type": "direct", "keywords": ["일반사망보장", "종신사망", "사망보장"]},
    "질병사망": {"type": "direct", "keywords": ["질병사망"]},
    "상해사망/재해사망": {"type": "aggregate", "key": "일반상해사망"},

    # 후유장해
    "상해후유장해3%": {"type": "direct", "keywords": ["상해3%이상후유장해", "상해후유장해3%"]},
    "질병후유장해3%": {"type": "direct", "keywords": ["질병3%이상후유장해", "질병후유장해3%"]},
    "상해후유장해": {"type": "direct", "keywords": ["상해후유장해", "일반상해80%이상후유장해"]},
    "질병후유장해": {"type": "direct", "keywords": ["질병후유장해", "질병80%이상후유장해"]},

    # 골절
    "골절진단": {"type": "aggregate", "key": "골절진단"},
    "골절수술": {"type": "aggregate", "key": "골절수술비"},
    "깁스": {"type": "direct", "keywords": ["깁스"]},

    # 치매
    "경도치매": {"type": "direct", "keywords": ["경도치매"]},
    "중증도치매": {"type": "direct", "keywords": ["중증도치매", "중등도치매"]},
    "중증치매": {"type": "direct", "keywords": ["중증치매진단비", "중증치매"]},
    "중증치매생활비": {"type": "direct", "keywords": ["중증치매생활비"]},

    # 운전자
    "벌금": {"type": "direct_exclude", "keywords": ["자동차사고벌금", "벌금"], "exclude": ["대물"]},
    "변호사선임비용": {"type": "direct", "keywords": [
        "변호사선임비용손해", "변호사선임비용", "변호사선임"
    ]},
    "교통사고처리지원금": {"type": "aggregate", "key": "교통사고처리지원금"},
    "대인형사합의금": {"type": "direct", "keywords": ["대인형사합의금"]},
    "자동차사고부상14등급": {"type": "aggregate", "key": "자동차사고부상14등급"},

    # 배상책임
    "가족일상배상책임": {"type": "aggregate", "key": "가족일상배상책임"},
}


def match_coverages(pdf_coverages, excel_coverages, threshold=70):
    """PDF 특약과 Excel 특약을 매칭"""
    results = []
    unmatched_pdf = []
    unmatched_excel = []

    aggregated = get_aggregated_amounts(pdf_coverages)
    surgery_grades = get_surgery_grade_amounts(pdf_coverages)
    pdf_simplified = {simplify_pdf_name(c["특약명"]): c for c in pdf_coverages}

    for excel_item in excel_coverages:
        excel_name = excel_item["특약명"]
        excel_norm = simplify_excel_name(excel_name)

        rule = MATCHING_RULES.get(excel_norm)

        if rule is None:
            unmatched_excel.append(excel_item)
            continue

        if rule == "special_premium":
            continue

        matched_amount = None
        matched_pdf_name = ""

        if isinstance(rule, dict):
            rule_type = rule["type"]

            if rule_type == "aggregate":
                key = rule["key"]
                if key in aggregated:
                    matched_amount = aggregated[key]
                    matched_pdf_name = f"[합산] {key}"

            elif rule_type == "surgery_grade":
                grade = rule["grade"]
                if grade in surgery_grades:
                    matched_amount = surgery_grades[grade]
                    matched_pdf_name = f"[최소값] {grade}종 수술"

            elif rule_type == "direct":
                keywords = rule["keywords"]
                for kw in keywords:
                    kw_clean = re.sub(r'\s+', '', kw)
                    for pdf_key, pdf_cov in pdf_simplified.items():
                        if kw_clean in pdf_key:
                            matched_amount = pdf_cov["가입금액"]
                            matched_pdf_name = pdf_cov["특약명"]
                            break
                    if matched_amount:
                        break

            elif rule_type == "direct_exclude":
                keywords = rule["keywords"]
                exclude = rule.get("exclude", [])
                for kw in keywords:
                    kw_clean = re.sub(r'\s+', '', kw)
                    for pdf_key, pdf_cov in pdf_simplified.items():
                        if kw_clean in pdf_key:
                            if any(ex in pdf_key for ex in exclude):
                                continue
                            matched_amount = pdf_cov["가입금액"]
                            matched_pdf_name = pdf_cov["특약명"]
                            break
                    if matched_amount:
                        break

        if matched_amount is not None:
            results.append({
                "excel_row": excel_item["row"],
                "excel_특약명": excel_item["특약명"],
                "pdf_특약명": matched_pdf_name,
                "가입금액": matched_amount,
                "유사도": 100.0,
                "amount_col": excel_item["amount_col"]
            })
        else:
            unmatched_excel.append(excel_item)

    for cov in pdf_coverages:
        is_used = False
        for r in results:
            if cov["특약명"] in r["pdf_특약명"] or simplify_pdf_name(cov["특약명"]) in r["pdf_특약명"]:
                is_used = True
                break
        if not is_used:
            unmatched_pdf.append(cov)

    return {
        "matched": results,
        "unmatched_excel": unmatched_excel,
        "unmatched_pdf": unmatched_pdf
    }
