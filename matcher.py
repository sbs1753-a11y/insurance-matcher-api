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

    # 현대해상 정리
    name = re.sub(r'\[맞춤고지[^\]]*\]', '', name)
    name = re.sub(r'담보$', '', name)
    name = re.sub(r'\(연간\d+회한\s*\)', '', name)
    name = re.sub(r'\(연간\d+회한,급여\)', '', name)
    name = re.sub(r'\(수술회당지급\)', '', name)
    name = re.sub(r'무배당', '', name)

    # 신한라이프 정리
    name = re.sub(r'\(무배당,\s*갱신형\)', '', name)
    name = re.sub(r',\s*갱신형\)', ')', name)
    name = re.sub(r'\[삭감없음용\]', '', name)
    name = re.sub(r'\[3-100%장해형\]', '', name)
    name = re.sub(r'비급여\(전액본인부담\s*포함\)', '', name)
    name = re.sub(r'\(치료별\s*연간\d*회?\)', '', name)
    name = re.sub(r'\(4대중증치료\)', '', name)
    name = re.sub(r'\(갑상선암및전립선암\s*제외\)', '', name)
    name = re.sub(r'\(갑상선암및전립선암\s*포함\)', '', name)

    # 라이나생명 정리
    name = re.sub(r'\(해약환급금미지급형2\)', '', name)
    name = re.sub(r'\(해약환급금미\s*지급형2\)', '', name)
    name = re.sub(r'\(해약환급금[^Ⅰ-ⅤI)]*형2?\)', '', name)
    name = re.sub(r'새로담는', '', name)
    name = re.sub(r'보장특약', '', name)  # 라이나 항암방사선약물치료"보장"특약 → "보장" 제거

    # DB손해보험 정리
    name = re.sub(r'^\(건강고지\)\s*', '', name)
    name = re.sub(r'^건강고지\)\s*', '', name)
    name = re.sub(r'\(동일사고당\d+회지급\)', '', name)
    name = re.sub(r'\(동일질병당\d+회지급\)', '', name)
    name = re.sub(r'\(매회지급\)', '', name)
    name = re.sub(r'\(\d+%체증형\)', '', name)
    name = re.sub(r'^체증형', '', name)  # 체증형뇌혈관질환수술비 → 뇌혈관질환수술비
    name = re.sub(r'\(최초\d+회한\)', '', name)

    # 공통 정리
    name = re.sub(r'\s+', '', name)
    name = re.sub(r'\(\s*,?\s*\)', '', name)  # (,) 빈 괄호 제거
    name = re.sub(r'\(\s*\)', '', name)  # () 빈 괄호 제거
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
    """1종~7종 수술비의 상해/질병별 금액을 모두 수집하고 각 종별 합산 반환
    
    규칙:
    1) 질병/상해(재해) 양쪽에 같은 종별 수술이 있으면 min(질병, 상해) 사용
       (흥국생명: 질병수술1종 + 재해수술1종 → min)
    2) 서로 다른 라이더 그룹(수술비Ⅱ(1-5종) vs 수술비(1-7종))이 있으면 sum
       (메리츠: 수술비Ⅱ[상해1종]=20 + 수술비[상해1종]=30 → 50)
    3) 미래에셋: 1-5종수술(1종) 20만 + [1-7종]1종 10만 → 30만
    """
    # {grade: {"group_name": {"상해": amt, "질병": amt}}}
    # 각 그룹 내에서 min(상해, 질병), 그룹 간 합산
    grade_groups = {i: {} for i in range(1, 8)}
    grade_extra = {i: [] for i in range(1, 8)}  # 미래에셋 1-7종 상세

    for cov in pdf_coverages:
        name = cov["특약명"]
        amount = cov["가입금액"]
        name_simplified = simplify_pdf_name(name)

        # 미래에셋 1-7종 상세 분류 ([1-7종]X종수술)
        for grade in range(1, 8):
            if name == f"[1-7종]{grade}종수술":
                grade_extra[grade].append(amount)
                break
        else:
            # 기본 종별 수술 (1-5종, 질병/상해/재해)
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
                    # 흥국생명: 1~5종질병수술(분리형,1종) 또는 1~5종재해수술
                    f"분리형,{grade}종)",
                    f",{grade}종)",
                    # 흥국생명: 보장내용 상세 페이지에서 추출된 재해 종별 수술
                    f"[재해]{grade}종수술",
                    # 라이나생명: 보장내역에서 추출된 종별 수술
                    f"[수술]{grade}종수술",
                    # 미래에셋: 1-5종수술(X종) 또는 1-5종수술특약(X종)
                    f"({grade}종)",
                ]
                matched = False
                for pattern in patterns:
                    if pattern in name:
                        # (N종) 패턴은 수술 관련 특약에서만 매칭
                        if pattern == f"({grade}종)" and "수술" not in name:
                            continue
                        # 그룹 결정: 수술비Ⅱ vs 수술비 vs 기타
                        if "수술비Ⅱ" in name or "수술비2" in name_simplified:
                            group = "수술비2"
                        elif re.match(r'^수술비\[', name):
                            group = "수술비1-7"
                        else:
                            group = "default"
                        # 상해/질병 구분
                        is_injury = "상해" in pattern or "재해" in pattern or "상해" in name
                        cat = "상해" if is_injury else "질병"
                        if group not in grade_groups[grade]:
                            grade_groups[grade][group] = {}
                        grade_groups[grade][group][cat] = amount
                        matched = True
                        break
                if not matched:
                    # 미래에셋 simplified: '1-5종수술(X종)무배당' 패턴
                    if f"({grade}종)" in name_simplified and "수술" in name_simplified:
                        group = "default"
                        grade_groups[grade].setdefault(group, {})["질병"] = amount
                        matched = True
                if matched:
                    break

    result = {}
    for grade in range(1, 8):
        total = 0
        # 각 그룹 내에서 min(상해, 질병) 적용, 그룹 간 합산
        for group, cats in grade_groups[grade].items():
            if len(cats) >= 2:
                total += min(cats.values())
            elif len(cats) == 1:
                total += list(cats.values())[0]
        # 미래에셋 1-7종 extra 추가
        extra_val = max(grade_extra[grade]) if grade_extra[grade] else 0
        total += extra_val
        if total > 0:
            result[grade] = total

    return result


def get_aggregated_amounts(pdf_coverages):
    """합산 규칙이 적용되는 특약들의 금액 계산"""
    simplified = {}
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        simplified[key] = cov["가입금액"]

    result = {}

    # 질병수술비
    # 삼성화재: "질병 입원 수술비Ⅱ" 30만원만 해당 (111대질병, 2대주요기관, 상급종합, 1~5종 등은 별도 항목)
    # 흥국생명: (무)질병수술(체증형)
    # 미래에셋: 질병수술무배당
    # 현대해상: 질병수술[맞춤고지2]담보 → simplified "질병수술" (종별/120대 제외)
    disease_surgery = 0
    # 제외 키워드: 삼성화재의 특수 수술 분류는 질병수술비에 포함하지 않음
    disease_surgery_exclude = [
        "130대", "131대", "5대질환", "111대", "2대주요기관", "상급종합병원",
        "1~5종", "1-5종", "충수염", "4대특정", "양성신생물",
        "통합양성", "관혈", "비관혈", "통원", "백내장",
        "다빈치", "스텐트", "풍선", "창상봉합",
        "120대", "26대", "58대", "24대", "치핵", "갑상선", "다발성",
        "1종", "2종", "3종", "4종", "5종",
        "특정5대질병",  # 메리츠: 질병수술비(특정5대질병 제외) 별도 특약
        "호흡기관련",  # 메리츠: 호흡기관련질병수술비 별도 특약
        "119대",  # DB손해: 119대질병수술비Ⅲ 별도 항목
        "특정경증",  # DB손해: 질병수술비(특정경증질환,백내장및대장용종제외) 별도 항목
    ]
    for key, amount in simplified.items():
        if "질병수술비" in key:
            if not any(ex in key for ex in disease_surgery_exclude):
                disease_surgery += amount
        elif "질병재해수술" in key:
            disease_surgery += amount
        # 흥국생명: (무)질병수술(체증형) → 질병수술비
        elif "질병수술(체증형)" in key or "질병수술(체증" in key:
            disease_surgery += amount
        # 미래에셋: 질병수술무배당 (백내장 제외 아닌 순수 질병수술)
        elif re.match(r'^질병수술[무배당최초]', key) and "백내장" not in key:
            disease_surgery += amount
        # 삼성화재: "질병입원수술비" or "질병통원수술비" (Ⅱ/Ⅳ) → 질병수술비 (입원+통원 중 입원만)
        elif re.match(r'^질병입원수술비', key) and "백내장" not in key:
            disease_surgery += amount
        # 현대해상: "질병수술" 단독 (종별/120대 등 제외)
        elif re.match(r'^질병수술[0-9]*$', key):
            disease_surgery += amount
    if disease_surgery > 0:
        result["질병수술비"] = disease_surgery

    # 상해수술비
    injury_surgery = 0
    injury_surgery_exclude = [
        "1종", "2종", "3종", "4종", "5종",
    ]
    for key, amount in simplified.items():
        if key == "상해수술비":
            injury_surgery = amount
            break
        elif "질병재해수술" in key and injury_surgery == 0:
            injury_surgery = amount
        # 흥국생명: (무)재해수술보장 → 상해수술비
        elif "재해수술보장" in key and injury_surgery == 0:
            injury_surgery = amount
        # 미래에셋: 재해수술무배당 또는 재해수술 (단독) → 상해수술비
        elif re.match(r'^재해수술([무배당최초]|$)', key) and injury_surgery == 0:
            injury_surgery = amount
        # 삼성화재: "상해입원수술비(당일입원제외)" → 상해수술비
        elif re.match(r'^상해입원수술비', key) and injury_surgery == 0:
            injury_surgery = amount
        # 현대해상: "상해수술" 단독 (종별 제외)
        elif re.match(r'^상해수술[0-9]*$', key) and injury_surgery == 0:
            if not any(ex in key for ex in injury_surgery_exclude):
                injury_surgery = amount
    if injury_surgery > 0:
        result["상해수술비"] = injury_surgery

    # 뇌혈관질환 수술비
    # 삼성화재: 2대주요기관질병 관혈수술비 + 비관혈수술비 = 1500 (뇌+심장 공통)
    # 라이나생명: 심뇌혈관질환수술특약 → 뇌혈관 + 허혈성심장 둘 다 해당
    # 메리츠: 뇌혈관질환수술비 + 131대질병수술비(뇌혈관질환) 합산 가능
    # DB손해: 체증형뇌혈관질환수술비 + 주요심,뇌,5대혈관및양성뇌종양수술비 합산
    brain_surgery = 0
    heart_surg_from_combined = 0
    brain_surg_exclude = []
    combined_major_surgery = 0  # DB손해: 주요심,뇌,5대혈관및양성뇌종양수술비 (뇌+심장 공통)
    for key, amount in simplified.items():
        # 라이나생명: 심뇌혈관질환수술 (뇌 + 심장 통합) — 반드시 뇌혈관질환수술보다 먼저 체크
        if "심뇌혈관질환수술" in key:
            brain_surgery += amount
            heart_surg_from_combined = amount  # 허혈성심장수술에도 사용
        elif "뇌혈관질환수술비" in key and "130대" not in key and "131대" not in key:
            if not any(ex in key for ex in brain_surg_exclude):
                brain_surgery += amount
        elif "130대질병수술비" in key and "뇌혈관질환" in key:
            brain_surgery += amount
        # 메리츠: 131대질병수술비(뇌혈관질환) → 뇌혈관질환수술비에 합산
        elif "131대질병수술비" in key and "뇌혈관질환" in key:
            brain_surgery += amount
        # 메리츠: 5대질환 수술비(심장,뇌혈관 포함) 비관혈 → 뇌혈관 합산
        elif "5대질환" in key and "뇌혈관" in key and "수술비" in key and "비관혈" in key:
            brain_surgery += amount
        # 미래에셋: 뇌혈관질환수술(최초1회한) → 뇌혈관질환수술비
        elif "뇌혈관질환수술" in key and "130대" not in key and "131대" not in key:
            if not any(ex in key for ex in brain_surg_exclude):
                brain_surgery += amount
        # DB손해: 주요심,뇌,5대혈관및양성뇌종양수술비 → 뇌혈관+허혈심장 공통
        elif "주요심" in key and "뇌" in key and "5대혈관" in key and "수술비" in key:
            combined_major_surgery = amount
    # DB손해: 주요심,뇌,5대혈관 수술비를 뇌혈관수술에 합산
    if combined_major_surgery > 0:
        brain_surgery += combined_major_surgery
    # 삼성화재: "2대주요기관질병 관혈수술비" + "비관혈수술비" (뇌+심장 공통)
    major_organ_surgery = 0
    for key, amount in simplified.items():
        if "2대주요기관질병" in key and ("관혈수술" in key or "비관혈수술" in key):
            major_organ_surgery += amount
    if major_organ_surgery > 0 and brain_surgery == 0:
        brain_surgery = major_organ_surgery
    # 현대해상: 120대질병수술(질병수술3(24대질병)) = 500만원 → 24대질병에 뇌혈관질환 포함
    if brain_surgery == 0:
        for key, amount in simplified.items():
            if "120대질병수술" in key and "24대질병" in key:
                brain_surgery = amount
                break
    if brain_surgery > 0:
        result["뇌혈수술비"] = brain_surgery
        result["뇌혈관질환수술비"] = brain_surgery

    # 허혈성심장질환수술비
    # 메리츠: 허혈성심장질환수술비 + 131대질병수술비(심장질환) 합산
    # DB손해: 체증형허혈심장질환수술비 → simplified '허혈심장질환수술비' + 주요심,뇌,5대혈관 공통
    heart_surgery = 0
    heart_surg_exclude = []
    for key, amount in simplified.items():
        if "허혈성심장질환수술비" in key and "130대" not in key and "131대" not in key:
            if not any(ex in key for ex in heart_surg_exclude):
                heart_surgery += amount
        # DB손해: 허혈심장질환수술비 (simplified: 체증형 제거 후)
        elif "허혈심장질환수술비" in key and "130대" not in key and "131대" not in key:
            if not any(ex in key for ex in heart_surg_exclude):
                heart_surgery += amount
        elif "130대질병수술비" in key and "심장질환" in key:
            heart_surgery += amount
        # 메리츠: 131대질병수술비(심장질환) → 허혈성심장질환수술비에 합산
        elif "131대질병수술비" in key and "심장질환" in key:
            heart_surgery += amount
        # 메리츠: 5대질환 수술비(심장,뇌혈관 포함) 비관혈 → 심장 합산
        elif "5대질환" in key and "심장" in key and "수술비" in key and "비관혈" in key:
            heart_surgery += amount
        # 미래에셋: 허혈성심장질환수술(최초1회한) → 허혈성심장질환수술비
        elif "허혈성심장질환수술" in key and "130대" not in key and "131대" not in key:
            if not any(ex in key for ex in heart_surg_exclude):
                heart_surgery += amount
    # DB손해: 주요심,뇌,5대혈관 수술비를 허혈심장수술에도 합산
    if combined_major_surgery > 0:
        heart_surgery += combined_major_surgery
    # 라이나생명: 심뇌혈관질환수술에서 허혈성심장수술 금액도 매핑
    if heart_surgery == 0 and heart_surg_from_combined > 0:
        heart_surgery = heart_surg_from_combined
    # 삼성화재: "2대주요기관질병 관혈수술비" + "비관혈수술비" (뇌+심장 공통)
    if major_organ_surgery > 0 and heart_surgery == 0:
        heart_surgery = major_organ_surgery
    # 현대해상: 120대질병수술(질병수술3(24대질병)) = 500만원 → 24대질병에 심장질환 포함
    if heart_surgery == 0:
        for key, amount in simplified.items():
            if "120대질병수술" in key and "24대질병" in key:
                heart_surgery = amount
                break
    if heart_surgery > 0:
        result["허혈성심장질환수술비"] = heart_surgery

    # 골절진단 (합산 - 5대골절 제외)
    # 중요: simplified dict는 동일 키를 덮어쓰므로 원본 coverages에서 직접 합산
    fracture_diag = 0
    fracture_diag_exclude = ["5대골절", "부목", "철심"]
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        if ("골절" in key and "진단" in key and "수술" not in key) or \
           ("재해골절치료" in key):
            if not any(ex in key for ex in fracture_diag_exclude):
                fracture_diag += cov["가입금액"]
    if fracture_diag > 0:
        result["골절진단"] = fracture_diag

    # 골절수술비 (5대골절수술 제외)
    # 중요: simplified dict는 동일 키를 덮어쓰므로 원본 coverages에서 직접 합산
    fracture_surgery_exclude = ["5대골절", "철심"]
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        if "골절수술" in key:
            if not any(ex in key for ex in fracture_surgery_exclude):
                result["골절수술비"] = cov["가입금액"]
                break

    # 뇌혈관질환 진단비
    # 삼성화재: 뇌혈관질환 진단비 + 뇌혈관질환(90일면책) 진단비 합산
    # 현대해상: 뇌혈관질환(Ⅰ)진단 = 넓은 범위(뇌혈관질환 전체), (Ⅱ)진단 = 좁은 범위(뇌졸중)
    #   → 뇌혈관질환 진단비는 (Ⅰ)만 사용 (넓은 범위)
    #   → (Ⅱ)는 뇌졸중진단비로 별도 매칭
    # 메리츠: 뇌혈관질환진단비 + 뇌혈관질환진단비Ⅱ 합산 (산정특례 제외)
    brain_diag_exclude = ["산정특례", "중증질환자"]
    brain_diag_items = []
    for key, amount in simplified.items():
        if "뇌혈관질환" in key and "진단" in key and "수술" not in key:
            if not any(ex in key for ex in brain_diag_exclude):
                brain_diag_items.append((key, amount))
    if brain_diag_items:
        # 현대해상 등: (1)/(2) 번호가 붙은 별개 담보 존재 시
        #   (1) = 넓은 범위 = 뇌혈관질환 진단비
        #   (2) = 좁은 범위(뇌졸중) = 별도 뇌졸중진단비로 매칭
        has_numbered = any(re.search(r'\(\d\)', k) for k, _ in brain_diag_items)
        if has_numbered:
            # (1)번 담보만 사용 (넓은 범위 = 뇌혈관질환 전체)
            for k, a in brain_diag_items:
                if '(1)' in k:
                    result["뇌혈관질환진단비"] = a
                    break
            # (1)이 없으면 최소값 fallback
            if "뇌혈관질환진단비" not in result:
                result["뇌혈관질환진단비"] = min(a for _, a in brain_diag_items)
        else:
            # 삼성화재 등: 90일면책 등 동일 보장 합산
            result["뇌혈관질환진단비"] = sum(a for _, a in brain_diag_items)

    # 허혈성심장질환 진단비
    # 삼성화재: 허혈성심장질환 + 90일면책 합산
    # 현대해상: 심혈관질환(특정Ⅱ)진단 = 허혈성심장질환 범위 (약관 근거)
    #   특정Ⅱ ≠ 특정2대 (특정2대 = 급성심근경색)
    # 메리츠: 허혈성심장질환진단비 + 허혈성심장질환진단비Ⅱ 합산 (산정특례 제외)
    heart_diag_exclude = ["산정특례", "중증질환자"]
    heart_diag = 0
    for key, amount in simplified.items():
        if ("허혈성심장질환" in key or "허혈심장질환" in key) and "진단" in key and "수술" not in key:
            if not any(ex in key for ex in heart_diag_exclude):
                heart_diag += amount
    # 현대해상 fallback: 심혈관질환(특정2)진단 → 허혈성심장질환 (특정2대 제외)
    if heart_diag == 0:
        for key, amount in simplified.items():
            if "심혈관질환(특정2)" in key and "2대" not in key and "진단" in key and "수술" not in key:
                heart_diag = amount
                break
    if heart_diag > 0:
        result["허혈성심장질환진단비"] = heart_diag

    # 항암방사선약물치료비
    # ─────────────────────────────────────────────────────────
    # 규칙:
    #   단일 특약 "항암방사선약물치료비" → 그 금액 그대로
    #   분리형 "항암방사선치료" + "항암약물치료" → min(방사선, 약물)
    #     (어차피 방사선 치료시/약물 치료시 각각 나오므로 합산이 아닌 최소값)
    #   통합형 "통합항암약물방사선치료" → 단일 금액 (종별 있어도 하나)
    #   분리형 + 통합형 공존 → min(분리 방사선, 분리 약물) + 통합형
    # ─────────────────────────────────────────────────────────
    exclude_kws = ["소액암", "호르몬", "세기조절", "양성자", "중입자", "표적", "카티", "주요치료", "종합병원", "기타피부암", "갑상선암", "계속받는"]
    
    # 0) 메리츠 26종 항암방사선및약물치료비 (종별 동일금액 → 단건 최대값)
    meritz_26_amount = 0
    for key, amount in simplified.items():
        if "26종" in key and "항암방사선" in key and "약물치료" in key:
            meritz_26_amount = max(meritz_26_amount, amount)
    
    # 1) 단일 통합 특약 확인 ("항암방사선약물치료비" 또는 "항암방사선약물치료" 이름 그대로)
    single_combined = 0
    for key, amount in simplified.items():
        if ("항암방사선약물치료" in key or "항암약물방사선치료" in key):
            # 분리형(항암방사선치료, 항암약물치료)은 제외 — 반드시 "방사선약물" 또는 "약물방사선" 포함
            if "방사선약물" in key or "약물방사선" in key:
                if not any(ex in key for ex in exclude_kws) and "통합" not in key:
                    single_combined = amount
                    break
    
    if single_combined > 0:
        # 단일 특약이 있으면 그 금액 그대로
        result["항암방사선약물치료비"] = single_combined
    else:
        # 2) 분리형 수집: 항암약물치료(방사선 미포함), 항암방사선치료(약물 미포함)
        drug_amount = 0
        radiation_amount = 0
        for key, amount in simplified.items():
            if any(ex in key for ex in exclude_kws):
                continue
            if "통합" in key:
                continue
            if "항암약물치료" in key and "방사선" not in key and "약물방사선" not in key:
                drug_amount = max(drug_amount, amount)  # 동명 특약 중 최대
            elif "항암방사선치료" in key and "약물" not in key and "방사선약물" not in key:
                radiation_amount = max(radiation_amount, amount)
        
        separated_value = 0
        if drug_amount > 0 and radiation_amount > 0:
            separated_value = min(drug_amount, radiation_amount)
        elif drug_amount > 0:
            separated_value = drug_amount
        elif radiation_amount > 0:
            separated_value = radiation_amount
        
        # 3) 통합형 수집: "통합항암약물방사선치료" (종별 있어도 금액 동일 → 하나로)
        integrated_amount = 0
        for key, amount in simplified.items():
            if "통합" in key and "항암" in key and ("약물" in key or "방사선" in key):
                if not any(ex in key for ex in exclude_kws):
                    integrated_amount = max(integrated_amount, amount)
        
        # 4) 합산: 분리형 min값 + 통합형 + 메리츠 26종
        antineoplastic_total = separated_value + integrated_amount
        # 메리츠 26종은 단일 항목이 없을 때만 사용
        if antineoplastic_total == 0 and meritz_26_amount > 0:
            antineoplastic_total = meritz_26_amount
        elif meritz_26_amount > 0:
            antineoplastic_total += meritz_26_amount
        if antineoplastic_total > 0:
            result["항암방사선약물치료비"] = antineoplastic_total

    # 암주요치료비
    # 여러 보험사에서 동일 키가 여러 값(덮어쓰기)으로 나올 수 있으므로 원본에서 max 사용
    # DB손해: 하이클래스암주요치료비Ⅱ(수술시) 1000만 / 500만 (동일 키, 다른 금액) → max
    # 흥국생명: 종합병원일반암주요치료
    # 삼성화재: 암전액본인부담, 암통합치료비
    # 메리츠: 암통합치료비
    cancer_treatment_exclude = ["소액암", "유사암", "전이암", "생활비", "진단및치료비", "통합치료비2", "통합치료비3"]
    cancer_treatment = 0
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        amount = cov["가입금액"]
        matched_kw = False
        for kw in ["암주요치료비", "하이클래스암주요치료비", "일반암주요치료",
                    "암전액본인부담", "암통합치료비"]:
            if kw in key:
                matched_kw = True
                break
        if not matched_kw:
            continue
        if any(ex in key for ex in cancer_treatment_exclude):
            continue
        cancer_treatment = max(cancer_treatment, amount)
    if cancer_treatment > 0:
        result["암주요치료비"] = cancer_treatment

    # 암진단(일반암) 합산
    # 라이나생명: 암진단특약 3000만 + 통합암진단특약 7000만 = 10000만원
    # 다른 보험사: 암진단(일반암) 단일 특약 = 그 금액 그대로
    # 삼성화재: 암진단비(유사암제외) = 일반암
    # 메리츠 또또암: 암종별(30종)통합암진단비 = 종별로 4000만원 각각 → 중복X, 단건 4000만원
    #   + 암진단및치료비[암진단비(유사암제외)] 1000만원은 별도 합산 가능
    cancer_diag = 0
    cancer_exclude_names = ["소액암진단", "전이암진단", "갑상선암", "기타피부암",
                            "제자리암", "경계성종양", "남녀특정암", "특정암진단"]
    # 메리츠: 암종별(30종)통합암진단비는 각 종별 동일금액이므로 최대값 1건만 사용
    cancer_种별_max = 0
    cancer_종별_found = False
    for key, amount in simplified.items():
        # 메리츠: 암종별(30종)통합암진단비 → 별도 처리
        if "암종별" in key and "통합암진단비" in key:
            cancer_종별_found = True
            cancer_种별_max = max(cancer_种별_max, amount)
            continue
        # "암진단" 또는 "암(...제외)진단" 패턴 매칭
        if "암진단" in key or ("암" in key and "진단" in key and "수술" not in key):
            # "유사암" 포함이더라도 "유사암제외" 패턴이면 일반암으로 봄
            is_excluded = False
            for ex_name in cancer_exclude_names:
                if ex_name in key:
                    is_excluded = True
                    break
            # "유사암" 단독 포함 (제외 패턴이 아닌 경우) = 유사암 특약 → 제외
            if not is_excluded and "유사암" in key and "제외" not in key:
                is_excluded = True
            # 메리츠: 암진단및치료비는 패키지 상품이므로 순수 암진단비에서 제외
            if not is_excluded and "암진단및치료비" in key:
                is_excluded = True
            if not is_excluded:
                cancer_diag += amount
    # 메리츠 종별 암진단비 추가 (중복X, 단건으로)
    if cancer_종별_found:
        cancer_diag += cancer_种별_max
    if cancer_diag > 0:
        result["암진단(일반암)"] = cancer_diag

    # 일반상해사망 / 재해사망
    death_amount = 0
    for key, amount in simplified.items():
        if any(kw in key for kw in ["일반상해사망", "재해사망", "주보험재해사망", "주계약(재해사망)"]):
            death_amount = amount
            break
    if death_amount > 0:
        result["일반상해사망"] = death_amount
    # 흥국생명: (무)재해사망 → 일반상해사망에 주계약 가입금액 합산
    for cov in pdf_coverages:
        name = simplify_pdf_name(cov["특약명"])
        if "재해사망" in name and "일반상해사망" not in result:
            result["일반상해사망"] = cov["가입금액"]
            break
    # 삼성화재: "상해 사망" 5000만원 = 상해사망/재해사망에 해당
    if "일반상해사망" not in result:
        for key, amount in simplified.items():
            if key == "상해사망":
                result["일반상해사망"] = amount
                break
    # 현대해상: 기본계약(상해사망) → 상해사망
    if "일반상해사망" not in result:
        for key, amount in simplified.items():
            if "기본계약" in key and "상해사망" in key:
                result["일반상해사망"] = amount
                break

    # 일반사망 (주계약 사망보험금)
    # 흥국생명: 주계약 가입금액의 50% (재해 이외 원인으로 사망시)
    # 미래에셋: 주계약(재해사망)만 있으면 일반사망 = 0
    has_main_contract = False
    for cov in pdf_coverages:
        name = cov["특약명"]
        # 흥국생명 통합보험 주계약 (일반사망 50% 포함)
        if "통합보험" in name:
            result["일반사망"] = cov["가입금액"] // 2  # 50%
            has_main_contract = True
            break
    # 별도 일반사망보장/종신사망 특약 우선
    for key, amount in simplified.items():
        if any(kw in key for kw in ["일반사망보장", "종신사망", "사망보장"]):
            result["일반사망"] = amount
            break
    # 라이나생명: 주계약(건강보험) = 사망보험금
    if "일반사망" not in result:
        for key, amount in simplified.items():
            if key == "건강보험":
                result["일반사망"] = amount
                break

    # 상해사망/재해사망 합산 (주계약 재해사망 + 재해사망 특약)
    total_death = 0
    for cov in pdf_coverages:
        name = simplify_pdf_name(cov["특약명"])
        if "재해사망" in name or "일반상해사망" in name:
            total_death += cov["가입금액"]
    # 흥국생명: 주계약(통합보험)에 재해사망 100% 포함시 합산
    if has_main_contract:
        for cov in pdf_coverages:
            name = cov["특약명"]
            if "통합보험" in name:
                total_death += cov["가입금액"]  # 재해 원인 100%
                break
    # 삼성화재: "상해 사망" = 상해사망/재해사망
    if total_death == 0:
        for key, amount in simplified.items():
            if key == "상해사망":
                total_death = amount
                break
    # 현대해상: 기본계약(상해사망) = 상해사망/재해사망
    if total_death == 0:
        for key, amount in simplified.items():
            if "기본계약" in key and "상해사망" in key:
                total_death = amount
                break
    if total_death > 0:
        result["상해사망/재해사망"] = total_death

    # 가족일상배상책임
    for key, amount in simplified.items():
        if "가족일상생활중배상책임" in key or "가족일상배상책임" in key:
            result["가족일상배상책임"] = amount
            break
    # 현대해상: 무배당일상생활중배상책임(가족) → 가족일상배상책임
    if "가족일상배상책임" not in result:
        for key, amount in simplified.items():
            if "일상생활중배상책임" in key and "가족" in key:
                result["가족일상배상책임"] = amount
                break

    # 상해후유장해3% (합산)
    # 신한라이프: 재해장해특약 7000만원 + 주계약(신한통합건강보험) 500만원 = 7500만원
    # 다른 보험사: 단일 특약이면 그 금액만 사용
    # 메리츠: 일반상해후유장해(3-100%) = 상해후유장해3%에 해당
    # DB손해: 상해후유장해(3-100%) — 동일 키가 두 번(1억, 1만) 있으므로 max 사용
    injury_disability = 0
    for cov in pdf_coverages:
        key = simplify_pdf_name(cov["특약명"])
        amount = cov["가입금액"]
        if any(kw in key for kw in [
            "상해3%이상후유장해", "상해후유장해3%",
            "재해후유장해보장", "재해후유장해",
            "일반상해후유장해(3-100%)",  # 메리츠
            "상해후유장해(3-100%)",  # DB손해
        ]):
            injury_disability = max(injury_disability, amount)
        elif "기본계약(상해후유장해" in key:
            injury_disability = max(injury_disability, amount)
    # 신한라이프/미래에셋: 재해장해 (simplified 이름) - 부분 일치 (미래에셋: "재해장해최초계")
    if injury_disability == 0:
        for key, amount in simplified.items():
            if "재해장해" in key and "수술" not in key and "사망" not in key:
                injury_disability = amount
                break
    # 신한라이프: 주계약도 장해급여금 포함 (재해장해특약 + 주계약)
    if injury_disability > 0:
        for key, amount in simplified.items():
            if "신한통합건강보험" in key or "통합건강보험" in key:
                injury_disability += amount
                break
    if injury_disability > 0:
        result["상해후유장해3%"] = injury_disability

    # 전이암진단비
    # 라이나생명: 통합전이암진단특약 → 직접 매칭
    # 메리츠 또또암: 암종별(30종)통합암진단비(전이포함) 4000만원 = 전이암에도 동일 금액 적용
    transfer_cancer = 0
    for key, amount in simplified.items():
        if "전이암진단" in key or "전이암진단비" in key:
            transfer_cancer = max(transfer_cancer, amount)
        elif "통합전이암" in key:
            transfer_cancer = max(transfer_cancer, amount)
    # 메리츠: 암종별 통합암진단비(전이포함)에서 최대값 사용 (전이포함이므로)
    if transfer_cancer == 0 and cancer_종별_found:
        transfer_cancer = cancer_种별_max
    if transfer_cancer > 0:
        result["전이암진단비"] = transfer_cancer

    # 표적항암약물치료비
    # 메리츠: 표적항암약물허가치료비(연간 약물종류 개수별)(비급여)[N종이상] 형태
    #   → 3종이상 > 2종이상 > 1종이상 우선순위로 최상위 등급 선택
    #   → 급여/비급여가 simplify 후 같은 키로 합쳐지므로 원본에서 직접 처리
    # 다른 보험사: 표적항암약물치료비/허가치료 단일 특약 → max 사용
    target_chemo = 0
    
    # 메리츠 종류별 체크: 원본 coverages에서 비급여 3종이상 우선 탐색
    grade_amounts = {}  # {등급: 비급여금액}
    grade_amounts_all = {}  # {등급: 모든금액 중 max}
    has_grade_items = False
    for cov in pdf_coverages:
        name = cov["특약명"]
        amount = cov["가입금액"]
        if "표적항암약물" not in name or "허가치료" not in name:
            continue
        if "약물종류" in name or "개수별" in name:
            has_grade_items = True
            for grade_label in ["3종이상", "2종이상", "1종이상"]:
                if grade_label in name:
                    # 비급여 항목 우선
                    if "비급여" in name or "전액본인부담" in name:
                        grade_amounts[grade_label] = max(grade_amounts.get(grade_label, 0), amount)
                    grade_amounts_all[grade_label] = max(grade_amounts_all.get(grade_label, 0), amount)
                    break
    
    if has_grade_items:
        # 3종이상 > 2종이상 > 1종이상 우선순위 (비급여 우선, 없으면 전체)
        for grade_label in ["3종이상", "2종이상", "1종이상"]:
            if grade_label in grade_amounts and grade_amounts[grade_label] > 0:
                target_chemo = grade_amounts[grade_label]
                break
            elif grade_label in grade_amounts_all and grade_amounts_all[grade_label] > 0:
                target_chemo = grade_amounts_all[grade_label]
                break
    
    if target_chemo == 0:
        # 종류별이 아닌 일반 표적항암 항목 (다른 보험사 또는 표적항암약물허가치료비Ⅱ 등)
        for key, amount in simplified.items():
            if "표적항암약물" in key and "허가치료" in key:
                if "약물종류" not in key and "개수별" not in key:
                    target_chemo = max(target_chemo, amount)
            elif "표적항암약물치료비" in key:
                if "약물종류" not in key and "개수별" not in key:
                    target_chemo = max(target_chemo, amount)
    
    if target_chemo > 0:
        result["표적항암약물치료비"] = target_chemo

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
    "암진단(일반암)": {"type": "aggregate", "key": "암진단(일반암)"},
    "소액/유사암": {"type": "direct_exclude", "keywords": [
        "유사암진단비", "소액암진단비", "유사암진단",
        "소액암진단",  # 신한라이프: 소액암진단특약 → simplified "소액암진단"
        "소액암New보장", "소액암보장"  # 흥국생명: (무)소액암New보장(갱신형)
    ], "exclude": ["제외"]},  # "암(유사암제외)진단" 제외
    "전이암진단비": {"type": "aggregate", "key": "전이암진단비"},
    "암주요치료비": {"type": "aggregate", "key": "암주요치료비"},
    "항암방사선약물치료비": {"type": "aggregate", "key": "항암방사선약물치료비"},
    "표적항암약물치료비": {"type": "aggregate", "key": "표적항암약물치료비"},
    "양성자방사선치료비": {"type": "direct", "keywords": [
        "양성자방사선치료비", "항암양성자방사선치료",
        "항암방사선(양성자)",  # 현대해상: 항암방사선(양성자)치료
    ]},
    "세기조절방사선치료비": {"type": "direct", "keywords": [
        "세기조절방사선치료비", "항암세기조절방사선치료",
        "항암방사선(세기조절)",  # 현대해상: 항암방사선(세기조절)치료
        "표적항암방사선치료비(항암세기조절방사선)",  # DB손해: 표적항암방사선치료비(항암세기조절방사선)
    ]},
    "면역항암약물치료비": {"type": "direct", "keywords": [
        "면역항암약물치료비", "특정면역항암약물허가치료",
        "면역항암약물허가치료", "특정면역항암",
        "카티항암약물허가치료",  # 흥국생명: 카티(CAR-T) = 면역항암약물치료
        "카티(CAR-T)항암약물허가치료",  # 현대해상: 카티(CAR-T)항암약물허가치료
    ]},
    "카티항암약물치료비": {"type": "direct", "keywords": [
        "카티항암약물치료비", "카티항암약물허가치료", "CAR-T"
    ]},
    "다빈치로봇수술비": {"type": "direct_exclude", "keywords": [
        "다빈치로봇수술비", "다빈치로봇수술",
        "암다빈치로봇수술",
        "로봇암(특정암제외)수술",  # 흥국생명: (무)로봇암(특정암제외)수술(다빈치,레보아이)
        "로봇암수술", "로봇수술",
        "로봇암수술(다빈치",  # 현대해상: 로봇암수술(다빈치및레보아이)
    ], "exclude": ["종합병원", "주요치료", "소액암"]},

    # 뇌
    "뇌혈관질환진단비": {"type": "aggregate", "key": "뇌혈관질환진단비"},
    "뇌혈관질환": {"type": "aggregate", "key": "뇌혈관질환진단비"},
    "뇌혈수술비": {"type": "aggregate", "key": "뇌혈수술비"},
    "뇌혈관질환수술비": {"type": "aggregate", "key": "뇌혈관질환수술비"},
    "뇌졸증진단비": {"type": "direct", "keywords": [
        "뇌졸중진단비", "뇌졸증진단비",
        "뇌혈관질환(2)진단",  # 현대해상: 뇌혈관질환(Ⅱ)진단 = 뇌졸중 범위 (거미막하출혈, 뇌내출혈, 뇌경색증)
    ]},
    "뇌졸증": {"type": "direct", "keywords": [
        "뇌졸중진단비", "뇌졸증진단비",
        "뇌혈관질환(2)진단",  # 현대해상: 뇌혈관질환(Ⅱ)진단 = 뇌졸중 범위
    ]},
    "뇌출혈진단비": {"type": "direct_exclude", "keywords": ["뇌출혈진단비", "뇌출혈진단"], "exclude": ["외상성"]},
    "뇌출혈": {"type": "direct_exclude", "keywords": ["뇌출혈진단비", "뇌출혈진단"], "exclude": ["외상성"]},

    # 심장
    "심혈관질환진단비": {"type": "direct", "keywords": [
        "심혈관질환진단비", "기타부정맥진단",  # 흥국생명: 기타부정맥진단
        "부정맥진단",  # 미래에셋: 부정맥진단특약
        "심혈관질환(특정",  # 현대해상: 심혈관질환(특정1,I49제외)진단
        "심혈관질환(주요심장염증)",  # 현대해상: 심혈관질환(주요심장염증)진단
    ]},
    "심혈관질환": {"type": "direct", "keywords": [
        "심혈관질환진단비", "기타부정맥진단",  # 흥국생명: 기타부정맥진단
        "부정맥진단",  # 미래에셋: 부정맥진단특약
        "심혈관질환(특정",  # 현대해상
        "심혈관질환(주요심장염증)",  # 현대해상
    ]},
    "허혈성심장질환진단비": {"type": "aggregate", "key": "허혈성심장질환진단비"},
    "허혈성심장질환": {"type": "aggregate", "key": "허혈성심장질환진단비"},
    "허혈성심장질환수술비": {"type": "aggregate", "key": "허혈성심장질환수술비"},
    "급성심근경색진단비": {"type": "direct", "keywords": [
        "급성심근경색진단비", "급성심근경색증진단비",
        "급성심근경색증진단",  # 라이나생명: 급성심근경색증진단특약
        "급성심근경색진단",  # 일반적인 패턴
        "심혈관질환(특정2대)",  # 현대해상: 심혈관질환(특정2대)진단 = 급성심근경색
    ]},
    "급성심근경색": {"type": "direct", "keywords": [
        "급성심근경색진단비", "급성심근경색증진단비",
        "급성심근경색증진단",  # 라이나생명
        "급성심근경색진단",
        "심혈관질환(특정2대)",  # 현대해상
    ]},

    # 입원/응급
    "질병입원": {"type": "direct_exclude", "keywords": ["질병입원일당", "질병입원비"], "exclude": ["다빈도", "특정", "수술", "중환자실", "상급병실", "간병인", "요양성", "2-3인실", "1인실"]},
    "상해입원": {"type": "direct_exclude", "keywords": ["상해입원일당", "상해입원비"], "exclude": ["수술", "중환자실", "상급병실", "간병인", "2-3인실", "1인실"]},
    "응급실내원(응급)": {"type": "direct", "keywords": [
        "응급실내원(응급)", "응급실내원",  # 미래에셋: 응급실내원특약 (응급/비응급 구분 없음)
        "응급실내원비(응급)",  # 메리츠: 응급실내원비(응급)
    ]},
    "응급실내원(비응급)": {"type": "direct", "keywords": ["응급실내원(비응급)"]},

    # 사망
    "일반사망": {"type": "special_general_death"},  # 특수 처리: 주계약 사망보험금
    "질병사망": {"type": "direct", "keywords": ["질병사망"]},
    "상해사망/재해사망": {"type": "aggregate", "key": "상해사망/재해사망"},

    # 후유장해
    # 신한라이프: 재해장해특약 7000만원 + 주계약(신한통합건강보험) 500만원 = 7500만원
    "상해후유장해3%": {"type": "aggregate", "key": "상해후유장해3%"},
    "질병후유장해3%": {"type": "direct_exclude", "keywords": [
        "질병3%이상후유장해", "질병후유장해3%",
        "질병후유장해보장", "질병후유장해(3-100%)",  # 메리츠: 질병후유장해(3-100%)
        "질병후유장해",  # 흥국생명: 질병후유장해보장
        "질병장해"  # 미래에셋/신한라이프: 질병장해보장특약
    ], "exclude": ["80%이상"]},
    "상해후유장해": {"type": "direct", "keywords": [
        "상해후유장해", "일반상해80%이상후유장해",
        "재해후유장해보장", "재해후유장해"
    ]},
    "질병후유장해": {"type": "direct", "keywords": ["질병후유장해", "질병80%이상후유장해"]},

    # 골절
    "골절진단": {"type": "aggregate", "key": "골절진단"},
    "골절수술": {"type": "aggregate", "key": "골절수술비"},
    "깁스": {"type": "direct", "keywords": ["깁스치료", "깁스"]},

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

            elif rule_type == "special_general_death":
                # 일반사망: aggregated에서 찾기
                if "일반사망" in aggregated:
                    matched_amount = aggregated["일반사망"]
                    matched_pdf_name = "[합산] 일반사망"
                else:
                    # 직접 매칭 시도
                    for kw in ["일반사망보장", "종신사망", "사망보장"]:
                        kw_clean = re.sub(r'\s+', '', kw)
                        for pdf_key, pdf_cov in pdf_simplified.items():
                            if kw_clean in pdf_key:
                                matched_amount = pdf_cov["가입금액"]
                                matched_pdf_name = pdf_cov["특약명"]
                                break
                        if matched_amount:
                            break

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
