"""
심평원(HIRA) PDF 파서 — 기본진료내역, 처방조제정보, 세부진료내역
pdfplumber 기반으로 PDF 테이블을 추출하여 프론트엔드 Excel 파서와 동일한 JSON 형식으로 변환
"""
import pdfplumber
import re


def _clean(s):
    """줄바꿈 제거 + 양끝 공백 제거 + 중복 공백 정리"""
    if not s:
        return ''
    return re.sub(r'\s+', ' ', str(s).replace('\n', ' ').replace('\r', '')).strip()


def _parse_int(val):
    """정수 파싱 (실패 시 0)"""
    if val is None:
        return 0
    s = _clean(str(val))
    if not s:
        return 0
    m = re.search(r'(\d+)', s.replace(',', ''))
    return int(m.group(1)) if m else 0


def _parse_date(val):
    """날짜 문자열 정규화 → YYYY-MM-DD"""
    if not val:
        return None
    s = _clean(str(val))
    if not s:
        return None
    # YYYY-MM-DD
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    # YYYY.MM.DD or YYYY/MM/DD
    m = re.match(r'^(\d{4})[./](\d{1,2})[./](\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    # YYYYMMDD
    m = re.match(r'^(\d{4})(\d{2})(\d{2})$', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None


def _is_data_row(row):
    """데이터 행인지 판별 (순번이 숫자인 행)"""
    if not row or not row[0]:
        return False
    seq = str(row[0]).strip()
    if not seq:
        return False
    # 헤더/주석 행 제거
    if seq == '순번' or seq.startswith('·') or seq.startswith('*'):
        return False
    if '본 자료는' in seq or '자료출처' in seq:
        return False
    try:
        int(seq)
        return True
    except ValueError:
        return False


def _detect_header(rows, required_keywords):
    """헤더 행 자동 탐지 — required_keywords 중 하나라도 매칭되는 행"""
    for i, row in enumerate(rows):
        if not row:
            continue
        text = ' '.join(_clean(str(c or '')) for c in row).lower()
        if all(kw in text for kw in required_keywords):
            return i
    return 0


def _extract_all_tables(pdf_path):
    """PDF에서 모든 페이지의 테이블 행 추출 (헤더 행 중복 제거)"""
    all_rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                all_rows.append(row)
    return all_rows


def parse_basic_care_pdf(pdf_path):
    """기본진료내역 PDF → treatRecords JSON
    
    PDF 컬럼: 순번 | 진료시작일 | 병·의원&약국 | 진단과 | 입원/외래 | 주상병코드 | 주상병명 | 내원일수 | 총진료비 | 건강보험혜택 | 내가낸의료비
    출력 형식: { 진료일자, 진료기관명, 진단코드, 진단명, 입원외래, 내원일수, 총진료비, 건강보험혜택, 본인부담금 }
    """
    all_rows = _extract_all_tables(pdf_path)
    if not all_rows:
        return []
    
    # 헤더 자동 탐지
    header_idx = _detect_header(all_rows, ['진료시작일', '병'])
    if header_idx == 0:
        header_idx = _detect_header(all_rows, ['순번', '진료'])
    
    header = all_rows[header_idx] if header_idx < len(all_rows) else []
    
    # 동적 컬럼 매핑 (PDF 양식 변경에 대응)
    col_map = {}
    for i, h in enumerate(header):
        h_clean = _clean(str(h or '')).lower().replace(' ', '')
        if '진료시작일' in h_clean or '진료일' in h_clean:
            col_map['date'] = i
        elif '병·의원' in h_clean or '병의원' in h_clean or '약국' in h_clean:
            col_map['org'] = i
        elif '진단과' in h_clean:
            col_map['dept'] = i
        elif '입원' in h_clean and '외래' in h_clean:
            col_map['type'] = i
        elif '주상병' in h_clean and ('코드' in h_clean or '코 드' in h_clean):
            col_map['code'] = i
        elif '주상병' in h_clean and ('명' in h_clean):
            col_map['name'] = i
        elif '내원' in h_clean and '일수' in h_clean:
            col_map['days'] = i
        elif '총진료비' in h_clean or '요양급여비용총액' in h_clean or ('요양급여' in h_clean and '총액' in h_clean):
            col_map['total_cost'] = i
        elif '건강보험' in h_clean or '혜택받은' in h_clean or '공단부담' in h_clean or '보험자부담' in h_clean:
            col_map['insurer_cost'] = i
        elif '내가낸' in h_clean or '내가 낸' in h_clean or '본인부담' in h_clean or '환자부담' in h_clean:
            col_map['self_cost'] = i
    
    # 폴백: 고정 위치 (표준 심평원 양식)
    if 'date' not in col_map:
        col_map['date'] = 1
    if 'org' not in col_map:
        col_map['org'] = 2
    if 'type' not in col_map:
        col_map['type'] = 4
    if 'code' not in col_map:
        col_map['code'] = 5
    if 'name' not in col_map:
        col_map['name'] = 6
    if 'days' not in col_map:
        col_map['days'] = 7
    if 'total_cost' not in col_map:
        col_map['total_cost'] = 8
    if 'insurer_cost' not in col_map:
        col_map['insurer_cost'] = 9
    if 'self_cost' not in col_map:
        col_map['self_cost'] = 10
    
    records = []
    for row in all_rows[header_idx + 1:]:
        if not _is_data_row(row):
            continue
        
        date_val = _parse_date(row[col_map['date']] if col_map['date'] < len(row) else None)
        org = _clean(row[col_map['org']] if col_map['org'] < len(row) else '')
        
        if not date_val or not org:
            continue
        
        type_val = _clean(row[col_map['type']] if col_map['type'] < len(row) else '')
        inout = '입원' if '입원' in type_val else '외래'
        
        records.append({
            '진료일자': date_val,
            '진료기관명': org,
            '진단코드': _clean(row[col_map['code']] if col_map['code'] < len(row) else ''),
            '진단명': _clean(row[col_map['name']] if col_map['name'] < len(row) else ''),
            '입원외래': inout,
            '내원일수': _parse_int(row[col_map['days']] if col_map['days'] < len(row) else 0),
            '총진료비': _parse_int(row[col_map['total_cost']] if col_map['total_cost'] < len(row) else 0),
            '건강보험혜택': _parse_int(row[col_map['insurer_cost']] if col_map['insurer_cost'] < len(row) else 0),
            '본인부담금': _parse_int(row[col_map['self_cost']] if col_map['self_cost'] < len(row) else 0),
        })
    
    # 날짜 역순 정렬
    records.sort(key=lambda r: r['진료일자'], reverse=True)
    return records


def parse_prescription_pdf(pdf_path):
    """처방조제정보 PDF → rxRecords JSON
    
    PDF 컬럼: 순번 | 진료시작일 | 병·의원&약국 | 처방/조제 | 약품명 | 성분명 | 1회투약량 | 1일투여횟수 | 총투약일수
    출력 형식: { 진료일자, 진료기관명, 총투약일수, 약품명 }
    """
    all_rows = _extract_all_tables(pdf_path)
    if not all_rows:
        return []
    
    header_idx = _detect_header(all_rows, ['진료시작일', '약품명'])
    if header_idx == 0:
        header_idx = _detect_header(all_rows, ['순번', '투약'])
    
    header = all_rows[header_idx] if header_idx < len(all_rows) else []
    
    col_map = {}
    for i, h in enumerate(header):
        h_clean = _clean(str(h or '')).lower().replace(' ', '')
        if '진료시작일' in h_clean or '진료일' in h_clean:
            col_map['date'] = i
        elif '병·의원' in h_clean or '병의원' in h_clean or '약국' in h_clean:
            col_map['org'] = i
        elif '약품명' in h_clean:
            col_map['drug'] = i
        elif '총' in h_clean and '투약일수' in h_clean:
            col_map['total_days'] = i
    
    # 폴백
    if 'date' not in col_map:
        col_map['date'] = 1
    if 'org' not in col_map:
        col_map['org'] = 2
    if 'drug' not in col_map:
        col_map['drug'] = 4
    if 'total_days' not in col_map:
        col_map['total_days'] = 8
    
    records = []
    for row in all_rows[header_idx + 1:]:
        if not _is_data_row(row):
            continue
        
        date_val = _parse_date(row[col_map['date']] if col_map['date'] < len(row) else None)
        org = _clean(row[col_map['org']] if col_map['org'] < len(row) else '')
        
        if not date_val or not org:
            continue
        
        records.append({
            '진료일자': date_val,
            '진료기관명': org,
            '총투약일수': _parse_int(row[col_map['total_days']] if col_map['total_days'] < len(row) else 0),
            '약품명': _clean(row[col_map['drug']] if col_map['drug'] < len(row) else ''),
        })
    
    return records


def parse_detail_care_pdf(pdf_path):
    """세부진료내역 PDF → detailRecords JSON
    
    PDF 컬럼: 순번 | 진료시작일 | 병·의원&약국 | 진료내역 | 코드명 | 1회투약량 | 1일투여횟수 | 총투약일수
    출력 형식: { 진료시작일, 병의원, 진료내역, 코드명, 투약량, 투여횟수, 총투약일수 }
    """
    all_rows = _extract_all_tables(pdf_path)
    if not all_rows:
        return []
    
    header_idx = _detect_header(all_rows, ['진료시작일', '진료내역'])
    if header_idx == 0:
        header_idx = _detect_header(all_rows, ['순번', '코드명'])
    if header_idx == 0:
        header_idx = _detect_header(all_rows, ['순번', '진료'])
    
    header = all_rows[header_idx] if header_idx < len(all_rows) else []
    
    col_map = {}
    for i, h in enumerate(header):
        h_clean = _clean(str(h or '')).lower().replace(' ', '')
        if '진료시작일' in h_clean or '진료일' in h_clean:
            col_map['date'] = i
        elif '병·의원' in h_clean or '병의원' in h_clean or '약국' in h_clean:
            col_map['org'] = i
        elif '진료내역' in h_clean:
            col_map['treatment'] = i
        elif '코드명' in h_clean:
            col_map['code'] = i
        elif '1회' in h_clean and '투약량' in h_clean:
            col_map['dose'] = i
        elif '1일' in h_clean and '투여횟수' in h_clean:
            col_map['freq'] = i
        elif '총' in h_clean and '투약일수' in h_clean:
            col_map['total_days'] = i
    
    # 폴백
    if 'date' not in col_map:
        col_map['date'] = 1
    if 'org' not in col_map:
        col_map['org'] = 2
    if 'treatment' not in col_map:
        col_map['treatment'] = 3
    if 'code' not in col_map:
        col_map['code'] = 4
    if 'dose' not in col_map:
        col_map['dose'] = 5
    if 'freq' not in col_map:
        col_map['freq'] = 6
    if 'total_days' not in col_map:
        col_map['total_days'] = 7
    
    records = []
    for row in all_rows[header_idx + 1:]:
        if not _is_data_row(row):
            continue
        
        date_val = _parse_date(row[col_map['date']] if col_map['date'] < len(row) else None)
        if not date_val:
            continue
        
        treatment = _clean(row[col_map['treatment']] if col_map['treatment'] < len(row) else '')
        code = _clean(row[col_map['code']] if col_map['code'] < len(row) else '')
        
        if not treatment and not code:
            continue
        
        records.append({
            '진료시작일': date_val,
            '병의원': _clean(row[col_map['org']] if col_map['org'] < len(row) else ''),
            '진료내역': treatment,
            '코드명': code,
            '투약량': _parse_int(row[col_map['dose']] if col_map['dose'] < len(row) else 0),
            '투여횟수': _parse_int(row[col_map['freq']] if col_map['freq'] < len(row) else 0),
            '총투약일수': _parse_int(row[col_map['total_days']] if col_map['total_days'] < len(row) else 0),
        })
    
    # 날짜 역순 정렬
    records.sort(key=lambda r: r['진료시작일'], reverse=True)
    return records


def detect_hira_pdf_type(pdf_path):
    """PDF 유형 자동 감지 — 기본진료 / 처방조제 / 세부진료"""
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return 'unknown'
        text = ''
        for page in pdf.pages[:2]:
            page_text = page.extract_text()
            if page_text:
                text += page_text
    
    text_lower = text.lower().replace(' ', '')
    
    # 키워드 기반 감지
    if '기본진료내역' in text_lower or '기본진료정보' in text_lower:
        return 'basic'
    if '처방조제' in text_lower:
        return 'prescription'
    if '세부진료내역' in text_lower or '세부진료정보' in text_lower:
        return 'detail'
    
    # 헤더 기반 감지 (키워드 없을 때)
    table_rows = _extract_all_tables(pdf_path)
    if table_rows:
        header_text = ' '.join(_clean(str(c or '')) for c in (table_rows[0] or []))
        if '주상병' in header_text and ('입원' in header_text or '외래' in header_text):
            return 'basic'
        if '약품명' in header_text or '성분명' in header_text:
            return 'prescription'
        if '진료내역' in header_text and '코드명' in header_text:
            return 'detail'
    
    return 'unknown'
