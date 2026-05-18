import os
import gc
import tempfile
import shutil
import traceback
import psutil
from urllib.parse import quote
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from typing import List, Optional

from pdf_parser import parse_pdf_all_in_one
from excel_handler import (
    read_excel_coverages, write_matched_amounts,
    write_insurer_info, write_premium, find_structure
)
from matcher import match_coverages
from hira_pdf_parser import (
    parse_basic_care_pdf, parse_prescription_pdf,
    parse_detail_care_pdf, detect_hira_pdf_type
)

app = FastAPI(title="보험 보장분석 자동매칭 API", version="1.1.0")

# CORS 설정 — 프론트엔드 도메인 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://alillae.com",
        "https://insurance-checker.pages.dev",
        "http://localhost:3000",
        "http://localhost:5173",
        "*",  # 개발용
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

INSURER_NAMES = {
    "meritz": "메리츠화재", "samsung": "삼성화재",
    "samsung_life": "삼성생명",
    "kb": "KB손해보험", "db": "DB손해",
    "mirae": "미래에셋생명", "abl": "ABL생명",
    "heungkuk": "흥국생명", "hanwha": "한화생명",
    "hyundai": "현대해상", "lotte": "롯데손해보험",
    "nh": "NH농협생명", "dongyang": "동양생명",
    "kyobo": "교보생명", "shinhan": "신한라이프",
}

# ── 메모리 관리 유틸리티 ──

def _get_memory_mb():
    """현재 프로세스 메모리 사용량 (MB)"""
    try:
        process = psutil.Process(os.getpid())
        return process.memory_info().rss / 1024 / 1024
    except Exception:
        return -1


def _force_gc():
    """강제 가비지 컬렉션 — PDF/Excel 처리 후 메모리 해제"""
    gc.collect()


def _cleanup_temp_files(paths):
    """임시파일 안전 삭제"""
    for p in paths:
        try:
            if p and os.path.exists(p):
                os.unlink(p)
        except Exception:
            pass


def _cleanup_stale_temp_files():
    """오래된 임시파일 정리 (1시간 이상 된 .pdf/.xlsx 파일)"""
    import time
    tmp_dir = tempfile.gettempdir()
    cutoff = time.time() - 3600  # 1시간
    try:
        for fname in os.listdir(tmp_dir):
            if fname.endswith(('.pdf', '.xlsx')):
                fpath = os.path.join(tmp_dir, fname)
                try:
                    if os.path.getmtime(fpath) < cutoff:
                        os.unlink(fpath)
                except Exception:
                    pass
    except Exception:
        pass


@app.get("/")
async def root():
    mem_mb = _get_memory_mb()
    return {
        "status": "ok",
        "service": "보험 보장분석 자동매칭 API",
        "version": "1.1.0",
        "memory_mb": round(mem_mb, 1),
    }


@app.get("/health")
async def health():
    mem_mb = _get_memory_mb()
    return {
        "status": "ok",
        "memory_mb": round(mem_mb, 1),
    }


@app.post("/api/parse-pdf")
async def parse_pdf(pdf_file: UploadFile = File(...)):
    """단일 PDF 파싱 — 보험사, 상품명, 보험료, 특약 목록 반환 (최적화: 1회 오픈)"""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            content = await pdf_file.read()
            tmp.write(content)
            tmp_path = tmp.name
        # content 참조 즉시 해제
        del content

        # 최적화: parse_pdf_all_in_one으로 PDF를 1회만 열어서 전체 정보 추출
        pdf_info = parse_pdf_all_in_one(tmp_path)

        response = {
            "success": True,
            "filename": pdf_file.filename,
            "insurer_code": pdf_info["insurer_code"],
            "insurer_name": pdf_info["insurer_name"],
            "product_name": pdf_info["product_name"],
            "premium": pdf_info["premium"],
            "coverages": pdf_info["coverages"],
            "coverage_count": len(pdf_info["coverages"]),
        }
        del pdf_info
        return response
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        _cleanup_temp_files([tmp_path])
        _force_gc()


@app.post("/api/parse-hira-pdf")
async def parse_hira_pdf(
    files: List[UploadFile] = File(...),
):
    """심평원 PDF (기본진료/처방조제/세부진료) → JSON 변환
    
    1~3개 PDF를 업로드하면 자동으로 유형을 감지하여 파싱합니다.
    프론트엔드 Excel 파서와 동일한 JSON 형식으로 반환합니다.
    """
    tmp_paths = []
    try:
        result = {
            "success": True,
            "treatRecords": [],
            "rxRecords": [],
            "detailRecords": [],
            "detected_types": [],
            "file_count": len(files),
        }
        
        for f in files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await f.read()
                tmp.write(content)
                tmp_path = tmp.name
                tmp_paths.append(tmp_path)
            del content
            
            # 파일명 또는 내용 기반 유형 감지
            filename_lower = (f.filename or '').lower().replace(' ', '')
            
            if '기본진료' in filename_lower:
                pdf_type = 'basic'
            elif '처방조제' in filename_lower:
                pdf_type = 'prescription'
            elif '세부진료' in filename_lower:
                pdf_type = 'detail'
            else:
                # 파일명으로 판별 불가 시 내용 기반 감지
                pdf_type = detect_hira_pdf_type(tmp_path)
            
            result["detected_types"].append({
                "filename": f.filename,
                "type": pdf_type,
            })
            
            if pdf_type == 'basic':
                records = parse_basic_care_pdf(tmp_path)
                result["treatRecords"] = records
            elif pdf_type == 'prescription':
                records = parse_prescription_pdf(tmp_path)
                result["rxRecords"] = records
            elif pdf_type == 'detail':
                records = parse_detail_care_pdf(tmp_path)
                result["detailRecords"] = records
            else:
                result["detected_types"][-1]["warning"] = "유형 감지 실패 — 파일을 확인해주세요."
        
        # 금액 합계 계산 (청구금 계산용)
        treat_recs = result["treatRecords"]
        total_cost_sum = sum(r.get('총진료비', 0) for r in treat_recs)
        self_cost_sum = sum(r.get('본인부담금', 0) for r in treat_recs)
        insurer_cost_sum = sum(r.get('건강보험혜택', 0) for r in treat_recs)
        
        result["summary"] = {
            "treat_count": len(treat_recs),
            "rx_count": len(result["rxRecords"]),
            "detail_count": len(result["detailRecords"]),
            "total_cost_sum": total_cost_sum,
            "self_cost_sum": self_cost_sum,
            "insurer_cost_sum": insurer_cost_sum,
        }
        
        return result
    
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        _cleanup_temp_files(tmp_paths)
        _force_gc()


@app.post("/api/match-with-summary")
async def match_with_summary(
    pdf_files: List[UploadFile] = File(...),
    excel_file: UploadFile = File(...),
    customer_name: Optional[str] = Form(None),
    threshold: int = Form(75),
    sheet_name: Optional[str] = Form(None),
):
    """PDF + Excel 업로드 → 매칭 결과 JSON 반환 (다운로드 없이 결과만)"""
    tmp_pdf_paths = []
    excel_path = None

    try:
        # Excel 파일 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            content = await excel_file.read()
            tmp.write(content)
            excel_path = tmp.name
        del content

        sn = sheet_name if sheet_name else None

        # 구조 자동 탐지 (A형/B형 자동 분기)
        structure = find_structure(excel_path, sn, 2)
        template_type = structure.get("template_type", "A")
        coverage_col = structure.get("coverage_col", 2)
        first_amount_col = structure.get("first_amount_col", 4)

        insurer_name_row = structure["insurer_row"] or (4 if template_type == "A" else 5)
        product_name_row = structure["product_row"] or (5 if template_type == "A" else 6)
        premium_row = structure["premium_row"] or (6 if template_type == "A" else 9)
        start_row = structure["start_row"] or (8 if template_type == "A" else 11)

        all_results = []

        for pdf_idx, pdf_file in enumerate(pdf_files):
            current_amount_col = first_amount_col + pdf_idx  # A형: D=4,E=5... / B형: F=6,G=7...

            # PDF 임시 저장
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await pdf_file.read()
                tmp.write(content)
                tmp_path = tmp.name
                tmp_pdf_paths.append(tmp_path)
            del content

            # PDF 파싱 (통합 1회 오픈)
            pdf_info = parse_pdf_all_in_one(tmp_path)
            insurer_code = pdf_info["insurer_code"]
            insurer_display = pdf_info["insurer_name"]
            product_name = pdf_info["product_name"]
            premium = pdf_info["premium"]
            pdf_coverages = pdf_info["coverages"]
            # pdf_info의 나머지 데이터는 더 이상 필요 없으므로 삭제
            del pdf_info

            # PDF 임시파일 즉시 삭제 (매칭 전에 메모리 확보)
            try:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
            except Exception:
                pass

            # Excel에서 특약명 읽기 (A형: B열, B형: C열)
            excel_coverages = read_excel_coverages(
                excel_path, sn, coverage_col, current_amount_col, start_row
            )

            # 매칭
            result = match_coverages(pdf_coverages, excel_coverages, threshold)
            matched = result["matched"]

            all_results.append({
                "pdf_name": pdf_file.filename,
                "pdf_index": pdf_idx,
                "column_letter": chr(ord('A') + first_amount_col - 1 + pdf_idx),
                "insurer_code": insurer_code,
                "insurer_name": insurer_display,
                "product_name": product_name,
                "premium": premium,
                "pdf_coverage_count": len(pdf_coverages),
                "pdf_coverages": [
                    {"특약명": c["특약명"], "가입금액": c["가입금액"]}
                    for c in pdf_coverages
                ],
                "matched_count": len(matched),
                "unmatched_excel_count": len(result["unmatched_excel"]),
                "unmatched_pdf_count": len(result["unmatched_pdf"]),
                "matched": [
                    {
                        "excel_row": m["excel_row"],
                        "excel_특약명": m["excel_특약명"],
                        "pdf_특약명": m["pdf_특약명"],
                        "가입금액": m["가입금액"],
                        "가입금액_만원": m["가입금액"] // 10000,
                        "유사도": m["유사도"],
                    }
                    for m in matched
                ],
                "unmatched_excel": [
                    {"특약명": u["특약명"], "row": u["row"]}
                    for u in result["unmatched_excel"]
                ],
                "unmatched_pdf": [
                    {"특약명": u["특약명"], "가입금액": u["가입금액"], "가입금액_만원": u["가입금액"] // 10000}
                    for u in result["unmatched_pdf"]
                ],
            })

            # 매칭 완료 후 중간 데이터 해제
            del pdf_coverages, excel_coverages, result, matched
            _force_gc()

        return {
            "success": True,
            "customer_name": customer_name,
            "template_type": template_type,
            "structure": structure,
            "total_pdfs": len(pdf_files),
            "results": all_results,
        }

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        _cleanup_temp_files([excel_path] + tmp_pdf_paths)
        _force_gc()


@app.post("/api/match")
async def match_and_download(
    pdf_files: List[UploadFile] = File(...),
    excel_file: UploadFile = File(...),
    customer_name: Optional[str] = Form(None),
    threshold: int = Form(75),
    sheet_name: Optional[str] = Form(None),
):
    """PDF + Excel 업로드 → 매칭 결과가 기록된 Excel 파일 다운로드"""
    tmp_pdf_paths = []
    excel_path = None
    output_path = None

    try:
        # Excel 파일 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            content = await excel_file.read()
            tmp.write(content)
            excel_path = tmp.name
        del content

        output_path = excel_path.replace(".xlsx", "_result.xlsx")
        shutil.copy2(excel_path, output_path)

        sn = sheet_name if sheet_name else None

        # 구조 자동 탐지 (A형/B형 자동 분기)
        structure = find_structure(output_path, sn, 2)
        template_type = structure.get("template_type", "A")
        coverage_col = structure.get("coverage_col", 2)
        first_amount_col = structure.get("first_amount_col", 4)

        insurer_name_row = structure["insurer_row"] or (4 if template_type == "A" else 5)
        product_name_row = structure["product_row"] or (5 if template_type == "A" else 6)
        premium_row = structure["premium_row"] or (6 if template_type == "A" else 9)
        start_row = structure["start_row"] or (8 if template_type == "A" else 11)

        for pdf_idx, pdf_file in enumerate(pdf_files):
            current_amount_col = first_amount_col + pdf_idx

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await pdf_file.read()
                tmp.write(content)
                tmp_path = tmp.name
                tmp_pdf_paths.append(tmp_path)
            del content

            # PDF 파싱 (통합 1회 오픈)
            pdf_info = parse_pdf_all_in_one(tmp_path)
            insurer_display = pdf_info["insurer_name"]
            product_name = pdf_info["product_name"]
            premium = pdf_info["premium"]
            pdf_coverages = pdf_info["coverages"]
            del pdf_info

            # PDF 임시파일 즉시 삭제
            try:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)
            except Exception:
                pass

            # 보험사명, 상품명 기록
            write_insurer_info(
                output_path, output_path,
                insurer_display, insurer_name_row,
                product_name, product_name_row,
                current_amount_col, sn
            )

            # 보험료 기록
            if premium:
                write_premium(
                    output_path, output_path,
                    premium, premium_row,
                    current_amount_col, sn
                )

            # 특약 추출 및 매칭 (A형: B열, B형: C열)
            excel_coverages = read_excel_coverages(
                output_path, sn, coverage_col, current_amount_col, start_row
            )

            result = match_coverages(pdf_coverages, excel_coverages, threshold)
            matched = result["matched"]

            # 매칭 결과 기록 (만원 단위)
            if matched:
                write_data = [{
                    "row": m["excel_row"],
                    "amount_col": m["amount_col"],
                    "가입금액": m["가입금액"] // 10000
                } for m in matched]
                write_matched_amounts(output_path, output_path, write_data, sn)

            # 중간 데이터 해제
            del pdf_coverages, excel_coverages, result, matched
            _force_gc()

        # 결과 파일명
        filename = "보장분석표_매칭결과.xlsx"
        if customer_name:
            filename = f"{customer_name}_보장분석표.xlsx"

        # 한글 파일명 인코딩 (RFC 5987)
        encoded_filename = quote(filename)

        def file_iterator(path):
            try:
                with open(path, "rb") as f:
                    while chunk := f.read(65536):
                        yield chunk
            finally:
                # 스트리밍 완료 후 임시파일 삭제
                try:
                    if os.path.exists(path):
                        os.unlink(path)
                except Exception:
                    pass
                _force_gc()

        return StreamingResponse(
            file_iterator(output_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.document",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}",
                "Access-Control-Expose-Headers": "Content-Disposition",
            }
        )

    except Exception as e:
        # 오류 시 output_path 정리
        if output_path and os.path.exists(output_path):
            try:
                os.unlink(output_path)
            except Exception:
                pass
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        _cleanup_temp_files([excel_path] + tmp_pdf_paths)
        _force_gc()
        # 주기적으로 오래된 임시파일 정리
        _cleanup_stale_temp_files()
        # Note: output_path is cleaned up inside file_iterator after streaming


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
