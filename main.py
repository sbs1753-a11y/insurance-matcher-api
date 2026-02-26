import os
import tempfile
import shutil
import traceback
from urllib.parse import quote
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse, JSONResponse
from typing import List, Optional

from pdf_parser import (
    extract_coverage_from_pdf, detect_insurer,
    detect_product_name, extract_premium,
    parse_pdf_all_in_one
)
from excel_handler import (
    read_excel_coverages, write_matched_amounts,
    write_insurer_info, write_premium, find_structure
)
from matcher import match_coverages

app = FastAPI(title="보험 보장분석 자동매칭 API", version="1.0.0")

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
    "kb": "KB손해보험", "db": "DB손해보험",
    "mirae": "미래에셋생명", "abl": "ABL생명",
    "heungkuk": "흥국생명", "hanwha": "한화생명",
    "hyundai": "현대해상", "lotte": "롯데손해보험",
    "nh": "NH농협생명", "dongyang": "동양생명",
    "kyobo": "교보생명", "shinhan": "신한라이프",
}


@app.get("/")
async def root():
    return {"status": "ok", "service": "보험 보장분석 자동매칭 API", "version": "1.0.0"}


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/api/parse-pdf")
async def parse_pdf(pdf_file: UploadFile = File(...)):
    """단일 PDF 파싱 — 보험사, 상품명, 보험료, 특약 목록 반환"""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            content = await pdf_file.read()
            tmp.write(content)
            tmp_path = tmp.name

        insurer_code = detect_insurer(tmp_path)
        insurer_name = INSURER_NAMES.get(insurer_code, insurer_code or "알 수 없음")
        product_name = detect_product_name(tmp_path)
        premium = extract_premium(tmp_path)
        coverages = extract_coverage_from_pdf(tmp_path)

        return {
            "success": True,
            "filename": pdf_file.filename,
            "insurer_code": insurer_code,
            "insurer_name": insurer_name,
            "product_name": product_name,
            "premium": premium,
            "coverages": coverages,
            "coverage_count": len(coverages),
        }
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


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

        sn = sheet_name if sheet_name else None

        # 구조 자동 탐지
        structure = find_structure(excel_path, sn, 2)
        insurer_name_row = structure["insurer_row"] or 4
        product_name_row = structure["product_row"] or 5
        premium_row = structure["premium_row"] or 6
        start_row = structure["start_row"] or 8

        all_results = []

        for pdf_idx, pdf_file in enumerate(pdf_files):
            current_amount_col = 4 + pdf_idx  # D=4, E=5, F=6, ...

            # PDF 임시 저장
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await pdf_file.read()
                tmp.write(content)
                tmp_path = tmp.name
                tmp_pdf_paths.append(tmp_path)

            # PDF 파싱 (통합 1회 오픈)
            pdf_info = parse_pdf_all_in_one(tmp_path)
            insurer_code = pdf_info["insurer_code"]
            insurer_display = pdf_info["insurer_name"]
            product_name = pdf_info["product_name"]
            premium = pdf_info["premium"]
            pdf_coverages = pdf_info["coverages"]

            # Excel에서 특약명 읽기
            excel_coverages = read_excel_coverages(
                excel_path, sn, 2, current_amount_col, start_row
            )

            # 매칭
            result = match_coverages(pdf_coverages, excel_coverages, threshold)
            matched = result["matched"]

            all_results.append({
                "pdf_name": pdf_file.filename,
                "pdf_index": pdf_idx,
                "column_letter": chr(ord('D') + pdf_idx),
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

        return {
            "success": True,
            "customer_name": customer_name,
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
        if excel_path and os.path.exists(excel_path):
            os.unlink(excel_path)
        for p in tmp_pdf_paths:
            if os.path.exists(p):
                os.unlink(p)


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

        output_path = excel_path.replace(".xlsx", "_result.xlsx")
        shutil.copy2(excel_path, output_path)

        sn = sheet_name if sheet_name else None

        # 구조 자동 탐지
        structure = find_structure(output_path, sn, 2)
        insurer_name_row = structure["insurer_row"] or 4
        product_name_row = structure["product_row"] or 5
        premium_row = structure["premium_row"] or 6
        start_row = structure["start_row"] or 8

        for pdf_idx, pdf_file in enumerate(pdf_files):
            current_amount_col = 4 + pdf_idx

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                content = await pdf_file.read()
                tmp.write(content)
                tmp_path = tmp.name
                tmp_pdf_paths.append(tmp_path)

            # PDF 파싱 (통합 1회 오픈)
            pdf_info = parse_pdf_all_in_one(tmp_path)
            insurer_display = pdf_info["insurer_name"]
            product_name = pdf_info["product_name"]
            premium = pdf_info["premium"]

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

            # 특약 추출 및 매칭
            pdf_coverages = pdf_info["coverages"]
            excel_coverages = read_excel_coverages(
                output_path, sn, 2, current_amount_col, start_row
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

        # 결과 파일명
        filename = "보장분석표_매칭결과.xlsx"
        if customer_name:
            filename = f"{customer_name}_보장분석표.xlsx"

        # 한글 파일명 인코딩 (RFC 5987)
        encoded_filename = quote(filename)

        def file_iterator(path):
            with open(path, "rb") as f:
                while chunk := f.read(65536):
                    yield chunk
            # 스트리밍 완료 후 임시파일 삭제
            if os.path.exists(path):
                os.unlink(path)

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
            os.unlink(output_path)
        return JSONResponse(
            status_code=500,
            content={"success": False, "error": str(e), "traceback": traceback.format_exc()}
        )
    finally:
        if excel_path and os.path.exists(excel_path):
            os.unlink(excel_path)
        for p in tmp_pdf_paths:
            if os.path.exists(p):
                os.unlink(p)
        # Note: output_path is cleaned up inside file_iterator after streaming


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 10000))
    uvicorn.run(app, host="0.0.0.0", port=port)
