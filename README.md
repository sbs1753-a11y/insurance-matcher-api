# 보험 보장분석 자동매칭 API

보험 가입제안서 PDF와 보장분석표 Excel을 매칭하여 가입금액을 자동 입력하는 FastAPI 서버.

## API 엔드포인트

| 메서드 | 경로 | 설명 |
|---|---|---|
| GET | `/health` | 헬스체크 |
| POST | `/api/parse-pdf` | 단일 PDF 파싱 (보험사/상품명/보험료/특약 추출) |
| POST | `/api/match-with-summary` | PDF+Excel 매칭 결과 JSON 반환 |
| POST | `/api/match` | PDF+Excel 매칭 결과 Excel 파일 다운로드 |

## 지원 보험사
- 메리츠화재 (범용 테이블 파서)
- 삼성생명 (전용 파서)
- 미래에셋생명 (전용 파서)
- KB손해보험 (전용 파서)
- 기타 보험사 (범용 파서)

## Render 배포
- Python 3.11
- Web Service (Free tier)
- Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
