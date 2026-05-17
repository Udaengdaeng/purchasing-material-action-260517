
# Material Action Dashboard

생산계획 변경에 따라 자재 필요량, 현재 재고, 부족 자재, 발주 액션을 한눈에 확인하기 위한 Streamlit 대시보드입니다.

## 주요 기능

- 엑셀 업로드 기반 분석
- 생산계획별 자재 필요량 자동 계산
- BOM 기준 소요량 계산
- 현재 재고와 비교하여 부족 자재 산출
- MOQ/Loss 반영 권장 발주량 계산
- 공급업체별 요청 액션 문장 생성
- PO 상세 화면에서 BOM/재고/입고/출고 원본 데이터 확인

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포

1. GitHub 레포지토리에 `app.py`, `requirements.txt` 업로드
2. Streamlit Cloud에서 New app 생성
3. Main file path를 `app.py`로 설정
4. Deploy
