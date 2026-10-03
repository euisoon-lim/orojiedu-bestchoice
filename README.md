# 오로지교육 베스트 초이스 — 월간 데이터 파이프라인

업로드 페이지(오로지교육 /bestchoice-upload) → Supabase Storage → GitHub Actions 변환 → `data/latest.json` → 오로지교육 베스트 초이스 화면

- `inbox/`  : 업로드된 PPTX (자동 저장)
- `tools/`  : 변환기 (`bc_convert.py` PPTX→데이터, `bc_logos.py` 로고 처리, `build_data.py` 통합)
- `logos/`  : 보험사 로고 원본. 새 보험사는 파일 추가 후 `tools/logos_map.json`에 `"보험사명": "파일명"` 한 줄 추가
- `data/`   : 결과 JSON (`latest.json` = 현재 반영본, `YYYY-MM.json` = 월별 이력)

수동 실행: Actions → "베스트 초이스 변환" → Run workflow (inbox 최신 파일 사용)
