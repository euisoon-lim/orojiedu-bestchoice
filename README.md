# 오로지교육 베스트 초이스 — 월간 데이터 파이프라인

업로드 페이지(오로지교육 /bestchoice-upload) → Supabase Storage → GitHub Actions 변환 → `data/latest.json` → 오로지교육 베스트 초이스 화면

- `inbox/`  : 업로드된 PPTX (자동 저장)
- `tools/`  : 변환기 (`bc_convert.py` PPTX→데이터, `bc_logos.py` 로고 처리, `build_data.py` 통합)
- `logos/`  : 보험사 로고 원본. 새 보험사는 파일 추가 후 `tools/logos_map.json`에 `"보험사명": "파일명"` 한 줄 추가
- `data/`   : 결과 JSON (`latest.json` = 현재 반영본, `YYYY-MM.json` = 월별 이력)

수동 실행: Actions → "베스트 초이스 변환" → Run workflow (inbox 최신 파일 사용)

## 아임웹 코드 (imweb/)
- `bestchoice_v6.4.txt`(최신) / 이전 버전 v5.2~v6.3 — 베스트 초이스 페이지. `data/latest.json`을 읽어 렌더(월간 교체 불필요)
- `베스트초이스_업로드페이지_v1.0_아임웹바디코드.txt` — /bestchoice-upload 페이지. 비밀번호 + PPTX 드래그 → 자동 반영

## Supabase (오로지교육 프로젝트)
- `supabase_setup.sql` — 버킷 `bestchoice-inbox`, 테이블 `bestchoice_config`(passcode/gh_pat/gh_repo), `bestchoice_uploads`
- `supabase_edge_function_bestchoice-upload.ts` — Edge Function `bestchoice-upload` (Verify JWT OFF)
- 비밀번호 변경: `update bestchoice_config set value='새비밀번호' where key='passcode';`
