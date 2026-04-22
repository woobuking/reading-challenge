# manager Apps Script 전환 가이드

## 목적
기존 manager 프론트에 하드코딩된 Google Sheets API 키, client secret, refresh token을 제거하고 Apps Script 웹앱 중계 방식으로 전환한다.

## 변경 사항
- 프론트에서 제거:
  - `API_KEY`
  - `CLIENT_ID`
  - `CLIENT_SECRET`
  - `REFRESH_TOKEN`
  - Google Sheets 직접 호출 로직
- 프론트에서 유지:
  - `APPS_SCRIPT_URL`
  - 로컬 상태 렌더링
  - 날짜 선택, 체크 토글 UI
- Apps Script에서 담당:
  - 시트 불러오기(`action=load`)
  - 시트 저장(`POST { date, record }`)
  - 날짜 열 자동 확장
  - 멤버 O/X, 성공 수, 성공률 기록

## 적용 순서
1. `manager/Code.gs`를 Apps Script 편집기에 붙여넣기
2. 웹앱으로 배포
3. 배포된 `/exec` URL을 `manager/index.html`의 `APPS_SCRIPT_URL`에 넣기
4. 저장 테스트
5. 정상 확인 후 manager repo 분리 진행

## 현재 해야 할 것
- `REPLACE_WITH_APPS_SCRIPT_WEBAPP_URL` 자리에 실제 웹앱 URL을 넣어야 한다.
