# reading-challenge-viewer

## 구조
- `manager/` : 관리자용 입력/구글시트 저장 페이지
- `viewer/` : 참가자용 조회 페이지
- `SHEET_SYNC_FIX.md` : 시트 열 확장 오류 수정 메모

## 현재 엔트리
- Manager: `manager/index.html`
- Viewer: `viewer/index.html`

## 배포 시 주의
- Netlify에서 viewer 배포 경로를 바꿀 경우 publish 대상 파일/폴더를 `viewer/` 기준으로 다시 맞춰야 한다.
- manager도 별도 배포한다면 `manager/` 기준으로 따로 잡는 것이 좋다.
