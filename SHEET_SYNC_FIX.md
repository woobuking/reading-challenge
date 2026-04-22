# Google Sheets 동기화 오류 수정 메모

## 문제
저장 시 아래 오류가 발생했다.

- `Unable to parse range: 시트2!W1`
- `Range ('시트2'!W1) exceeds grid limits`

즉, 프론트 코드가 W열 같은 더 뒤쪽 열에 값을 쓰려 했지만 시트의 실제 columnCount가 그만큼 확보되지 않은 상태였다.

## 원인
`challenger manager.html`의 `syncToSheets()`는 날짜 열이 없을 때 바로 특정 셀 주소에 값을 쓰도록 만들었고, 저장 전에 시트 열 수를 확장하는 로직이 없었다.

## 수정 내용
- 저장 전에 Sheets 메타데이터를 조회해 현재 `columnCount`를 읽는다.
- 필요한 열 수가 현재 열 수보다 크면 `spreadsheets.batchUpdate`와 `appendDimension`으로 먼저 열을 확장한다.
- 그 다음 `values:batchUpdate`를 호출한다.

## 수정 파일
- `challenger manager.html`
- `SHEET_SYNC_FIX.md`

## 기대 효과
- 날짜가 뒤쪽 열로 밀려도 저장 시 자동으로 시트 열이 늘어난다.
- `W1 exceeds grid limits` 류 오류가 재발하지 않는다.

## 참고
이 문제는 Apps Script가 아니라 프론트의 Google Sheets 직접 저장 로직 문제였다.
