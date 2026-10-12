/**
 * 사후평가 운영 관리표 - 입력시각 자동 기록 (구글 시트용)
 *
 * 설치 방법
 *  1) 관리표를 구글 드라이브에 올려 '구글 스프레드시트로 열기'
 *  2) 메뉴 > 확장 프로그램 > Apps Script
 *  3) 기존 코드를 지우고 이 파일 내용을 붙여넣기 > 저장
 *  4) 처음 한 번 권한 승인 창이 뜨면 허용 (onEdit 은 별도 트리거 설정이 필요 없습니다)
 *  5) 입력_운동 탭에서 번호 칸(B열)에 숫자를 입력해 A열에 시각이 찍히는지 확인 (리허설 후 지우기)
 *
 * 동작
 *  - 입력_운동 / 입력_측정 / 입력_상담 탭에서 '번호'(B열)를 입력하면, 같은 줄 A열(입력시각)이 비어 있을 때
 *    현재 시각을 기록합니다. 이미 시각이 있으면 덮어쓰지 않습니다.
 *  - 번호를 지우면 그 줄의 시각도 지웁니다(줄을 재사용할 때 이전 시각이 남지 않도록).
 *  - 여러 줄을 붙여넣어도 줄마다 처리됩니다.
 *  - 탭 이름이나 열 위치를 바꾸면 아래 상수도 같이 고쳐야 합니다.
 */
const LOG_SHEETS = ['입력_운동', '입력_측정', '입력_상담'];
const FIRST_ROW = 5;  // 데이터 시작 행
const ID_COL = 2;     // B: 번호
const STAMP_COL = 1;  // A: 입력시각

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (LOG_SHEETS.indexOf(sheet.getName()) === -1) return;

  // 번호 칸이 이번 편집 범위에 포함된 경우만 처리
  if (ID_COL < range.getColumn() || ID_COL > range.getLastColumn()) return;

  const firstRow = Math.max(range.getRow(), FIRST_ROW);
  const lastRow = range.getLastRow();
  if (lastRow < firstRow) return;
  const n = lastRow - firstRow + 1;

  const ids = sheet.getRange(firstRow, ID_COL, n, 1).getValues();
  const stamps = sheet.getRange(firstRow, STAMP_COL, n, 1).getValues();
  const now = new Date();
  let changed = false;

  for (let i = 0; i < n; i++) {
    const hasId = String(ids[i][0]).trim() !== '';
    const hasStamp = stamps[i][0] !== '';
    if (hasId && !hasStamp) {
      stamps[i][0] = now;
      changed = true;
    } else if (!hasId && hasStamp) {
      stamps[i][0] = '';
      changed = true;
    }
  }
  if (changed) sheet.getRange(firstRow, STAMP_COL, n, 1).setValues(stamps);
}
