/**
 * 사후평가 운영 관리표 - 단계별 시각 자동 기록 (구글 시트용)
 *
 * 설치 방법
 *  1) 관리표를 구글 드라이브에 올려 '구글 스프레드시트로 열기'
 *  2) 메뉴 > 확장 프로그램 > Apps Script
 *  3) 기존 코드를 지우고 이 파일 내용을 붙여넣기 > 저장
 *  4) 처음 한 번 권한 승인 창이 뜨면 허용 (onEdit 은 별도 트리거 설정이 필요 없습니다)
 *  5) 연습용 줄(예: 번호 399)로 '접수' 칸에 O를 선택해 '접수 시각'이 찍히는지 확인
 *
 * 동작
 *  - 명단 시트에서 단계 칸을 O로 바꾸는 순간, 해당 단계의 시각 칸이 비어 있으면 현재 시각을 기록합니다.
 *  - 이미 시각이 있으면 덮어쓰지 않습니다(처음 기록 유지). 다시 찍고 싶으면 시각 칸을 지우세요.
 *  - 열 위치(명단 시트 기준): I=접수 J=운동평가 K~M=혈압·혈당 O=상담 Q=추가상담
 *                              X=접수시각 Y=운동평가시각 Z=측정시각 AA=상담시각 AB=추가상담시각
 *  - 명단 시트의 열을 옮기거나 추가하면 아래 STEPS 의 열 번호(A=1, B=2 ...)도 같이 고쳐야 합니다.
 */
const SHEET_NAME = '명단';
const FIRST_ROW = 5;      // 데이터 시작 행
const LAST_STAMP_COL = 28; // AB

const STEPS = [
  { inputCols: [9],          needO: true,  stampCol: 24 }, // 접수
  { inputCols: [10],         needO: true,  stampCol: 25 }, // 운동평가
  { inputCols: [11, 12, 13], needO: false, stampCol: 26 }, // 수축기·이완기·혈당 세 칸이 모두 채워지면
  { inputCols: [15],         needO: true,  stampCol: 27 }, // 선생님 상담
  { inputCols: [17],         needO: true,  stampCol: 28 }, // 추가상담
];

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const firstRow = Math.max(range.getRow(), FIRST_ROW);
  const lastRow = range.getLastRow();
  if (lastRow < firstRow) return;
  const firstCol = range.getColumn();
  const lastCol = range.getLastColumn();

  const rows = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, LAST_STAMP_COL).getValues();
  const now = new Date();

  rows.forEach(function (row, i) {
    STEPS.forEach(function (step) {
      const touched = step.inputCols.some(function (c) { return c >= firstCol && c <= lastCol; });
      if (!touched) return;
      if (row[step.stampCol - 1] !== '') return; // 이미 기록됨
      const done = step.inputCols.every(function (c) {
        const v = String(row[c - 1]).trim();
        return step.needO ? v.toUpperCase() === 'O' : v !== '';
      });
      if (done) sheet.getRange(firstRow + i, step.stampCol).setValue(now);
    });
  });
}
