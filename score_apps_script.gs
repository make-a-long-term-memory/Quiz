/**
 * 퀴즈 점수 받기 (구글 Apps Script)
 * - 학생이 '선생님께 점수 보내기'를 누르면 이 시트의 '퀴즈 점수' 탭에 한 줄씩 시간 순으로 쌓여요.
 * - 처음 한 번은 위쪽 함수 칸에서 setup 을 골라 ▶ 실행하고 권한을 허용하세요. '퀴즈 점수' 탭이 만들어져요.
 * - 학생이 퀴즈를 끝까지 풀면 3초 뒤 자동으로 보내져요. 같은 학번은 퀴즈마다 3번까지만 받아요.
 * 설정 방법은 README.md의 '점수 받기 설정'을 보세요.
 */
const SHEET_NAME = '퀴즈 점수';
const LIMIT = 3;
const AREAS = ['IT 지식', 'HTML', '파이썬'];
const QUIZZES = ['인터넷과 IT 서비스', 'HTML OX', 'HTML 빈칸 채우기', 'HTML 첫걸음', 'CSS 레이아웃', '파이썬 OX', '파이썬 빈칸 채우기', '파이썬 첫걸음', '파이썬 선택문'];
const HEADER = ['보낸 시각', '학번', '학년', '반', '번호', '이름', '영역', '퀴즈', '맞힌 수', '문제 수', '100점 환산', '영역 누적(100점)', '영역 퀴즈별 점수', '몇 번째'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const grade = Number(d.grade), cls = Number(d.cls), num = Number(d.num);
    const name = String(d.name || '').trim().slice(0, 20);
    const area = String(d.area || '').trim();
    if (!(grade >= 1 && grade <= 3 && (cls === 5 || cls === 6) && num >= 1 && num <= 20)) return reply({ ok: false, error: '학년·반·번호가 올바르지 않음' });
    const id = String(grade) + cls + (num < 10 ? '0' : '') + num;
    if (!name) return reply({ ok: false, error: '이름이 없음' });
    const quiz = String(d.quiz || '').trim();
    if (AREAS.indexOf(area) < 0) return reply({ ok: false, error: '영역이 올바르지 않음' });
    if (QUIZZES.indexOf(quiz) < 0) return reply({ ok: false, error: '퀴즈 이름이 올바르지 않음' });

    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === id && String(rows[i][7]) === quiz) count++;
    }
    if (count >= LIMIT) return reply({ ok: false, error: 'limit', count: count });

    sheet.appendRow([
      new Date(), id, grade, cls, num, name, area, quiz,
      Number(d.correct) || 0, Number(d.total) || 0, Number(d.score) || 0, Number(d.areaScore) || 0,
      String(d.detail || '').slice(0, 300), count + 1
    ]);
    return reply({ ok: true, count: count + 1 });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function setup() {
  getSheet();
}

function doGet() {
  return reply({ ok: true, message: '점수 받기 준비됨' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
  } else if (sheet.getRange(1, 1, 1, HEADER.length).getValues()[0].join('|') !== HEADER.join('|')) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);  // 머리글이 예전 것이면 새것으로
  }
  return sheet;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
