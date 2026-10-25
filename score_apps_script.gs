/**
 * 퀴즈 점수 받기 (구글 Apps Script)
 * - 학생이 '선생님께 점수 보내기'를 누르면 이 시트의 '점수' 탭에 한 줄씩 시간 순으로 쌓여요.
 * - 같은 학생(학년·번호·이름)은 영역(IT 지식 · HTML · 파이썬)마다 3번까지만 받아요.
 * 설정 방법은 README.md의 '점수 받기 설정'을 보세요.
 */
const SHEET_NAME = '퀴즈 점수';
const LIMIT = 3;
const AREAS = ['IT 지식', 'HTML', '파이썬'];
const HEADER = ['보낸 시각', '학년', '번호', '이름', '영역', '100점 환산', '맞힌 수', '문제 수', '퀴즈별 점수', '몇 번째'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const grade = Number(d.grade), num = Number(d.num);
    const name = String(d.name || '').trim().slice(0, 20);
    const area = String(d.area || '').trim();
    if (!(grade >= 1 && grade <= 3 && num >= 1 && num <= 20)) return reply({ ok: false, error: '학년·번호가 올바르지 않음' });
    if (!name) return reply({ ok: false, error: '이름이 없음' });
    if (AREAS.indexOf(area) < 0) return reply({ ok: false, error: '영역이 올바르지 않음' });

    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    let count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (Number(rows[i][1]) === grade && Number(rows[i][2]) === num && String(rows[i][3]) === name && String(rows[i][4]) === area) count++;
    }
    if (count >= LIMIT) return reply({ ok: false, error: 'limit', count: count });

    sheet.appendRow([
      new Date(), grade, num, name, area,
      Number(d.score) || 0, Number(d.correct) || 0, Number(d.total) || 0,
      String(d.detail || '').slice(0, 300), count + 1
    ]);
    return reply({ ok: true, count: count + 1 });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
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
  }
  return sheet;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
