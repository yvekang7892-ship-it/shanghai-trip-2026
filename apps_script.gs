// ── 상하이 여행 예산 연동 Apps Script ──
// "예산" 탭에 지출을 저장/조회하고, 총예산을 보관합니다.
// 사이트(GitHub Pages)에서 이 웹앱 URL로 호출합니다.

const SHEET_NAME = '예산';
const HEADERS = ['id', '날짜', '분류', '내용', '위안(CNY)', '원화(KRW)', '입력자', '기록시각'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f0ece8');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 사이트가 로드될 때 전체 데이터를 읽어갑니다.
function doGet(e) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const expenses = values.slice(1)
    .filter(function (r) { return r[0] !== '' && r[0] != null; })
    .map(function (r) {
      return { id: Number(r[0]), date: r[1], cat: r[2], desc: r[3],
               cny: Number(r[4]) || 0, krw: Number(r[5]) || 0, who: r[6] || '' };
    });
  const budget = Number(PropertiesService.getDocumentProperties().getProperty('totalBudget') || 0);
  return json_({ totalBudget: budget, expenses: expenses });
}

// 사이트에서 추가/삭제/예산설정 시 호출됩니다.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    if (body.action === 'add') {
      sh.appendRow([body.id, body.date, body.cat, body.desc,
                    body.cny || 0, body.krw || 0, body.who || '', new Date()]);

    } else if (body.action === 'delete') {
      const values = sh.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) {
        if (String(values[i][0]) === String(body.id)) { sh.deleteRow(i + 1); break; }
      }

    } else if (body.action === 'budget') {
      PropertiesService.getDocumentProperties()
        .setProperty('totalBudget', String(body.totalBudget || 0));
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
