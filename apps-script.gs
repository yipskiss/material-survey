// 재질 시각 평가 — Google Sheets 백엔드 (Apps Script)
// 사용법: 구글 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// → 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포 → URL 복사

var SHEET_NAME = 'responses';
var API_KEY = 'hxlab-material-2026'; // index.html의 APIKEY와 같아야 함. 원하면 바꿔도 됨(양쪽 동일하게)

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow([
      'participant', 'trial', 'material', 'category',
      'q1_플라스틱같다-종이같다', 'q2_인공적이다-자연스럽다', 'q3_거칠다-매끄럽다',
      'q4_무광이다-유광이다', 'q5_차갑다-따뜻하다', 'q6_불투명하다-투명하다',
      'q7_무채색이다-다채롭다', 'q8_색이약하다-색이강하다', 'q9_선호하지않는다-선호한다',
      'timestamp', 'client'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 저장: 앱이 재료 하나 끝날 때마다 호출
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    if (API_KEY && data.key !== API_KEY) return json_({ ok: false, error: 'bad_key' });

    var sh = getSheet_();

    // 피험자 초기화 요청
    if (data.action === 'reset' && data.pid) {
      var vals = sh.getDataRange().getValues();
      for (var i = vals.length - 1; i >= 1; i--) {
        if (String(vals[i][0]) === String(data.pid)) sh.deleteRow(i + 1);
      }
      return json_({ ok: true, reset: data.pid });
    }

    // 응답 저장 (중복 pid+trial은 최초 것 유지)
    var existing = {};
    var all = sh.getDataRange().getValues();
    for (var r = 1; r < all.length; r++) existing[all[r][0] + '#' + all[r][1]] = true;

    var rows = [];
    (data.records || []).forEach(function (rec) {
      if (!rec || !rec.pid || !rec.trial || !rec.ratings || rec.ratings.length !== 9) return;
      if (existing[rec.pid + '#' + rec.trial]) return;
      existing[rec.pid + '#' + rec.trial] = true;
      rows.push([rec.pid, rec.trial, rec.material, rec.category]
        .concat(rec.ratings).concat([rec.timestamp, rec.client || '']));
    });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return json_({ ok: true, added: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 진행 현황 조회: 앱이 열릴 때 호출 (피험자별로 몇 번째까지 응답했는지)
function doGet() {
  var sh = getSheet_();
  var vals = sh.getDataRange().getValues();
  var seen = {}, answered = {};
  for (var i = 1; i < vals.length; i++) {
    var pid = String(vals[i][0]), trial = Number(vals[i][1]);
    if (!pid || !trial) continue;
    var k = pid + '#' + trial;
    if (seen[k]) continue;
    seen[k] = true;
    (answered[pid] = answered[pid] || []).push(trial);
  }
  return json_({ ok: true, answered: answered });
}
