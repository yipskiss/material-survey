// 재질 시각 평가 — Google Sheets 백엔드 (Apps Script) v2
// 사용법: 구글 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
// → 배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포  (반드시 새 버전으로!)

var SHEET_NAME = 'responses_v2';   // 실험 응답 (4문항 버전, 자동 생성)
var PAY_SHEET = 'participants';    // 참여자 인적사항 (자동 생성)
var API_KEY = 'hxlab-material-2026-v2';  // index.html의 APIKEY와 동일해야 함
var ADMIN_KEY = 'hxlab611';      // 관리자 키 — 인적사항 열람/저장용 (각 기기의 관리자 화면에서 입력)

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow([
      'participant', 'trial', 'material', 'category',
      'a1_종이처럼보인다', 'a2_플라스틱처럼보인다', 'a3_주된재료', 'a4_외관선호',
      'timestamp', 'client'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getPaySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PAY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PAY_SHEET);
    sh.appendRow([
      'id', '피험자', '성명', '주민등록번호', '소속', '직위', '활용일시',
      '연락처', '주소', '청구금액', '은행', '예금주', '계좌번호',
      '동의', '동의시각', 'client', '저장시각'
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 시트의 데이터 행을 전부 삭제할 수 없으므로, 마지막 남는 행은 내용만 비운다
function deleteRows_(sh, rowsDesc) {
  var dataRows = sh.getLastRow() - 1;
  rowsDesc.forEach(function (row, j) {
    if (j === rowsDesc.length - 1 && rowsDesc.length === dataRows) {
      sh.getRange(row, 1, 1, sh.getLastColumn()).clearContent();
    } else {
      sh.deleteRow(row);
    }
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var act = data.action || '';

    // ── 인적사항 저장: 누구나 가능 (사이트 기본 키) ──
    if (act === 'addinfo') {
      if (API_KEY && data.key !== API_KEY) return json_({ ok: false, error: 'bad_key' });
      var ps0 = getPaySheet_();
      var f = data.info || {};
      if (!f.id || !f.name) return json_({ ok: false, error: 'missing_fields' });
      var vals = ps0.getDataRange().getValues();
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][0]) === String(f.id)) return json_({ ok: true, dup: true });
      }
      ps0.appendRow([f.id, f.pid || '', f.name, f.rrn || '', f.org || '', f.pos || '', f.dt || '',
        f.phone || '', f.addr || '', f.amt || '', f.bank || '', f.holder || '', f.acct || '',
        f.agree || '', f.agreedAt || '', f.client || '', new Date().toISOString()]);
      return json_({ ok: true, added: 1 });
    }

    // ── 인적사항 열람·삭제 (관리자 키 필요) ──
    if (act === 'listinfo' || act === 'delinfo' || act === 'clearinfo') {
      if (!ADMIN_KEY || data.key !== ADMIN_KEY) return json_({ ok: false, error: 'bad_admin_key' });
      var ps = getPaySheet_();

      if (act === 'listinfo') {
        var pv = ps.getDataRange().getValues();
        var list = [];
        for (var r = 1; r < pv.length; r++) {
          if (!pv[r][0]) continue;
          list.push({ id: pv[r][0], pid: pv[r][1], name: pv[r][2], rrn: String(pv[r][3]), org: pv[r][4],
            pos: pv[r][5], dt: String(pv[r][6]), phone: String(pv[r][7]), addr: pv[r][8], amt: String(pv[r][9]),
            bank: pv[r][10], holder: pv[r][11], acct: String(pv[r][12]), agree: pv[r][13], agreedAt: String(pv[r][14]) });
        }
        return json_({ ok: true, list: list });
      }
      if (act === 'delinfo') {
        var dv = ps.getDataRange().getValues();
        var targets = [];
        for (var d = dv.length - 1; d >= 1; d--) {
          if (String(dv[d][0]) === String(data.id)) targets.push(d + 1);
        }
        deleteRows_(ps, targets);
        return json_({ ok: true, removed: targets.length });
      }
      if (act === 'clearinfo') {
        var cv = ps.getDataRange().getValues();
        var all = [];
        for (var c = cv.length - 1; c >= 1; c--) { if (cv[c][0]) all.push(c + 1); }
        deleteRows_(ps, all);
        return json_({ ok: true, removed: all.length });
      }
    }

    // ── 실험 응답 (API 키) ──
    if (API_KEY && data.key !== API_KEY) return json_({ ok: false, error: 'bad_key' });
    var sh = getSheet_();

    // 특정 재료(시행) 하나만 삭제 — 관리자 순서표에서 사용
    if (act === 'deltrial' && data.pid && data.trial) {
      var tv = sh.getDataRange().getValues();
      var tt = [];
      for (var y = tv.length - 1; y >= 1; y--) {
        if (String(tv[y][0]) === String(data.pid) && Number(tv[y][1]) === Number(data.trial)) tt.push(y + 1);
      }
      deleteRows_(sh, tt);
      return json_({ ok: true, removed: tt.length });
    }

    if (act === 'reset' && data.pid) {
      var rv = sh.getDataRange().getValues();
      var rt = [];
      for (var x = rv.length - 1; x >= 1; x--) {
        if (String(rv[x][0]) === String(data.pid)) rt.push(x + 1);
      }
      deleteRows_(sh, rt);
      return json_({ ok: true, reset: data.pid, removed: rt.length });
    }

    // 재평가(replace) 레코드는 기존 행을 지우고 새로 저장
    var delKeys = {};
    (data.records || []).forEach(function (rec) {
      if (rec && rec.replace && rec.pid && rec.trial) delKeys[rec.pid + '#' + rec.trial] = true;
    });
    var allv = sh.getDataRange().getValues();
    var delRows = [];
    for (var dr = allv.length - 1; dr >= 1; dr--) {
      if (delKeys[allv[dr][0] + '#' + allv[dr][1]]) delRows.push(dr + 1);
    }
    if (delRows.length) { deleteRows_(sh, delRows); allv = sh.getDataRange().getValues(); }

    // 응답 저장 (문항 4개, 중복 pid+trial은 최초 것 유지)
    var existing = {};
    for (var a = 1; a < allv.length; a++) existing[allv[a][0] + '#' + allv[a][1]] = true;

    var rows = [];
    (data.records || []).forEach(function (rec) {
      if (!rec || !rec.pid || !rec.trial || !rec.ratings || rec.ratings.length !== 4) return;
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

// 진행 현황 조회 (피험자별로 몇 번째까지 응답했는지)
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
