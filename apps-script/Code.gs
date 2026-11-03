/**
 * 경제정책 시뮬레이터 → 구글 시트 저장용 Apps Script
 * 시트에서 [확장 프로그램] > [Apps Script]에 이 코드를 붙여 넣고
 * [배포] > [새 배포] > 유형: 웹 앱, 실행: 나, 액세스: 모든 사용자 로 배포하세요.
 */
const PROG = '진행', SUBMIT = '제출';
const PROG_HEAD = ['갱신 시각','학번','이름','모둠','상황','진행 분기','성장률(%)','물가(%)','실업률(%)','환율(원)','국가채무(%)','현재 점수','최근 정책 근거','세션 ID'];
const SUBMIT_HEAD = ['제출 시각','학번','이름','모둠','상황','점수','목표 결과','진단1: 문제와 원인','진단2: 충돌 목표와 순서',
  '1분기','2분기','3분기','4분기','5분기','6분기','7분기','8분기','성찰1: 효과적 정책','성찰2: 예상과 다른 결과','성찰3: 다시 한다면','보고서 전문','세션 ID'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (d.type === 'progress') upsertProgress(ss, d);
    else appendSubmit(ss, d);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
function doGet() { return out({ ok: true, msg: '경제정책 시뮬레이터 저장소가 작동 중입니다.' }); }

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean(v) { // 학생 글이 수식으로 실행되지 않도록
  if (typeof v !== 'string') return v;
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function sheet(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1,1,1,head.length).setFontWeight('bold'); }
  return sh;
}
function upsertProgress(ss, d) {
  const sh = sheet(ss, PROG, PROG_HEAD);
  const row = [new Date(), d.gid, d.name, d.team, d.scenario, d.q + '/8', d.g, d.pi, d.u, d.fx, d.debt, d.score, d.lastWhy, d.sid].map(clean);
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, PROG_HEAD.length, last - 1, 1).getValues().map(r => r[0]);
    const i = ids.indexOf(d.sid);
    if (i >= 0) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}
function appendSubmit(ss, d) {
  const sh = sheet(ss, SUBMIT, SUBMIT_HEAD);
  const q = (d.decisions || []).slice(0, 8);
  while (q.length < 8) q.push('');
  const row = [new Date(), d.gid, d.name, d.team, d.scenario, d.score, d.goals, d.diag1, d.diag2]
    .concat(q).concat([d.ref1, d.ref2, d.ref3, d.report, d.sid]).map(clean);
  sh.appendRow(row);
  sh.getRange(sh.getLastRow(), 1, 1, row.length).setWrap(true).setVerticalAlignment('top');
}
