/**
 * 학교폭력 예방 플래시몹 경진대회 — 심사 점수 저장용 Google Apps Script
 *
 * 구글 시트의 '점수' 탭이 데이터베이스 역할을 합니다.
 *  - 한 줄 = 심사위원 1명이 학급 1곳에 매긴 점수 (심사위원번호+학년+반이 같으면 덮어씀)
 *  - 웹앱은 GET 요청 + URL 파라미터로만 호출합니다 (POST의 CORS 문제를 피하기 위함)
 *
 * 처음 한 번: 편집기 위쪽에서 setup 함수를 선택해 [실행] → 권한 승인
 */

// ▼ 관리자 PIN — 반드시 바꾸세요 (집계·순위 화면 접속용)
const ADMIN_PIN = '0000';

const SHEET_NAME = '점수';
const N_ITEMS = 20;
const ITEM_MAX = [6,6,6,6,6, 6,6,6,6,6, 6,6,6,6,6, 2,2,2,2,2];
const ITEM_NAMES = [
  '협동-전원 참여','협동-동작 일치도','협동-대형 전환','협동-역할 분담·배려','협동-연습 완성도',
  '창의-주제 전달력','창의-자유곡 선곡·편곡','창의-안무 독창성','창의-구성·연출','창의-소품·구호 활용',
  '질서-무대 입장','질서-시작 준비','질서-공간·간격 유지','질서-안전·바른 언행','질서-퇴장·무대 정리',
  '호응-관객 반응','호응-호응 유도','호응-표정·에너지','호응-다른 반 응원','호응-존중하는 응원'
];
const HEADERS = ['저장 시각','심사위원','성명','학년','반'].concat(ITEM_NAMES)
  .concat(['협동심','창의성','질서','호응도','합계','입력 항목 수']);
const COL_ITEMS = 6;                 // F열부터 20개 항목
const COL_SUMS = COL_ITEMS + N_ITEMS; // 소계 시작 열

function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    switch (p.action) {
      case 'ping':     out = { ok: true, time: new Date().toISOString() }; break;
      case 'save':     out = saveRecord_(JSON.parse(p.data || '{}')); break;
      case 'mine':     out = { ok: true, records: readRecords_().filter(r => String(r.judge) === String(p.judge)) }; break;
      case 'checkPin': checkPin_(p.pin); out = { ok: true }; break;
      case 'all':      checkPin_(p.pin); out = { ok: true, records: readRecords_() }; break;
      default: throw new Error('알 수 없는 요청입니다: ' + p.action);
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkPin_(pin) {
  if (String(pin || '') !== String(ADMIN_PIN)) throw new Error('PIN이 맞지 않습니다.');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#dbeafe');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 저장: 같은 심사위원·학년·반이 있으면 그 줄을 고치고, 없으면 새 줄을 추가 */
function saveRecord_(rec) {
  const judge = Number(rec.judge), grade = Number(rec.grade), cls = Number(rec.cls);
  if (!judge || !grade || !cls) throw new Error('심사위원·학년·반 정보가 없습니다.');
  if (!Array.isArray(rec.items) || rec.items.length !== N_ITEMS) throw new Error('항목 수가 맞지 않습니다.');
  const items = rec.items.map((v, i) => {
    if (v === null || v === '' || v === undefined) return '';
    const n = Number(v);
    if (!Number.isInteger(n) || n < 0 || n > ITEM_MAX[i]) throw new Error((i + 1) + '번 항목 점수가 범위를 벗어났습니다.');
    return n;
  });
  const sums = [0, 5, 10, 15].map(s => {
    const part = items.slice(s, s + 5).filter(v => v !== '');
    return part.length ? part.reduce((a, b) => a + b, 0) : '';
  });
  const total = sums.reduce((a, b) => a + (b === '' ? 0 : b), 0);
  const count = items.filter(v => v !== '').length;
  const now = new Date();
  const row = [now, judge, String(rec.name || '').slice(0, 20), grade, cls].concat(items).concat(sums).concat([total, count]);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);               // 여러 심사위원이 동시에 저장해도 줄이 겹치지 않도록
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    let target = -1;
    if (last >= 2) {
      const keys = sh.getRange(2, 2, last - 1, 4).getValues(); // 심사위원, 성명, 학년, 반
      for (let i = 0; i < keys.length; i++) {
        if (Number(keys[i][0]) === judge && Number(keys[i][2]) === grade && Number(keys[i][3]) === cls) { target = i + 2; break; }
      }
    }
    if (target > 0) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, ts: now.toISOString() };
}

function readRecords_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, COL_SUMS - 1).getValues();
  return vals.filter(r => r[1] !== '').map(r => ({
    ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    judge: Number(r[1]), name: String(r[2] || ''), grade: Number(r[3]), cls: Number(r[4]),
    items: r.slice(COL_ITEMS - 1, COL_ITEMS - 1 + N_ITEMS).map(v => v === '' ? null : Number(v))
  }));
}

/** 처음 한 번 실행: '점수' 탭 만들기 + 권한 승인 */
function setup() {
  sheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('점수 탭 준비 완료. 이제 [배포] → [새 배포]로 웹앱을 만드세요.');
}
