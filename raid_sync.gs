/**
 * 로아 레이드 일정 — 공대원 가능 시간 공유 (Google Apps Script)
 *
 * 설치 방법 (매장 지도 team_sync.gs 때와 같아요)
 * 1) 구글 시트를 새로 만들고 → 메뉴 [확장 프로그램] → [Apps Script]
 * 2) 원래 있던 코드를 지우고 이 파일 내용을 전부 붙여넣기
 * 3) 아래 TEAM_CODE 를 공대원끼리 쓸 코드로 바꾸기 (4자 이상, 예: '별빛공대2026')
 * 4) 저장(디스크 아이콘) → 오른쪽 위 [배포] → [새 배포] → 톱니바퀴에서 [웹 앱] 선택
 *    - 다음 사용자 인증 정보로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] → 액세스 승인 → 나오는 '웹 앱 URL'(…/exec)을 config.js 의 teamSyncUrl 에 넣기
 *
 * 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전: '새 버전' → [배포] 를 해야 반영돼요. (URL은 그대로)
 *
 * 시트 구성 (처음 쓸 때 탭이 자동으로 생겨요)
 * - '일정' 탭: 레이드 일정 하나가 한 줄 (ID · 제목 · 레이드 · 기간 · 시간대 …). 상태 칸이 '삭제'면 목록에서 숨겨져요.
 * - '가능시간' 탭: 공대원 한 명이 한 줄 (일정 ID · 이름 · 가능 시간 · 데이터 · 수정 시간)
 *   '가능 시간' 칸은 사람이 읽는 요약이고, '데이터' 칸은 사이트가 읽는 값이라 손대지 마세요.
 * 시트를 직접 고쳤다면 시트 메뉴 [레이드 일정] → [사이트에 바로 반영] 을 누르세요. (안 눌러도 2분 안에 반영돼요)
 */
const TEAM_CODE = '여기에_팀코드'; // 공대원들이 쓰는 코드. 사이트의 '링크 복사'로 보낸 링크에 들어가 있어요.

const API_VERSION = 1;
const SCHED_SHEET = '일정';
const SCHED_HEADER = ['ID', '제목', '레이드', '난이도', '시작일', '종료일', '시작 시각', '종료 시각',
  '칸(분)', '인원', '소요(분)', '메모', '만든 사람', '만든 시간', '수정 시간', '상태'];
const MEM_SHEET = '가능시간';
const MEM_HEADER = ['일정 ID', '이름', '가능 시간', '데이터', '수정 시간'];
const CACHE_SEC = 120; // 시트를 직접 고쳤을 때 사이트에 반영되기까지 최대 시간(초)

// ───────────────────────── 시트 ─────────────────────────

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@'); // 날짜·시각이 자동 변환되지 않게 '텍스트'
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (!sh.getRange(1, 1).getDisplayValue()) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function schedSheet_() { return sheet_(SCHED_SHEET, SCHED_HEADER); }
function memSheet_() { return sheet_(MEM_SHEET, MEM_HEADER); }

// 헤더 아래 전체 줄 (항상 문자열로)
function rows_(sh, width) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, width).getDisplayValues();
}

// r = 시트의 줄 번호(1부터). 텍스트 형식으로 맞춘 뒤 씀 (시트 줄이 모자라면 100줄 늘림)
function writeRow_(sh, r, values) {
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), Math.max(100, r - sh.getMaxRows()));
  sh.getRange(r, 1, 1, values.length).setNumberFormat('@').setValues([values]);
}

// ───────────────────────── 값 정리 ─────────────────────────

function str_(v, n) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, n);
}
// 시트에서 수식으로 해석되지 않게 (=, +, -, @ 로 시작하면 앞에 ' 를 붙임)
function cell_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function name_(v) { return str_(v, 24); }
function sameName_(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }

// '2026-11-11', '2026. 11. 11', '2026/11/11' → '2026-11-11'
function normDate_(s) {
  const m = /(\d{4})\D+(\d{1,2})\D+(\d{1,2})/.exec(String(s || ''));
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}
// '19:00', '19:00:00', '오후 7:00:00' → '19:00'
function normTime_(s) {
  s = String(s || '');
  const m = /(\d{1,2}):(\d{2})/.exec(s);
  if (!m) return '';
  let h = Number(m[1]);
  if (/오후|PM/i.test(s) && h < 12) h += 12;
  if (/오전|AM/i.test(s) && h === 12) h = 0;
  return ('0' + h).slice(-2) + ':' + m[2];
}
function dayNum_(ymd) {
  const p = ymd.split('-').map(Number);
  return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

// 일정 입력값 검사 → 정리된 값 또는 { error }
function cleanSchedule_(s) {
  const raid = str_(s.raid, 40);
  const level = str_(s.level, 20);
  const v = {
    title: str_(s.title, 60) || str_(raid + ' ' + level, 60),
    raid: raid,
    level: level,
    start: normDate_(s.start),
    end: normDate_(s.end),
    from: normTime_(s.from),
    to: normTime_(s.to),
    step: Number(s.step) === 60 ? 60 : 30,
    size: Math.round(Number(s.size) || 8),
    dur: Math.round(Number(s.dur) || 120),
    memo: str_(s.memo, 200),
  };
  if (!v.title) return { error: 'invalid', message: '제목이 비어 있어요' };
  if (!v.start || !v.end) return { error: 'invalid', message: '날짜 형식이 맞지 않아요' };
  const span = dayNum_(v.end) - dayNum_(v.start);
  if (span < 0 || span > 30) return { error: 'invalid', message: '기간은 1~31일이어야 해요' };
  if (!v.from || !v.to) return { error: 'invalid', message: '시간 형식이 맞지 않아요' };
  if (Number(v.from.slice(3)) % 30 || Number(v.to.slice(3)) % 30) return { error: 'invalid', message: '시간은 30분 단위예요' };
  if (v.size < 1 || v.size > 40) return { error: 'invalid', message: '인원은 1~40명이에요' };
  if (v.dur < 30 || v.dur > 720) v.dur = 120;
  return v;
}

function schedFromRow_(r) {
  return {
    id: r[0], title: r[1], raid: r[2], level: r[3],
    start: normDate_(r[4]), end: normDate_(r[5]), from: normTime_(r[6]), to: normTime_(r[7]),
    step: Number(r[8]) === 60 ? 60 : 30, size: Number(r[9]) || 8, dur: Number(r[10]) || 120,
    memo: r[11], by: r[12], created: r[13], updated: r[14], status: r[15],
  };
}
function schedToRow_(id, v, by, created, updated, status) {
  return [id, cell_(v.title), cell_(v.raid), cell_(v.level), v.start, v.end, v.from, v.to,
    String(v.step), String(v.size), String(v.dur), cell_(v.memo), cell_(by || ''), created, updated, status || ''];
}

// ───────────────────────── 캐시 ─────────────────────────

function cache_() { return CacheService.getScriptCache(); }
function cacheGet_(key) {
  try { const v = cache_().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function cachePut_(key, obj) {
  try { const s = JSON.stringify(obj); if (s.length < 95000) cache_().put(key, s, CACHE_SEC); } catch (e) { /* 캐시는 없어도 됨 */ }
}
function cacheDrop_(keys) {
  try { cache_().removeAll(keys); } catch (e) { /* 무시 */ }
}

// ───────────────────────── 동작 ─────────────────────────

function findSched_(id) {
  if (!id) return null;
  const sh = schedSheet_();
  const rows = rows_(sh, SCHED_HEADER.length);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === id) return { row: i + 2, s: schedFromRow_(rows[i]) };
  }
  return null;
}

function list_() {
  const hit = cacheGet_('list');
  if (hit) return hit;
  const counts = {};
  rows_(memSheet_(), MEM_HEADER.length).forEach(function (r) { counts[r[0]] = (counts[r[0]] || 0) + 1; });
  const list = rows_(schedSheet_(), SCHED_HEADER.length)
    .filter(function (r) { return r[0] && r[15] !== '삭제'; })
    .map(function (r) { const s = schedFromRow_(r); s.count = counts[s.id] || 0; return s; });
  list.sort(function (a, b) { return (b.start + b.created).localeCompare(a.start + a.created); });
  const res = { ok: true, api: API_VERSION, schedules: list.slice(0, 80) };
  cachePut_('list', res);
  return res;
}

function get_(id) {
  const hit = cacheGet_('get:' + id);
  if (hit) return hit;
  const f = findSched_(id);
  if (!f || f.s.status === '삭제') return { error: 'notfound' };
  const members = rows_(memSheet_(), MEM_HEADER.length)
    .filter(function (r) { return r[0] === id && r[1]; })
    .map(function (r) { return { name: r[1], summary: r[2], data: r[3], updated: r[4] }; });
  const res = { ok: true, api: API_VERSION, schedule: f.s, members: members };
  cachePut_('get:' + id, res);
  return res;
}

function create_(req) {
  const v = cleanSchedule_(req.schedule || {});
  if (v.error) return v;
  const id = newId_();
  const t = now_();
  const by = name_(req.name);
  const sh = schedSheet_();
  writeRow_(sh, sh.getLastRow() + 1, schedToRow_(id, v, by, t, t, ''));
  cacheDrop_(['list']);
  const s = { id: id, by: by, created: t, updated: t, status: '' };
  Object.keys(v).forEach(function (k) { s[k] = v[k]; });
  return { ok: true, api: API_VERSION, id: id, schedule: s };
}

function update_(req) {
  const id = str_(req.id, 20);
  const f = findSched_(id);
  if (!f || f.s.status === '삭제') return { error: 'notfound' };
  const v = cleanSchedule_(req.schedule || {});
  if (v.error) return v;
  writeRow_(schedSheet_(), f.row, schedToRow_(id, v, f.s.by, f.s.created, now_(), ''));
  cacheDrop_(['list', 'get:' + id]);
  return get_(id);
}

function delete_(req) {
  const id = str_(req.id, 20);
  const f = findSched_(id);
  if (!f) return { error: 'notfound' };
  schedSheet_().getRange(f.row, 15, 1, 2).setNumberFormat('@').setValues([[now_(), '삭제']]);
  cacheDrop_(['list', 'get:' + id]);
  return { ok: true, api: API_VERSION };
}

// 내 가능 시간 저장 (이름 기준으로 덮어씀). prev 가 있으면 이름 바꾸기
function save_(req) {
  const id = str_(req.id, 20);
  const name = name_(req.name);
  const prev = name_(req.prev);
  if (!name) return { error: 'name' };
  const f = findSched_(id);
  if (!f || f.s.status === '삭제') return { error: 'notfound' };
  const data = String(req.data || '').replace(/[^0-9a-fA-F:,]/g, '').slice(0, 40000);
  const summary = str_(req.summary, 3000);
  const sh = memSheet_();
  const rows = rows_(sh, MEM_HEADER.length);
  let at = -1, prevAt = -1;
  rows.forEach(function (r, i) {
    if (r[0] !== id) return;
    if (sameName_(r[1], name)) at = i;
    if (prev && sameName_(r[1], prev)) prevAt = i;
  });
  if (prev && !sameName_(prev, name) && at >= 0 && prevAt >= 0) return { error: 'taken' };
  if (at < 0) at = prevAt;
  const r = at >= 0 ? at + 2 : sh.getLastRow() + 1;
  writeRow_(sh, r, [id, cell_(name), cell_(summary), data, now_()]);
  cacheDrop_(['list', 'get:' + id]);
  return get_(id);
}

// 내 기록 지우기
function leave_(req) {
  const id = str_(req.id, 20);
  const name = name_(req.name);
  const sh = memSheet_();
  const rows = rows_(sh, MEM_HEADER.length);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][0] === id && sameName_(rows[i][1], name)) sh.deleteRow(i + 2);
  }
  cacheDrop_(['list', 'get:' + id]);
  return get_(id);
}

function handle_(req, isPost) {
  if (TEAM_CODE.indexOf('여기에') === 0 || TEAM_CODE.length < 4) return { error: 'setup' };
  if (String(req.code || '') !== TEAM_CODE) return { error: 'code' };
  const op = String(req.op || 'list');
  if (op === 'ping') return { ok: true, api: API_VERSION };
  if (op === 'list') return list_();
  if (op === 'get') return get_(str_(req.id, 20));
  if (!isPost) return { error: 'op' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (op === 'create') return create_(req);
    if (op === 'update') return update_(req);
    if (op === 'delete') return delete_(req);
    if (op === 'save') return save_(req);
    if (op === 'leave') return leave_(req);
    return { error: 'op' };
  } finally {
    lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    return out_(handle_((e && e.parameter) || {}, false));
  } catch (err) {
    return out_({ error: 'server', message: String((err && err.message) || err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ error: 'bad request' });
  }
  try {
    return out_(handle_(body || {}, true));
  } catch (err) {
    return out_({ error: 'server', message: String((err && err.message) || err) });
  }
}

// 시트 메뉴: [레이드 일정] → [사이트에 바로 반영] (시트를 직접 고쳤을 때)
function onOpen() {
  SpreadsheetApp.getUi().createMenu('레이드 일정').addItem('사이트에 바로 반영', 'clearCache').addToUi();
}

function clearCache() {
  const keys = ['list'];
  rows_(schedSheet_(), 1).forEach(function (r) { if (r[0]) keys.push('get:' + r[0]); });
  cacheDrop_(keys);
}
