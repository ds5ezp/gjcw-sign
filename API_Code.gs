/**
 * 전광판 API (Code.gs) - 전광판 표시 페이지(GitHub)가 콘텐츠 목록과 이미지를 받아 가는 곳
 *
 * 배포: 실행 계정 = 나(viewer@geojecwc.kr), 액세스 = 모든 사용자(익명 포함)
 *   전광판 플레이어는 구글 로그인을 할 수 없으므로 '모든 사용자' 로 배포하고,
 *   대신 화면마다 정해진 접속 키(전광판 관리 시트 '화면' 탭)가 맞아야만 내용을 돌려줌
 *
 * 요청
 *   ?fn=list&screen=lobby&key=XXXX   → { ok, now, w, h, items: [...] }
 *   ?fn=image&screen=lobby&key=XXXX&id=파일ID → { ok, data: 'data:image/...;base64,...' }
 *   키가 틀리거나 화면이 없으면 { ok: false } (표시 페이지는 검은 화면)
 *
 * 콘텐츠 등록·수정은 업무포털 '전광판 관리' 메뉴에서 함 (이 프로젝트는 읽기만 함)
 */

// 업무포털 '전광판 관리 > 화면·주소' 탭에 나오는 시트 ID
const SHEET_ID = '여기에_전광판_관리_시트_ID';

const SCREEN_TAB = '화면';
const ITEM_TAB = '콘텐츠';
const LIST_CACHE_SEC = 10;   // 같은 화면 목록은 10초 동안 저장해 두고 씀 (전광판 2대가 20초마다 묻기 때문)

function doGet(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    const screen = findScreen_(String(p.screen || ''), String(p.key || ''));
    if (!screen) result = { ok: false };
    else if (p.fn === 'image') result = imageData_(screen, String(p.id || ''));
    else result = listFor_(screen);
  } catch (err) {
    // 시트를 잠시 못 읽는 경우 등. 표시 페이지는 error 를 받으면 지금 보여주던 내용을 그대로 유지함
    result = { ok: true, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function findScreen_(code, key) {
  if (!code || !key) return null;
  const sheet = ss_().getSheetByName(SCREEN_TAB);
  if (!sheet) return null;
  const hit = sheet.getDataRange().getValues().slice(1).filter(v =>
    String(v[0] || '').trim() === code && String(v[5] || '').trim() === key)[0];
  if (!hit) return null;
  return { code: code, w: Number(hit[2]) || 0, h: Number(hit[3]) || 0 };
}

// 시트 칸 값 → 밀리초 (빈 칸은 0 = 제한 없음). 시간대는 한국
function toMs_(value) {
  if (value instanceof Date) return value.getTime();
  const s = String(value || '').trim();
  if (!s) return 0;
  try {
    return Utilities.parseDate(s, 'Asia/Seoul', 'yyyy-MM-dd HH:mm').getTime();
  } catch (e) {
    return 0;   // 형식이 틀린 칸은 제한 없음으로 봄
  }
}

function parseFiles_(text) {
  const files = {};
  String(text || '').split(',').forEach(part => {
    const i = part.indexOf(':');
    if (i > 0) files[part.slice(0, i).trim()] = part.slice(i + 1).trim();
  });
  return files;
}

// 이 화면에 나갈 콘텐츠 (사용 중이고 아직 끝나지 않은 것, 순서대로)
// 시작 전인 것도 보내 두면 표시 페이지가 시각에 맞춰 알아서 보여줌
function screenItems_(screen) {
  const now = Date.now();
  const sheet = ss_().getSheetByName(ITEM_TAB);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map((v, i) => ({
    id: String(v[0] || '').trim(),
    type: String(v[1] || '').trim() === '이미지' ? 'image' : 'text',
    targets: String(v[3] || '').split(',').map(s => s.trim()),
    order: Number(v[4]) || 0,
    row: i,
    enabled: String(v[5] || '').trim().toUpperCase() !== 'N',
    start: toMs_(v[6]),
    end: toMs_(v[7]),
    sec: Number(v[8]) || 10,
    text: String(v[9] || ''),
    color: String(v[10] || '#ffffff'),
    bg: String(v[11] || '#000000'),
    size: Number(v[12]) || 60,
    scroll: String(v[13] || '').trim() === '흐름',
    file: parseFiles_(v[16])[screen.code] || '',
    v: v[19] instanceof Date ? v[19].getTime() : String(v[19] || '')
  }))
    .filter(it => it.id && it.enabled && it.targets.indexOf(screen.code) !== -1 && (!it.end || it.end > now))
    .filter(it => it.type === 'text' ? it.text.trim() : it.file)
    .sort((a, b) => a.order - b.order || a.row - b.row)
    .map(it => {
      const out = { id: it.id, type: it.type, sec: it.sec, start: it.start, end: it.end, v: it.v };
      if (it.type === 'image') out.file = it.file;
      else { out.text = it.text; out.color = it.color; out.bg = it.bg; out.size = it.size; out.scroll = it.scroll; }
      return out;
    });
}

function listFor_(screen) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'LIST_' + screen.code;
  let items = null;
  const cached = cache.get(cacheKey);
  if (cached) items = JSON.parse(cached);
  else {
    items = screenItems_(screen);
    const text = JSON.stringify(items);
    if (text.length < 90000) cache.put(cacheKey, text, LIST_CACHE_SEC);
  }
  return { ok: true, now: Date.now(), w: screen.w, h: screen.h, items: items };
}

// 이 화면 콘텐츠에 쓰이는 이미지만 돌려줌 (다른 드라이브 파일은 읽지 않음)
function imageData_(screen, fileId) {
  const used = screenItems_(screen).some(it => it.file === fileId);
  if (!fileId || !used) return { ok: true, missing: true };
  const blob = DriveApp.getFileById(fileId).getBlob();
  return { ok: true, data: 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

// 편집기에서 실행해 시트 연결과 권한을 확인 (실행 로그에 화면별 콘텐츠 수가 나옴)
function testApi() {
  const sheet = ss_().getSheetByName(SCREEN_TAB);
  sheet.getDataRange().getValues().slice(1).forEach(v => {
    const screen = { code: String(v[0]), w: Number(v[2]), h: Number(v[3]) };
    Logger.log(v[1] + ' (' + screen.w + '×' + screen.h + '): 콘텐츠 ' + screenItems_(screen).length + '개');
  });
  DriveApp.getRootFolder();   // 드라이브 권한 승인용
}
