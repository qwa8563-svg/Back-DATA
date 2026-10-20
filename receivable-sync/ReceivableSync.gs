// ===========================================================================
// 주간회의록 시트 → 데이터통합 시트 [M] 채권 / [D] 채권  "부품 채권" 자동 동기화 (독립 Apps Script 프로젝트용)
//
//  * 포털(Code.gs) 프로젝트와는 별개입니다. 이 파일 하나만 새 Apps Script 프로젝트에 붙여넣으세요.
//  * 원본: 주간회의록 시트의 '7-2. 메디칼 채권 미회수 상황'(→ [M] 채권), '7-1. 덴탈 채권 미회수 상황'(→ [D] 채권)
//  * 조건: N열(모델 및 수량)에 "부품", "SPARE PART(S)"가 들어 있는 행 (대소문자·띄어쓰기 무시)
//         통화 기호($, €, ￥)가 없는 행(덴탈의 천원 단위 KRW 행)은 제외합니다.
//  * 중복 판단(포털의 [M]/[D] 채권과 같은 규칙): 채널구분 + 법인/거래처 + 국가 + 청구금액 + 발송일
//    - 같은 건이 이미 있고 값이 다르면 → 그 행의 입금액·잔액·입금예정일·연체여부·비고 등을 최신 값으로 갱신
//      (변경 전 내용은 '_변경이력' 탭에 포털과 같은 형식으로 남김)
//    - 같은 건이 없으면 → 새 행으로 추가 (적용환율·원화환산 같은 수식 열은 기존 수식을 복사)
//    - 원본에서 사라진 행은 지우지 않습니다.
//  * 사용법
//    1) 아래 TARGET_SPREADSHEET_ID / SOURCE_SPREADSHEET_ID 확인 (이미 들어 있음)
//    2) syncPartsReceivables 를 편집기에서 한 번 실행 → 권한 승인 → '_채권동기화로그' 탭에서 결과 확인
//    3) installReceivableSyncTrigger 를 한 번 실행 → 매일 아침 SYNC_HOUR 시경 자동 실행
//  * 이 프로젝트를 만든 계정이 통합시트 "편집" 권한과 주간회의록 시트 "보기" 권한을 가지고 있어야 합니다.
// ===========================================================================

const TARGET_SPREADSHEET_ID = '1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q';   // 데이터통합 시트
const SOURCE_SPREADSHEET_ID = '1X4aYpEk-zJpdU37vPNSBHcnTQVw3x4E09C1K6NxdPDU';   // 주간회의록 시트
const SYNC_HOUR = 8;
const ACTOR = '자동 동기화(주간회의록)';

const SYNC_JOBS = [
  { target: '[M] 채권', sourceGid: 1044151971, label: '7-2. 메디칼 채권 미회수 상황' },
  { target: '[D] 채권', sourceGid: 1745860698, label: '7-1. 덴탈 채권 미회수 상황' },
];

// 반영할 열 (통합시트 헤더 이름). A열 No.와 수식 열(적용환율, 원화환산)은 건드리지 않습니다.
const UPLOAD_HEADERS = ['채널구분', '법인/거래처', '국가', '통화', '청구금액', '입금액', '잔액', '발송일', '입금예정일', '연체여부', '비고'];
// 중복 판단에 쓰는 열 (포털 Code.gs의 DEDUPE_KEYS와 같은 값)
const KEY_HEADERS = ['채널구분', '법인/거래처', '국가', '청구금액', '발송일'];

const HISTORY_LOG = { name: '_변경이력', headers: ['변경시각', '작업자', '탭', '시트 행번호', '구분', '변경 내용', '변경 전 행(JSON)'] };
const SYNC_LOG = { name: '_채권동기화로그', headers: ['실행시각', '대상 탭', '원본 읽은 행', '부품 채권 행', '결과'] };

// ---------------------------------------------------------------------------
// 실행 진입점
// ---------------------------------------------------------------------------
function syncPartsReceivables() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('이전 동기화가 아직 실행 중입니다. 잠시 후 다시 시도하세요.');
  try {
    const target = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const tz = target.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Seoul';
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    let source = null, sourceError = '';
    try {
      source = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    } catch (e) {
      sourceError = '주간회의록 시트를 열 수 없습니다. 이 프로젝트를 실행하는 계정에 해당 시트의 보기 권한이 있는지 확인하세요. (' + e.message + ')';
    }
    const logSheet = rsGetOrCreateLog_(target, SYNC_LOG);
    const results = [];
    SYNC_JOBS.forEach(job => {
      let row;
      try {
        if (!source) throw new Error(sourceError);
        const r = syncOne_(target, source, job, tz, today);
        row = [new Date(), job.target, r.read, r.parts, r.message];
        results.push(`${job.target}: ${r.message}`);
      } catch (e) {
        row = [new Date(), job.target, '', '', '오류: ' + e.message];
        results.push(`${job.target}: 오류 - ${e.message}`);
      }
      rsAppendLogRows_(logSheet, [row]);
    });
    const summary = results.join('\n');
    Logger.log(summary);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function syncOne_(target, source, job, tz, today) {
  const sh = source.getSheets().filter(s => s.getSheetId() === job.sourceGid)[0];
  if (!sh) throw new Error(`원본 탭(gid ${job.sourceGid}, '${job.label}')을 찾지 못했습니다. 탭이 삭제됐거나 주소가 바뀌었는지 확인하세요.`);
  const values = sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 15)).getValues();
  const rows = parseReceivableRows_(values, tz, today);
  if (!rows.length) return { read: values.length, parts: 0, message: '부품 채권 행이 없어 반영할 내용이 없습니다.' };
  const msg = upsertReceivables_(target, job.target, rows, tz);
  return { read: values.length, parts: rows.length, message: msg };
}

// 매일 아침 SYNC_HOUR시경에 syncPartsReceivables를 실행하는 트리거를 하나만 만듭니다. (이미 있으면 다시 만들지 않음)
function installReceivableSyncTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'syncPartsReceivables');
  if (exists) return '이미 자동 동기화 트리거가 설치되어 있습니다.';
  ScriptApp.newTrigger('syncPartsReceivables').timeBased().everyDays(1).atHour(SYNC_HOUR).create();
  return `매일 ${SYNC_HOUR}시 경에 자동 동기화하도록 설치했습니다.`;
}

function removeReceivableSyncTrigger() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncPartsReceivables') { ScriptApp.deleteTrigger(t); n++; }
  });
  return `자동 동기화 트리거 ${n}개를 삭제했습니다.`;
}

// ---------------------------------------------------------------------------
// 원본 표 읽기 (시트를 수정하지 않는 순수 함수)
//  열: C 법인코드/구분, D 거래처, E 국가, F 통화기호, G 금액, I 입금액, K 잔액, L 선적(출하)일, M 입금예정일, N 모델 및 수량, O 내용
//  표 구조: 법인코드(C)·거래처(D)·국가(E)는 블록의 첫 줄에만 적혀 있어서 아래 줄들은 위 값을 이어받습니다.
// ---------------------------------------------------------------------------
function recvText_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

// "26,218" / "22,568.65" / 1200 / "$ 1,000" → 숫자. 비어 있거나 숫자가 아니면 null
function recvNum_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = recvText_(v).replace(/[,\s$€￥¥₩]/g, '');
  if (!s || !/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Number(s);
}

// Date 객체 또는 "26-1-21", "25/5/26", "2026-04-01" → 'yyyy-MM-dd'. 'installment' 등 날짜가 아니면 ''
function recvDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = recvText_(v).match(/^(\d{2}|\d{4})\s*[-\/.]\s*(\d{1,2})\s*[-\/.]\s*(\d{1,2})\.?$/);
  if (!m) return '';
  const y = m[1].length === 2 ? 2000 + Number(m[1]) : Number(m[1]);
  const mo = Number(m[2]), d = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return `${y}-${('0' + mo).slice(-2)}-${('0' + d).slice(-2)}`;
}

function isPartsText_(s) {
  const t = recvText_(s).toLowerCase().replace(/\s+/g, '');
  return t.indexOf('부품') !== -1 || t.indexOf('sparepart') !== -1;
}

const CURRENCY_SYMBOLS = ['$', '€', '￥', '¥', '₩'];

// 제품과 부품이 한 청구서에 같이 나간 건은 "부품" 뒤에 적힌 금액만 우리 채권으로 잡습니다.
//  예) "P2 50대 $52,250, + D테크부품 $1,505.75" → 1505.75 / "P3D 2대, P+7대외 Spare Part(1,140 USD)" → 1140
//  부품 키워드 바로 뒤(6자 이내)에 통화 기호($ € ￥ ¥ ₩) 또는 단위(USD EUR JPY KRW 원 달러 유로 엔)가 붙은 금액이 있거나,
//  "부품(9,830)"처럼 괄호 안에 금액만 적힌 경우에만 인정합니다. ("부품 2대" 같은 수량 표기는 무시)
function extractPartsAmount_(text) {
  const t = recvText_(text);
  const kw = /부품|spare\s*parts?/ig;
  let m, found = null;
  while ((m = kw.exec(t)) !== null) {
    const rest = t.slice(m.index + m[0].length);
    const a = rest.match(/^[^0-9$€￥¥₩]{0,6}?([$€￥¥₩])?\s*(\d[\d,]*(?:\.\d+)?)\s*(USD|EUR|JPY|KRW|원|달러|유로|엔)?/i);
    let n = null;
    if (a && (a[1] || a[3])) n = Number(a[2].replace(/,/g, ''));
    else {
      const b = rest.match(/^\s*\(\s*(\d[\d,]*(?:\.\d+)?)\s*\)/);     // "부품(9,830)" — 통화 표시 없이 괄호 안에 금액만 있는 경우 (통화는 행의 통화)
      if (b) n = Number(b[1].replace(/,/g, ''));
    }
    if (n !== null && isFinite(n) && n > 0) found = n;
  }
  return found;
}

// 반환: UPLOAD_HEADERS 순서의 행 배열 (+ 통화 기호가 있는 데이터 행 수는 배열의 scanned 속성)
function parseReceivableRows_(values, tz, today) {
  const C = 2, D = 3, E = 4, F = 5, G = 6, I = 8, K = 10, L = 11, M = 12, N = 13, O = 14;
  const hasLayout = values.some(r => recvText_(r[D]) === '거래처' && recvText_(r[N]).replace(/\s+/g, '') === '모델및수량');
  if (!hasLayout) {
    throw new Error("원본 시트에서 헤더(D열 '거래처' / N열 '모델 및 수량')를 찾지 못했습니다. 원본 표의 열 위치가 바뀐 것 같습니다.");
  }
  let channel = '', client = '', country = '';
  let scanned = 0;
  const out = [];
  values.forEach(r => {
    const c = recvText_(r[C]), d = recvText_(r[D]), e = recvText_(r[E]), f = recvText_(r[F]);
    const b = recvText_(r[1]);
    if (/total/i.test(b) || /total/i.test(d) || /^(USD|EUR|JPY|KRW|CNY)$/i.test(f)) return;   // 소계/총계 줄
    if (c && c !== '법인' && c !== '구분') channel = c;
    if (d === '거래처') return;                                                                 // 반복되는 헤더 줄
    if (d) { client = d; country = e; } else if (e) { country = e; }
    if (CURRENCY_SYMBOLS.indexOf(f) === -1) return;                                             // 통화 기호 없는 행(천원 단위 KRW 등)은 제외
    const claim = recvNum_(r[G]);
    if (claim === null) return;
    scanned++;
    if (!isPartsText_(r[N])) return;
    let paid = recvNum_(r[I]) || 0;
    let bal = recvNum_(r[K]);
    if (bal === null) bal = claim - paid;
    const ship = recvDate_(r[L], tz), due = recvDate_(r[M], tz);
    const noteParts = [recvText_(r[N]).replace(/\s*\n\s*/g, ' '), recvText_(r[O]).replace(/\s*\n\s*/g, ' ')].filter(x => x);
    // 제품과 함께 나간 건: 부품 금액만 반영. 입금액/잔액은 청구금액 대비 비율로 나눕니다. (전액 입금/전액 미입금이면 그대로 부품 금액)
    let claimOut = claim;
    const partsAmt = extractPartsAmount_(r[N]);
    if (partsAmt !== null && partsAmt < claim) {
      const origClaim = claim;
      claimOut = partsAmt;
      if (bal <= 0) { paid = partsAmt; bal = 0; }
      else if (paid <= 0) { paid = 0; bal = partsAmt; }
      else { paid = Math.round(paid * partsAmt / origClaim * 100) / 100; bal = Math.round((partsAmt - paid) * 100) / 100; }
      noteParts.push(`부품분만 반영(원 청구 ${origClaim.toLocaleString('en-US')} 중)`);
    }
    out.push([channel, client, country, f, claimOut, paid, bal, ship, due, (bal > 0 && due && due < today) ? '연체' : '정상', noteParts.join(' / ')]);
  });
  out.scanned = scanned;
  return out;
}

// ---------------------------------------------------------------------------
// 통합시트 반영: 신규는 추가, 같은 건(키 일치)은 값이 다르면 갱신, 같으면 건너뜀
// ---------------------------------------------------------------------------
function normHeader_(s) {
  return (s === null || s === undefined) ? '' : String(s).trim().replace(/\s+/g, '');
}

function headerIndex_(headers, name) {
  const n = normHeader_(name);
  for (let i = 0; i < headers.length; i++) if (normHeader_(headers[i]) === n) return i;
  return -1;
}

let tzCache_ = null;
function normalizeKeyPart_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tzCache_ || 'Asia/Seoul', 'yyyy-MM-dd');
  return (v === null || v === undefined) ? '' : String(v).trim();
}

// 비교용 정규화: 날짜는 yyyy-MM-dd, "1,200,000"과 1200000은 같은 값으로 봅니다(앞자리 0이 있는 코드류는 제외).
function cmpNorm_(v) {
  const s = normalizeKeyPart_(v);
  const dm = s.match(/^(\d{4})\s*[.\-\/]\s*(\d{1,2})\s*[.\-\/]\s*(\d{1,2})\.?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);
  if (dm) {
    const mm = Number(dm[2]), dd = Number(dm[3]);
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) return `${dm[1]}-${('0' + mm).slice(-2)}-${('0' + dd).slice(-2)}`;
  }
  if (/^-?\d+(\.\d+)?$/.test(s) || /^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
    if (!/^-?0\d/.test(s)) {
      const n = Number(s.replace(/,/g, ''));
      if (!isNaN(n)) return String(n);
    }
  }
  return s;
}

function contiguousRuns_(idxs) {
  const runs = [];
  let s = null, p = null;
  idxs.forEach(i => {
    if (s === null) { s = i; p = i; }
    else if (i === p + 1) { p = i; }
    else { runs.push([s, p]); s = i; p = i; }
  });
  if (s !== null) runs.push([s, p]);
  return runs;
}

function plainRow_(row, tz) {
  return row.map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v);
}

function shortText_(v, max) {
  const s = normalizeKeyPart_(v) || '(빈값)';
  return s.length > max ? s.slice(0, max) + '…' : s;
}

function diffSummary_(headers, oldRow, newRow, idxs) {
  const LIMIT = 6;
  const parts = [];
  idxs.forEach(i => {
    if (cmpNorm_(oldRow[i]) !== cmpNorm_(newRow[i])) {
      parts.push(`${headers[i] || ('열' + (i + 1))}: ${shortText_(oldRow[i], 30)} → ${shortText_(newRow[i], 30)}`);
    }
  });
  return parts.slice(0, LIMIT).join('; ') + (parts.length > LIMIT ? ` 외 ${parts.length - LIMIT}개` : '');
}

// 수식 열(No., 적용환율, 원화환산 등)과 데이터 끝 행을 알아냅니다.
//  - 수식 열: 첫 데이터 행(수식 템플릿 행)과 맨 아래 5개 행에서 수식이 걸린 열
//  - 데이터 끝 행: 수식 열을 뺀 칸에 값이 하나라도 있는 마지막 행 (수식만 채워 둔 빈 행은 데이터로 보지 않음)
function layoutInfo_(sheet, headerRow, width, allValues) {
  const firstDataRow = headerRow + 1;
  const lastRow = sheet.getLastRow();
  const info = { firstDataRow, formulaCols: {}, srcFormulas: {}, dataEnd: firstDataRow - 1 };
  if (lastRow < firstDataRow) return info;
  const addFormulas = (fromRow, toRow) => {
    sheet.getRange(fromRow, 1, toRow - fromRow + 1, width).getFormulasR1C1()
      .forEach(row => row.forEach((f, c) => { if (f) { info.formulaCols[c] = true; info.srcFormulas[c] = f; } }));
  };
  const scanFrom = Math.max(firstDataRow, lastRow - 4);
  if (scanFrom > firstDataRow) addFormulas(firstDataRow, firstDataRow);
  addFormulas(scanFrom, lastRow);
  for (let r = Math.min(allValues.length, lastRow); r >= firstDataRow; r--) {
    const row = allValues[r - 1];
    if (row.some((v, c) => c < width && !info.formulaCols[c] && v !== '' && v !== null)) { info.dataEnd = r; break; }
  }
  return info;
}

function upsertReceivables_(ss, tabName, rows, tz) {
  tzCache_ = tz;
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error(`'${tabName}' 탭을 찾을 수 없습니다.`);
  const existing = sheet.getDataRange().getValues();

  // 헤더 행: 키 열 이름이 전부 있는 첫 행(상위 5행 이내)
  let headerIdx = -1;
  for (let r = 0; r < Math.min(5, existing.length); r++) {
    if (KEY_HEADERS.every(n => headerIndex_(existing[r], n) !== -1)) { headerIdx = r; break; }
  }
  if (headerIdx === -1) throw new Error(`'${tabName}' 탭 상위 5행에서 헤더(${KEY_HEADERS.join(', ')})를 찾지 못했습니다.`);
  const headers = existing[headerIdx];
  const width = headers.length;
  const colOf = {};
  const missing = [];
  UPLOAD_HEADERS.forEach(h => { const i = headerIndex_(headers, h); if (i === -1) missing.push(h); else colOf[h] = i; });
  if (missing.length) throw new Error(`'${tabName}' 탭 헤더에서 [${missing.join(', ')}] 열을 찾지 못했습니다.`);

  const layout = layoutInfo_(sheet, headerIdx + 1, width, existing);
  const formulaKey = KEY_HEADERS.filter(h => layout.formulaCols[colOf[h] !== undefined ? colOf[h] : headerIndex_(headers, h)]);
  if (formulaKey.length) throw new Error(`'${tabName}' 탭의 중복 판단 열 [${formulaKey.join(', ')}]이(가) 수식 열입니다. 키로 쓸 수 없습니다.`);

  // 이번에 값을 채우는 열 = 반영 대상 열 중 수식이 아닌 열
  const providedIdx = UPLOAD_HEADERS.map(h => colOf[h]).filter(i => !layout.formulaCols[i]).sort((a, b) => a - b);
  const keyIdx = KEY_HEADERS.map(h => headerIndex_(headers, h));
  const buildKey = r => keyIdx.map(i => cmpNorm_(r[i])).join('|');

  // 기존 시트: 키 → 행 번호(1-based) 목록
  const keyMap = {};
  for (let i = headerIdx + 1; i < existing.length; i++) {
    const k = buildKey(existing[i]);
    if (k.replace(/\|/g, '')) (keyMap[k] = keyMap[k] || []).push(i + 1);
  }

  const toAppend = [], appendByKey = {};
  const toUpdate = [], updateByKey = {};
  let identical = 0, skippedKey = 0;
  rows.forEach(src => {
    const aligned = new Array(width).fill('');
    UPLOAD_HEADERS.forEach((h, j) => { if (!layout.formulaCols[colOf[h]]) aligned[colOf[h]] = src[j]; });
    const filledKeys = keyIdx.filter(i => normalizeKeyPart_(aligned[i]) !== '').length;
    if (filledKeys < Math.ceil(keyIdx.length / 2)) { skippedKey++; return; }   // 키가 절반 미만만 채워진 행은 제외
    const key = buildKey(aligned);
    const same = r => providedIdx.every(i => cmpNorm_(existing[r - 1][i]) === cmpNorm_(aligned[i]));
    const hit = keyMap[key];
    if (hit) {
      if (hit.some(same)) { identical++; return; }
      if (updateByKey[key] !== undefined) toUpdate[updateByKey[key]].data = aligned;       // 원본 안에서 키가 같은 행은 마지막 값 사용
      else { updateByKey[key] = toUpdate.length; toUpdate.push({ row: hit[hit.length - 1], data: aligned }); }
      return;
    }
    if (appendByKey[key] !== undefined) { toAppend[appendByKey[key]] = aligned; return; }
    appendByKey[key] = toAppend.length;
    toAppend.push(aligned);
  });

  // 변경 이력(포털과 같은 형식)을 먼저 남깁니다. 여기서 실패하면 본문은 아직 그대로입니다.
  if (toUpdate.length) {
    const now = new Date();
    rsAppendLogRows_(rsGetOrCreateLog_(ss, HISTORY_LOG), toUpdate.map(u => [
      now, ACTOR, tabName, u.row, '업로드 덮어쓰기',
      diffSummary_(headers, existing[u.row - 1], u.data, providedIdx),
      JSON.stringify(plainRow_(existing[u.row - 1], tz)),
    ]));
  }
  toUpdate.forEach(u => {
    contiguousRuns_(providedIdx).forEach(run => {
      sheet.getRange(u.row, run[0] + 1, 1, run[1] - run[0] + 1).setValues([u.data.slice(run[0], run[1] + 1)]);
    });
  });
  if (toAppend.length) appendRowsKeepingFormulas_(sheet, layout, toAppend, width);
  SpreadsheetApp.flush();

  return `신규 ${toAppend.length}건 추가, 기존 ${toUpdate.length}건 갱신, 이미 동일 ${identical}건`
    + (skippedKey ? `, 키 값 부족으로 제외 ${skippedKey}건` : '');
}

// 신규 행 추가: 수식 열은 값을 쓰지 않고 맨 아래 행의 수식을 상대참조로 복사해 채웁니다.
function appendRowsKeepingFormulas_(sheet, info, rows, width) {
  const startRow = info.dataEnd + 1;
  const n = rows.length;
  // 마지막 안전장치: 값을 쓸 칸에 이미 수식이 들어 있으면 그 열은 값으로 덮어쓰지 않습니다.
  const targetEnd = Math.min(sheet.getLastRow(), startRow + n - 1);
  if (targetEnd >= startRow) {
    sheet.getRange(startRow, 1, targetEnd - startRow + 1, width).getFormulasR1C1()
      .forEach(row => row.forEach((f, c) => { if (f) { info.formulaCols[c] = true; if (!info.srcFormulas[c]) info.srcFormulas[c] = f; } }));
  }
  const valueCols = [];
  for (let c = 0; c < width; c++) if (!info.formulaCols[c]) valueCols.push(c);
  contiguousRuns_(valueCols).forEach(run => {
    sheet.getRange(startRow, run[0] + 1, n, run[1] - run[0] + 1).setValues(rows.map(r => r.slice(run[0], run[1] + 1)));
  });
  Object.keys(info.formulaCols).forEach(c => {
    if (Number(c) < width) sheet.getRange(startRow, Number(c) + 1, n, 1).setFormulaR1C1(info.srcFormulas[c]);
  });
}

// ---------------------------------------------------------------------------
// 로그 탭
// ---------------------------------------------------------------------------
function rsGetOrCreateLog_(ss, def) {
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rsAppendLogRows_(sh, rows) {
  if (!rows.length) return;
  const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  const padded = rows.map(r => r.concat(new Array(width - r.length).fill('')));
  sh.getRange(sh.getLastRow() + 1, 1, padded.length, width).setValues(padded);
}
