// 데이터가 들어있는 구글 시트 ID
const SPREADSHEET_ID = '1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('AS/매출 통합 관리 포털')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------------
// 1. 7개 탭 수동 입력 저장
//    - 서버에서 입력값을 검증하고(필수/날짜/숫자/날짜 순서/드롭다운 허용값),
//    - 같은 건이 이미 있으면 저장하지 않고 "중복" 결과를 돌려줍니다. (화면에서 덮어쓰기/새 행 추가/취소 선택)
// 규칙의 키(d1, d2 ...)는 화면 입력칸 id(t1_d1 → d1)의 번호입니다.
// ---------------------------------------------------------------------------
const MD_BOND_RULES_ = {
  labels: { d2: '법인/거래처', d5: '청구금액', d6: '입금액', d7: '발송일', d8: '입금예정일' },
  required: ['d2', 'd7'], dates: ['d7', 'd8'], numbers: ['d5', 'd6'], order: [['d7', 'd8']],
};
const MANUAL_RULES = {
  '[K] AS/매출': {
    labels: { d1: '접수일', d2: '조치일', d6: '제조번호/코드', d9: '증상/내용', d13: '수리비용' },
    required: ['d1', 'd6', 'd9'], dates: ['d1', 'd2'], numbers: ['d13'], order: [['d1', 'd2']],
  },
  '[K] 유지보수': {
    labels: { d1: '병원명', d3: 'S/N', d4: '장비 납품일', d5: '유지보수 계약일', d6: '계약 만료일',
              d9: '유지보수 금액(월납)', d11: '유효 계약 수', d12: '유지보수 합계(원)' },
    required: ['d1', 'd3'], dates: ['d4', 'd5', 'd6'], numbers: ['d9', 'd11', 'd12'], order: [['d5', 'd6']],
  },
  '[K] 채권': {
    labels: { d2: '거래처', d3: '청구금액', d4: '입금액', d5: '발생일', d6: '회수일자' },
    required: ['d2', 'd5'], dates: ['d5', 'd6'], numbers: ['d3', 'd4'], order: [['d5', 'd6']],
  },
  '[M/D] AS': {
    labels: { d1: '팀', d3: '접수일', d4: '조치일', d9: '제조번호/코드', d12: '증상/내용' },
    required: ['d1', 'd3', 'd9', 'd12'], dates: ['d3', 'd4'], numbers: [], order: [['d3', 'd4']],
  },
  '[M/D] 매출': {
    labels: { d2: '발송일', d7: '제조번호/코드' },
    required: ['d2', 'd7'], dates: ['d2'], numbers: [], order: [],
  },
  '[M] 채권': MD_BOND_RULES_,
  '[D] 채권': MD_BOND_RULES_,
};

function isValidIsoDate_(str) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str)) return false;
  const p = str.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
  return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
}

// 입력값을 다듬고(공백 제거, 숫자 변환) 문제가 있으면 한꺼번에 모아 예외로 알립니다.
function validateManualInput_(data, rules) {
  const clean = {};
  Object.keys(data || {}).forEach(k => {
    const v = data[k];
    clean[k] = (typeof v === 'string') ? v.trim() : (v === null || v === undefined ? '' : v);
  });
  rules.required.concat(rules.dates, rules.numbers).forEach(k => { if (clean[k] === undefined) clean[k] = ''; });
  const label = k => rules.labels[k] || k;
  const errors = [];

  rules.required.forEach(k => {
    if (clean[k] === '') errors.push(`${label(k)}은(는) 필수 입력입니다.`);
  });
  rules.dates.forEach(k => {
    if (clean[k] !== '' && !isValidIsoDate_(String(clean[k]))) {
      errors.push(`${label(k)}: 올바른 날짜(yyyy-mm-dd)가 아닙니다. (입력값: ${clean[k]})`);
    }
  });
  rules.numbers.forEach(k => {
    if (clean[k] === '') return;
    const n = Number(clean[k]);
    if (!isFinite(n)) errors.push(`${label(k)}: 숫자가 아닙니다. (입력값: ${clean[k]})`);
    else if (n < 0) errors.push(`${label(k)}: 0 이상이어야 합니다.`);
    else clean[k] = n;
  });
  rules.order.forEach(pair => {
    const a = clean[pair[0]], b = clean[pair[1]];
    if (a !== '' && b !== '' && isValidIsoDate_(String(a)) && isValidIsoDate_(String(b)) && String(b) < String(a)) {
      errors.push(`${label(pair[1])}은(는) ${label(pair[0])}보다 빠를 수 없습니다.`);
    }
  });

  if (errors.length) throw new Error('입력값을 확인해주세요:\n- ' + errors.join('\n- '));
  return clean;
}

// 탭별 저장용 행(A열 No.는 수식 유지를 위해 비워둠)
function buildManualRow_(targetSheetName, data) {
  switch (targetSheetName) {
    case '[K] AS/매출':
      return ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11, data.d12, data.d13, data.d14];
    case '[K] 유지보수':
      // 시트에는 '유지보수 금액(월납)'(J열)과 '월'(L열) 사이에 빈 열(K열)이 있어서 한 칸을 비워 둡니다.
      return ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, "", data.d10, data.d11, data.d12, data.d13];
    case '[K] 채권':
      // 열 순서: 장비구분, 거래처, 청구금액, 입금액, 잔액(자동계산), 발생일, 회수일자, 연체여부, 비고
      return ["", data.d1, data.d2, data.d3, data.d4, (Number(data.d3) || 0) - (Number(data.d4) || 0), data.d5, data.d6, data.d7, data.d8];
    case '[M/D] AS':
      return ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11, data.d12, data.d13, data.d14, data.d15];
    case '[M/D] 매출':
      // 시트에서 '증상/내용'은 K:L 두 칸이 병합되어 있고 '통화'는 M열이라서, L열은 비워 둡니다.
      return ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, "", data.d11];
    case '[M] 채권':
    case '[D] 채권':
      // 잔액 = 청구금액(d5) - 입금액(d6). (d4는 통화 기호라서 계산에 쓰면 안 됩니다)
      return ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, (Number(data.d5) || 0) - (Number(data.d6) || 0), data.d7, data.d8, data.d9, data.d10];
    default:
      throw new Error("알 수 없는 양식입니다.");
  }
}

// 키 컬럼명이 전부 등장하는 첫 행(상위 5행 이내)을 헤더 행으로 봅니다. 없으면 -1.
function detectHeaderRow_(data, targetSheetName) {
  const def = DEDUPE_KEYS[targetSheetName];
  const names = Array.isArray(def) ? def : (def && def.exclude ? def.exclude : []);
  for (let r = 0; r < Math.min(5, data.length); r++) {
    if (names.every(n => headerIndex_(data[r], n) !== -1)) return r;
  }
  return -1;
}

// mode: 생략 = 검사 후 저장 / 'overwrite' = 같은 건이 있으면 그 행을 덮어씀 / 'append' = 중복 검사 없이 새 행으로 추가
// 반환: {status: 'saved' | 'duplicate' | 'identical', message, ...}
function saveMultiData(data, targetSheetName, mode) {
  const rules = MANUAL_RULES[targetSheetName];
  if (!rules) throw new Error("알 수 없는 양식입니다.");
  const clean = validateManualInput_(data, rules);

  return withLock_(function() {
    const ss = openSpreadsheet_();
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

    const rowData = buildManualRow_(targetSheetName, clean).map(v => v === undefined ? '' : v);
    const existingData = sheet.getDataRange().getValues();
    const headerRowIdx = detectHeaderRow_(existingData, targetSheetName);
    if (headerRowIdx === -1) {
      throw new Error(`'${targetSheetName}' 탭에서 중복 검사용 헤더를 찾지 못했습니다. 시트 헤더 이름과 Code.gs의 DEDUPE_KEYS를 확인하세요.`);
    }
    const headers = existingData[headerRowIdx];
    const colLabel = i => (headers[i] || `${i + 1}번째 열`);

    // 드롭다운(목록) 열은 허용값 확인 — 저장 시 알 수 없는 시트 오류가 나는 대신 어느 항목이 문제인지 알려줍니다.
    {
      const lastRow = sheet.getLastRow();
      const errors = [];
      for (let i = 1; i < rowData.length; i++) {
        const v = rowData[i];
        if (v === '' || lastRow <= headerRowIdx + 1) continue;
        const allowed = getAllowedValuesForColumn_(sheet, i + 1, lastRow);
        if (allowed && allowed.indexOf(v) === -1) errors.push(`${colLabel(i)}: '${v}'은(는) 선택 가능한 목록에 없습니다.`);
      }
      if (errors.length) throw new Error('입력값을 확인해주세요:\n- ' + errors.join('\n- '));
    }

    // 값이 입력된 열(A열 제외)만 비교/덮어쓰기 대상으로 봅니다. (비워 둔 칸이 기존 값을 지우지 않도록)
    const filledIdx = [];
    for (let i = 1; i < rowData.length; i++) if (rowData[i] !== '') filledIdx.push(i);

    if (mode !== 'append') {
      const keyIdx = resolveDedupeIdx_(targetSheetName, headers, true);
      if (keyIdx) {
        const buildKey = r => keyIdx.map(i => cmpNorm_(r[i])).join('|');
        const key = buildKey(rowData);
        const matches = [];
        if (key.replace(/\|/g, '')) {
          for (let i = headerRowIdx + 1; i < existingData.length; i++) {
            if (buildKey(existingData[i]) === key) matches.push(i + 1);
          }
        }
        if (matches.length > 0) {
          const diffOf = r => filledIdx.filter(i => cmpNorm_(existingData[r - 1][i]) !== cmpNorm_(rowData[i]));
          const sameRow = matches.find(r => diffOf(r).length === 0);
          if (sameRow) {
            return { status: 'identical', message: `${sameRow}행에 이미 같은 내용이 등록되어 있어 저장하지 않았습니다.` };
          }
          const target = matches[matches.length - 1];
          const diff = diffSummary_(headers, existingData[target - 1], rowData, diffOf(target));
          if (mode !== 'overwrite') {
            return {
              status: 'duplicate', row: target, matchCount: matches.length, diff: diff,
              message: `같은 건이 이미 ${target}행에 있습니다.`,
            };
          }
          // 덮어쓰기: 변경 전 행을 _변경이력에 남긴 뒤, 입력한 열만 갱신합니다.
          appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.history), [[
            new Date(), currentUser_(), targetSheetName, target, '수동 입력 덮어쓰기', diff,
            JSON.stringify(plainRow_(existingData[target - 1])),
          ]]);
          writeProvidedCols_(sheet, target, rowData, filledIdx);
          SpreadsheetApp.flush();
          return { status: 'saved', message: `${targetSheetName} ${target}행을 덮어썼습니다. (변경 전 내용은 '${LOG_SHEETS.history.name}' 탭에 기록)` };
        }
      }
    }

    appendRows_(sheet, [rowData], rowData.length);
    SpreadsheetApp.flush();
    return { status: 'saved', message: `${targetSheetName} 시트에 성공적으로 저장되었습니다!` };
  });
}

// 탭별 "중복 판단 기준" 컬럼명. 여기 정의된 탭만 중복 체크 + 업데이트(덮어쓰기)를 하고,
// 정의 안 된 탭은 기존처럼 무조건 새 행으로 추가합니다.
// - 배열: 그 컬럼들만 비교 (예: [K] AS/매출은 금액만 달라도 같은 건으로 보고 업데이트해야 하므로
//   접수일/제조번호/증상 3개만 기준으로 씁니다)
// - '__ALL__': No.(A열)를 제외한 나머지 모든 열 값이 전부 같아야 중복으로 판단합니다
//   (완전히 똑같은 행이 두 번 들어온 경우만 중복 처리)
// - { exclude: [...] }: No.(A열)와 exclude에 적은 컬럼(들)만 빼고 나머지 전부가 같으면 중복으로 판단
// 채권 탭의 키는 "청구서를 식별하는 열"(거래처/청구금액/발생·발송일 등)만 씁니다. 입금이 들어오면 입금액·잔액뿐 아니라
// 회수일자·연체여부·입금예정일·비고도 함께 바뀌는 게 정상이라, 그 열들을 키에 넣으면 갱신이 아니라 새 행으로 쌓입니다.
// (M/D 채권은 통화 표기가 '$'와 'USD'로 섞일 수 있어 키에서 뺐습니다.)
// 한계: 거래처·금액·일자가 모두 같은 별개 청구서는 같은 건으로 취급됩니다. (덮어쓰기 전 내용은 _변경이력에 남습니다)
const DEDUPE_KEYS = {
  '[K] AS/매출': ['접수일', '제조번호/코드', '증상/내용'],
  '[K] 유지보수': ['병원명', 'S/N', '유지보수 계약일'],
  '[K] 채권': ['장비구분', '거래처', '청구금액', '발생일'],
  '[M/D] AS': ['팀', '접수일', '제조번호/코드', '증상/내용'],
  '[M/D] 매출': { exclude: ['수리비용'] },
  '[M] 채권': ['채널구분', '법인/거래처', '국가', '청구금액', '발송일'],
  '[D] 채권': ['채널구분', '법인/거래처', '국가', '청구금액', '발송일'],
};

// 이미 시트에 쌓인 중복을 "삭제"하는 정리 기능(previewDuplicateCleanup/runDuplicateCleanup)용 기준.
// 업로드 매칭(DEDUPE_KEYS)은 틀려도 '_검토대기'로 갈 뿐이지만, 삭제는 데이터가 사라지므로
// [K] 유지보수/[M/D] AS는 "전체 열이 완전히 같은 행"만 중복으로 봅니다. (같은 날 같은 기기의 별개 접수건 보호)
// 채권 3개 탭도 정리(삭제)는 예전의 "입금액/잔액만 빼고 전부 같아야 중복" 기준을 그대로 씁니다.
// 업로드 매칭용 키(청구서를 식별하는 열만 사용)로 삭제하면, 같은 거래처·금액·일자의 별개 청구서가 지워질 수 있기 때문입니다.
const CLEANUP_KEYS = Object.assign({}, DEDUPE_KEYS, {
  '[K] 유지보수': '__ALL__',
  '[M/D] AS': '__ALL__',
  '[K] 채권': { exclude: ['입금액', '잔액'] },
  '[M] 채권': { exclude: ['입금액', '잔액'] },
  '[D] 채권': { exclude: ['입금액', '잔액'] },
});

// 키(위 DEDUPE_KEYS)가 같은데 값이 다른 행이 업로드됐을 때의 처리 방식.
// - 'review'   : 덮어쓰지 않고 '_검토대기' 탭에 기존값/새값을 나란히 보관 → 담당자가 반영/신규추가/무시 선택
//                (같은 날 같은 기기·같은 증상의 별개 접수건이 조용히 사라지는 것을 막기 위함)
// - 'overwrite': 최신 값으로 기존 행을 갱신 (변경 전 내용은 '_변경이력' 탭에 자동 기록)
//                채권의 입금액, [M/D] 매출의 수리비용처럼 값이 바뀌는 게 정상 흐름인 탭
const DUPLICATE_POLICY = {
  '[K] AS/매출': 'review',
  '[K] 유지보수': 'review',
  '[M/D] AS': 'review',
  '[M/D] 매출': 'overwrite',
  '[K] 채권': 'overwrite',
  '[M] 채권': 'overwrite',
  '[D] 채권': 'overwrite',
};

// 날짜(Date 객체)/문자열을 동일한 형식으로 맞춰서 비교 가능하게 만듭니다.
function normalizeKeyPart_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, spreadsheetTz_(), 'yyyy-MM-dd');
  }
  return (v === null || v === undefined) ? '' : String(v).trim();
}

function normHeader_(s) {
  return (s === null || s === undefined) ? '' : String(s).trim().replace(/\s+/g, '');
}

function headerIndex_(headers, name) {
  const n = normHeader_(name);
  for (let i = 0; i < headers.length; i++) if (normHeader_(headers[i]) === n) return i;
  return -1;
}

// 탭의 DEDUPE_KEYS 정의를 시트 헤더 기준 "키 열 인덱스 배열"로 바꿉니다. 키가 없는 탭이면 null.
// 키로 지정한 컬럼명이 시트 헤더에 없으면(오타/헤더 변경) 조용히 중복 검사를 끄지 않고 오류로 알립니다.
function resolveDedupeIdx_(targetSheetName, sheetHeaders, useNameMapping) {
  const def = DEDUPE_KEYS[targetSheetName] || null;
  if (!def) return null;
  const idx = [];
  if (def === '__ALL__') {
    for (let s = 1; s < sheetHeaders.length; s++) idx.push(s);
    return idx;
  }
  if (!useNameMapping) return null;
  if (def.exclude) {
    const excludeSet = new Set(def.exclude.map(normHeader_));
    for (let s = 1; s < sheetHeaders.length; s++) {
      if (!excludeSet.has(normHeader_(sheetHeaders[s]))) idx.push(s);
    }
    return idx;
  }
  const list = def.map(name => headerIndex_(sheetHeaders, name));
  const missing = def.filter((name, k) => list[k] === -1);
  if (missing.length) {
    throw new Error(`'${targetSheetName}' 탭 헤더에서 중복 판단 컬럼 [${missing.join(', ')}]을(를) 찾지 못했습니다. 시트 헤더 이름을 확인하거나 Code.gs의 DEDUPE_KEYS를 수정하세요.`);
  }
  return list;
}

// 특정 열에 "목록에서 선택" 데이터 확인 규칙이 걸려있으면 허용값 배열을, 없으면 null을 반환합니다.
// (기존 데이터가 있는 마지막 행을 기준으로 확인 — 보통 같은 규칙이 열 전체에 적용되어 있습니다.)
function getAllowedValuesForColumn_(sheet, colIndex1Based, checkRow) {
  if (checkRow < 1) return null;
  try {
    const dv = sheet.getRange(checkRow, colIndex1Based).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      return dv.getCriteriaValues()[0];
    }
  } catch (e) {}
  return null;
}

// ---------------------------------------------------------------------------
// 공통 헬퍼: 로그 탭(_변경이력 / _검토대기 / _삭제로그), 락, 값 비교
// ---------------------------------------------------------------------------
const LOG_SHEETS = {
  history: {
    name: '_변경이력',
    headers: ['변경시각', '작업자', '탭', '시트 행번호', '구분', '변경 내용', '변경 전 행(JSON)'],
  },
  review: {
    name: '_검토대기',
    headers: ['등록시각', '작업자', '탭', '기존 행번호', '상태', '변경 내용', '기존값(JSON)', '새값(JSON)', '반영컬럼(JSON)', '처리결과'],
    // 상태 열(E)에 드롭다운을 걸어둡니다. (스크립트가 '완료(반영)' 같은 값을 쓰므로 목록 밖 값도 허용)
    onCreate: function(sh) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['대기', '반영', '신규추가', '무시'], true)
        .setAllowInvalid(true)
        .build();
      sh.getRange(2, 5, 2000, 1).setDataValidation(rule);
    },
  },
  deleted: {
    name: '_삭제로그',
    headers: ['삭제시각', '작업자', '탭', '원래 행번호', '중복 키', '삭제된 행 값 (A열부터 순서대로 →)'],
  },
};

function openSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 시트의 날짜 셀은 "스프레드시트 시간대" 기준 자정입니다. 스크립트 시간대와 다르면 하루씩 밀려 비교가 틀어지므로 시트 기준을 씁니다.
let spreadsheetTzCache_ = null;
function spreadsheetTz_() {
  if (!spreadsheetTzCache_) {
    let tz = null;
    try {
      tz = openSpreadsheet_().getSpreadsheetTimeZone();
    } catch (e) {}
    // 시트 시간대가 비어 있는 경우(값이 null/빈 문자열)에도 항상 문자열을 돌려주도록 단계별로 대체합니다.
    spreadsheetTzCache_ = (typeof tz === 'string' && tz) ? tz : (Session.getScriptTimeZone() || 'Asia/Seoul');
  }
  return spreadsheetTzCache_;
}

function getOrCreateLogSheet_(ss, def) {
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (def.onCreate) def.onCreate(sh);
  }
  return sh;
}

// 길이가 서로 다른 행들도 한 번에 로그 탭 맨 아래에 붙입니다.
function appendLogRows_(sh, rows) {
  if (!rows.length) return;
  const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  const padded = rows.map(r => r.concat(new Array(width - r.length).fill('')));
  sh.getRange(sh.getLastRow() + 1, 1, padded.length, width).setValues(padded);
}

function currentUser_() {
  try {
    return Session.getActiveUser().getEmail() || '(확인 불가)';
  } catch (e) {
    return '(확인 불가)';
  }
}

// 업로드/삭제/검토반영이 동시에 행 번호를 건드리지 않도록 한 번에 하나씩만 실행합니다.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('다른 작업이 진행 중입니다. 잠시 후 다시 시도해주세요.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// 비교용 정규화: 날짜는 yyyy-MM-dd, "1,200,000"과 1200000은 같은 값으로 봅니다(앞자리 0이 있는 코드류는 제외).
function cmpNorm_(v) {
  const s = normalizeKeyPart_(v);
  if (/^-?\d+(\.\d+)?$/.test(s) || /^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
    if (!/^-?0\d/.test(s)) {
      const n = Number(s.replace(/,/g, ''));
      if (!isNaN(n)) return String(n);
    }
  }
  return s;
}

function rowsEqual_(a, b, idxs) {
  return idxs.every(i => cmpNorm_(a[i]) === cmpNorm_(b[i]));
}

// 시트에서 읽은 행을 JSON으로 남기기 좋게 (Date → 문자열) 변환합니다.
function plainRow_(row) {
  return row.map(v => v instanceof Date
    ? Utilities.formatDate(v, spreadsheetTz_(), 'yyyy-MM-dd') : v);
}

function shortText_(v, max) {
  const s = normalizeKeyPart_(v) || '(빈값)';
  return s.length > max ? s.slice(0, max) + '…' : s;
}

// "컬럼명: 기존값 → 새값" 형태의 변경 요약.
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

function describeRow_(row) {
  const vals = [];
  for (let i = 1; i < row.length && vals.length < 4; i++) {
    const s = normalizeKeyPart_(row[i]);
    if (s) vals.push(s.length > 20 ? s.slice(0, 20) + '…' : s);
  }
  return vals.join(' | ');
}

// 오름차순 인덱스 배열을 연속 구간으로 묶습니다. 예: [1,2,3,5] → [[1,3],[5,5]]
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

// 엑셀에 있던(=providedIdx) 열만 덮어씁니다. 엑셀에 없는 열(수식 등)은 그대로 둡니다. A열(No.)은 항상 건드리지 않습니다.
function writeProvidedCols_(sheet, row, data, providedIdx) {
  contiguousRuns_(providedIdx).forEach(run => {
    sheet.getRange(row, run[0] + 1, 1, run[1] - run[0] + 1).setValues([data.slice(run[0], run[1] + 1)]);
  });
}

// 신규 행 일괄 추가 + A열(No.) 수식 복사.
function appendRows_(sheet, rows, width) {
  const startRow = sheet.getLastRow() + 1;
  // A열(No.) 수식은 행마다 개별적으로 걸려있어 새 행엔 복사되어 있지 않으므로,
  // 바로 위 행의 수식을 그대로 복사해 붙여넣습니다(상대참조라 행 번호는 자동으로 맞춰집니다).
  const aboveFormula = startRow > 1 ? sheet.getRange(startRow - 1, 1).getFormulaR1C1() : '';
  sheet.getRange(startRow, 1, rows.length, width).setValues(rows);
  if (aboveFormula) {
    sheet.getRange(startRow, 1, rows.length, 1).setFormulasR1C1(rows.map(() => [aboveFormula]));
  }
}

// ---------------------------------------------------------------------------
// 엑셀 업로드: 분석(analyzeUpload_) → 미리보기(previewExcelUpload) / 실행(uploadExcelData)
// 분석 단계는 시트를 전혀 수정하지 않습니다.
// ---------------------------------------------------------------------------
function analyzeUpload_(newDataArray, targetSheetName) {
  const ss = openSpreadsheet_();
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) {
    throw new Error(`진단 에러: '${targetSheetName}' 탭을 구글 시트에서 찾을 수 없습니다.`);
  }

  const existingData = sheet.getDataRange().getValues();
  if (existingData.length === 0) {
    throw new Error(`진단 에러: '${targetSheetName}' 탭에 1행(제목줄) 데이터가 존재하지 않습니다.`);
  }

  // 엑셀 1행(제목줄)을 시트 헤더와 "이름"으로 매칭합니다.
  // (엑셀 파일에 No./순번 열이 포함되어 있거나 열 순서가 달라도 밀리지 않도록 하기 위함)
  const normalize = s => (s === null || s === undefined) ? '' : String(s).trim().replace(/\s+/g, '');
  const excelHeaders = (newDataArray[0] || []).map(normalize);

  // 시트 1행이 "AS 매출 K테크" 같은 제목줄이고 실제 컬럼명은 2행 이후에 있는 경우가 있으므로,
  // 상위 몇 개 행 중 엑셀 헤더와 가장 많이 일치하는 행을 실제 헤더 행으로 채택합니다.
  const HEADER_SCAN_ROWS = Math.min(5, existingData.length);
  let sheetHeaders = existingData[0];
  let colMap = [];
  let nameMatchCount = 0;
  let headerRowIdx = 0; // existingData 배열 안에서 헤더 행의 0-based 인덱스
  for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
    const candidateHeaders = existingData[r];
    const candidateColMap = new Array(candidateHeaders.length).fill(-1);
    let candidateMatchCount = 0;
    for (let s = 1; s < candidateHeaders.length; s++) {
      const idx = excelHeaders.indexOf(normalize(candidateHeaders[s]));
      candidateColMap[s] = idx;
      if (idx !== -1) candidateMatchCount++;
    }
    if (candidateMatchCount > nameMatchCount) {
      nameMatchCount = candidateMatchCount;
      sheetHeaders = candidateHeaders;
      colMap = candidateColMap;
      headerRowIdx = r;
    }
  }
  // 헤더명이 2개 이상 일치하면 이름 매칭 사용, 그렇지 않으면(제목줄이 없는 단순 엑셀 등) 기존 순서 매칭으로 대체
  const useNameMapping = nameMatchCount >= 2;

  // 이번 업로드가 실제로 값을 채우는 열(A열 제외). 비교/덮어쓰기는 이 열들만 대상으로 합니다.
  const providedIdx = [];
  for (let s = 1; s < sheetHeaders.length; s++) {
    if (!useNameMapping || colMap[s] !== -1) providedIdx.push(s);
  }

  // ---- 드롭다운(목록) 허용값을 미리 읽어둡니다 (열마다 한 번씩만, 기존 방식처럼
  //      "일단 써보고 실패하면 한 줄씩 재시도"하지 않기 위함 — 대용량 업로드에서 시간초과의 원인이었습니다). ----
  const lastRow = sheet.getLastRow();
  const validationCheckRow = lastRow > headerRowIdx + 1 ? lastRow : -1;
  const allowedValuesByCol = new Array(sheetHeaders.length).fill(null);
  if (validationCheckRow > 0) {
    for (let s = 1; s < sheetHeaders.length; s++) {
      allowedValuesByCol[s] = getAllowedValuesForColumn_(sheet, s + 1, validationCheckRow);
    }
  }

  // ---- 중복 체크 준비: 이 탭에 대해 정의된 키 컬럼들의 시트 내 위치를 찾습니다. ----
  const policy = DUPLICATE_POLICY[targetSheetName] || 'overwrite';
  const dedupeIdx = resolveDedupeIdx_(targetSheetName, sheetHeaders, useNameMapping);
  const buildKey_ = (rowArray) => dedupeIdx.map(i => cmpNorm_(rowArray[i])).join('|');

  // 기존 시트 데이터로 "키 → 시트 행 번호(1-based) 목록" 맵을 만듭니다. (같은 키가 여러 행일 수 있음)
  const existingKeyMap = {};
  if (dedupeIdx) {
    for (let i = headerRowIdx + 1; i < existingData.length; i++) {
      const key = buildKey_(existingData[i]);
      if (key.replace(/\|/g, '')) (existingKeyMap[key] = existingKeyMap[key] || []).push(i + 1);
    }
  }

  const rowsToAppend = [];
  const keyToAppendIdx = {};   // overwrite 정책: 파일 내 같은 키는 마지막 값으로 교체
  const newByKey = {};         // review 정책: 파일 내 같은 키라도 값이 다르면 둘 다 추가
  const rowsToUpdate = [];     // {row, data}  (overwrite 정책: 엑셀 열 전체 덮어쓰기)
  const rowsToFill = [];       // {row, data, cols} (review 정책: 기존 행의 빈 칸만 채움 — 기존 값은 그대로)
  const fillByKey = {};
  const keyToUpdateIdx = {};
  const conflicts = [];        // {row, data, excelRow}
  const conflictsByKey = {};
  const invalidRows = [];
  const MAX_REPORTED = 30;
  let blankCount = 0;
  let invalidCount = 0;
  let identicalCount = 0;
  let alreadyPendingCount = 0;

  // 이미 '_검토대기'에 '대기' 상태로 올라가 있는 충돌(같은 탭/행/새값)은 재업로드 때 또 쌓지 않습니다.
  const pendingKeys = new Set();
  const reviewSheet = ss.getSheetByName(LOG_SHEETS.review.name);
  if (reviewSheet && reviewSheet.getLastRow() >= 2) {
    reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, 8).getValues().forEach(r => {
      if (String(r[4]).trim() === '대기') pendingKeys.add(`${r[2]}|${r[3]}|${r[7]}`);
    });
  }

  // 엑셀 데이터 행 순회 (0번째 줄은 엑셀 제목이므로 1번째 줄부터 시작)
  for (let i = 1; i < newDataArray.length; i++) {
    let excelRow = newDataArray[i];
    if (!excelRow || excelRow.length === 0) {
      blankCount++;
      continue;
    }

    // 구글 시트 양식 크기의 빈 배열 생성
    let alignedRow = new Array(sheetHeaders.length).fill("");
    // (A열인 인덱스 0은 구글 시트의 자동 넘버링/수식을 위해 항상 비워둡니다)

    if (useNameMapping) {
      for (let s = 1; s < sheetHeaders.length; s++) {
        const excelIdx = colMap[s];
        if (excelIdx !== -1 && excelIdx < excelRow.length) {
          let val = excelRow[excelIdx];
          // 드롭다운(데이터 확인) 규칙과의 공백 불일치를 막기 위해 문자열은 트림 처리
          alignedRow[s] = (typeof val === 'string') ? val.trim() : val;
        }
      }
    } else {
      // 제목줄이 시트 헤더와 매칭되지 않을 때의 대체 로직: 엑셀 0번째 칸부터 순서대로 B열부터 채움
      for (let j = 0; j < excelRow.length; j++) {
        let targetIdx = j + 1;
        if (targetIdx < sheetHeaders.length) {
          let val = excelRow[j];
          alignedRow[targetIdx] = (typeof val === 'string') ? val.trim() : val;
        }
      }
    }

    // 빈 줄 검사 (B열 이후에 내용이 하나라도 있으면 통과)
    let hasContent = alignedRow.some((val, idx) => idx > 0 && val !== "");
    if (!hasContent) {
      blankCount++;
      continue;
    }

    // 드롭다운 허용값 사전 검증 (값이 비어있지 않은데 목록에 없으면 이 행은 제외)
    let invalidReason = null;
    for (let s = 1; s < sheetHeaders.length; s++) {
      const allowed = allowedValuesByCol[s];
      const val = alignedRow[s];
      if (allowed && val !== "" && allowed.indexOf(val) === -1) {
        invalidReason = `${sheetHeaders[s]} 값 '${val}'이(가) 허용 목록에 없음`;
        break;
      }
    }
    if (invalidReason) {
      if (invalidRows.length < MAX_REPORTED) {
        invalidRows.push(`엑셀 ${i + 1}행: ${invalidReason}`);
      }
      invalidCount++;
      continue;
    }

    // 중복 체크
    if (dedupeIdx) {
      const key = buildKey_(alignedRow);
      if (key.replace(/\|/g, '')) {
        const existingRows = existingKeyMap[key];
        if (existingRows) {
          const target = existingRows[existingRows.length - 1];
          if (policy === 'overwrite') {
            // 같은 키의 기존 행 중 값까지 전부 같은 게 있으면 할 일이 없습니다.
            if (existingRows.some(r => rowsEqual_(existingData[r - 1], alignedRow, providedIdx))) {
              identicalCount++;
              continue;
            }
            // 덮어쓰기 탭(채권 등): 값이 바뀌는 게 정상 흐름이므로 마지막 기존 행을 최신 값으로 갱신 (이력은 _변경이력에 남김)
            if (keyToUpdateIdx[key] !== undefined) {
              rowsToUpdate[keyToUpdateIdx[key]].data = alignedRow;
            } else {
              keyToUpdateIdx[key] = rowsToUpdate.length;
              rowsToUpdate.push({ row: target, data: alignedRow });
            }
          } else {
            // 검토 탭: 엑셀에 값이 있는 열만 비교합니다. (엑셀이 비어있는 칸은 "정보 없음"이므로 기존 값을 유지)
            const diffOf = r => providedIdx.filter(i =>
              normalizeKeyPart_(alignedRow[i]) !== '' && cmpNorm_(existingData[r - 1][i]) !== cmpNorm_(alignedRow[i]));
            if (existingRows.some(r => diffOf(r).length === 0)) {
              identicalCount++;
              continue;
            }
            const diffIdx = diffOf(target);
            const targetRow = existingData[target - 1];
            const fillable = diffIdx.every(i => normalizeKeyPart_(targetRow[i]) === '');
            const queuedFill = fillByKey[key];
            if (queuedFill && rowsEqual_(queuedFill.data, alignedRow, providedIdx)) {
              identicalCount++;
              continue;
            }
            if (fillable && !queuedFill) {
              // 기존 행에서 비어있던 칸(조치일/조치사항/결과 등)만 새로 채워지는 경우: 잃는 정보가 없으므로 자동 반영
              const f = { row: target, data: alignedRow, cols: diffIdx };
              fillByKey[key] = f;
              rowsToFill.push(f);
              continue;
            }
            // 기존에 값이 있는데 다른 값이 온 경우: 덮어쓰지 않고 _검토대기로 보냅니다.
            const list = conflictsByKey[key] || (conflictsByKey[key] = []);
            if (list.some(c => rowsEqual_(c.data, alignedRow, providedIdx))) {
              identicalCount++;
              continue;
            }
            const c = {
              row: target, data: alignedRow, excelRow: i + 1,
              diffIdx: diffIdx,
              writeIdx: providedIdx.filter(k => normalizeKeyPart_(alignedRow[k]) !== ''),
            };
            list.push(c);
            if (pendingKeys.has(`${targetSheetName}|${target}|${JSON.stringify(alignedRow)}`)) {
              alreadyPendingCount++;
              continue;
            }
            conflicts.push(c);
          }
          continue;
        }
        // 기존 시트에 없는 키
        if (policy === 'overwrite') {
          if (keyToAppendIdx[key] !== undefined) {
            rowsToAppend[keyToAppendIdx[key]] = alignedRow;
            continue;
          }
          keyToAppendIdx[key] = rowsToAppend.length;
        } else {
          const list = newByKey[key] || (newByKey[key] = []);
          if (list.some(r => rowsEqual_(r, alignedRow, providedIdx))) {
            identicalCount++;
            continue;
          }
          list.push(alignedRow);
        }
      }
    }

    rowsToAppend.push(alignedRow);
  }

  return {
    ss, sheet, targetSheetName, sheetHeaders, existingData, providedIdx, policy,
    rowsToAppend, rowsToUpdate, rowsToFill, conflicts,
    invalidRows, invalidCount, blankCount, identicalCount, alreadyPendingCount, totalRows: newDataArray.length,
    MAX_REPORTED,
  };
}

// 업로드 전 미리보기: 시트를 수정하지 않고 "신규 / 덮어쓰기 / 검토대기 / 동일 / 제외" 건수와 샘플을 돌려줍니다.
function previewExcelUpload(newDataArray, targetSheetName) {
  if (!newDataArray || newDataArray.length === 0) {
    throw new Error("진단 에러: 프론트엔드에서 엑셀 데이터를 서버로 전혀 전달하지 못했습니다.");
  }
  const plan = analyzeUpload_(newDataArray, targetSheetName);
  const SAMPLE = 10;
  const h = plan.sheetHeaders;
  return {
    sheetName: targetSheetName,
    policy: plan.policy,
    totalRows: plan.totalRows,
    appendCount: plan.rowsToAppend.length,
    updateCount: plan.rowsToUpdate.length,
    fillCount: plan.rowsToFill.length,
    conflictCount: plan.conflicts.length,
    identicalCount: plan.identicalCount,
    alreadyPendingCount: plan.alreadyPendingCount,
    invalidCount: plan.invalidCount,
    blankCount: plan.blankCount,
    appendSamples: plan.rowsToAppend.slice(0, SAMPLE).map(describeRow_),
    updateSamples: plan.rowsToUpdate.slice(0, SAMPLE).map(u =>
      `${u.row}행: ${diffSummary_(h, plan.existingData[u.row - 1], u.data, plan.providedIdx)}`),
    fillSamples: plan.rowsToFill.slice(0, SAMPLE).map(f =>
      `${f.row}행: ${diffSummary_(h, plan.existingData[f.row - 1], f.data, f.cols)}`),
    conflictSamples: plan.conflicts.slice(0, SAMPLE).map(c =>
      `${c.row}행 (엑셀 ${c.excelRow}행): ${diffSummary_(h, plan.existingData[c.row - 1], c.data, c.diffIdx)}`),
    invalidSamples: plan.invalidRows.slice(0, SAMPLE),
  };
}

// 엑셀 파일 일괄 업로드 및 유연한 데이터 매칭 함수
function uploadExcelData(newDataArray, targetSheetName) {
  if (!newDataArray || newDataArray.length === 0) {
    throw new Error("진단 에러: 프론트엔드에서 엑셀 데이터를 서버로 전혀 전달하지 못했습니다.");
  }

  return withLock_(function() {
    const plan = analyzeUpload_(newDataArray, targetSheetName);
    const { rowsToAppend, rowsToUpdate, rowsToFill, conflicts, invalidRows, invalidCount, blankCount, identicalCount, alreadyPendingCount, MAX_REPORTED } = plan;

    if (rowsToAppend.length === 0 && rowsToUpdate.length === 0 && rowsToFill.length === 0 && conflicts.length === 0 && identicalCount === 0 && alreadyPendingCount === 0) {
      const reasonMsg = invalidRows.length
        ? `\n\n제외 사유:\n${invalidRows.join('\n')}`
        : '';
      throw new Error(`엑셀에서 유효한 데이터를 찾지 못했습니다. (읽은 총 행 수: ${plan.totalRows}행)${reasonMsg}`);
    }

    try {
      const { ss, sheet, sheetHeaders, existingData, providedIdx } = plan;
      const now = new Date();
      const who = currentUser_();

      // ---- 0) 로그를 먼저 남깁니다. (여기서 실패하면 시트 본문은 아직 아무것도 바뀌지 않은 상태) ----
      if (rowsToUpdate.length > 0) {
        appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.history), rowsToUpdate.map(u => [
          now, who, targetSheetName, u.row, '업로드 덮어쓰기',
          diffSummary_(sheetHeaders, existingData[u.row - 1], u.data, providedIdx),
          JSON.stringify(plainRow_(existingData[u.row - 1])),
        ]));
      }
      if (rowsToFill.length > 0) {
        appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.history), rowsToFill.map(f => [
          now, who, targetSheetName, f.row, '업로드 빈칸 채움',
          diffSummary_(sheetHeaders, existingData[f.row - 1], f.data, f.cols),
          JSON.stringify(plainRow_(existingData[f.row - 1])),
        ]));
      }
      if (conflicts.length > 0) {
        appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.review), conflicts.map(c => [
          now, who, targetSheetName, c.row, '대기',
          diffSummary_(sheetHeaders, existingData[c.row - 1], c.data, c.diffIdx),
          JSON.stringify(plainRow_(existingData[c.row - 1])),
          JSON.stringify(c.data),
          // verify: 반영 시 "그 사이 행이 안 바뀌었는지" 확인할 열 / write: 실제로 덮어쓸 열(엑셀에 값이 있는 열만)
          JSON.stringify({ verify: providedIdx, write: c.writeIdx }),
          '',
        ]));
      }

      // ---- 1) 기존 행 업데이트 (엑셀에 있던 열만 덮어씀 — A열 No. 수식과 엑셀에 없는 열은 그대로) ----
      rowsToUpdate.forEach(u => writeProvidedCols_(sheet, u.row, u.data, providedIdx));
      rowsToFill.forEach(f => writeProvidedCols_(sheet, f.row, f.data, f.cols));

      // ---- 2) 신규 행 추가 (한 번에 일괄 저장) ----
      if (rowsToAppend.length > 0) {
        appendRows_(sheet, rowsToAppend, sheetHeaders.length);
      }

      SpreadsheetApp.flush();
    } catch (writeError) {
      throw new Error(`저장 중 오류가 발생했습니다: ${writeError.message}\n(사전 검증을 통과했는데도 실패했다면, 예상치 못한 데이터 확인 규칙이 걸려있을 수 있습니다.)`);
    }

    let msg = `성공! 신규 ${rowsToAppend.length}건 추가`
      + (rowsToUpdate.length > 0 ? `, 기존 ${rowsToUpdate.length}건 덮어쓰기(변경 전 내용은 '${LOG_SHEETS.history.name}' 탭에 기록)` : '')
      + (rowsToFill.length > 0 ? `, 기존 ${rowsToFill.length}건은 빈 칸만 채움(기존 값 유지)` : '')
      + (conflicts.length > 0 ? `, 값이 달라 덮어쓰지 않은 ${conflicts.length}건은 '${LOG_SHEETS.review.name}' 탭에 보관` : '')
      + (alreadyPendingCount > 0 ? `, 이미 검토 대기 중인 ${alreadyPendingCount}건은 중복 등록하지 않음` : '')
      + (identicalCount > 0 ? `, 이미 동일한 ${identicalCount}건 건너뜀` : '')
      + (invalidCount > 0 ? `, 제외 ${invalidCount}건(드롭다운 값 오류)` : '')
      + ` · 빈 행 ${blankCount}개 제외`;
    if (invalidRows.length) {
      msg += `\n\n제외된 행(드롭다운 목록에 없는 값, 최대 ${MAX_REPORTED}건 표시):\n${invalidRows.join('\n')}`
        + (invalidCount > invalidRows.length ? `\n...외 ${invalidCount - invalidRows.length}건 더` : '');
    }
    return msg;
  });
}

// ---------------------------------------------------------------------------
// _검토대기 탭 처리: 담당자가 E열(상태)을 '반영'(기존 행 덮어쓰기) / '신규추가'(별개 건이라 새 행으로 추가) /
// '무시'(엑셀 값 버림) 중 하나로 바꾼 행만 처리합니다. '대기'인 행은 그대로 남습니다.
// 웹 화면의 "검토 대기 반영" 버튼이나, Apps Script 편집기에서 applyReviewedChanges()를 직접 실행하면 됩니다.
// ---------------------------------------------------------------------------
function applyReviewedChanges() {
  return withLock_(function() {
    const ss = openSpreadsheet_();
    const reviewSheet = ss.getSheetByName(LOG_SHEETS.review.name);
    if (!reviewSheet || reviewSheet.getLastRow() < 2) return '검토 대기 중인 항목이 없습니다.';

    const COLS = 10;
    const rows = reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, COLS).getValues();
    const now = new Date();
    const who = currentUser_();
    const historyRows = [];
    const pendingWrites = [];
    const statusOut = rows.map(r => [r[4], r[9]]); // [상태, 처리결과]
    let applied = 0, added = 0, ignored = 0, failed = 0, pending = 0;

    rows.forEach((r, i) => {
      const status = String(r[4]).trim();
      if (status === '대기' || status === '') { pending++; return; }
      if (['반영', '신규추가', '무시'].indexOf(status) === -1) return; // '완료(...)' 등 이미 처리된 행

      const tabName = r[2];
      const sheetRow = Number(r[3]);
      try {
        if (status === '무시') {
          statusOut[i] = ['완료(무시)', `${now.toLocaleString()} ${who}`];
          ignored++;
          return;
        }
        const sheet = ss.getSheetByName(tabName);
        if (!sheet) throw new Error(`'${tabName}' 탭을 찾을 수 없음`);
        const oldSnapshot = JSON.parse(r[6]);
        const newData = JSON.parse(r[7]);
        const cfg = JSON.parse(r[8]);
        // 예전 형식(배열)과 새 형식({verify, write}) 모두 지원
        const verifyIdx = Array.isArray(cfg) ? cfg : cfg.verify;
        const writeIdx = Array.isArray(cfg) ? cfg : cfg.write;

        if (status === '신규추가') {
          appendRows_(sheet, [newData], newData.length);
          statusOut[i] = ['완료(신규추가)', `${now.toLocaleString()} ${who}`];
          added++;
          return;
        }

        // '반영': 그 사이 시트가 바뀌어 행 번호가 밀리거나 값이 달라졌으면 잘못 덮어쓰지 않도록 중단합니다.
        const current = sheet.getRange(sheetRow, 1, 1, newData.length).getValues()[0];
        if (!rowsEqual_(current, oldSnapshot, verifyIdx)) {
          throw new Error(`${sheetRow}행이 검토 등록 이후 변경되어 덮어쓰지 않았습니다. 내용 확인 후 다시 처리하세요.`);
        }
        historyRows.push([
          now, who, tabName, sheetRow, '검토 반영',
          r[5], // 업로드 때 계산해 둔 변경 요약
          JSON.stringify(plainRow_(current)),
        ]);
        // 이력을 먼저 확정(아래 appendLogRows_)해야 하므로 실제 쓰기는 이력 기록 뒤로 미룹니다.
        pendingWrites.push({ sheet, sheetRow, newData, providedIdx: writeIdx });
        statusOut[i] = ['완료(반영)', `${now.toLocaleString()} ${who}`];
        applied++;
      } catch (e) {
        statusOut[i] = [status, `실패: ${e.message}`];
        failed++;
      }
    });

    appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.history), historyRows);
    pendingWrites.forEach(p => writeProvidedCols_(p.sheet, p.sheetRow, p.newData, p.providedIdx));
    reviewSheet.getRange(2, 5, rows.length, 1).setValues(statusOut.map(s => [s[0]]));
    reviewSheet.getRange(2, 10, rows.length, 1).setValues(statusOut.map(s => [s[1]]));
    SpreadsheetApp.flush();

    const msg = `검토 반영 결과: 덮어쓰기 ${applied}건, 신규추가 ${added}건, 무시 ${ignored}건, 실패 ${failed}건 (아직 '대기' ${pending}건)`;
    Logger.log(msg);
    return msg;
  });
}

// ---------------------------------------------------------------------------
// 이미 시트에 쌓여있는 중복 항목 정리용 (반복 업로드로 생긴 과거 중복 데이터 청소).
// DEDUPE_KEYS에 등록된 모든 탭을 한 번에 확인/정리합니다. 탭마다 판단 기준이 다릅니다:
// - [K] AS/매출: 접수일+제조번호/코드+증상/내용이 같으면 중복(금액 등 나머지 값은 달라도 됨)
// - 그 외 탭: No.(A열) 제외 전체 열 값이 완전히 같아야 중복
// 각 탭에서 같은 키를 가진 행이 여러 개면 가장 마지막(아래쪽, 최신) 행만 남기고 나머지는 삭제합니다.
//
// Apps Script 편집기에서 함수 목록 중 아래 두 개를 골라 "실행"하면 됩니다 (인자 입력 불필요):
//   1) previewDuplicateCleanup()  → 실제로 지우지 않고, 몇 건이 삭제 대상인지만 미리 확인
//   2) runDuplicateCleanup()      → 실제 삭제 실행
// 실행 결과는 Apps Script 편집기의 "실행 로그"에서 확인할 수 있습니다.
// 삭제되는 행의 값은 삭제 직전에 '_삭제로그' 탭에 자동 백업됩니다(백업 실패 시 삭제하지 않음).
// 잘못 지웠다면 그 탭에서 해당 행을 복사해 되돌리세요. (No. 열의 수식은 값으로만 남습니다.)
// 그래도 대량 삭제 전에는 스프레드시트 사본(파일 > 사본 만들기)을 만들어두면 더 안전합니다.
// ---------------------------------------------------------------------------

// 여기 등록된 모든 탭(DEDUPE_KEYS에 정의된 전부)을 한 번에 확인/정리합니다.
function previewDuplicateCleanup() {
  const msg = Object.keys(CLEANUP_KEYS)
    .map(name => `■ ${name}\n` + cleanupDuplicates_(name, false))
    .join('\n\n');
  Logger.log(msg);
  return msg;
}

function runDuplicateCleanup() {
  const msg = withLock_(() => Object.keys(CLEANUP_KEYS)
    .map(name => `■ ${name}\n` + cleanupDuplicates_(name, true))
    .join('\n\n'));
  Logger.log(msg);
  return msg;
}

function cleanupDuplicates_(targetSheetName, actuallyDelete) {
  const dedupeCols = CLEANUP_KEYS[targetSheetName];
  if (!dedupeCols) {
    throw new Error(`'${targetSheetName}' 탭에는 중복 판단 기준(CLEANUP_KEYS)이 정의되어 있지 않습니다.`);
  }

  const ss = openSpreadsheet_();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return '데이터가 없습니다.';

  // 헤더 행 자동 탐지.
  let headerRowIdx = -1;
  let idxList = null;
  const HEADER_SCAN_ROWS = Math.min(5, data.length);
  if (dedupeCols === '__ALL__') {
    // 전체 열 기준 탭은 이름으로 찾을 컬럼이 없으므로 1행을 헤더로 간주합니다.
    headerRowIdx = 0;
    idxList = [];
    for (let s = 1; s < data[0].length; s++) idxList.push(s);
  } else if (dedupeCols.exclude) {
    // 제외 컬럼 기준: exclude에 적힌 컬럼명이 전부 등장하는 첫 행을 헤더로 간주하고,
    // 그 컬럼들만 빼고 나머지 전부를 키로 사용합니다.
    const excludeSet = new Set(dedupeCols.exclude);
    for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
      const hasAllExcluded = dedupeCols.exclude.every(name => data[r].indexOf(name) !== -1);
      if (hasAllExcluded) {
        headerRowIdx = r;
        idxList = [];
        for (let s = 1; s < data[r].length; s++) {
          if (!excludeSet.has(data[r][s])) idxList.push(s);
        }
        break;
      }
    }
    if (!idxList) {
      throw new Error(`헤더 행에서 제외 대상 컬럼(${dedupeCols.exclude.join(', ')})을 찾지 못했습니다.`);
    }
  } else {
    // dedupeCols 이름이 전부 등장하는 첫 행을 헤더로 간주합니다.
    for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
      const candidate = dedupeCols.map(name => data[r].indexOf(name));
      if (candidate.every(i => i !== -1)) {
        headerRowIdx = r;
        idxList = candidate;
        break;
      }
    }
    if (!idxList) {
      throw new Error(`헤더 행에서 중복 판단 기준 컬럼(${dedupeCols.join(', ')})을 찾지 못했습니다.`);
    }
  }

  const buildKey = (row) => idxList.map(i => cmpNorm_(row[i])).join('|');

  // 같은 키를 가진 행 번호들을 등장 순서대로 모읍니다.
  const keyToRows = {}; // key -> [시트 행 번호(1-based), ...]
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const key = buildKey(data[i]);
    if (!key.replace(/\|/g, '')) continue; // 키 값이 전부 비어있으면 무시
    if (!keyToRows[key]) keyToRows[key] = [];
    keyToRows[key].push(i + 1);
  }

  // 각 중복 그룹에서 마지막(최신) 행만 남기고 나머지를 삭제 대상으로 표시합니다.
  const rowsToDelete = [];
  let dupGroups = 0;
  Object.keys(keyToRows).forEach(key => {
    const rows = keyToRows[key];
    if (rows.length > 1) {
      dupGroups++;
      rowsToDelete.push(...rows.slice(0, -1));
    }
  });
  rowsToDelete.sort((a, b) => a - b);

  if (!actuallyDelete) {
    const preview = rowsToDelete.slice(0, 50).join(', ') + (rowsToDelete.length > 50 ? ' ...' : '');
    return `[미리보기] 중복 그룹 ${dupGroups}개, 삭제 대상 ${rowsToDelete.length}행\n`
      + `삭제될 행 번호: ${preview}\n\n`
      + `실제로 삭제하려면 runDuplicateCleanup()을 실행하세요.`;
  }

  // 삭제 전 백업: 지울 행의 값을 '_삭제로그' 탭에 먼저 남깁니다. (여기서 실패하면 예외가 나서 삭제는 일어나지 않습니다.)
  if (rowsToDelete.length > 0) {
    const now = new Date();
    const who = currentUser_();
    appendLogRows_(getOrCreateLogSheet_(ss, LOG_SHEETS.deleted), rowsToDelete.map(r =>
      [now, who, targetSheetName, r, buildKey(data[r - 1])].concat(plainRow_(data[r - 1]))));
  }

  // 실제 삭제: 큰 행 번호부터 지워야 앞쪽 행 번호가 밀리지 않습니다.
  // 연속된 행 번호는 묶어서 한 번에 삭제합니다.
  let i = rowsToDelete.length - 1;
  let deletedCount = 0;
  while (i >= 0) {
    let start = i;
    while (start > 0 && rowsToDelete[start - 1] === rowsToDelete[start] - 1) {
      start--;
    }
    const count = i - start + 1;
    sheet.deleteRows(rowsToDelete[start], count);
    deletedCount += count;
    i = start - 1;
  }

  return `삭제 완료: 중복 그룹 ${dupGroups}개, 총 ${deletedCount}행 삭제됨. (삭제된 행 값은 '${LOG_SHEETS.deleted.name}' 탭에 백업됨)`;
}

// ---------------------------------------------------------------------------
// [K] AS/매출 탭의 "제품명(사양)" 드롭다운 목록에 새로 확인된 정식 모델명을 추가합니다.
// (표기 오타로 보이는 값은 제외하고, 실제로 다른 모델로 판단되는 것만 포함했습니다.)
// Apps Script 편집기에서 addAllowedProductModels() 를 선택해 실행하면 됩니다.
// ---------------------------------------------------------------------------
function addAllowedProductModels() {
  const NEW_VALUES = [
    'GIX-1 (S2)', 'GIX-1 (S1)', 'MX-600', 'ZEN-5000 Z3', 'ZEN-2060P Z1', 'ZEN-2060P',
    'E2V(S2)', 'MX-500', 'PAPAYA-CUST', 'ZEN-2060P Z2', 'TOSHIBA(S2)', 'GDP-1C',
    'PORT-X 4', 'GDP-1', 'HESTIA(L)', 'GX-D1',
  ];

  const ss = openSpreadsheet_();
  const sheet = ss.getSheetByName('[K] AS/매출');
  if (!sheet) throw new Error("'[K] AS/매출' 탭을 찾을 수 없습니다.");

  const data = sheet.getDataRange().getValues();
  let colIdx = -1;
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const idx = data[r].indexOf('제품명(사양)');
    if (idx !== -1) { colIdx = idx + 1; break; }
  }
  if (colIdx === -1) throw new Error("'제품명(사양)' 헤더를 찾지 못했습니다.");

  const checkRow = sheet.getLastRow() > 1 ? sheet.getLastRow() : 2;
  const existingRule = sheet.getRange(checkRow, colIdx).getDataValidation();
  if (!existingRule || existingRule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    throw new Error("'제품명(사양)' 열에 목록 검증 규칙이 걸려있지 않습니다.");
  }

  const currentValues = existingRule.getCriteriaValues()[0];
  const merged = currentValues.slice();
  let addedCount = 0;
  NEW_VALUES.forEach(v => {
    if (merged.indexOf(v) === -1) {
      merged.push(v);
      addedCount++;
    }
  });

  const newRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(merged, true)
    .setAllowInvalid(false)
    .build();

  // 기존 규칙과 같은 방식으로, 2행부터 시트 최대 행까지 전체 열에 재적용합니다.
  sheet.getRange(2, colIdx, sheet.getMaxRows() - 1, 1).setDataValidation(newRule);

  const msg = `완료: 드롭다운에 ${addedCount}개 값 신규 추가 (전체 허용값 ${merged.length}개).`;
  Logger.log(msg);
  return msg;
}

// ---------------------------------------------------------------------------
// A열(No.)이 비어있는데 다른 칸엔 내용이 있는 행을 찾아서 목록으로 보여줍니다.
// (예전 코드가 No. 수식을 못 채우던 시절 생긴 부분 기록/잔여 테스트 데이터일 수 있어서,
// 삭제 전에 먼저 내용을 확인하기 위한 진단용입니다. 이 함수는 아무것도 지우지 않습니다.)
// Apps Script 편집기에서 previewBlankNoRows() 를 선택해 실행하세요.
// ---------------------------------------------------------------------------
function previewBlankNoRows(targetSheetName) {
  targetSheetName = targetSheetName || '[K] AS/매출';
  const ss = openSpreadsheet_();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return '데이터가 없습니다.';

  // 헤더 행 탐지 (No., 접수일 등 표준 헤더명이 있는 행)
  const HEADER_SCAN_ROWS = Math.min(5, data.length);
  let headerRowIdx = 0;
  for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
    if (data[r].indexOf('접수일') !== -1) { headerRowIdx = r; break; }
  }

  const lines = [];
  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = data[i];
    const noVal = row[0];
    const hasOtherContent = row.some((v, idx) => idx > 0 && v !== '' && v !== null);
    if ((noVal === '' || noVal === null) && hasOtherContent) {
      const rowNum = i + 1;
      // 접수일(1), 제조번호/코드(6), 증상/내용(9) 위주로 요약 표시
      const summary = [row[1], row[6], row[9]].map(v => v instanceof Date
        ? Utilities.formatDate(v, spreadsheetTz_(), 'yyyy-MM-dd') : v).join(' | ');
      lines.push(`${rowNum}행: ${summary}`);
    }
  }

  if (lines.length === 0) return 'No.가 비어있는 데이터 행이 없습니다.';

  const msg = `No.가 비어있는 행 ${lines.length}건:\n` + lines.slice(0, 100).join('\n')
    + (lines.length > 100 ? `\n...외 ${lines.length - 100}건 더 (Apps Script 실행 로그에 전체 출력)` : '');
  Logger.log(`No.가 비어있는 행 ${lines.length}건:\n` + lines.join('\n'));
  return msg;
}

// ---------------------------------------------------------------------------
// previewBlankNoRows()로 확인한 "No.가 비어있는 행"은 대부분 중복이 아니라, 예전에
// (이번 수정 이전에) 다른 방식으로 추가되면서 A열 번호 수식만 채워지지 않은 정상 데이터입니다.
// 이 함수는 내용은 전혀 건드리지 않고, 그 행들의 A열에 번호 수식만 채워 넣습니다.
// (진짜 중복 여부는 이 함수가 아니라 previewDuplicateCleanup()으로 확인하세요 —
//  그 함수는 No. 값과 상관없이 접수일/제조번호/증상 기준으로 실제 중복을 찾습니다.)
// Apps Script 편집기에서 backfillNoFormulas() 를 선택해 실행하세요.
// ---------------------------------------------------------------------------
function backfillNoFormulas(targetSheetName) {
  targetSheetName = targetSheetName || '[K] AS/매출';
  const ss = openSpreadsheet_();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return '데이터가 없습니다.';

  const HEADER_SCAN_ROWS = Math.min(5, data.length);
  let headerRowIdx = 0;
  for (let r = 0; r < HEADER_SCAN_ROWS; r++) {
    if (data[r].indexOf('접수일') !== -1) { headerRowIdx = r; break; }
  }

  const lastRow = sheet.getLastRow();
  const startRow = headerRowIdx + 2; // 데이터 첫 행(1-based)
  if (startRow > lastRow) return '데이터가 없습니다.';

  // A열 수식을 한 번에 통째로 읽고/씁니다 (행마다 개별 호출하면 대량 데이터에서 시간초과 위험).
  const numRows = lastRow - startRow + 1;
  const colAFormulas = sheet.getRange(startRow, 1, numRows, 1).getFormulasR1C1();

  let lastKnownFormula = '';
  const newFormulas = [];
  let filledCount = 0;

  for (let i = 0; i < numRows; i++) {
    const row = data[startRow - 1 + i];
    const existingFormula = colAFormulas[i][0];

    if (existingFormula) {
      lastKnownFormula = existingFormula;
      newFormulas.push([existingFormula]);
      continue;
    }

    const hasOtherContent = row.some((v, idx) => idx > 0 && v !== '' && v !== null);
    const noVal = row[0];
    if ((noVal === '' || noVal === null) && hasOtherContent && lastKnownFormula) {
      newFormulas.push([lastKnownFormula]);
      filledCount++;
    } else {
      newFormulas.push([existingFormula || '']);
    }
  }

  if (filledCount > 0) {
    sheet.getRange(startRow, 1, numRows, 1).setFormulasR1C1(newFormulas);
  }

  const msg = `완료: A열(No.) 수식을 ${filledCount}개 행에 채워 넣었습니다.`;
  Logger.log(msg);
  return msg;
}
