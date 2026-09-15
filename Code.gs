function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('AS/매출 통합 관리 포털')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. 7개 탭 수동 입력 저장 함수
function saveMultiData(data, targetSheetName) {
  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

  // A열(No.) 수식 유지를 위해 B열부터 데이터 저장
  let rowData = [];

  switch (targetSheetName) {
    case '[K] AS/매출':
      rowData = ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11, data.d12, data.d13, data.d14];
      break;
    case '[K] 유지보수':
      rowData = ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11, data.d12, data.d13];
      break;
    case '[K] 채권':
      const balanceK = (Number(data.d3) || 0) - (Number(data.d4) || 0);
      rowData = ["", data.d1, data.d2, data.d3, data.d4, balanceK, data.d6, data.d7, data.d8, data.d9];
      break;
    case '[M/D] AS':
      rowData = ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11, data.d12, data.d13, data.d14, data.d15];
      break;
    case '[M/D] 매출':
      rowData = ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, data.d7, data.d8, data.d9, data.d10, data.d11];
      break;
    case '[M] 채권':
    case '[D] 채권':
      const balanceMD = (Number(data.d4) || 0) - (Number(data.d5) || 0);
      rowData = ["", data.d1, data.d2, data.d3, data.d4, data.d5, data.d6, balanceMD, data.d7, data.d8, data.d9, data.d10];
      break;
    default:
      throw new Error("알 수 없는 양식입니다.");
  }

  const prevLastRow = sheet.getLastRow();
  sheet.appendRow(rowData);
  const newRow = sheet.getLastRow();

  // A열(No.) 수식은 행마다 개별적으로 걸려있어 새 행엔 복사되어 있지 않으므로,
  // 바로 위 행의 수식을 그대로 복사해 붙여넣습니다(상대참조라 행 번호는 자동으로 맞춰집니다).
  if (prevLastRow >= 1) {
    const aboveFormula = sheet.getRange(prevLastRow, 1).getFormulaR1C1();
    if (aboveFormula) {
      sheet.getRange(newRow, 1).setFormulaR1C1(aboveFormula);
    }
  }

  return `${targetSheetName} 시트에 성공적으로 저장되었습니다!`;
}

// 탭별 "중복 판단 기준" 컬럼명. 여기 정의된 탭만 중복 체크 + 업데이트(덮어쓰기)를 하고,
// 정의 안 된 탭은 기존처럼 무조건 새 행으로 추가합니다.
const DEDUPE_KEYS = {
  '[K] AS/매출': ['접수일', '제조번호/코드', '증상/내용'],
};

// 날짜(Date 객체)/문자열을 동일한 형식으로 맞춰서 비교 가능하게 만듭니다.
function normalizeKeyPart_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return (v === null || v === undefined) ? '' : String(v).trim();
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

// 엑셀 파일 일괄 업로드 및 유연한 데이터 매칭 함수
function uploadExcelData(newDataArray, targetSheetName) {
  if (!newDataArray || newDataArray.length === 0) {
    throw new Error("진단 에러: 프론트엔드에서 엑셀 데이터를 서버로 전혀 전달하지 못했습니다.");
  }

  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
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
  const dedupeCols = DEDUPE_KEYS[targetSheetName] || null;
  let dedupeIdx = null;
  if (dedupeCols && useNameMapping) {
    const idxList = dedupeCols.map(name => sheetHeaders.indexOf(name));
    if (idxList.every(i => i !== -1)) dedupeIdx = idxList;
  }
  const buildKey_ = (rowArray) => dedupeIdx.map(i => normalizeKeyPart_(rowArray[i])).join('|');

  // 기존 시트 데이터로 "키 → 시트 행 번호(1-based)" 맵을 만듭니다.
  const existingKeyMap = {};
  if (dedupeIdx) {
    for (let i = headerRowIdx + 1; i < existingData.length; i++) {
      const key = buildKey_(existingData[i]);
      if (key.replace(/\|/g, '')) existingKeyMap[key] = i + 1;
    }
  }

  const rowsToAppend = [];
  const keyToAppendIdx = {};
  const rowsToUpdate = []; // {row, data}
  const keyToUpdateIdx = {};
  const invalidRows = [];
  const MAX_REPORTED = 30;
  let blankCount = 0;
  let invalidCount = 0;
  let updateCount = 0;

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

    // 중복 체크: 같은 키가 기존 시트에 있으면 그 행을 업데이트, 이번 파일 안에서 같은 키가
    // 먼저 나왔으면 그 대기 항목을 최신 값으로 교체, 둘 다 아니면 신규 추가로 담습니다.
    if (dedupeIdx) {
      const key = buildKey_(alignedRow);
      if (key.replace(/\|/g, '') && existingKeyMap[key] !== undefined) {
        const rowNum = existingKeyMap[key];
        if (keyToUpdateIdx[key] !== undefined) {
          rowsToUpdate[keyToUpdateIdx[key]].data = alignedRow;
        } else {
          keyToUpdateIdx[key] = rowsToUpdate.length;
          rowsToUpdate.push({ row: rowNum, data: alignedRow });
          updateCount++;
        }
        continue;
      } else if (key.replace(/\|/g, '') && keyToAppendIdx[key] !== undefined) {
        rowsToAppend[keyToAppendIdx[key]] = alignedRow;
        continue;
      } else if (key.replace(/\|/g, '')) {
        keyToAppendIdx[key] = rowsToAppend.length;
      }
    }

    rowsToAppend.push(alignedRow);
  }

  if (rowsToAppend.length === 0 && rowsToUpdate.length === 0) {
    const reasonMsg = invalidRows.length
      ? `\n\n제외 사유:\n${invalidRows.join('\n')}`
      : '';
    throw new Error(`엑셀에서 유효한 데이터를 찾지 못했습니다. (읽은 총 행 수: ${newDataArray.length}행)${reasonMsg}`);
  }

  try {
    // ---- 1) 기존 행 업데이트 (덮어쓰기) — A열(No. 수식)은 건드리지 않고 B열부터만 씁니다. ----
    rowsToUpdate.forEach(item => {
      sheet.getRange(item.row, 2, 1, sheetHeaders.length - 1).setValues([item.data.slice(1)]);
    });

    // ---- 2) 신규 행 추가 (한 번에 일괄 저장) ----
    if (rowsToAppend.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      // A열(No.) 수식은 행마다 개별적으로 걸려있어 새 행엔 복사되어 있지 않으므로,
      // 바로 위 행의 수식을 그대로 복사해 붙여넣습니다(상대참조라 행 번호는 자동으로 맞춰집니다).
      const aboveFormula = startRow > 1 ? sheet.getRange(startRow - 1, 1).getFormulaR1C1() : '';
      sheet.getRange(startRow, 1, rowsToAppend.length, sheetHeaders.length).setValues(rowsToAppend);
      if (aboveFormula) {
        const formulas = rowsToAppend.map(() => [aboveFormula]);
        sheet.getRange(startRow, 1, rowsToAppend.length, 1).setFormulasR1C1(formulas);
      }
    }

    SpreadsheetApp.flush();
  } catch (writeError) {
    throw new Error(`저장 중 오류가 발생했습니다: ${writeError.message}\n(사전 검증을 통과했는데도 실패했다면, 예상치 못한 데이터 확인 규칙이 걸려있을 수 있습니다.)`);
  }

  let msg = `성공! 신규 ${rowsToAppend.length}건 추가`
    + (updateCount > 0 ? `, 기존 ${updateCount}건 업데이트(덮어쓰기)` : '')
    + (invalidCount > 0 ? `, 제외 ${invalidCount}건(드롭다운 값 오류)` : '')
    + ` · 빈 행 ${blankCount}개 제외`;
  if (invalidRows.length) {
    msg += `\n\n제외된 행(드롭다운 목록에 없는 값, 최대 ${MAX_REPORTED}건 표시):\n${invalidRows.join('\n')}`
      + (invalidCount > invalidRows.length ? `\n...외 ${invalidCount - invalidRows.length}건 더` : '');
  }
  return msg;
}

// ---------------------------------------------------------------------------
// 이미 시트에 쌓여있는 중복 항목 정리용 (반복 업로드로 생긴 과거 중복 데이터 청소).
// DEDUPE_KEYS에 정의된 키(예: 접수일+제조번호/코드+증상/내용)가 같은 행이 여러 개면
// 가장 마지막(아래쪽, 최신) 행만 남기고 나머지는 삭제합니다.
//
// Apps Script 편집기에서 함수 목록 중 아래 두 개를 골라 "실행"하면 됩니다 (인자 입력 불필요):
//   1) previewDuplicateCleanup()  → 실제로 지우지 않고, 몇 건이 삭제 대상인지만 미리 확인
//   2) runDuplicateCleanup()      → 실제 삭제 실행
// 실행 결과는 Apps Script 편집기의 "실행 로그"에서 확인할 수 있습니다.
// ⚠️ 삭제는 되돌릴 수 없으니, runDuplicateCleanup() 실행 전에 스프레드시트를
//    사본으로 백업해두는 것을 권장합니다 (파일 > 사본 만들기).
// ---------------------------------------------------------------------------

function previewDuplicateCleanup() {
  const msg = cleanupDuplicates_('[K] AS/매출', false);
  Logger.log(msg);
  return msg;
}

function runDuplicateCleanup() {
  const msg = cleanupDuplicates_('[K] AS/매출', true);
  Logger.log(msg);
  return msg;
}

function cleanupDuplicates_(targetSheetName, actuallyDelete) {
  const dedupeCols = DEDUPE_KEYS[targetSheetName];
  if (!dedupeCols) {
    throw new Error(`'${targetSheetName}' 탭에는 중복 판단 기준(DEDUPE_KEYS)이 정의되어 있지 않습니다.`);
  }

  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error(`'${targetSheetName}' 탭을 찾을 수 없습니다.`);

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return '데이터가 없습니다.';

  // 헤더 행 자동 탐지: dedupeCols 이름이 전부 등장하는 첫 행을 헤더로 간주합니다.
  const HEADER_SCAN_ROWS = Math.min(5, data.length);
  let headerRowIdx = -1;
  let idxList = null;
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

  const buildKey = (row) => idxList.map(i => normalizeKeyPart_(row[i])).join('|');

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

  return `삭제 완료: 중복 그룹 ${dupGroups}개, 총 ${deletedCount}행 삭제됨.`;
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

  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
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
  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
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
        ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v).join(' | ');
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
  const ss = SpreadsheetApp.openById('1EqKrXRWWuDZv9j11iUHDOQmN0cDwAp47dBWvv1SyV7Q');
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
