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

  sheet.appendRow(rowData);
  return `${targetSheetName} 시트에 성공적으로 저장되었습니다!`;
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

  const rowsToAppend = [];
  let skippedCount = 0;

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
    }
  }
  // 헤더명이 2개 이상 일치하면 이름 매칭 사용, 그렇지 않으면(제목줄이 없는 단순 엑셀 등) 기존 순서 매칭으로 대체
  const useNameMapping = nameMatchCount >= 2;

  // 엑셀 데이터 행 순회 (0번째 줄은 엑셀 제목이므로 1번째 줄부터 시작)
  for (let i = 1; i < newDataArray.length; i++) {
    let excelRow = newDataArray[i];
    if (!excelRow || excelRow.length === 0) {
      skippedCount++;
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
      skippedCount++;
      continue;
    }

    // 무조건 추가 목록에 담습니다!
    rowsToAppend.push(alignedRow);
  }

  if (rowsToAppend.length === 0) {
    throw new Error(`엑셀에서 유효한 데이터를 찾지 못했습니다. (읽은 총 행 수: ${newDataArray.length}행)`);
  }

  const startRow = sheet.getLastRow() + 1;

  // 1차: 한 번에 일괄 저장 시도 (빠른 경로)
  // setValues()는 대기열에만 쌓이고 스크립트 종료 시점에야 실제 반영되므로,
  // flush()로 즉시 반영시켜야 검증 규칙 위반 에러를 이 자리에서 catch할 수 있습니다.
  try {
    sheet.getRange(startRow, 1, rowsToAppend.length, sheetHeaders.length).setValues(rowsToAppend);
    SpreadsheetApp.flush();
    return `[코드 v2] 성공! 총 ${rowsToAppend.length}건의 데이터가 시트 맨 아래에 추가되었습니다. (제외된 빈 행: ${skippedCount}개)`;
  } catch (bulkError) {
    // 일괄 저장이 실패하면(주로 드롭다운 등 데이터 확인 규칙 위반), 한 줄씩 다시 시도해서
    // 문제 없는 행은 저장하고, 규칙에 위반되는 행만 걸러내어 사용자에게 알려줍니다.
    let successCount = 0;
    let currentRow = startRow;
    const failedRows = [];
    const MAX_REPORTED = 30;

    for (let i = 0; i < rowsToAppend.length; i++) {
      try {
        sheet.getRange(currentRow, 1, 1, sheetHeaders.length).setValues([rowsToAppend[i]]);
        SpreadsheetApp.flush();
        successCount++;
        currentRow++;
      } catch (rowError) {
        if (failedRows.length < MAX_REPORTED) {
          failedRows.push(`엑셀 ${i + 2}행: ${rowError.message}`);
        }
      }
    }

    const omitted = (rowsToAppend.length - successCount) - failedRows.length;
    let msg = `[코드 v2] 일부만 저장되었습니다. 성공 ${successCount}건 / 실패 ${rowsToAppend.length - successCount}건 (제외된 빈 행: ${skippedCount}개)\n\n`
      + `실패 사유(주로 드롭다운 목록에 없는 값): \n${failedRows.join('\n')}`
      + (omitted > 0 ? `\n...외 ${omitted}건 더` : '')
      + `\n\n[진단] 시트 헤더: ${sheetHeaders.map((h, idx) => `${idx}:${h}`).join(' | ')}`
      + `\n[진단] 엑셀 헤더: ${excelHeaders.map((h, idx) => `${idx}:${h}`).join(' | ')}`
      + `\n[진단] 이름매칭 사용: ${useNameMapping} (일치 ${nameMatchCount}개)`
      + `\n[진단] 실패한 행의 실제 기록값: ${JSON.stringify(rowsToAppend[0])}`;

    if (successCount === 0) {
      throw new Error(msg);
    }
    return msg;
  }
}
