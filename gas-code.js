// ============================================================
// Google Apps Script — 미네랄 주문서 (수정본)
// ============================================================
// 변경사항:
// 1. 우편번호 / 주소 셀 분리 (12열: 우편번호, 13열: 주소)
// 2. 운송장번호에 아무 텍스트나 입력해도 완료 시트로 이동
// 3. 완료 시트 이동 시 17열(발송날짜)에 날짜 자동 기록
// 4. 배송용 사이트를 위한 getShippingData 함수 추가
// ============================================================
// ★ 컬럼 구조 (변경됨):
//  1:주문번호  2:접수시간  3:주문상태  4:송금일시  5:송금금액
//  6:입금내역  7:DMAX지갑  8:주문자이름  9:주문자전화
// 10:수령인이름 11:수령인전화 12:우편번호 13:주소
// 14:배송요청  15:택배사  16:운송장번호  17:발송날짜
// ============================================================

function doGet(e) {
  var result;
  try {
    var action = e.parameter.action || '';
    if (action === 'order') {
      result = saveOrder(e.parameter);
    } else if (action === 'search') {
      result = searchOrders(e.parameter.name, e.parameter.phone);
    } else if (action === 'shipping') {
      result = getShippingData();
    } else {
      result = { success: true, message: 'API OK' };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  
  var callback = e.parameter.callback;
  var output = callback 
    ? callback + '(' + JSON.stringify(result) + ')'
    : JSON.stringify(result);
  
  return ContentService
    .createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function saveOrder(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  
  // 신규주문 시트가 없으면 생성
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow([
      '주문번호','접수시간','주문상태','송금일시','송금금액',
      '입금내역','DMAX지갑','주문자이름','주문자전화',
      '수령인이름','수령인전화',
      '우편번호','주소',          // ← 분리됨
      '배송요청','택배사','운송장번호','발송날짜'
    ]);
    sheet.setFrozenRows(1);
    setupCourierDropdown(sheet);
  }
  
  var orderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  
  // ★ 주소 분리: 우편번호 따로, 기본주소+상세주소 따로
  var postcode = (p.postcode || '').trim();
  var address = '';
  if (p.addr) address += p.addr;
  if (p.addrDetail) address += ' ' + p.addrDetail;
  address = address.trim();
  
  var payProof = (p.payProof || '').trim();
  var dmaxWallet = (p.dmaxWallet || '').trim();
  
  sheet.appendRow([
    orderNum,
    timestamp,
    '주문접수',
    p.payDate || '',
    p.payAmount || '',
    payProof,
    dmaxWallet,
    p.ordererName || '',
    p.ordererPhone || '',
    p.receiverName || '',
    p.receiverPhone || '',
    postcode,           // 12열: 우편번호
    address,            // 13열: 주소
    p.deliveryNote || '',
    '',  // 택배사 (15열)
    '',  // 운송장번호 (16열)
    ''   // 발송날짜 (17열)
  ]);
  
  // 마지막 행에 링크 설정
  var lastRow = sheet.getLastRow();
  
  // 입금내역 (F열, 6번째) - TxID를 BscScan 트랜잭션 링크로
  if (payProof) {
    var txId = payProof;
    if (txId.indexOf('0x') === 0 || txId.length >= 60) {
      var txLink = 'https://bscscan.com/tx/' + txId;
      var richText1 = SpreadsheetApp.newRichTextValue()
        .setText(txId)
        .setLinkUrl(txLink)
        .build();
      sheet.getRange(lastRow, 6).setRichTextValue(richText1);
    }
  }
  
  // DMAX지갑 (G열, 7번째) - 지갑 주소를 BscScan 주소 링크로
  if (dmaxWallet) {
    if (dmaxWallet.indexOf('0x') === 0) {
      var walletLink = 'https://bscscan.com/address/' + dmaxWallet;
      var richText2 = SpreadsheetApp.newRichTextValue()
        .setText(dmaxWallet)
        .setLinkUrl(walletLink)
        .build();
      sheet.getRange(lastRow, 7).setRichTextValue(richText2);
    }
  }
  
  return { success: true, orderNumber: orderNum };
}

function searchOrders(name, phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = [];
  var searchName = (name || '').trim();
  var searchPhone = (phone || '').replace(/[^0-9]/g, '');
  
  // 신규주문 시트 검색
  var sheet1 = ss.getSheetByName('신규주문');
  if (sheet1 && sheet1.getLastRow() > 1) {
    var data1 = sheet1.getDataRange().getValues();
    for (var i = 1; i < data1.length; i++) {
      var row = data1[i];
      var rowName = String(row[7] || '').trim();
      var rowPhone = String(row[8] || '').replace(/[^0-9]/g, '');
      
      if (rowName === searchName && rowPhone === searchPhone) {
        orders.push({
          orderNumber: row[0] || '',
          timestamp: row[1] || '',
          status: row[2] || '주문접수',
          amount: row[4] || '',
          address: (row[11] ? row[11] + ' ' : '') + (row[12] || ''),  // 우편번호 + 주소
          courier: row[14] || '',      // 15열 → index 14
          trackingNumber: row[15] || '' // 16열 → index 15
        });
      }
    }
  }
  
  // 완료 시트 검색
  var sheet2 = ss.getSheetByName('완료');
  if (sheet2 && sheet2.getLastRow() > 1) {
    var data2 = sheet2.getDataRange().getValues();
    for (var j = 1; j < data2.length; j++) {
      var row2 = data2[j];
      var rowName2 = String(row2[7] || '').trim();
      var rowPhone2 = String(row2[8] || '').replace(/[^0-9]/g, '');
      
      if (rowName2 === searchName && rowPhone2 === searchPhone) {
        orders.push({
          orderNumber: row2[0] || '',
          timestamp: row2[1] || '',
          status: row2[2] || '배송중',
          amount: row2[4] || '',
          address: (row2[11] ? row2[11] + ' ' : '') + (row2[12] || ''),
          courier: row2[14] || '',
          trackingNumber: row2[15] || ''
        });
      }
    }
  }
  
  // 최신순 정렬
  orders.sort(function(a, b) {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });
  
  return { success: true, orders: orders };
}

/**
 * 배송용 사이트 데이터 조회
 * 신규주문 시트에서 아직 발송되지 않은 주문 목록 반환
 */
function getShippingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  var orders = [];
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, orders: [] };
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var payAmount = parseFloat(row[4]) || 0;
    var lastDigit = String(row[4] || '0').replace(/[^0-9]/g, '');
    lastDigit = lastDigit.charAt(lastDigit.length - 1);
    
    // 착불/선불 판별: 끝자리 0이면 착불, 그 외 선불
    var payType = (lastDigit === '0') ? '착불' : '선불';
    
    // 갯수 = 송금금액 / 10
    var quantity = Math.floor(payAmount / 10);
    
    orders.push({
      orderNumber: row[0] || '',
      timestamp: row[1] || '',
      status: row[2] || '',
      payAmount: row[4] || '',
      ordererName: row[7] || '',
      ordererPhone: row[8] || '',
      receiverName: row[9] || '',
      receiverPhone: row[10] || '',
      postcode: row[11] || '',
      address: row[12] || '',
      deliveryNote: row[13] || '',
      courier: row[14] || '',
      trackingNumber: row[15] || '',
      shippingDate: row[16] || '',
      payType: payType,
      quantity: quantity
    });
  }
  
  return { success: true, orders: orders };
}

/**
 * 택배사 드롭다운 설정
 * ★ 택배사가 O열(15열)로 이동
 */
function setupCourierDropdown(sheet) {
  var courierList = ['CJ대한통운', '롯데택배', '한진택배', '우체국', '로젠택배', '직접전달'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(courierList, true)
    .setAllowInvalid(false)
    .build();
  
  // O열 전체에 드롭다운 적용 (2행부터 1000행까지)
  sheet.getRange('O2:O1000').setDataValidation(rule);
}

/**
 * 시트 설정 함수 - 처음 한 번 실행
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow([
      '주문번호','접수시간','주문상태','송금일시','송금금액',
      '입금내역','DMAX지갑','주문자이름','주문자전화',
      '수령인이름','수령인전화',
      '우편번호','주소',
      '배송요청','택배사','운송장번호','발송날짜'
    ]);
    sheet.setFrozenRows(1);
  }
  
  setupCourierDropdown(sheet);
  Logger.log('택배사 드롭다운이 설정되었습니다.');
}

/**
 * 운송장번호 입력 시 자동으로 완료 시트로 이동
 * ★ 숫자뿐 아니라 아무 텍스트도 허용
 * ★ 이동 시 17열(발송날짜)에 현재 날짜 기록
 */
function onEdit(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return;
  }
  
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // ★ 운송장번호가 16열로 변경됨
    if (sheet.getName() !== '신규주문' || range.getColumn() !== 16) {
      return;
    }
    
    var trackingNumber = range.getValue();
    var row = range.getRow();
    
    // 헤더 행이거나 빈값이면 무시
    if (row === 1 || !trackingNumber || String(trackingNumber).trim() === '') {
      return;
    }
    
    var ss = e.source;
    var sourceSheet = ss.getSheetByName('신규주문');
    var targetSheet = ss.getSheetByName('완료');
    
    // 완료 시트가 없으면 생성
    if (!targetSheet) {
      targetSheet = ss.insertSheet('완료');
      targetSheet.appendRow([
        '주문번호','접수시간','주문상태','송금일시','송금금액',
        '입금내역','DMAX지갑','주문자이름','주문자전화',
        '수령인이름','수령인전화',
        '우편번호','주소',
        '배송요청','택배사','운송장번호','발송날짜'
      ]);
      targetSheet.setFrozenRows(1);
    }
    
    // ★ 17열까지 가져오기
    var rowData = sourceSheet.getRange(row, 1, 1, 17).getValues()[0];
    var orderNumber = rowData[0];
    
    // 이미 완료 시트에 같은 주문번호가 있으면 중복 추가 안함
    var completedData = targetSheet.getDataRange().getValues();
    for (var i = 1; i < completedData.length; i++) {
      if (completedData[i][0] === orderNumber) {
        sourceSheet.deleteRow(row);
        return;
      }
    }
    
    // 주문상태를 '배송중'으로 변경
    rowData[2] = '배송중';
    
    // ★ 발송날짜(17열, index 16)에 현재 날짜 기록
    rowData[16] = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    
    // 완료 시트에 추가
    targetSheet.appendRow(rowData);
    
    // 신규주문 시트에서 삭제
    sourceSheet.deleteRow(row);
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * 트리거 설정 함수 - 한 번만 실행
 */
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  Logger.log('트리거가 생성되었습니다. (기존 트리거 삭제 후 새로 생성)');
}

// ============================================================
// ★★★ 마이그레이션 함수 — 최초 1회만 실행 ★★★
// 실행 방법: Apps Script 편집기 > 함수 선택: migrateColumns > ▶ 실행
// ============================================================
// 기존 구조 (15열):
//   ... | 12:주소 | 13:배송요청 | 14:택배사 | 15:운송장번호
// 변환 후 (17열):
//   ... | 12:우편번호 | 13:주소 | 14:배송요청 | 15:택배사 | 16:운송장번호 | 17:발송날짜
// ============================================================
function migrateColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ---- 신규주문 시트 변환 ----
  var sheet1 = ss.getSheetByName('신규주문');
  if (sheet1) {
    migrateSheet(sheet1);
    // 택배사 드롭다운을 O열(15열)로 재설정
    setupCourierDropdown(sheet1);
    Logger.log('신규주문 시트 변환 완료');
  }
  
  // ---- 완료 시트 변환 ----
  var sheet2 = ss.getSheetByName('완료');
  if (sheet2 && sheet2.getLastRow() >= 1) {
    migrateSheet(sheet2);
    Logger.log('완료 시트 변환 완료');
  }
  
  Logger.log('=== 마이그레이션 완료! ===');
  Logger.log('반드시 createTrigger() 도 다시 실행하세요.');
}

function migrateSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  
  // ── 현재 상태 감지 ──
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // 이미 마이그레이션 완료된 경우 → 건너뛰기
  if (String(headerRow[11]).trim() === '우편번호' && String(headerRow[12]).trim() === '주소' && lastCol >= 17) {
    Logger.log(sheet.getName() + ': 이미 변환 완료된 상태 → 건너뜀');
    return;
  }
  
  // ── 전체 데이터를 한 번에 읽기 ──
  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  // ── 상태 판별 ──
  // Case A: 15열 (원본) — 12열에 "12345 주소" 합쳐진 상태
  // Case B: 16열 (이전 실행에서 insertColumn만 된 상태) — 12열에 "12345 주소", 13열 비어있음
  var isPartial = (lastCol === 16);
  
  Logger.log(sheet.getName() + ': 현재 ' + lastCol + '열 감지 → ' + (isPartial ? '부분변환 복구' : '신규변환'));
  
  // ── 새 데이터 배열 생성 (메모리에서 처리) ──
  var newData = [];
  
  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    
    if (i === 0) {
      // 헤더 행
      newData.push([
        '주문번호','접수시간','주문상태','송금일시','송금금액',
        '입금내역','DMAX지갑','주문자이름','주문자전화',
        '수령인이름','수령인전화',
        '우편번호','주소',
        '배송요청','택배사','운송장번호','발송날짜'
      ]);
      continue;
    }
    
    // 12열(index 11)에서 우편번호 분리
    var oldAddr = String(row[11] || '');
    var postcode = '';
    var address = oldAddr;
    
    var match = oldAddr.match(/^(\d{5})\s+(.*)$/);
    if (match) {
      postcode = match[1];
      address = match[2];
    }
    
    var newRow = [];
    // 1~11열 그대로
    for (var c = 0; c < 11; c++) {
      newRow.push(row[c] !== undefined ? row[c] : '');
    }
    newRow.push(postcode);  // 12열: 우편번호
    newRow.push(address);   // 13열: 주소
    
    if (isPartial) {
      // Case B: 16열 — insertColumn으로 밀린 상태
      // 기존 13열(비어있음) 건너뛰고, 14=배송요청, 15=택배사, 16=운송장번호
      newRow.push(row[13] || '');  // 14열: 배송요청
      newRow.push(row[14] || '');  // 15열: 택배사
      newRow.push(row[15] || '');  // 16열: 운송장번호
    } else {
      // Case A: 15열 — 원본 상태
      // 기존 13=배송요청, 14=택배사, 15=운송장번호
      newRow.push(row[12] || '');  // 14열: 배송요청
      newRow.push(row[13] || '');  // 15열: 택배사
      newRow.push(row[14] || '');  // 16열: 운송장번호
    }
    newRow.push('');  // 17열: 발송날짜 (신규)
    
    newData.push(newRow);
  }
  
  // ── 시트 초기화 후 일괄 쓰기 ──
  sheet.clear();
  SpreadsheetApp.flush();
  
  sheet.getRange(1, 1, newData.length, 17).setValues(newData);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  
  Logger.log(sheet.getName() + ': ' + (newData.length - 1) + '행 변환 완료');
}
