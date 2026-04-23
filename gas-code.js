// ============================================================
// Google Apps Script — 미네랄 주문서 (최종본)
// ============================================================
// 컬럼 구조 (17열):
//  A:주문번호  B:접수시간  C:주문상태  D:송금일시  E:송금금액
//  F:입금내역  G:DMAX지갑  H:주문자이름  I:주문자전화
//  J:수령인이름  K:수령인전화  L:우편번호  M:주소
//  N:배송요청  O:택배사  P:운송장번호  Q:발송날짜
// ============================================================

// ──────────────────────────────────────
// 1) 웹 요청 처리
// ──────────────────────────────────────
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

// ──────────────────────────────────────
// 2) 주문 저장
// ──────────────────────────────────────
function saveOrder(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow(getHeaders());
    sheet.setFrozenRows(1);
    setupCourierDropdown(sheet);
  }
  
  var orderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  
  // 우편번호와 주소를 분리해서 저장
  var postcode = (p.postcode || '').trim();
  var address = '';
  if (p.addr) address += p.addr;
  if (p.addrDetail) address += ' ' + p.addrDetail;
  address = address.trim();
  
  var payProof = (p.payProof || '').trim();
  var dmaxWallet = (p.dmaxWallet || '').trim();
  
  sheet.appendRow([
    orderNum,          // A: 주문번호
    timestamp,         // B: 접수시간
    '주문접수',         // C: 주문상태
    p.payDate || '',   // D: 송금일시
    p.payAmount || '', // E: 송금금액
    payProof,          // F: 입금내역
    dmaxWallet,        // G: DMAX지갑
    p.ordererName || '',   // H: 주문자이름
    p.ordererPhone || '',  // I: 주문자전화
    p.receiverName || '',  // J: 수령인이름
    p.receiverPhone || '', // K: 수령인전화
    postcode,          // L: 우편번호
    address,           // M: 주소
    p.deliveryNote || '', // N: 배송요청
    '',                // O: 택배사
    '',                // P: 운송장번호
    ''                 // Q: 발송날짜
  ]);
  
  var lastRow = sheet.getLastRow();
  
  // 입금내역 → BscScan 링크
  if (payProof && (payProof.indexOf('0x') === 0 || payProof.length >= 60)) {
    var richText1 = SpreadsheetApp.newRichTextValue()
      .setText(payProof)
      .setLinkUrl('https://bscscan.com/tx/' + payProof)
      .build();
    sheet.getRange(lastRow, 6).setRichTextValue(richText1);
  }
  
  // DMAX지갑 → BscScan 링크
  if (dmaxWallet && dmaxWallet.indexOf('0x') === 0) {
    var richText2 = SpreadsheetApp.newRichTextValue()
      .setText(dmaxWallet)
      .setLinkUrl('https://bscscan.com/address/' + dmaxWallet)
      .build();
    sheet.getRange(lastRow, 7).setRichTextValue(richText2);
  }
  
  return { success: true, orderNumber: orderNum };
}

// ──────────────────────────────────────
// 3) 주문 조회
// ──────────────────────────────────────
function searchOrders(name, phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orders = [];
  var searchName = (name || '').trim();
  var searchPhone = (phone || '').replace(/[^0-9]/g, '');
  
  var sheetNames = ['신규주문', '완료'];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowName = String(row[7] || '').trim();
      var rowPhone = String(row[8] || '').replace(/[^0-9]/g, '');
      
      if (rowName === searchName && rowPhone === searchPhone) {
        orders.push({
          orderNumber: row[0] || '',
          timestamp: row[1] || '',
          status: row[2] || (s === 0 ? '주문접수' : '배송중'),
          amount: row[4] || '',
          address: ((row[11] || '') + ' ' + (row[12] || '')).trim(),
          courier: row[14] || '',
          trackingNumber: row[15] || ''
        });
      }
    }
  }
  
  orders.sort(function(a, b) {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });
  
  return { success: true, orders: orders };
}

// ──────────────────────────────────────
// 4) 배송용 사이트 데이터 조회
// ──────────────────────────────────────
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
    
    // 착불/선불: 금액 끝자리가 0이면 착불
    var amountStr = String(row[4] || '0').replace(/[^0-9]/g, '');
    var lastDigit = amountStr.charAt(amountStr.length - 1);
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
      postcode: String(row[11] || ''),
      address: String(row[12] || ''),
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

// ──────────────────────────────────────
// 5) 운송장번호 입력 → 완료 시트 이동
//    아무 텍스트나 입력해도 이동됨
//    이동 시 Q열(17열)에 발송날짜 자동 기록
// ──────────────────────────────────────
function onEdit(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // P열(16열) = 운송장번호
    if (sheet.getName() !== '신규주문' || range.getColumn() !== 16) return;
    
    var trackingNumber = range.getValue();
    var row = range.getRow();
    
    if (row === 1 || !trackingNumber || String(trackingNumber).trim() === '') return;
    
    var ss = e.source;
    var sourceSheet = ss.getSheetByName('신규주문');
    var targetSheet = ss.getSheetByName('완료');
    
    if (!targetSheet) {
      targetSheet = ss.insertSheet('완료');
      targetSheet.appendRow(getHeaders());
      targetSheet.setFrozenRows(1);
    }
    
    var rowData = sourceSheet.getRange(row, 1, 1, 17).getValues()[0];
    var orderNumber = rowData[0];
    
    // 중복 체크
    var completedData = targetSheet.getDataRange().getValues();
    for (var i = 1; i < completedData.length; i++) {
      if (completedData[i][0] === orderNumber) {
        sourceSheet.deleteRow(row);
        return;
      }
    }
    
    rowData[2] = '배송중';
    rowData[16] = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    
    targetSheet.appendRow(rowData);
    sourceSheet.deleteRow(row);
    
  } finally {
    lock.releaseLock();
  }
}

// ──────────────────────────────────────
// 6) 유틸리티
// ──────────────────────────────────────
function getHeaders() {
  return [
    '주문번호','접수시간','주문상태','송금일시','송금금액',
    '입금내역','DMAX지갑','주문자이름','주문자전화',
    '수령인이름','수령인전화',
    '우편번호','주소',
    '배송요청','택배사','운송장번호','발송날짜'
  ];
}

function setupCourierDropdown(sheet) {
  var courierList = ['CJ대한통운', '롯데택배', '한진택배', '우체국', '로젠택배', '직접전달'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(courierList, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('O2:O1000').setDataValidation(rule);
}

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow(getHeaders());
    sheet.setFrozenRows(1);
  }
  setupCourierDropdown(sheet);
  Logger.log('설정 완료');
}

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
  Logger.log('트리거 생성 완료');
}

// ──────────────────────────────────────
// 7) ★ 마이그레이션 — 기존 데이터 변환 ★
//    실행: 함수 선택 > migrateColumns > ▶ 실행
// ──────────────────────────────────────
function migrateColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet1 = ss.getSheetByName('신규주문');
  if (sheet1) {
    migrateOneSheet(sheet1);
    setupCourierDropdown(sheet1);
  }
  
  var sheet2 = ss.getSheetByName('완료');
  if (sheet2 && sheet2.getLastRow() >= 1) {
    migrateOneSheet(sheet2);
  }
  
  Logger.log('=== 완료! createTrigger()도 실행하세요 ===');
}

function migrateOneSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  
  var lastCol = sheet.getLastColumn();
  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  // 이미 변환됨?
  if (String(allData[0][11] || '').trim() === '우편번호' && lastCol >= 17) {
    Logger.log(sheet.getName() + ' → 이미 변환됨');
    return;
  }
  
  Logger.log(sheet.getName() + ' → ' + lastCol + '열 감지, 변환 시작');
  
  var newData = [];
  for (var i = 0; i < allData.length; i++) {
    var row = allData[i];
    
    if (i === 0) {
      newData.push(getHeaders());
      continue;
    }
    
    // 우편번호 분리: "12345 서울시 강남구..." → "12345" + "서울시 강남구..."
    var combined = String(row[11] || '');
    var postcode = '';
    var addr = combined;
    var m = combined.match(/^(\d{5})\s+(.+)$/);
    if (m) {
      postcode = m[1];
      addr = m[2];
    }
    
    var newRow = [];
    for (var c = 0; c < 11; c++) newRow.push(row[c] !== undefined ? row[c] : '');
    
    newRow.push(postcode);  // L: 우편번호
    newRow.push(addr);      // M: 주소
    
    if (lastCol >= 16) {
      // 16열: 이전 실행에서 insertColumn된 상태
      newRow.push(row[13] || '');  // N: 배송요청
      newRow.push(row[14] || '');  // O: 택배사
      newRow.push(row[15] || '');  // P: 운송장번호
    } else {
      // 15열: 원본 상태
      newRow.push(row[12] || '');  // N: 배송요청
      newRow.push(row[13] || '');  // O: 택배사
      newRow.push(row[14] || '');  // P: 운송장번호
    }
    newRow.push('');  // Q: 발송날짜
    
    newData.push(newRow);
  }
  
  sheet.clear();
  SpreadsheetApp.flush();
  sheet.getRange(1, 1, newData.length, 17).setValues(newData);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  
  Logger.log(sheet.getName() + ' → ' + (newData.length - 1) + '행 완료!');
}

// ──────────────────────────────────────
// 8) 테스트 — shipping API 동작 확인
//    함수 선택 > testShipping > ▶ 실행
//    실행 로그에서 결과 확인
// ──────────────────────────────────────
function testShipping() {
  var result = getShippingData();
  Logger.log('주문 수: ' + result.orders.length);
  if (result.orders.length > 0) {
    var o = result.orders[0];
    Logger.log('첫 주문 → 우편번호: ' + o.postcode + ' | 주소: ' + o.address + ' | 구분: ' + o.payType + ' | 수량: ' + o.quantity);
  }
}
