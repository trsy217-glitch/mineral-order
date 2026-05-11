/**
 * ============================================================
 * 오륜 미네랄 주문 시스템 v3 (최종)
 * ============================================================
 * 
 * ★ 기존 GoogleAppsScript.js 코드를 전부 지우고
 *   이 코드를 통째로 붙여넣으세요 ★
 * 
 * 시트 구조: 18열
 * 주문번호 | 접수시간 | 주문상태 | 송금일시 | 주문수량 | 송금금액
 * 입금내역 | DMAX지갑 | 주문자이름 | 주문자전화
 * 수령인이름 | 수령인전화 | 우편번호 | 주소
 * 배송요청 | 택배사 | 운송장번호 | 발송날짜
 * 
 * 적용 순서:
 * 1. 기존 코드 전부 지우고 이 코드 붙여넣기 → 저장
 * 2. fixMisalignedRows 실행 (밀린 행 복구)
 * 3. createTrigger 실행
 * 4. 배포 > 배포 관리 > 연필 > 새 버전 > 배포
 * ============================================================
 */

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
    } else if (action === 'txcheck') {
      result = checkDuplicateTx(e.parameter.txid);
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
// ★ 핵심: 시트 구조 자동 감지 ★
// ──────────────────────────────────────
function detectColumns(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  if (lastCol >= 18 && String(headers[4] || '').trim() === '주문수량') {
    return {
      version: 18,
      orderQty: 4, payAmount: 5, payProof: 6, dmaxWallet: 7,
      ordererName: 8, ordererPhone: 9,
      receiverName: 10, receiverPhone: 11,
      postcode: 12, address: 13, deliveryNote: 14,
      courier: 15, tracking: 16, shippingDate: 17,
      totalCols: 18
    };
  }
  
  return {
    version: 17,
    orderQty: -1, payAmount: 4, payProof: 5, dmaxWallet: 6,
    ordererName: 7, ordererPhone: 8,
    receiverName: 9, receiverPhone: 10,
    postcode: 11, address: 12, deliveryNote: 13,
    courier: 14, tracking: 15, shippingDate: 16,
    totalCols: 17
  };
}

// ──────────────────────────────────────
// 주문 저장 — 10병 단위로 분할 저장
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
  
  var col = detectColumns(sheet);
  
  var baseOrderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  
  var postcode = (p.postcode || '').trim();
  var address = '';
  if (p.addr) address += p.addr;
  if (p.addrDetail) address += ' ' + p.addrDetail;
  address = address.trim();
  
  var payProof = (p.payProof || '').trim();
  var dmaxWallet = (p.dmaxWallet || '').trim();
  
  // ★ 서버 측 TxID 중복 검사 (클라이언트 우회 방지) ★
  if (payProof) {
    var dupResult = checkDuplicateTx(payProof);
    if (dupResult.duplicate) {
      return { success: false, error: 'DUPLICATE_TX', message: '이미 사용된 TxID입니다.' };
    }
  }
  
  var totalQty = parseInt(p.orderQty) || 1;
  var totalAmount = parseFloat(p.payAmount) || 0;
  
  // 10병 단위로 분할: 30병 → [10, 10, 10], 25병 → [10, 10, 5], 3병 → [3]
  var chunks = [];
  var remaining = totalQty;
  while (remaining > 0) {
    var chunkQty = Math.min(remaining, 10);
    chunks.push(chunkQty);
    remaining -= chunkQty;
  }
  
  var numChunks = chunks.length;
  
  for (var c = 0; c < numChunks; c++) {
    // 주문번호: 1건이면 그대로, 여러 건이면 -1, -2, -3 붙임
    var orderNum = (numChunks === 1) ? baseOrderNum : baseOrderNum + '-' + (c + 1);
    var chunkQty = chunks[c];
    
    var rowData;
    
    if (col.version === 18) {
      rowData = [
        orderNum, timestamp, '주문접수', p.payDate || '',
        chunkQty,
        totalAmount,       // 전체 송금금액 (첫 행에만 의미, 참고용으로 모든 행에 기록)
        payProof, dmaxWallet,
        p.ordererName || '', p.ordererPhone || '',
        p.receiverName || '', p.receiverPhone || '',
        postcode, address, p.deliveryNote || '',
        '', '', ''
      ];
    } else {
      rowData = [
        orderNum, timestamp, '주문접수', p.payDate || '',
        totalAmount,
        payProof, dmaxWallet,
        p.ordererName || '', p.ordererPhone || '',
        p.receiverName || '', p.receiverPhone || '',
        postcode, address, p.deliveryNote || '',
        '', '', ''
      ];
    }
    
    sheet.appendRow(rowData);
    
    var lastRow = sheet.getLastRow();
    
    // 첫 행에만 입금내역/DMAX 링크 설정
    if (c === 0) {
      var proofCol = col.payProof + 1;
      if (payProof && (payProof.indexOf('0x') === 0 || payProof.length >= 60)) {
        var richText1 = SpreadsheetApp.newRichTextValue()
          .setText(payProof)
          .setLinkUrl('https://bscscan.com/tx/' + payProof)
          .build();
        sheet.getRange(lastRow, proofCol).setRichTextValue(richText1);
      }
      
      var walletCol = col.dmaxWallet + 1;
      if (dmaxWallet && dmaxWallet.indexOf('0x') === 0) {
        var richText2 = SpreadsheetApp.newRichTextValue()
          .setText(dmaxWallet)
          .setLinkUrl('https://bscscan.com/address/' + dmaxWallet)
          .build();
        sheet.getRange(lastRow, walletCol).setRichTextValue(richText2);
      }
    }
  }
  
  // ★ 지갑 주소 추적 시트 업데이트 ★
  if (dmaxWallet && dmaxWallet.indexOf('0x') === 0) {
    try {
      updateWalletTracking(dmaxWallet, p.ordererName || '', totalQty, totalAmount, timestamp);
    } catch(e) {
      // 지갑 추적 실패해도 주문은 정상 처리
      Logger.log('지갑 추적 오류: ' + e.toString());
    }
  }
  
  return { 
    success: true, 
    orderNumber: baseOrderNum,
    totalQty: totalQty,
    chunks: numChunks
  };
}

// ──────────────────────────────────────
// TxID 중복 검사
// ──────────────────────────────────────
function checkDuplicateTx(txid) {
  if (!txid) return { success: true, duplicate: false };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 입력값에서 URL 접두사 제거 후 비교
  var searchTx = stripTxPrefix(txid.trim()).toLowerCase();
  if (!searchTx) return { success: true, duplicate: false };
  
  var sheetNames = ['신규주문', '완료'];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    
    var col = detectColumns(sheet);
    // getDisplayValues로 읽어야 하이퍼링크가 아닌 표시 텍스트를 가져옴
    var data = sheet.getDataRange().getDisplayValues();
    
    for (var i = 1; i < data.length; i++) {
      var rowTx = stripTxPrefix(String(data[i][col.payProof] || '').trim()).toLowerCase();
      if (rowTx && rowTx === searchTx) {
        return { success: true, duplicate: true };
      }
    }
  }
  
  return { success: true, duplicate: false };
}

// URL 접두사 제거 헬퍼
function stripTxPrefix(val) {
  var prefixes = ['https://bscscan.com/tx/', 'http://bscscan.com/tx/', 'https://www.bscscan.com/tx/'];
  for (var i = 0; i < prefixes.length; i++) {
    if (val.indexOf(prefixes[i]) === 0) {
      return val.substring(prefixes[i].length);
    }
  }
  return val;
}

// ──────────────────────────────────────
// 주문 조회
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
    
    var col = detectColumns(sheet);
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowName = String(row[col.ordererName] || '').trim();
      var rowPhone = String(row[col.ordererPhone] || '').replace(/[^0-9]/g, '');
      
      if (rowName === searchName && rowPhone === searchPhone) {
        // 운송장번호: Date 객체나 Invalid Date 처리
        var tn = row[col.tracking];
        if (tn instanceof Date) {
          // 시트에서 displayValue로 다시 읽기
          try { tn = sheet.getRange(i + 1, col.tracking + 1).getDisplayValue(); } catch(e2) { tn = ''; }
        }
        tn = String(tn || '').trim();
        if (tn === 'Invalid Date') tn = '';
        
        // 주문번호에서 부모 주문번호 추출 (ORN260506...-2 → ORN260506...)
        var orderNum = String(row[0] || '');
        var parentOrder = orderNum;
        var subIndex = 0;
        var dashMatch = orderNum.match(/^(.+)-(\d+)$/);
        if (dashMatch) {
          parentOrder = dashMatch[1];
          subIndex = parseInt(dashMatch[2]);
        }
        
        orders.push({
          orderNumber: orderNum,
          parentOrder: parentOrder,
          subIndex: subIndex,
          timestamp: row[1] || '',
          status: row[2] || (s === 0 ? '주문접수' : '배송중'),
          quantity: col.orderQty >= 0 ? (row[col.orderQty] || '') : '',
          amount: row[col.payAmount] || '',
          address: ((row[col.postcode] || '') + ' ' + (row[col.address] || '')).trim(),
          courier: row[col.courier] || '',
          trackingNumber: tn
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
// 배송용 데이터 조회
// ──────────────────────────────────────
function getShippingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  var orders = [];
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, orders: [] };
  }
  
  var col = detectColumns(sheet);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var payAmount = parseFloat(row[col.payAmount]) || 0;
    var quantity = (col.orderQty >= 0) ? (parseInt(row[col.orderQty]) || 0) : 0;
    
    var amountStr = String(row[col.payAmount] || '0').replace(/[^0-9]/g, '');
    var lastDigit = amountStr.charAt(amountStr.length - 1);
    var payType = (lastDigit === '0') ? '착불' : '선불';
    
    if (quantity === 0 && payAmount > 0) {
      quantity = Math.floor(payAmount / 10);
    }
    
    orders.push({
      orderNumber: row[0] || '',
      timestamp: row[1] || '',
      status: row[2] || '',
      orderQty: quantity,
      payAmount: row[col.payAmount] || '',
      ordererName: row[col.ordererName] || '',
      ordererPhone: row[col.ordererPhone] || '',
      receiverName: row[col.receiverName] || '',
      receiverPhone: row[col.receiverPhone] || '',
      postcode: String(row[col.postcode] || ''),
      address: String(row[col.address] || ''),
      deliveryNote: row[col.deliveryNote] || '',
      courier: row[col.courier] || '',
      trackingNumber: row[col.tracking] || '',
      shippingDate: row[col.shippingDate] || '',
      payType: payType,
      quantity: quantity
    });
  }
  
  return { success: true, orders: orders };
}

// ──────────────────────────────────────
// 운송장번호 입력 → 완료 시트 이동
// ──────────────────────────────────────
function onEdit(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== '신규주문') return;
    
    var col = detectColumns(sheet);
    var courierCol1 = col.courier + 1;
    var trackingCol1 = col.tracking + 1;
    
    var range = e.range;
    var startCol = range.getColumn();
    var endCol = startCol + range.getNumColumns() - 1;
    
    if (endCol < courierCol1 || startCol > trackingCol1) return;
    
    var startRow = range.getRow();
    var numRows = range.getNumRows();
    
    var ss = e.source;
    var sourceSheet = ss.getSheetByName('신규주문');
    var targetSheet = ss.getSheetByName('완료');
    
    if (!targetSheet) {
      targetSheet = ss.insertSheet('완료');
      targetSheet.appendRow(getHeaders());
      targetSheet.setFrozenRows(1);
    }
    
    var rowsToDelete = [];
    
    for (var r = startRow + numRows - 1; r >= startRow; r--) {
      if (r === 1) continue;
      
      var trackingNumber = sourceSheet.getRange(r, trackingCol1).getValue();
      if (!trackingNumber || String(trackingNumber).trim() === '') continue;
      
      var rowData = sourceSheet.getRange(r, 1, 1, col.totalCols).getValues()[0];
      var orderNumber = rowData[0];
      
      if (!orderNumber || String(orderNumber).trim() === '') continue;
      
      var isDuplicate = false;
      var completedData = targetSheet.getDataRange().getValues();
      for (var i = 1; i < completedData.length; i++) {
        // 정확히 같은 주문번호(서브번호 포함)로 중복 체크
        if (String(completedData[i][0]) === String(orderNumber)) {
          isDuplicate = true;
          break;
        }
      }
      
      if (isDuplicate) {
        rowsToDelete.push(r);
        continue;
      }
      
      rowData[col.tracking] = String(rowData[col.tracking]).replace(/\.0$/, '');
      // Invalid Date 방지: Date 객체면 원본 셀에서 다시 읽기
      var trackVal = rowData[col.tracking];
      if (trackVal instanceof Date || String(trackVal) === 'Invalid Date') {
        trackVal = sourceSheet.getRange(r, trackingCol1).getDisplayValue();
      }
      rowData[col.tracking] = String(trackVal).replace(/\.0$/, '');
      
      rowData[2] = '배송중';
      rowData[col.shippingDate] = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
      
      targetSheet.appendRow(rowData);
      
      // 완료 시트에서 운송장번호 셀을 텍스트 서식으로 강제 설정
      var targetLastRow = targetSheet.getLastRow();
      targetSheet.getRange(targetLastRow, col.tracking + 1).setNumberFormat('@').setValue(String(trackVal).replace(/\.0$/, ''));
      
      rowsToDelete.push(r);
    }
    
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var d = 0; d < rowsToDelete.length; d++) {
      sourceSheet.deleteRow(rowsToDelete[d]);
    }
    
  } finally {
    lock.releaseLock();
  }
}

// ──────────────────────────────────────
// 밀린 주문 일괄 이동
// ──────────────────────────────────────
function moveStuckOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('신규주문');
  var targetSheet = ss.getSheetByName('완료');
  
  if (!sourceSheet) { Logger.log('신규주문 시트 없음'); return; }
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet('완료');
    targetSheet.appendRow(getHeaders());
    targetSheet.setFrozenRows(1);
  }
  
  var lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) { Logger.log('이동할 주문 없음'); return; }
  
  var col = detectColumns(sourceSheet);
  var data = sourceSheet.getRange(2, 1, lastRow - 1, col.totalCols).getValues();
  var completedData = targetSheet.getDataRange().getValues();
  
  var completedOrders = {};
  for (var c = 1; c < completedData.length; c++) {
    completedOrders[completedData[c][0]] = true;
  }
  
  var rowsToDelete = [];
  var movedCount = 0;
  
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var orderNumber = row[0];
    var trackingNumber = row[col.tracking];
    
    if (!orderNumber || String(orderNumber).trim() === '') {
      var isEmpty = true;
      for (var k = 0; k < row.length; k++) {
        if (row[k] !== '' && row[k] !== null && row[k] !== undefined) { isEmpty = false; break; }
      }
      if (isEmpty) rowsToDelete.push(i + 2);
      continue;
    }
    
    if (!trackingNumber || String(trackingNumber).trim() === '') continue;
    if (completedOrders[orderNumber]) { rowsToDelete.push(i + 2); continue; }
    
    row[col.tracking] = String(row[col.tracking]).replace(/\.0$/, '');
    row[2] = '배송중';
    row[col.shippingDate] = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    
    targetSheet.appendRow(row);
    completedOrders[orderNumber] = true;
    rowsToDelete.push(i + 2);
    movedCount++;
  }
  
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var d = 0; d < rowsToDelete.length; d++) {
    sourceSheet.deleteRow(rowsToDelete[d]);
  }
  
  Logger.log('=== ' + movedCount + '건 이동, ' + rowsToDelete.length + '행 정리 ===');
}

// ──────────────────────────────────────
// 빈 행 정리
// ──────────────────────────────────────
function cleanEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  var col = detectColumns(sheet);
  var data = sheet.getRange(2, 1, lastRow - 1, col.totalCols).getValues();
  var deleteCount = 0;
  for (var i = data.length - 1; i >= 0; i--) {
    if (!data[i][0] || String(data[i][0]).trim() === '') {
      sheet.deleteRow(i + 2);
      deleteCount++;
    }
  }
  Logger.log('빈 행 ' + deleteCount + '개 삭제');
}

// ──────────────────────────────────────
// 유틸리티
// ──────────────────────────────────────
function getHeaders() {
  return [
    '주문번호','접수시간','주문상태','송금일시','주문수량','송금금액',
    '입금내역','DMAX지갑','주문자이름','주문자전화',
    '수령인이름','수령인전화','우편번호','주소',
    '배송요청','택배사','운송장번호','발송날짜'
  ];
}

function setupCourierDropdown(sheet) {
  var col = detectColumns(sheet);
  var courierList = ['CJ대한통운', '롯데택배', '한진택배', '우체국', '로젠택배', '직접전달'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(courierList, true)
    .setAllowInvalid(false)
    .build();
  var colLetter = col.version === 18 ? 'P' : 'O';
  sheet.getRange(colLetter + '2:' + colLetter + '1000').setDataValidation(rule);
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
// ★ 밀린 행 복구 ★
// 18열 시트에 17열 데이터로 저장된 행을 자동 감지하여 수정
// ──────────────────────────────────────
function fixMisalignedRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('신규주문');
  if (!sheet) return;
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  var data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
  var fixedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var val5 = String(row[5] || '').trim();
    var val8 = String(row[8] || '').trim();
    
    var isMisaligned = false;
    if (val5.indexOf('0x') === 0) isMisaligned = true;
    if (val8.match(/^010-\d{4}-\d{4}$/)) isMisaligned = true;
    
    if (!isMisaligned) continue;
    
    var fixedRow = [];
    fixedRow.push(row[0]); // 주문번호
    fixedRow.push(row[1]); // 접수시간
    fixedRow.push(row[2]); // 주문상태
    fixedRow.push(row[3]); // 송금일시
    
    var payAmount = parseFloat(row[4]) || 0;
    var qty = Math.floor(payAmount / 10);
    fixedRow.push(qty);    // 주문수량 (역산)
    
    for (var c = 4; c <= 16; c++) {
      fixedRow.push(row[c] !== undefined ? row[c] : '');
    }
    
    sheet.getRange(i + 2, 1, 1, 18).setValues([fixedRow]);
    fixedCount++;
    Logger.log('행 ' + (i + 2) + ' 복구: ' + row[0] + ', 수량=' + qty);
  }
  
  Logger.log('=== 총 ' + fixedCount + '건 복구 완료 ===');
}

// ──────────────────────────────────────
// ★ 완료 시트 운송장번호 복구 ★
// Invalid Date나 깨진 운송장번호를 displayValue로 복구
// ──────────────────────────────────────
function fixTrackingNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('완료');
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log('완료 시트 없음'); return; }
  
  var col = detectColumns(sheet);
  var trackCol1 = col.tracking + 1; // 1-based
  var lastRow = sheet.getLastRow();
  var fixedCount = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    var cell = sheet.getRange(r, trackCol1);
    var val = cell.getValue();
    var display = cell.getDisplayValue();
    
    // Invalid Date 또는 Date 객체인 경우
    if (val instanceof Date || String(val) === 'Invalid Date' || display === 'Invalid Date') {
      // 원본 값을 복구할 수 없으면 빈값으로
      Logger.log('행 ' + r + ': 운송장 깨짐 (val=' + val + ', display=' + display + ')');
      // 셀을 텍스트 서식으로 변경
      cell.setNumberFormat('@');
      if (display && display !== 'Invalid Date') {
        cell.setValue(display);
      }
      fixedCount++;
    } else if (val && typeof val === 'number') {
      // 숫자로 저장된 운송장번호를 텍스트로 변환
      cell.setNumberFormat('@').setValue(String(val).replace(/\.0$/, ''));
      fixedCount++;
    }
  }
  
  // 운송장번호 열 전체를 텍스트 서식으로 설정
  sheet.getRange(2, trackCol1, lastRow - 1, 1).setNumberFormat('@');
  
  Logger.log('=== ' + fixedCount + '건 운송장번호 복구 완료 ===');
}

// ──────────────────────────────────────
// 진단용
// ──────────────────────────────────────
function diagnosRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('신규주문');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  Logger.log('=== 헤더 ===');
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var h = 0; h < headers.length; h++) {
    Logger.log('  [' + h + '] ' + headers[h]);
  }
  
  var startRow = Math.max(2, lastRow - 4);
  Logger.log('');
  Logger.log('=== 행 ' + startRow + '~' + lastRow + ' ===');
  
  for (var r = startRow; r <= lastRow; r++) {
    var row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    Logger.log('--- 행 ' + r + ' ---');
    Logger.log('  [0] 주문번호: ' + row[0]);
    Logger.log('  [3] 송금일시: ' + row[3]);
    Logger.log('  [4] 주문수량: ' + row[4]);
    Logger.log('  [5] 송금금액: ' + row[5]);
    Logger.log('  [6] 입금내역: ' + String(row[6]).substring(0, 30));
    Logger.log('  [8] 주문자이름: ' + row[8]);
    Logger.log('  [15] 택배사: ' + row[15]);
    Logger.log('  [16] 운송장번호: ' + row[16]);
  }
}

function checkSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('신규주문');
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('열 수: ' + lastCol);
  Logger.log('행 수: ' + lastRow);
  Logger.log('헤더: ' + headers.join(' | '));
  Logger.log('감지 버전: ' + detectColumns(sheet).version + '열 구조');
}

// ──────────────────────────────────────
// 택배사별 출력 시트
// ──────────────────────────────────────
function generateCourierSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcSheet = ss.getSheetByName('완료');
  
  if (!srcSheet || srcSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('완료 시트에 데이터가 없습니다.');
    return;
  }
  
  var col = detectColumns(srcSheet);
  var data = srcSheet.getDataRange().getValues();
  var grouped = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var courier = String(row[col.courier] || '').trim();
    var trackingRaw = row[col.tracking];
    
    if (!courier && String(trackingRaw).trim() === '공장배송') courier = '공장배송';
    if (!courier) courier = '기타';
    
    var rawPostcode = String(row[col.postcode] || '').trim();
    var rawAddr = String(row[col.address] || '').trim();
    var postcode = '', address = '';
    
    var m = rawPostcode.match(/^(\d{5})\s+(.+)$/);
    if (m) { postcode = m[1]; address = m[2]; }
    else if (rawPostcode.match(/^\d{4,5}(\.0)?$/)) {
      postcode = rawPostcode.replace(/\.0$/, '');
      while (postcode.length < 5) postcode = '0' + postcode;
      address = rawAddr;
    } else { postcode = rawPostcode.replace(/\.0$/, ''); address = rawAddr; }
    
    var tracking = '';
    if (trackingRaw && String(trackingRaw).trim() !== '공장배송') {
      tracking = String(trackingRaw).replace(/\.0$/, '').trim();
      if (tracking.indexOf('E') !== -1 || tracking.indexOf('e') !== -1)
        tracking = String(Number(trackingRaw).toFixed(0));
    }
    
    var quantity = (col.orderQty >= 0) ? (parseInt(row[col.orderQty]) || 0) : 0;
    var amount = parseFloat(row[col.payAmount]) || 0;
    if (quantity === 0 && amount > 0) quantity = Math.floor(amount / 10);
    
    var amountStr = String(row[col.payAmount] || '0').replace(/[^0-9]/g, '');
    var lastDigit = amountStr.charAt(amountStr.length - 1);
    var payType = (lastDigit === '0') ? '착불' : '선불';
    
    if (!grouped[courier]) grouped[courier] = [];
    grouped[courier].push({
      orderNumber: row[0],
      receiverName: String(row[col.receiverName] || '').trim(),
      receiverPhone: String(row[col.receiverPhone] || '').trim(),
      postcode: postcode, address: address,
      deliveryNote: String(row[col.deliveryNote] || '').trim(),
      tracking: tracking, quantity: quantity, payType: payType,
      amount: amount, shippingDate: row[col.shippingDate] || ''
    });
  }
  
  var outSheet = ss.getSheetByName('택배출력');
  if (outSheet) outSheet.clear();
  else outSheet = ss.insertSheet('택배출력');
  
  var currentRow = 1;
  var courierOrder = ['한진택배','CJ대한통운','롯데택배','우체국','로젠택배','공장배송','직접전달','기타'];
  var sortedKeys = [];
  for (var k = 0; k < courierOrder.length; k++) {
    if (grouped[courierOrder[k]]) sortedKeys.push(courierOrder[k]);
  }
  for (var key in grouped) {
    if (sortedKeys.indexOf(key) === -1) sortedKeys.push(key);
  }
  
  for (var s = 0; s < sortedKeys.length; s++) {
    var cn = sortedKeys[s];
    var ords = grouped[cn];
    if (cn === '공장배송' || cn === '직접전달')
      currentRow = writeFactorySection(outSheet, currentRow, cn, ords);
    else
      currentRow = writeCourierSection(outSheet, currentRow, cn, ords);
    currentRow += 1;
  }
  
  outSheet.setColumnWidth(1,140); outSheet.setColumnWidth(2,80);
  outSheet.setColumnWidth(3,120); outSheet.setColumnWidth(4,70);
  outSheet.setColumnWidth(5,300); outSheet.setColumnWidth(6,50);
  outSheet.setColumnWidth(7,60);  outSheet.setColumnWidth(8,180);
  outSheet.setColumnWidth(9,90);
  ss.setActiveSheet(outSheet);
  SpreadsheetApp.getUi().alert('택배출력 시트 생성 완료\n' + sortedKeys.join(', '));
}

function writeCourierSection(sheet, startRow, courierName, orders) {
  var row = startRow;
  sheet.getRange(row,1).setValue('📦 '+courierName+' ('+orders.length+'건)');
  sheet.getRange(row,1,1,9).merge();
  sheet.getRange(row,1).setFontWeight('bold').setFontSize(12).setBackground('#e8f5e9').setFontColor('#2e7d32');
  row++;
  var h = ['운송장번호','받는분','전화번호','우편번호','주소','수량','운임','배송메모','발송일'];
  sheet.getRange(row,1,1,h.length).setValues([h]);
  sheet.getRange(row,1,1,h.length).setFontWeight('bold').setBackground('#c8e6c9').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true);
  row++;
  for (var i=0;i<orders.length;i++) {
    var o=orders[i]; var ds='';
    if(o.shippingDate){try{ds=Utilities.formatDate(new Date(o.shippingDate),'Asia/Seoul','yyyy-MM-dd');}catch(e){ds=String(o.shippingDate);}}
    sheet.getRange(row,1,1,9).setValues([[o.tracking,o.receiverName,o.receiverPhone,o.postcode,o.address,o.quantity,o.payType,o.deliveryNote,ds]]);
    sheet.getRange(row,1).setNumberFormat('@'); sheet.getRange(row,4).setNumberFormat('@');
    row++;
  }
  if(orders.length>0) sheet.getRange(startRow+2,1,orders.length,9).setBorder(true,true,true,true,true,true,'#bdbdbd',SpreadsheetApp.BorderStyle.SOLID);
  return row;
}

function writeFactorySection(sheet, startRow, courierName, orders) {
  var row = startRow;
  var emoji = courierName==='공장배송'?'🏭':'🤝';
  sheet.getRange(row,1).setValue(emoji+' '+courierName+' ('+orders.length+'건)');
  sheet.getRange(row,1,1,8).merge();
  sheet.getRange(row,1).setFontWeight('bold').setFontSize(12).setBackground('#fff3e0').setFontColor('#e65100');
  row++;
  var h=['받는분','전화번호','우편번호','주소','수량','운임','배송메모','발송일'];
  sheet.getRange(row,1,1,h.length).setValues([h]);
  sheet.getRange(row,1,1,h.length).setFontWeight('bold').setBackground('#ffe0b2').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true);
  row++;
  for(var i=0;i<orders.length;i++){
    var o=orders[i]; var ds='';
    if(o.shippingDate){try{ds=Utilities.formatDate(new Date(o.shippingDate),'Asia/Seoul','yyyy-MM-dd');}catch(e){ds=String(o.shippingDate);}}
    sheet.getRange(row,1,1,8).setValues([[o.receiverName,o.receiverPhone,o.postcode,o.address,o.quantity,o.payType,o.deliveryNote,ds]]);
    sheet.getRange(row,3).setNumberFormat('@'); row++;
  }
  if(orders.length>0) sheet.getRange(startRow+2,1,orders.length,8).setBorder(true,true,true,true,true,true,'#bdbdbd',SpreadsheetApp.BorderStyle.SOLID);
  return row;
}

// ──────────────────────────────────────
// ★ 지갑 주소 추적 시트 ★
// ──────────────────────────────────────

// 지갑추적 시트 헤더
function getWalletHeaders() {
  return ['지갑주소', '주문자이름', '최초등록일', '총주문횟수', '총주문수량', '총송금금액', '스테이킹금액', '주문가능수량', '잔여수량', '상태', '최근주문일'];
}

// 지갑별 주문내역 헤더
function getWalletDetailHeaders() {
  return ['지갑주소', '주문번호', '주문일시', '주문자이름', '주문수량', '송금금액', 'TxID'];
}

// 주문 시 지갑 추적 업데이트
function updateWalletTracking(walletAddr, ordererName, qty, amount, timestamp) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // === 지갑요약 시트 ===
  var summarySheet = ss.getSheetByName('지갑추적');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('지갑추적');
    summarySheet.appendRow(getWalletHeaders());
    summarySheet.setFrozenRows(1);
    summarySheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#e8f5e9');
    setWalletColumnWidths(summarySheet);
  }
  
  var wallet = walletAddr.trim().toLowerCase();
  
  // 기존 지갑 찾기
  var summaryData = summarySheet.getDataRange().getDisplayValues();
  var foundRow = -1;
  for (var i = 1; i < summaryData.length; i++) {
    if (String(summaryData[i][0] || '').trim().toLowerCase() === wallet) {
      foundRow = i + 1;
      break;
    }
  }
  
  if (foundRow > 0) {
    // 기존 지갑 업데이트
    var prevCount = parseInt(summarySheet.getRange(foundRow, 4).getValue()) || 0;
    var prevQty = parseInt(summarySheet.getRange(foundRow, 5).getValue()) || 0;
    var prevAmount = parseFloat(summarySheet.getRange(foundRow, 6).getValue()) || 0;
    
    var newQty = prevQty + qty;
    
    summarySheet.getRange(foundRow, 2).setValue(ordererName);
    summarySheet.getRange(foundRow, 4).setValue(prevCount + 1);
    summarySheet.getRange(foundRow, 5).setValue(newQty);
    summarySheet.getRange(foundRow, 6).setValue(prevAmount + amount);
    // G열(스테이킹금액)은 수기 입력이므로 건드리지 않음
    // H,I,J열은 수식이 자동 계산
    summarySheet.getRange(foundRow, 11).setValue(timestamp);
  } else {
    // 신규 지갑 등록
    var newRowNum = summarySheet.getLastRow() + 1;
    summarySheet.getRange(newRowNum, 1, 1, 7).setValues([[
      walletAddr, ordererName, timestamp, 1, qty, amount, 0
    ]]);
    // H열: 주문가능수량 수식
    summarySheet.getRange(newRowNum, 8).setFormula('=IF(G' + newRowNum + '=0,0,FLOOR(G' + newRowNum + '/100))');
    // I열: 잔여수량 수식
    summarySheet.getRange(newRowNum, 9).setFormula('=H' + newRowNum + '-E' + newRowNum);
    // J열: 상태 수식
    summarySheet.getRange(newRowNum, 10).setFormula('=IF(G' + newRowNum + '=0,"미설정",IF(E' + newRowNum + '>H' + newRowNum + ',"초과","정상"))');
    // K열: 최근주문일
    summarySheet.getRange(newRowNum, 11).setValue(timestamp);
    // G열 숫자 서식
    summarySheet.getRange(newRowNum, 7).setNumberFormat('#,##0');
    
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(walletAddr)
      .setLinkUrl('https://bscscan.com/address/' + walletAddr)
      .build();
    summarySheet.getRange(newRowNum, 1).setRichTextValue(richText);
  }
  
  // === 지갑주문내역 시트 (상세 로그) ===
  var detailSheet = ss.getSheetByName('지갑주문내역');
  if (!detailSheet) {
    detailSheet = ss.insertSheet('지갑주문내역');
    detailSheet.appendRow(getWalletDetailHeaders());
    detailSheet.setFrozenRows(1);
    detailSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e3f2fd');
    detailSheet.setColumnWidth(1, 380);
    detailSheet.setColumnWidth(2, 160);
    detailSheet.setColumnWidth(3, 150);
    detailSheet.setColumnWidth(4, 100);
    detailSheet.setColumnWidth(5, 80);
    detailSheet.setColumnWidth(6, 100);
    detailSheet.setColumnWidth(7, 380);
  }
  
  // 주문번호는 saveOrder에서 생성된 것을 사용 (여기서는 시간 기반으로 매칭)
  var orderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  
  detailSheet.appendRow([
    walletAddr,
    orderNum,
    timestamp,
    ordererName,
    qty,
    amount,
    ''  // TxID는 아래에서 별도 설정
  ]);
}

// 기존 주문 데이터로 지갑추적 시트 일괄 생성
// 실행: 함수 선택 > buildWalletTracking > ▶ 실행
function buildWalletTracking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 기존 스테이킹 금액 보존
  var existingStaking = {};
  var oldSheet = ss.getSheetByName('지갑추적');
  if (oldSheet && oldSheet.getLastRow() > 1) {
    var oldData = oldSheet.getDataRange().getValues();
    for (var o = 1; o < oldData.length; o++) {
      var addr = String(oldData[o][0] || '').trim().toLowerCase();
      var staking = parseFloat(oldData[o][6]) || 0;
      if (addr && staking > 0) {
        existingStaking[addr] = staking;
      }
    }
  }
  
  // 시트 초기화
  var summarySheet = oldSheet;
  if (summarySheet) {
    summarySheet.clear();
    summarySheet.clearConditionalFormatRules();
  }
  else summarySheet = ss.insertSheet('지갑추적');
  summarySheet.appendRow(getWalletHeaders());
  summarySheet.setFrozenRows(1);
  summarySheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#e8f5e9');
  
  var detailSheet = ss.getSheetByName('지갑주문내역');
  if (detailSheet) detailSheet.clear();
  else detailSheet = ss.insertSheet('지갑주문내역');
  detailSheet.appendRow(getWalletDetailHeaders());
  detailSheet.setFrozenRows(1);
  detailSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e3f2fd');
  
  // 데이터 수집
  var wallets = {};
  var sheetNames = ['신규주문', '완료'];
  for (var s = 0; s < sheetNames.length; s++) {
    var sheet = ss.getSheetByName(sheetNames[s]);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var col = detectColumns(sheet);
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var walletAddr = String(row[col.dmaxWallet] || '').trim();
      if (!walletAddr || walletAddr.indexOf('0x') !== 0) continue;
      var walletKey = walletAddr.toLowerCase();
      var orderNum = row[0];
      var orderTime = row[1];
      var ordererName = row[col.ordererName] || '';
      var qty = parseInt(row[col.orderQty >= 0 ? col.orderQty : col.payAmount]) || 0;
      var amount = parseFloat(row[col.payAmount]) || 0;
      var txid = row[col.payProof] || '';
      var parentOrder = orderNum.replace(/-\d+$/, '');
      if (!wallets[walletKey]) {
        wallets[walletKey] = {
          address: walletAddr, name: ordererName,
          firstDate: orderTime, lastDate: orderTime,
          orders: {}, totalQty: 0, totalAmount: 0, details: []
        };
      }
      var w = wallets[walletKey];
      if (ordererName) w.name = ordererName;
      if (orderTime && orderTime < w.firstDate) w.firstDate = orderTime;
      if (orderTime && orderTime > w.lastDate) w.lastDate = orderTime;
      w.totalQty += qty;
      if (!w.orders[parentOrder]) { w.orders[parentOrder] = true; w.totalAmount += amount; }
      w.details.push([walletAddr, orderNum, orderTime, ordererName, qty, amount, txid]);
    }
  }
  
  // 요약 시트 작성
  var walletKeys = Object.keys(wallets);
  var detailRows = [];
  
  for (var k = 0; k < walletKeys.length; k++) {
    var w = wallets[walletKeys[k]];
    var orderCount = Object.keys(w.orders).length;
    var staking = existingStaking[walletKeys[k]] || 0;
    var rowNum = k + 2; // 1-based, 헤더 다음
    
    // A~G열: 데이터 (G열=스테이킹금액은 수기)
    summarySheet.getRange(rowNum, 1, 1, 7).setValues([[
      w.address, w.name, w.firstDate, orderCount, w.totalQty, w.totalAmount, staking
    ]]);
    
    // H열: 주문가능수량 = 스테이킹금액/100 (수식)
    summarySheet.getRange(rowNum, 8).setFormula('=IF(G' + rowNum + '=0,0,FLOOR(G' + rowNum + '/100))');
    // I열: 잔여수량 = 주문가능수량 - 총주문수량 (수식)
    summarySheet.getRange(rowNum, 9).setFormula('=H' + rowNum + '-E' + rowNum);
    // J열: 상태 (수식)
    summarySheet.getRange(rowNum, 10).setFormula('=IF(G' + rowNum + '=0,"미설정",IF(E' + rowNum + '>H' + rowNum + ',"초과","정상"))');
    // K열: 최근주문일
    summarySheet.getRange(rowNum, 11).setValue(w.lastDate);
    
    // 지갑주소 링크
    if (w.address.indexOf('0x') === 0) {
      var rt = SpreadsheetApp.newRichTextValue()
        .setText(w.address)
        .setLinkUrl('https://bscscan.com/address/' + w.address)
        .build();
      summarySheet.getRange(rowNum, 1).setRichTextValue(rt);
    }
    
    for (var d = 0; d < w.details.length; d++) {
      detailRows.push(w.details[d]);
    }
  }
  
  // G열(스테이킹금액) 서식을 숫자로 강제 (날짜 자동변환 방지)
  if (walletKeys.length > 0) {
    summarySheet.getRange(2, 7, walletKeys.length, 1).setNumberFormat('#,##0');
  }
  
  // 상세 내역
  if (detailRows.length > 0) {
    detailSheet.getRange(2, 1, detailRows.length, 7).setValues(detailRows);
  }
  
  // 조건부 서식: 초과 시 빨간색, 미설정 시 노란색
  applyWalletConditionalFormatting(summarySheet);
  
  setWalletColumnWidths(summarySheet);
  detailSheet.setColumnWidth(1, 380); detailSheet.setColumnWidth(2, 160);
  detailSheet.setColumnWidth(3, 150); detailSheet.setColumnWidth(4, 100);
  detailSheet.setColumnWidth(5, 80);  detailSheet.setColumnWidth(6, 100);
  detailSheet.setColumnWidth(7, 380);
  
  Logger.log('=== 지갑추적 완료: ' + walletKeys.length + '개 지갑, ' + detailRows.length + '건 주문 ===');
}

// 조건부 서식 설정 (스테이킹금액 수정 시 자동으로 색상 반영)
function applyWalletConditionalFormatting(sheet) {
  var rules = [];
  
  // 초과: J열이 "초과"이면 행 전체 빨간색
  var overRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="초과"')
    .setBackground('#ffcdd2')
    .setFontColor('#b71c1c')
    .setRanges([sheet.getRange('A2:K1000')])
    .build();
  rules.push(overRule);
  
  // 미설정: J열이 "미설정"이면 행 전체 노란색
  var unsetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2="미설정"')
    .setBackground('#fff8e1')
    .setFontColor('#e65100')
    .setRanges([sheet.getRange('A2:K1000')])
    .build();
  rules.push(unsetRule);
  
  sheet.setConditionalFormatRules(rules);
}

// 열 너비 설정 헬퍼
function setWalletColumnWidths(sheet) {
  sheet.setColumnWidth(1, 380);  // 지갑주소
  sheet.setColumnWidth(2, 100);  // 주문자이름
  sheet.setColumnWidth(3, 130);  // 최초등록일
  sheet.setColumnWidth(4, 90);   // 총주문횟수
  sheet.setColumnWidth(5, 90);   // 총주문수량
  sheet.setColumnWidth(6, 100);  // 총송금금액
  sheet.setColumnWidth(7, 120);  // 스테이킹금액
  sheet.setColumnWidth(8, 100);  // 주문가능수량
  sheet.setColumnWidth(9, 90);   // 잔여수량
  sheet.setColumnWidth(10, 80);  // 상태
  sheet.setColumnWidth(11, 130); // 최근주문일
}

// 스테이킹 금액 변경 시 전체 상태 재계산
// 관리자가 스테이킹금액(F열)을 수정하면 자동 재계산
// 실행: 함수 선택 > recalcWalletStatus > ▶ 실행
// 지갑추적 시트 수식/서식 재적용
// G열(스테이킹금액) 입력 후 수식이 깨졌을 때 복구용
function recalcWalletStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('지갑추적');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var lastRow = sheet.getLastRow();
  var updated = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    // G열 숫자 서식 강제
    sheet.getRange(r, 7).setNumberFormat('#,##0');
    // H,I,J열 수식 재설정
    sheet.getRange(r, 8).setFormula('=IF(G' + r + '=0,0,FLOOR(G' + r + '/100))');
    sheet.getRange(r, 9).setFormula('=H' + r + '-E' + r);
    sheet.getRange(r, 10).setFormula('=IF(G' + r + '=0,"미설정",IF(E' + r + '>H' + r + ',"초과","정상"))');
    updated++;
  }
  
  // 조건부 서식 재적용
  applyWalletConditionalFormatting(sheet);
  
  Logger.log('=== ' + updated + '건 수식/서식 재적용 완료 ===');
}

// ──────────────────────────────────────
// 메뉴
// ──────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 주문관리')
    .addItem('택배출력 시트 생성/갱신', 'generateCourierSheet')
    .addItem('밀린 주문 일괄 이동', 'moveStuckOrders')
    .addItem('빈 행 정리', 'cleanEmptyRows')
    .addSeparator()
    .addItem('지갑추적 시트 생성/갱신', 'buildWalletTracking')
    .addItem('지갑 상태 재계산', 'recalcWalletStatus')
    .addSeparator()
    .addItem('밀린 행 복구 (fixMisalignedRows)', 'fixMisalignedRows')
    .addItem('운송장번호 복구 (fixTrackingNumbers)', 'fixTrackingNumbers')
    .addItem('시트 진단 (diagnosRows)', 'diagnosRows')
    .addSeparator()
    .addItem('트리거 재설정', 'createTrigger')
    .addItem('시트 초기설정', 'setupSheet')
    .addToUi();
}
