/**
 * 오륜 미네랄 주문 시스템 v2
 * 
 * ========== 시트 구조 ==========
 * 
 * [신규주문] 시트 헤더 (A1에 붙여넣기):
 * 주문번호	접수시간	주문상태	송금일시	송금금액	입금내역	DMAX지갑	주문자이름	주문자전화	수령인이름	수령인전화	주소	배송요청	택배사	운송장번호
 * 
 * [완료] 시트: 자동 생성됨 (동일한 헤더)
 * 
 * ========== 설정 순서 ==========
 * 
 * 1. 이 코드 붙여넣기 후 저장
 * 2. 상단 메뉴: 실행 > 함수 실행 > setupSheet 선택 (택배사 드롭다운 설정)
 * 3. 상단 메뉴: 실행 > 함수 실행 > createTrigger 선택 (자동 이동 설정)
 * 4. 배포 > 배포 관리 > 새 버전 배포
 * 
 */

function doGet(e) {
  var result;
  try {
    var action = e.parameter.action || '';
    if (action === 'order') {
      result = saveOrder(e.parameter);
    } else if (action === 'search') {
      result = searchOrders(e.parameter.name, e.parameter.phone);
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
    sheet.appendRow(['주문번호','접수시간','주문상태','송금일시','송금금액','입금내역','DMAX지갑','주문자이름','주문자전화','수령인이름','수령인전화','주소','배송요청','택배사','운송장번호']);
    sheet.setFrozenRows(1);
    setupCourierDropdown(sheet);
  }
  
  var orderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  
  // 주소 합치기 (우편번호 + 기본주소 + 상세주소)
  var fullAddress = '';
  if (p.postcode) fullAddress += p.postcode + ' ';
  if (p.addr) fullAddress += p.addr;
  if (p.addrDetail) fullAddress += ' ' + p.addrDetail;
  fullAddress = fullAddress.trim();
  
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
    fullAddress,
    p.deliveryNote || '',
    '',  // 택배사
    ''   // 운송장번호
  ]);
  
  // 마지막 행에 링크 설정
  var lastRow = sheet.getLastRow();
  
  // 입금내역 (F열, 6번째) - TxID를 BscScan 트랜잭션 링크로
  if (payProof) {
    var txId = payProof;
    // 0x로 시작하면 트랜잭션 해시로 판단
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
    // 0x로 시작하면 지갑 주소로 판단
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
          address: row[11] || '',
          courier: row[13] || '',
          trackingNumber: row[14] || ''
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
          address: row2[11] || '',
          courier: row2[13] || '',
          trackingNumber: row2[14] || ''
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
 * 택배사 드롭다운 설정
 */
function setupCourierDropdown(sheet) {
  var courierList = ['CJ대한통운', '롯데택배', '한진택배', '우체국', '로젠택배', '직접전달'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(courierList, true)
    .setAllowInvalid(false)
    .build();
  
  // N열 전체에 드롭다운 적용 (2행부터 1000행까지)
  sheet.getRange('N2:N1000').setDataValidation(rule);
}

/**
 * 시트 설정 함수 - 처음 한 번 실행
 * 실행 > 함수 실행 > setupSheet
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow(['주문번호','접수시간','주문상태','송금일시','송금금액','입금내역','DMAX지갑','주문자이름','주문자전화','수령인이름','수령인전화','주소','배송요청','택배사','운송장번호']);
    sheet.setFrozenRows(1);
  }
  
  setupCourierDropdown(sheet);
  Logger.log('택배사 드롭다운이 설정되었습니다.');
}

/**
 * 운송장번호 입력 시 자동으로 완료 시트로 이동
 */
function onEdit(e) {
  // Lock으로 중복 실행 방지
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return; // 이미 다른 실행 중이면 종료
  }
  
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // 신규주문 시트의 O열(15번째, 운송장번호)이 수정된 경우만
    if (sheet.getName() !== '신규주문' || range.getColumn() !== 15) {
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
      targetSheet.appendRow(['주문번호','접수시간','주문상태','송금일시','송금금액','입금내역','DMAX지갑','주문자이름','주문자전화','수령인이름','수령인전화','주소','배송요청','택배사','운송장번호']);
      targetSheet.setFrozenRows(1);
    }
    
    // 해당 행 데이터 가져오기
    var rowData = sourceSheet.getRange(row, 1, 1, 15).getValues()[0];
    var orderNumber = rowData[0];
    
    // 이미 완료 시트에 같은 주문번호가 있으면 중복 추가 안함
    var completedData = targetSheet.getDataRange().getValues();
    for (var i = 1; i < completedData.length; i++) {
      if (completedData[i][0] === orderNumber) {
        // 이미 있으면 신규주문에서만 삭제
        sourceSheet.deleteRow(row);
        return;
      }
    }
    
    // 주문상태를 '배송중'으로 변경
    rowData[2] = '배송중';
    
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
 * 실행 > 함수 실행 > createTrigger
 */
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  
  // 기존 모든 onEdit 트리거 삭제
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 새 트리거 생성 (하나만)
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  Logger.log('트리거가 생성되었습니다. (기존 트리거 삭제 후 새로 생성)');
}
