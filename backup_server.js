/**
 * 🛡️ 백업 주문 서버 (Backup Order Server)
 * 
 * 메인 서버와 동시에 주문을 받아 이중 백업을 수행합니다.
 * 메인과 동일한 포맷으로 데이터를 저장하므로, 메인이 다운되어도
 * 이 시트에서 주문 조회가 가능합니다.
 * 
 * 주요 기능:
 * - 주문 받기 (saveOrder)
 * - TxID 중복 검사 (checkDuplicateTx)
 * - 주문 조회 (searchOrders)
 * - 기존 데이터 일괄 import (importLegacy)
 */

// ──────────────────────────────────────
// 진입점: 모든 요청 처리
// ──────────────────────────────────────
function doGet(e) {
  var params = e.parameter || {};
  var action = params.action || '';
  var callback = params.callback || '';
  var result;
  
  try {
    if (action === 'order') {
      result = saveOrder(params);
    } else if (action === 'txcheck') {
      result = checkDuplicateTx(params.txid);
    } else if (action === 'search') {
      result = searchOrders(params.name, params.phone);
    } else if (action === 'ping') {
      // 서버 생존 확인용
      result = { success: true, server: 'backup', timestamp: new Date().toISOString() };
    } else if (action === 'notify') {
      // 한쪽 서버 실패 시 관리자 알림 (백업이 살아있을 때 호출됨)
      result = handleNotification_(params);
    } else {
      result = { success: false, error: 'UNKNOWN_ACTION', message: '알 수 없는 요청' };
    }
  } catch(err) {
    result = { success: false, error: 'EXCEPTION', message: err.toString() };
    Logger.log('❌ 오류: ' + err.toString());
  }
  
  // JSONP 응답 (브라우저에서 직접 호출 시)
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST 요청 (메인 서버에서 importLegacy 호출 시 사용)
function doPost(e) {
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'importLegacy') {
      result = importLegacyOrder(payload);
    } else {
      result = { success: false, error: 'UNKNOWN_ACTION' };
    }
  } catch(err) {
    result = { success: false, error: 'EXCEPTION', message: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────
// 헤더 정의 (메인과 동일 - 17컬럼 버전)
// ──────────────────────────────────────
function getHeaders() {
  return [
    '주문번호', '주문일시', '주문상태', '송금일자',
    '결제금액(USDT)', 'TxID', 'DMAX지갑',
    '주문자', '주문자전화',
    '받는분', '받는분전화',
    '우편번호', '주소', '배송메모',
    '택배사', '운송장번호', '발송날짜'
  ];
}

// 컬럼 인덱스 (0부터 시작)
var COL = {
  orderNum: 0, timestamp: 1, status: 2, payDate: 3,
  payAmount: 4, payProof: 5, dmaxWallet: 6,
  ordererName: 7, ordererPhone: 8,
  receiverName: 9, receiverPhone: 10,
  postcode: 11, address: 12, deliveryNote: 13,
  courier: 14, tracking: 15, shippingDate: 16
};

// ──────────────────────────────────────
// 시트 가져오기 (없으면 생성)
// ──────────────────────────────────────
function getOrderSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  
  if (!sheet) {
    sheet = ss.insertSheet('신규주문');
    sheet.appendRow(getHeaders());
    sheet.setFrozenRows(1);
    
    // 헤더 스타일
    var headerRange = sheet.getRange(1, 1, 1, 17);
    headerRange.setBackground('#1565c0')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
  }
  
  return sheet;
}

// ──────────────────────────────────────
// 주문 저장 (메인과 동일 로직)
// ──────────────────────────────────────
function saveOrder(p) {
  var sheet = getOrderSheet_();
  
  var baseOrderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);  // 랜덤 3자리로 같은 초 주문번호 겹침 방지
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  
  var postcode = (p.postcode || '').trim();
  var address = '';
  if (p.addr) address += p.addr;
  if (p.addrDetail) address += ' ' + p.addrDetail;
  address = address.trim();
  
  var payProof = (p.payProof || '').trim();
  var dmaxWallet = (p.dmaxWallet || '').trim();

  // ★★ 서버 측 입력 검증 (주문서를 거치지 않은 직접 호출·탐침·스팸 차단) ★★
  //   주문서 HTML의 필수항목을 서버에서도 동일하게 검사.
  //   하나라도 어긋나면 시트에 아무것도 쓰지 않고 거절한다.
  var vName = (p.ordererName || '').trim();
  var vPhone = (p.ordererPhone || '').trim();
  var vRecv = (p.receiverName || '').trim();
  var phoneOK = /^010-?\d{4}-?\d{4}$/.test(vPhone);      // 실제 휴대폰 형식
  var walletOK = /^0x[0-9a-fA-F]{40}$/.test(dmaxWallet); // 지갑주소 42자(0x+40)
  var txOK = (payProof.indexOf('0x') === 0 && payProof.length >= 60); // TxID 60자 이상
  var addrOK = address.length >= 5;                      // 주소 존재
  if (!vName || !vRecv || !phoneOK || !addrOK || !walletOK || !txOK) {
    return { success: false, error: 'INVALID_INPUT', message: '필수 입력값이 올바르지 않습니다.' };
  }

  // TxID 중복 검사
  if (payProof) {
    var dupResult = checkDuplicateTx(payProof);
    if (dupResult.duplicate) {
      return { success: false, error: 'DUPLICATE_TX', message: '이미 사용된 TxID입니다.' };
    }
  }
  
  var totalQty = parseInt(p.orderQty) || 1;
  var totalAmount = parseFloat(p.payAmount) || 0;
  
  // 10병 단위로 분할
  var chunks = [];
  var remaining = totalQty;
  while (remaining > 0) {
    var chunkQty = Math.min(remaining, 10);
    chunks.push(chunkQty);
    remaining -= chunkQty;
  }
  
  var numChunks = chunks.length;
  
  for (var c = 0; c < numChunks; c++) {
    var orderNum = (numChunks === 1) ? baseOrderNum : baseOrderNum + '-' + (c + 1);
    var chunkQty = chunks[c];
    
    var rowData = [
      orderNum, timestamp, '주문접수', p.payDate || '',
      totalAmount,
      payProof, dmaxWallet,
      p.ordererName || '', p.ordererPhone || '',
      p.receiverName || '', p.receiverPhone || '',
      postcode, address, p.deliveryNote || '',
      '', '', ''
    ];
    
    // ⚠️ 수량 컬럼이 있는 18컬럼 버전인 경우 변환 (필요시)
    sheet.appendRow(rowData);
    
    // 메모에 수량 기록 (백업이라 단순화)
    var lastRow = sheet.getLastRow();
    if (chunkQty !== totalQty) {
      sheet.getRange(lastRow, 1).setNote('분할 ' + (c+1) + '/' + numChunks + ' (' + chunkQty + '병)');
    }
  }
  
  return {
    success: true,
    server: 'backup',
    orderNumber: baseOrderNum,
    totalQty: totalQty,
    chunks: numChunks
  };
}

// ──────────────────────────────────────
// TxID 중복 검사
// ──────────────────────────────────────
function stripTxPrefix(txid) {
  if (!txid) return '';
  var t = String(txid).trim();
  // BSCScan URL 형태면 마지막 path 추출
  var match = t.match(/\/tx\/([0-9a-fA-Fx]+)/);
  if (match) return match[1];
  return t;
}

function checkDuplicateTx(txid) {
  if (!txid) return { success: true, duplicate: false };
  
  var searchTx = stripTxPrefix(txid.trim()).toLowerCase();
  if (!searchTx) return { success: true, duplicate: false };
  
  var sheet = getOrderSheet_();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var existing = data[i][COL.payProof];
    if (existing) {
      var existingTx = stripTxPrefix(String(existing).trim()).toLowerCase();
      if (existingTx && existingTx === searchTx) {
        return { 
          success: true, 
          duplicate: true, 
          orderNumber: data[i][COL.orderNum] 
        };
      }
    }
  }
  
  return { success: true, duplicate: false };
}

// ──────────────────────────────────────
// 주문 조회 (메인 다운 시 백업으로 조회 가능)
// ──────────────────────────────────────
function searchOrders(name, phone) {
  if (!name || !phone) {
    return { success: false, message: '이름과 연락처가 필요합니다.' };
  }
  
  var searchName = String(name).trim();
  var searchPhone = String(phone).replace(/[^0-9]/g, '');
  if (searchPhone.length < 4) {
    return { success: false, message: '연락처는 최소 4자리 이상' };
  }
  
  var orders = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('신규주문');
  if (!sheet) return { success: true, orders: [] };
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.orderNum]) continue;
    
    var rowName = String(row[COL.ordererName] || '').trim();
    var rowPhone = String(row[COL.ordererPhone] || '').replace(/[^0-9]/g, '');
    
    if (rowName === searchName && rowPhone.indexOf(searchPhone) >= 0) {
      var orderNum = String(row[COL.orderNum] || '');
      var parentOrder = orderNum;
      var subIndex = 0;
      var match = orderNum.match(/^(ORN\d+)-(\d+)$/);
      if (match) {
        parentOrder = match[1];
        subIndex = parseInt(match[2]);
      }
      
      var tn = String(row[COL.tracking] || '').trim();
      var courier = String(row[COL.courier] || '').trim();
      if (!courier && tn === '공장배송') courier = '공장배송';
      
      var shipDateRaw = row[COL.shippingDate];
      var shipDateStr = '';
      if (shipDateRaw) {
        if (shipDateRaw instanceof Date) {
          shipDateStr = Utilities.formatDate(shipDateRaw, 'Asia/Seoul', 'yyyy-MM-dd');
        } else {
          shipDateStr = String(shipDateRaw).trim();
        }
      }
      
      orders.push({
        orderNumber: orderNum,
        parentOrder: parentOrder,
        subIndex: subIndex,
        timestamp: row[COL.timestamp] || '',
        status: row[COL.status] || '주문접수',
        amount: row[COL.payAmount] || '',
        address: ((row[COL.postcode] || '') + ' ' + (row[COL.address] || '')).trim(),
        courier: courier,
        trackingNumber: tn,
        shippingDate: shipDateStr
      });
    }
  }
  
  return { success: true, server: 'backup', orders: orders };
}

// ──────────────────────────────────────
// 기존 데이터 import (메인에서 일괄 이관 시)
// ──────────────────────────────────────
function importLegacyOrder(payload) {
  var sheet = getOrderSheet_();
  
  // 이미 같은 주문번호 있으면 스킵 (중복 방지)
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][COL.orderNum] === payload.orderNumber) {
      return { success: true, skipped: true, message: '이미 존재' };
    }
  }
  
  var rowData = [
    payload.orderNumber || '',
    payload.timestamp || '',
    payload.status || '주문접수',
    payload.payDate || '',
    payload.payAmount || '',
    payload.payProof || '',
    payload.dmaxWallet || '',
    payload.ordererName || '',
    payload.ordererPhone || '',
    payload.receiverName || '',
    payload.receiverPhone || '',
    payload.postcode || '',
    payload.address || '',
    payload.deliveryNote || '',
    payload.courier || '',
    payload.tracking || '',
    payload.shippingDate || ''
  ];
  
  sheet.appendRow(rowData);
  
  // 원본 시트 정보를 메모에 기록
  if (payload.sourceSheet) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNote('원본: ' + payload.sourceSheet);
  }
  
  return { success: true, imported: true };
}

// ──────────────────────────────────────
// 📧 관리자 알림 처리 (메인 서버 다운 시)
// ──────────────────────────────────────
function handleNotification_(params) {
  var type = params.type || '';
  var msg = params.message || '';
  
  if (type === 'server_failure') {
    sendAdminAlert_('🚨 메인 서버 문제 발생', 
      '메인 서버에서 주문 처리가 실패했습니다.\n\n' +
      '상세 내용:\n' + msg + '\n\n' +
      '✅ 주문은 백업 서버에 안전하게 저장되었습니다.\n' +
      '⚠️ 빠르게 메인 서버 상태를 확인해주세요.\n\n' +
      '시간: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
    );
  }
  
  return { success: true };
}

function sendAdminAlert_(subject, body) {
  try {
    var props = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty('ADMIN_EMAIL');
    
    if (!adminEmail) {
      adminEmail = Session.getEffectiveUser().getEmail();
    }
    
    // 중복 알림 방지
    var cacheKey = 'alert_' + subject;
    var cache = CacheService.getScriptCache();
    if (cache.get(cacheKey)) return;
    cache.put(cacheKey, '1', 300);
    
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body
    });
  } catch(e) {
    Logger.log('알림 발송 실패: ' + e.toString());
  }
}

// 한 번만 실행: 관리자 이메일 설정 (백업 서버용)
function setAdminEmail() {
  var email = 'aamarketing250611@gmail.com';
  if (email === 'aamarketing250611@gmail.com') {
    Logger.log('aamarketing250611@gmail.com');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAIL', email);
  Logger.log('✅ 관리자 이메일 설정 완료: ' + email);
}

// ──────────────────────────────────────
// 🧪 테스트: 백업 서버 동작 확인
// ──────────────────────────────────────
function testBackupServer() {
  var testData = {
    orderQty: '3',
    payDate: '2026-05-15',
    payAmount: '36',
    payProof: '0xTEST_BACKUP_' + Date.now(),
    dmaxWallet: '',
    ordererName: '테스트백업',
    ordererPhone: '010-1234-5678',
    receiverName: '테스트수령',
    receiverPhone: '010-1234-5678',
    postcode: '12345',
    addr: '서울시 강남구 테스트로',
    addrDetail: '101호',
    deliveryNote: '백업 서버 테스트'
  };
  
  Logger.log('백업 서버 테스트 시작...');
  var result = saveOrder(testData);
  Logger.log('결과: ' + JSON.stringify(result));
}
