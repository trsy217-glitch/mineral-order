/**
 * 오륜 미네랄 주문 시스템 - Google Apps Script
 * 
 * ========== 설정 방법 ==========
 * 
 * 1. 새 Google 스프레드시트 만들기
 * 
 * 2. 시트 이름을 "주문목록"으로 변경 (하단 탭 우클릭 > 이름 바꾸기)
 * 
 * 3. 첫 번째 행(A1)부터 아래 헤더 복사해서 붙여넣기:
 *    주문번호 | 접수시간 | 주문상태 | 송금일시 | 송금금액 | 입금내역 | 주문자이름 | 주문자전화 | 수령인이름 | 수령인전화 | 주소 | 상세주소 | 배송요청 | 운송장번호
 * 
 * 4. 메뉴 > 확장 프로그램 > Apps Script 클릭
 * 
 * 5. 기존 코드 전부 지우고 이 파일 내용 복사 붙여넣기
 * 
 * 6. 저장 (Ctrl+S)
 * 
 * 7. 배포하기:
 *    - 오른쪽 상단 "배포" 버튼 클릭
 *    - "새 배포" 선택
 *    - 톱니바퀴 아이콘 > "웹 앱" 선택
 *    - 설명: 오륜미네랄 주문
 *    - 다음 사용자 인증 정보로 실행: "나"
 *    - 액세스 권한이 있는 사용자: "모든 사용자"  ⚠️ 중요!
 *    - "배포" 클릭
 *    - 권한 승인 (Google 계정 선택 > 고급 > 안전하지 않은 페이지로 이동 > 허용)
 * 
 * 8. 웹 앱 URL 복사 (https://script.google.com/macros/s/xxxxx/exec 형태)
 * 
 * 9. index.html과 order-status.html의 SHEET_URL에 붙여넣기
 * 
 * ========== 끝 ==========
 */

// 시트 이름
const SHEET_NAME = '주문목록';

/**
 * GET 요청 처리 - 주문 조회
 */
function doGet(e) {
  let result;
  
  try {
    const action = e.parameter.action;
    
    if (action === 'search') {
      result = searchOrders(e.parameter.name, e.parameter.phone);
    } else {
      result = { success: true, message: 'API 정상 작동 중' };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  
  // JSONP 콜백 지원
  const callback = e.parameter.callback;
  const output = callback 
    ? callback + '(' + JSON.stringify(result) + ')'
    : JSON.stringify(result);
  
  return ContentService
    .createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

/**
 * POST 요청 처리 - 주문 등록
 */
function doPost(e) {
  let result;
  
  try {
    // 데이터 파싱
    let data = {};
    
    if (e.postData && e.postData.contents) {
      const contents = e.postData.contents;
      
      // JSON 파싱 시도
      try {
        data = JSON.parse(contents);
      } catch (parseError) {
        // URL 인코딩된 데이터 파싱 시도
        const pairs = contents.split('&');
        pairs.forEach(function(pair) {
          const parts = pair.split('=');
          if (parts.length === 2) {
            data[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1].replace(/\+/g, ' '));
          }
        });
      }
    }
    
    // 파라미터에서도 데이터 가져오기 (GET 방식 폴백)
    if (e.parameter) {
      Object.keys(e.parameter).forEach(function(key) {
        if (!data[key]) {
          data[key] = e.parameter[key];
        }
      });
    }
    
    // 시트에 저장
    const sheet = getOrCreateSheet();
    const orderNumber = generateOrderNumber();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    
    const fullAddress = [
      data['우편번호'] || '',
      data['주소'] || '',
    ].filter(Boolean).join(' ');
    
    sheet.appendRow([
      orderNumber,
      timestamp,
      '주문접수',
      data['송금일시'] || '',
      data['송금금액'] || '',
      data['입금내역'] || '',
      data['주문자이름'] || '',
      data['주문자전화'] || '',
      data['수령인이름'] || '',
      data['수령인전화'] || '',
      fullAddress,
      data['상세주소'] || '',
      data['배송요청'] || '',
      ''
    ]);
    
    result = { success: true, orderNumber: orderNumber };
    
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 주문 검색
 */
function searchOrders(name, phone) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, orders: [] };
  }
  
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function(header, idx) {
    colIndex[header] = idx;
  });
  
  const orders = [];
  const searchName = (name || '').trim();
  const searchPhone = normalizePhone(phone || '');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowName = String(row[colIndex['주문자이름']] || '').trim();
    const rowPhone = normalizePhone(String(row[colIndex['주문자전화']] || ''));
    
    if (rowName === searchName && rowPhone === searchPhone) {
      orders.push({
        orderNumber: row[colIndex['주문번호']] || '',
        timestamp: row[colIndex['접수시간']] || '',
        status: row[colIndex['주문상태']] || '주문접수',
        amount: row[colIndex['송금금액']] || '',
        address: row[colIndex['주소']] || '',
        trackingNumber: row[colIndex['운송장번호']] || ''
      });
    }
  }
  
  orders.reverse();
  return { success: true, orders: orders };
}

/**
 * 전화번호 정규화
 */
function normalizePhone(phone) {
  return String(phone).replace(/[^0-9]/g, '');
}

/**
 * 주문번호 생성
 */
function generateOrderNumber() {
  const now = new Date();
  const y = now.getFullYear().toString().slice(-2);
  const m = ('0' + (now.getMonth() + 1)).slice(-2);
  const d = ('0' + now.getDate()).slice(-2);
  const h = ('0' + now.getHours()).slice(-2);
  const min = ('0' + now.getMinutes()).slice(-2);
  const rand = ('0' + Math.floor(Math.random() * 100)).slice(-2);
  return 'ORN' + y + m + d + h + min + rand;
}

/**
 * 시트 가져오기 (없으면 생성)
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      '주문번호', '접수시간', '주문상태', '송금일시', '송금금액',
      '입금내역', '주문자이름', '주문자전화', '수령인이름', '수령인전화',
      '주소', '상세주소', '배송요청', '운송장번호'
    ]);
    sheet.getRange(1, 1, 1, 14).setBackground('#43a047').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * 테스트 함수 - Apps Script 편집기에서 실행해서 테스트
 */
function testPost() {
  const testEvent = {
    postData: {
      contents: JSON.stringify({
        '송금일시': '2024/01/15',
        '송금금액': '100',
        '입금내역': '테스트 TxID',
        '주문자이름': '테스트',
        '주문자전화': '010-1234-5678',
        '수령인이름': '테스트',
        '수령인전화': '010-1234-5678',
        '주소': '서울시 테스트구',
        '상세주소': '101호',
        '배송요청': '테스트'
      })
    },
    parameter: {}
  };
  
  const result = doPost(testEvent);
  Logger.log(result.getContent());
}

function testGet() {
  const testEvent = {
    parameter: {
      action: 'search',
      name: '테스트',
      phone: '010-1234-5678'
    }
  };
  
  const result = doGet(testEvent);
  Logger.log(result.getContent());
}
