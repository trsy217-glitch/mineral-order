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
  var sheet = ss.getSheets()[0];
  
  var orderNum = 'ORN' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss');
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  var addr = ((p.postcode || '') + ' ' + (p.addr || '')).trim();
  
  sheet.appendRow([
    orderNum,
    timestamp,
    '주문접수',
    p.payDate || '',
    p.payAmount || '',
    p.payProof || '',
    p.dmaxWallet || '',
    p.ordererName || '',
    p.ordererPhone || '',
    p.receiverName || '',
    p.receiverPhone || '',
    addr,
    p.addrDetail || '',
    p.deliveryNote || '',
    ''
  ]);
  
  return { success: true, orderNumber: orderNum };
}

function searchOrders(name, phone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  var data = sheet.getDataRange().getValues();
  var orders = [];
  var searchName = (name || '').trim();
  var searchPhone = (phone || '').replace(/[^0-9]/g, '');
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowName = String(row[7] || '').trim();
    var rowPhone = String(row[8] || '').replace(/[^0-9]/g, '');
    
    if (rowName === searchName && rowPhone === searchPhone) {
      orders.push({
        orderNumber: row[0] || '',
        timestamp: row[1] || '',
        status: row[2] || '주문접수',
        amount: row[4] || '',
        address: row[11] || '',
        trackingNumber: row[14] || ''
      });
    }
  }
  
  orders.reverse();
  return { success: true, orders: orders };
}
