// ============================================================
// Code.gs — Server-side: CRUD + Email Notification
// Google Sheets as database:
//   Departments : Dept_ID | Dept_Name | Emails
//   Customers   : Customer_ID | Customer_Name | Phone | Member_Level
//   Products    : Product_ID  | Product_Name  | Price
//   Orders      : Order_ID | Customer_ID | Product_ID | Quantity | Total_Amount | Status | Created_At
// ============================================================

// ── Entry point ───────────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Order Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Sheet accessors ───────────────────────────────────────────
function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Check sheet name.');
  return sh;
}

// ── Generic: sheet → array of objects ────────────────────────
function sheetToObjects_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Normalise Date objects
      if (val instanceof Date) {
        obj[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      } else {
        obj[headers[j]] = (val === null || val === undefined) ? '' : val;
      }
    }
    rows.push(obj);
  }
  return rows;
}

// ── ID generator ──────────────────────────────────────────────
function generateId_(prefix, sheet) {
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0]);
    var num = parseInt(id.replace(prefix, ''));
    if (!isNaN(num) && num > max) max = num;
  }
  var next = max + 1;
  return prefix + String(next).padStart(3, '0');
}

// ── Find row index by ID (column 0) ──────────────────────────
function findRowIndex_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

// ============================================================
//  CUSTOMERS
// ============================================================
function getCustomers() {
  return sheetToObjects_(getSheet_('Customers'));
}

function addCustomer(name, phone, level) {
  var sheet = getSheet_('Customers');
  var id = generateId_('C', sheet);
  sheet.appendRow([id, name, phone, level]);
  return { success: true, id: id };
}

function updateCustomer(id, name, phone, level) {
  var sheet = getSheet_('Customers');
  var row = findRowIndex_(sheet, id);
  if (row === -1) throw new Error('Customer ' + id + ' not found.');
  sheet.getRange(row, 1, 1, 4).setValues([[id, name, phone, level]]);
  return { success: true };
}

function deleteCustomer(id) {
  var sheet = getSheet_('Customers');
  var row = findRowIndex_(sheet, id);
  if (row === -1) throw new Error('Customer ' + id + ' not found.');
  // Check if referenced in Orders
  var orders = sheetToObjects_(getSheet_('Orders'));
  var inUse = orders.some(function(o) { return o.Customer_ID === id; });
  if (inUse) throw new Error('Không thể xóa: khách hàng đang có đơn hàng.');
  sheet.deleteRow(row);
  return { success: true };
}

// ============================================================
//  PRODUCTS
// ============================================================
function getProducts() {
  return sheetToObjects_(getSheet_('Products'));
}

function addProduct(name, price) {
  var sheet = getSheet_('Products');
  var id = generateId_('P', sheet);
  sheet.appendRow([id, name, parseFloat(price)]);
  return { success: true, id: id };
}

function updateProduct(id, name, price) {
  var sheet = getSheet_('Products');
  var row = findRowIndex_(sheet, id);
  if (row === -1) throw new Error('Product ' + id + ' not found.');
  sheet.getRange(row, 1, 1, 3).setValues([[id, name, parseFloat(price)]]);
  return { success: true };
}

function deleteProduct(id) {
  var sheet = getSheet_('Products');
  var row = findRowIndex_(sheet, id);
  if (row === -1) throw new Error('Product ' + id + ' not found.');
  var orders = sheetToObjects_(getSheet_('Orders'));
  var inUse = orders.some(function(o) { return o.Product_ID === id; });
  if (inUse) throw new Error('Không thể xóa: sản phẩm đang có trong đơn hàng.');
  sheet.deleteRow(row);
  return { success: true };
}

// ============================================================
//  ORDERS
// ============================================================
function getOrders() {
  return sheetToObjects_(getSheet_('Orders'));
}

function createOrder(customerId, productId, quantity) {
  // Validate inputs
  var qty = parseInt(quantity);
  if (isNaN(qty) || qty < 1) throw new Error('Số lượng không hợp lệ.');

  // Lookup product price
  var products = sheetToObjects_(getSheet_('Products'));
  var product = null;
  for (var i = 0; i < products.length; i++) {
    if (products[i].Product_ID === productId) { product = products[i]; break; }
  }
  if (!product) throw new Error('Sản phẩm ' + productId + ' không tồn tại.');

  // Lookup customer
  var customers = sheetToObjects_(getSheet_('Customers'));
  var customer = null;
  for (var i = 0; i < customers.length; i++) {
    if (customers[i].Customer_ID === customerId) { customer = customers[i]; break; }
  }
  if (!customer) throw new Error('Khách hàng ' + customerId + ' không tồn tại.');

  var totalAmount = parseFloat(product.Price) * qty;
  var orderSheet = getSheet_('Orders');
  var orderId = generateId_('O', orderSheet);
  var createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  orderSheet.appendRow([orderId, customerId, productId, qty, totalAmount, 'Pending', createdAt]);

  // Send email notification
  var emailResult = sendOrderNotification_({
    orderId     : orderId,
    customerName: customer.Customer_Name,
    productName : product.Product_Name,
    quantity    : qty,
    totalAmount : totalAmount,
    createdAt   : createdAt,
    status      : 'Pending'
  });

  return { success: true, orderId: orderId, emailSent: emailResult.sent, emailError: emailResult.error };
}

function updateOrderStatus(orderId, status) {
  var sheet = getSheet_('Orders');
  var row = findRowIndex_(sheet, orderId);
  if (row === -1) throw new Error('Order ' + orderId + ' not found.');
  // Status is column 6 (index 5, 1-based = 6)
  sheet.getRange(row, 6).setValue(status);
  return { success: true };
}

function deleteOrder(orderId) {
  var sheet = getSheet_('Orders');
  var row = findRowIndex_(sheet, orderId);
  if (row === -1) throw new Error('Order ' + orderId + ' not found.');
  sheet.deleteRow(row);
  return { success: true };
}

// ============================================================
//  EMAIL NOTIFICATION
// ============================================================
function sendOrderNotification_(orderInfo) {
  try {
    var depts = sheetToObjects_(getSheet_('Departments'));
    var emails = [];
    depts.forEach(function(d) {
      var raw = String(d.Emails || '').trim();
      if (raw.length > 0) {
        raw.split(',').forEach(function(e) {
          var trimmed = e.trim();
          if (trimmed.length > 0) emails.push(trimmed);
        });
      }
    });

    if (emails.length === 0) return { sent: false, error: 'No recipient emails found in Departments sheet.' };

    var formattedAmount = new Intl.NumberFormat('vi-VN').format(orderInfo.totalAmount) + ' VND';

    var subject = '[Đơn hàng mới] ' + orderInfo.orderId + ' - ' + orderInfo.customerName;

    var htmlBody = [
      '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">',
      '  <div style="background:#1a1f5e;padding:20px 24px;">',
      '    <h2 style="color:#fff;margin:0;font-size:18px;">🛒 Thông báo đơn hàng mới</h2>',
      '  </div>',
      '  <div style="padding:24px;">',
      '    <table style="width:100%;border-collapse:collapse;font-size:14px;">',
      '      <tr style="background:#f5f7ff;"><td style="padding:10px 12px;font-weight:600;width:40%;border-bottom:1px solid #eee;">Mã đơn hàng</td><td style="padding:10px 12px;border-bottom:1px solid #eee;">' + orderInfo.orderId + '</td></tr>',
      '      <tr><td style="padding:10px 12px;font-weight:600;border-bottom:1px solid #eee;">Khách hàng</td><td style="padding:10px 12px;border-bottom:1px solid #eee;">' + orderInfo.customerName + '</td></tr>',
      '      <tr style="background:#f5f7ff;"><td style="padding:10px 12px;font-weight:600;border-bottom:1px solid #eee;">Sản phẩm</td><td style="padding:10px 12px;border-bottom:1px solid #eee;">' + orderInfo.productName + '</td></tr>',
      '      <tr><td style="padding:10px 12px;font-weight:600;border-bottom:1px solid #eee;">Số lượng</td><td style="padding:10px 12px;border-bottom:1px solid #eee;">' + orderInfo.quantity + '</td></tr>',
      '      <tr style="background:#f5f7ff;"><td style="padding:10px 12px;font-weight:600;border-bottom:1px solid #eee;">Tổng tiền</td><td style="padding:10px 12px;border-bottom:1px solid #eee;color:#e53e3e;font-weight:700;">' + formattedAmount + '</td></tr>',
      '      <tr><td style="padding:10px 12px;font-weight:600;border-bottom:1px solid #eee;">Trạng thái</td><td style="padding:10px 12px;border-bottom:1px solid #eee;"><span style="background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:12px;font-size:12px;">Pending</span></td></tr>',
      '      <tr style="background:#f5f7ff;"><td style="padding:10px 12px;font-weight:600;">Ngày tạo</td><td style="padding:10px 12px;">' + orderInfo.createdAt + '</td></tr>',
      '    </table>',
      '  </div>',
      '  <div style="background:#f8f9fa;padding:12px 24px;font-size:12px;color:#6c757d;text-align:center;">',
      '    Email được gửi tự động từ hệ thống Order Management. Vui lòng không reply.',
      '  </div>',
      '</div>'
    ].join('\n');

    // Send to each recipient (BCC alternative: send once to all)
    var toList = emails.join(',');
    MailApp.sendEmail({
      to      : toList,
      subject : subject,
      htmlBody: htmlBody
    });

    return { sent: true, error: null, recipients: toList };
  } catch(e) {
    Logger.log('Email error: ' + e.toString());
    return { sent: false, error: e.toString() };
  }
}

// ── Expose for manual test in Apps Script editor ──────────────
function testSendEmail() {
  var result = sendOrderNotification_({
    orderId     : 'O-TEST',
    customerName: 'Test Customer',
    productName : 'Test Product',
    quantity    : 2,
    totalAmount : 36000000,
    createdAt   : new Date().toISOString(),
    status      : 'Pending'
  });
  Logger.log(JSON.stringify(result));
}
