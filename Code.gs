/**
 * ระบบสั่งจองของที่ระลึก (เสื้อ / แก้ว / แหวน)
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีติดตั้งดูใน README.md
 */

var SHEET_PRODUCTS = 'Products';
var SHEET_ORDERS = 'Orders';
var STATUSES = ['สำรวจ', 'รอชำระเงิน', 'ชำระแล้ว', 'กำลังผลิต', 'พร้อมรับ', 'รับแล้ว', 'ยกเลิก'];

/* ---------- Web App ---------- */
var API = ['getShop', 'submitOrder', 'attachSlip', 'trackOrders', 'cancelMyOrder',
  'adminLogin', 'adminGetData', 'adminSetStatus', 'adminSaveProduct', 'adminSetMode', 'adminSetRingSize'];

function doGet() {
  return ContentService.createTextOutput('Souvenir API OK');
}

// Frontend (GitHub Pages) ส่ง POST แบบ text/plain: {"fn":"...","args":[...]}
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    if (API.indexOf(req.fn) < 0) throw new Error('unknown function');
    out = { ok: true, data: this[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- ติดตั้งครั้งแรก: รันฟังก์ชันนี้ 1 ครั้ง ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ต้องสร้างสคริปต์จาก Google Sheets (Extensions > Apps Script)');

  var p = ss.getSheetByName(SHEET_PRODUCTS) || ss.insertSheet(SHEET_PRODUCTS);
  if (p.getLastRow() === 0) {
    p.appendRow(['ID', 'ชื่อสินค้า', 'ราคา', 'ตัวเลือก (คั่นด้วย ,)', 'ชื่อตัวเลือก', 'สต็อก (ว่าง=ไม่จำกัด)', 'เปิดขาย', 'รูป (URL)', 'รายละเอียด', 'รูปใหญ่ (URL)']);
    productRows_().forEach(function (r) { p.appendRow(r); });
    p.getRange('1:1').setFontWeight('bold').setBackground('#e8eefc');
    p.setFrozenRows(1);
  }

  var o = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
  if (o.getLastRow() === 0) {
    o.appendRow(['เลขที่ออเดอร์', 'เวลา', 'ชื่อ-สกุล', 'เบอร์โทร', 'รุ่น', 'รายการ', 'ยอดรวม', 'สถานะ', 'สลิป', 'หมายเหตุ', 'รายการ(JSON)', 'ตัดสต็อก', 'ขนาดแหวน']);
    o.getRange('1:1').setFontWeight('bold').setBackground('#e8eefc');
    o.setFrozenRows(1);
    o.getRange('H2:H').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  }

  if (!o.getRange(1, 13).getValue()) o.getRange(1, 13).setValue('ขนาดแหวน').setFontWeight('bold').setBackground('#e8eefc');
  o.getRange('M2:M').setNumberFormat('@');

  var props = PropertiesService.getScriptProperties();
  var newPass = '';
  if (!props.getProperty('ADMIN_PASS')) {
    newPass = Math.random().toString(36).slice(2, 10);
    props.setProperty('ADMIN_PASS', newPass);
  }
  if (!props.getProperty('PAY_INFO')) props.setProperty('PAY_INFO', 'ธนาคาร xxx เลขที่ xxx-x-xxxxx-x ชื่อบัญชี xxxxxxxx');
  if (!props.getProperty('SHOP_TITLE')) props.setProperty('SHOP_TITLE', 'สั่งจองของที่ระลึก');
  // MODE: survey = สำรวจความต้องการ (ยังไม่ชำระเงิน) | pay = เปิดชำระเงิน | order = จองพร้อมชำระเงินทันที | closed = ปิด
  if (!props.getProperty('MODE')) props.setProperty('MODE', 'survey');
  o.getRange('H2:H').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  return 'setup เสร็จแล้ว' + (newPass ? ' — รหัสแอดมิน: ' + newPass : '') + ' (เปลี่ยนได้ที่ Project Settings > Script properties > ADMIN_PASS)';
}

/* ---------- สินค้าเริ่มต้น ----------
 * ราคา 0 = "ราคาแจ้งภายหลัง" (แก้ราคาในชีต Products หรือหน้าแอดมิน)
 * ตัวเลือกแบบ "สี / ไซซ์" ใช้ " / " คั่น หน้าเว็บจะแยกเป็น 2 ช่องให้อัตโนมัติ */
var SHIRT_SIZES = ['S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL'];
var RING_MEASURE = 'ขอวัดที่แผนกวิชา';
function ringOptions_() {
  var o = [RING_MEASURE];
  for (var n = 42; n <= 66; n++) o.push(String(n));
  return o.join(',');
}
var RING_DESC = 'เงินแท้ / สแตนเลส 316L (ตามงบประมาณ) ลงยาสีฟ้า หน้าแหวน 16-18 มม. — เลือกเบอร์นิ้ว 42-66 หรือเลือก "ขอวัดที่แผนกวิชา" แล้วไปวัดที่แผนกภายหลัง (กดปุ่มด้านล่างเพื่อดูวิธีวัดไซซ์ด้วยตัวเอง)';

function productRows_() {
  var colors = ['ดำ', 'ขาว'], sizes = SHIRT_SIZES, combo = [];
  colors.forEach(function (c) { sizes.forEach(function (s) { combo.push(c + ' / ' + s); }); });
  var cs = combo.join(',');
  return [
    ['polo', 'เสื้อโปโล (ปักโลโก้)', 0, cs, 'สี / ไซซ์', '', true, 'img/polo.jpg',
      'เสื้อโปโล ปักโลโก้ มีกระเป๋า ปกโปโลปักเส้นคารู หลังปักโลโก้ ปลายแขนปัก RAIKHING', 'img/poster-polo.jpg'],
    ['tee', 'เสื้อยืด (สกรีน)', 0, cs, 'สี / ไซซ์', '', true, 'img/tee.jpg',
      'เสื้อยืดสกรีน มีกระเป๋า คอกลมใส่สบาย ลายกราฟิกด้านข้าง', 'img/poster-tee.jpg'],
    ['mug', 'แก้วเก็บความเย็น', 0, 'ขาว,ดำ', 'สี', '', true, 'img/mug.jpg',
      'เก็บความเย็นได้ 12 ชั่วโมง เก็บความร้อนได้ 8 ชั่วโมง วัสดุสแตนเลส SUS304 ฝาพลาสติก PC', 'img/poster-mug.jpg'],
    ['ring', 'แหวนช่างไฟฟ้ากำลัง 35 ปี', 0, ringOptions_(), 'ขนาดแหวน', '', true, 'img/ring.jpg',
      RING_DESC, 'img/poster-ring.jpg']
  ];
}

/* อัปเดตเฉพาะตัวเลือกไซซ์ (เสื้อถึง 6XL / แหวน 42-66 + ขอวัดที่แผนก) โดยไม่แตะราคา สต็อก หรือค่าอื่น */
function updateSizes() {
  var ps = sheet_(SHEET_PRODUCTS), v = ps.getDataRange().getValues(), rows = productRows_(), by = {};
  rows.forEach(function (r) { by[r[0]] = r; });
  var done = [];
  for (var i = 1; i < v.length; i++) {
    var r = by[String(v[i][0])];
    if (!r || (r[0] !== 'polo' && r[0] !== 'tee' && r[0] !== 'ring' && r[0] !== 'mug')) continue;
    if (r[0] !== 'mug') ps.getRange(i + 1, 4, 1, 2).setValues([[r[3], r[4]]]);
    if (r[0] === 'ring' || r[0] === 'mug') ps.getRange(i + 1, 9).setValue(r[8]);
    done.push(r[0]);
  }
  return 'อัปเดตแล้ว: ' + done.join(', ');
}

/* รันเมื่อต้องการรีเซ็ตเฉพาะชีต Products เป็นค่าเริ่มต้นใหม่ (ไม่แตะออเดอร์) */
function resetProducts() {
  var p = sheet_(SHEET_PRODUCTS);
  p.clear();
  p.appendRow(['ID', 'ชื่อสินค้า', 'ราคา', 'ตัวเลือก (คั่นด้วย ,)', 'ชื่อตัวเลือก', 'สต็อก (ว่าง=ไม่จำกัด)', 'เปิดขาย', 'รูป (URL)', 'รายละเอียด', 'รูปใหญ่ (URL)']);
  productRows_().forEach(function (r) { p.appendRow(r); });
  p.getRange('1:1').setFontWeight('bold').setBackground('#e8eefc');
  p.setFrozenRows(1);
  return 'รีเซ็ตสินค้าแล้ว — อย่าลืมกรอกราคาในคอลัมน์ C';
}

/* ---------- Helpers ---------- */
function prop_(k, d) { return PropertiesService.getScriptProperties().getProperty(k) || d; }
function sheet_(n) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  if (!s) throw new Error('ยังไม่ได้รัน setup()');
  return s;
}
function checkAdmin_(pass) {
  var real = prop_('ADMIN_PASS', '');
  if (!real || String(pass) !== String(real)) throw new Error('รหัสแอดมินไม่ถูกต้อง');
}
function readProducts_() {
  var v = sheet_(SHEET_PRODUCTS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (!r[0]) continue;
    out.push({
      row: i + 1, id: String(r[0]), name: String(r[1]), price: Number(r[2]) || 0,
      options: String(r[3] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
      optionLabel: String(r[4] || 'ตัวเลือก'),
      stock: r[5] === '' || r[5] === null ? null : Number(r[5]),
      active: r[6] === true || String(r[6]).toUpperCase() === 'TRUE',
      image: String(r[7] || ''), desc: String(r[8] || ''), poster: String(r[9] || '')
    });
  }
  return out;
}

/* ---------- Public API (ลูกค้า) ---------- */
function getShop() {
  return {
    title: prop_('SHOP_TITLE', 'สั่งจองของที่ระลึก'),
    payInfo: prop_('PAY_INFO', ''),
    mode: prop_('MODE', 'survey'),
    products: readProducts_().filter(function (p) { return p.active; }).map(function (p) {
      return { id: p.id, name: p.name, price: p.price, options: p.options, optionLabel: p.optionLabel,
               soldOut: p.stock !== null && p.stock <= 0, stock: p.stock, image: p.image, desc: p.desc, poster: p.poster };
    })
  };
}

function submitOrder(data) {
  var mode = prop_('MODE', 'survey');
  if (mode !== 'survey' && mode !== 'order') throw new Error(mode === 'pay' ? 'ปิดรับสำรวจแล้ว อยู่ในขั้นตอนชำระเงิน' : 'ปิดรับจองแล้ว');
  var useStock = mode === 'order';
  var name = String(data.name || '').trim(), phone = String(data.phone || '').replace(/\D/g, '');
  if (!name) throw new Error('กรุณากรอกชื่อ-สกุล');
  if (phone.length < 9) throw new Error('กรุณากรอกเบอร์โทรให้ถูกต้อง');
  if (!data.items || !data.items.length) throw new Error('ยังไม่ได้เลือกสินค้า');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var products = readProducts_(), byId = {};
    products.forEach(function (p) { byId[p.id] = p; });
    var need = {}, lines = [], total = 0, clean = [];

    data.items.forEach(function (it) {
      var p = byId[it.id];
      var qty = Math.floor(Number(it.qty));
      if (!p || !p.active) throw new Error('ไม่พบสินค้า: ' + it.id);
      if (!(qty > 0) || qty > 50) throw new Error('จำนวนไม่ถูกต้อง: ' + p.name);
      if (p.options.length && p.options.indexOf(it.opt) < 0) throw new Error('กรุณาเลือก' + p.optionLabel + ': ' + p.name);
      need[p.id] = (need[p.id] || 0) + qty;
      var sub = p.price * qty; total += sub;
      var label = p.name + (p.options.length ? ' (' + p.optionLabel + ' ' + it.opt + ')' : '') + ' x' + qty;
      lines.push(label);
      clean.push({ id: p.id, name: p.name, opt: it.opt || '', qty: qty, price: p.price });
    });

    // ตรวจสต็อก
    if (useStock) Object.keys(need).forEach(function (id) {
      var p = byId[id];
      if (p.stock !== null && p.stock < need[id]) throw new Error(p.name + ' เหลือ ' + Math.max(p.stock, 0) + ' ชิ้น');
    });
    // ตัดสต็อก
    var ps = sheet_(SHEET_PRODUCTS);
    if (useStock) Object.keys(need).forEach(function (id) {
      var p = byId[id];
      if (p.stock !== null) ps.getRange(p.row, 6).setValue(p.stock - need[id]);
    });

    var os = sheet_(SHEET_ORDERS);
    var orderNo = 'SV' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd') + '-' +
      ('000' + (os.getLastRow())).slice(-3);

    var slipUrl = '';
    if (useStock && data.slip && data.slip.base64) slipUrl = saveSlip_(orderNo, data.slip);

    os.appendRow([orderNo, new Date(), name, "'" + phone, String(data.group || '').trim(),
      lines.join('\n'), total, useStock ? 'รอชำระเงิน' : 'สำรวจ', slipUrl, String(data.note || '').trim(), JSON.stringify(clean), useStock, '']);
    os.getRange(os.getLastRow(), 4).setNumberFormat('@').setValue(phone);
    return { orderNo: orderNo, total: total, lines: lines, survey: !useStock };
  } finally {
    lock.releaseLock();
  }
}

function saveSlip_(orderNo, slip) {
  var folderName = 'สลิปสั่งจองของที่ระลึก';
  var it = DriveApp.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  var blob = Utilities.newBlob(Utilities.base64Decode(slip.base64), slip.mime || 'image/jpeg', orderNo + '.jpg');
  var f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getUrl();
}

function attachSlip(orderNo, phone, slip) {
  var os = sheet_(SHEET_ORDERS), v = os.getDataRange().getValues();
  phone = String(phone).replace(/\D/g, '');
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === orderNo && String(v[i][3]).replace(/\D/g, '') === phone) {
      os.getRange(i + 1, 9).setValue(saveSlip_(orderNo, slip));
      return true;
    }
  }
  throw new Error('ไม่พบออเดอร์');
}

function trackOrders(phone) {
  phone = String(phone || '').replace(/\D/g, '');
  if (phone.length < 9) throw new Error('กรุณากรอกเบอร์โทร');
  var v = sheet_(SHEET_ORDERS).getDataRange().getValues(), out = [];
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][3]).replace(/\D/g, '') === phone) {
      out.push({ orderNo: v[i][0], time: Utilities.formatDate(new Date(v[i][1]), 'Asia/Bangkok', 'd/M/yyyy HH:mm'),
        items: v[i][5], total: v[i][6], status: v[i][7], hasSlip: !!v[i][8], ringSize: String(v[i][12] || '') });
    }
  }
  return { orders: out, payInfo: prop_('PAY_INFO', '') };
}

function cancelMyOrder(orderNo, phone) {
  phone = String(phone || '').replace(/\D/g, '');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var os = sheet_(SHEET_ORDERS), v = os.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (v[i][0] === orderNo && String(v[i][3]).replace(/\D/g, '') === phone) {
        if (v[i][7] !== 'สำรวจ') throw new Error('ยกเลิกเองได้เฉพาะช่วงสำรวจ กรุณาติดต่อแอดมิน');
        os.getRange(i + 1, 8).setValue('ยกเลิก');
        return true;
      }
    }
    throw new Error('ไม่พบออเดอร์');
  } finally { lock.releaseLock(); }
}

/* ---------- Admin API ---------- */
function adminLogin(pass) { checkAdmin_(pass); return { statuses: STATUSES }; }

function adminGetData(pass) {
  checkAdmin_(pass);
  var v = sheet_(SHEET_ORDERS).getDataRange().getValues(), orders = [];
  for (var i = v.length - 1; i >= 1; i--) {
    if (!v[i][0]) continue;
    orders.push({ orderNo: v[i][0], time: Utilities.formatDate(new Date(v[i][1]), 'Asia/Bangkok', 'd/M/yy HH:mm'),
      name: v[i][2], phone: v[i][3], group: v[i][4], items: v[i][5], total: v[i][6],
      status: v[i][7], slip: v[i][8], note: v[i][9], json: v[i][10], ringSize: String(v[i][12] || '') });
  }
  return { orders: orders, products: readProducts_(), mode: prop_('MODE', 'survey') };
}

function adminSetStatus(pass, orderNo, status) {
  checkAdmin_(pass);
  if (STATUSES.indexOf(status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var os = sheet_(SHEET_ORDERS), v = os.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (v[i][0] !== orderNo) continue;
      var old = v[i][7];
      os.getRange(i + 1, 8).setValue(status);
      // ยกเลิก -> คืนสต็อก
      var deducted = v[i][11] === true || String(v[i][11]).toUpperCase() === 'TRUE';
      if (deducted && status === 'ยกเลิก' && old !== 'ยกเลิก') restock_(v[i][10], 1);
      if (deducted && old === 'ยกเลิก' && status !== 'ยกเลิก') restock_(v[i][10], -1);
      return true;
    }
    throw new Error('ไม่พบออเดอร์');
  } finally { lock.releaseLock(); }
}

function restock_(json, sign) {
  var items; try { items = JSON.parse(json); } catch (e) { return; }
  var products = readProducts_(), ps = sheet_(SHEET_PRODUCTS);
  items.forEach(function (it) {
    products.forEach(function (p) {
      if (p.id === it.id && p.stock !== null) {
        p.stock += sign * it.qty;
        ps.getRange(p.row, 6).setValue(p.stock);
      }
    });
  });
}

function adminSaveProduct(pass, row, patch) {
  checkAdmin_(pass);
  var ps = sheet_(SHEET_PRODUCTS);
  if (patch.price !== undefined) ps.getRange(row, 3).setValue(Number(patch.price));
  if (patch.stock !== undefined) ps.getRange(row, 6).setValue(patch.stock === '' ? '' : Number(patch.stock));
  if (patch.active !== undefined) ps.getRange(row, 7).setValue(!!patch.active);
  return true;
}

function adminSetMode(pass, mode) {
  checkAdmin_(pass);
  if (['survey', 'pay', 'order', 'closed'].indexOf(mode) < 0) throw new Error('โหมดไม่ถูกต้อง');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var converted = 0;
    // เปลี่ยนเป็นขั้นชำระเงิน: ออเดอร์ที่ยังเป็น "สำรวจ" -> "รอชำระเงิน"
    if (mode === 'pay') {
      var os = sheet_(SHEET_ORDERS), v = os.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        if (v[i][7] === 'สำรวจ') { os.getRange(i + 1, 8).setValue('รอชำระเงิน'); converted++; }
      }
    }
    PropertiesService.getScriptProperties().setProperty('MODE', mode);
    return { converted: converted };
  } finally { lock.releaseLock(); }
}

/* บันทึกขนาดแหวนหลังวัดนิ้วที่แผนก (พิมพ์เป็นข้อความ เช่น "56" หรือ "56, 58" ถ้าสั่งหลายวง) */
function adminSetRingSize(pass, orderNo, size) {
  checkAdmin_(pass);
  var os = sheet_(SHEET_ORDERS), v = os.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === orderNo) {
      if (!os.getRange(1, 13).getValue()) os.getRange(1, 13).setValue('ขนาดแหวน');
      os.getRange(i + 1, 13).setNumberFormat('@').setValue(String(size || '').trim());
      return true;
    }
  }
  throw new Error('ไม่พบออเดอร์');
}
