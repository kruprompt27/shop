/**
 * ============================================================
 * 🛒 ระบบสั่งซื้อแผนการสอน ครูพร้อมสอน — ป.1–ม.6 (v9.1 — GAS Web App)
 * ============================================================
 * สถาปัตยกรรม: Google Apps Script Web App (ฝัง index.html)
 *             + Google Sheets เป็นฐานข้อมูล + Google Drive เก็บไฟล์
 *
 * วิธีติดตั้ง:
 *   1) สร้าง Google Sheet ใหม่ → Extensions → Apps Script
 *   2) วางโค้ดนี้ใน Code.gs และสร้างไฟล์ HTML ชื่อ "index" วางโค้ดหน้าเว็บ
 *   3) Deploy → New deployment → Web app
 *      Execute as: Me / Who has access: Anyone
 *   4) เปิด Web App URL ที่ได้ — ระบบจะสร้างชีตให้อัตโนมัติในครั้งแรก
 * ============================================================
 */

// ============================================================
// ⚙️ ค่าคงที่ของระบบ (แก้ไขได้ที่นี่)
// ============================================================
const PAGE_NAME = 'ครูพร้อมสอน';
const SHEET_SETTINGS = 'Settings';
const SHEET_PRODUCTS   = 'แผนประถม';  // ระดับประถม ป.1–ป.6 (เปลี่ยนชื่อจากชีต "แผนการสอน"/"Products" เดิมให้อัตโนมัติ)
const SHEET_PRODUCTS_M = 'แผนมัธยม';  // ระดับมัธยม ม.1–ม.6 — แยกชีตเพื่อจัดการง่าย (หน้าร้านเห็นรวมกัน)
const SHEET_ORDERS   = 'Orders';
const SHEET_SLIPHASH = 'SlipHash';            // ลายนิ้วมือสลิป (กันสลิปซ้ำ) — ชีตซ่อน
const BACKUP_FOLDER_NAME = 'KruPromSorn_Backups'; // โฟลเดอร์สำรองข้อมูลรายคืน
const BACKUP_KEEP = 14;                       // เก็บไฟล์สำรองย้อนหลังกี่ชุด
const SHOP_URL = 'https://kruprompt27.github.io/shop/'; // หน้าร้านสำหรับลูกค้า
const PREVIEW_PAGE_URL = 'https://kruprompt27.github.io/shop/preview.html'; // หน้าตัวอย่างสื่อ (โฮสต์โฟลเดอร์เดียวกับร้าน)

const SLIPS_FOLDER_NAME  = 'KruPromSorn_Slips';   // โฟลเดอร์เก็บสลิปใน Drive
const IMAGES_FOLDER_NAME = 'KruPromSorn_Images';  // โฟลเดอร์เก็บโลโก้/QR

// สถานะคำสั่งซื้อ (คงชุดเดียวกับระบบเดิมเพื่อความคุ้นเคย)
const STATUSES = {
  PENDING:   'รอชำระเงิน',
  PAID:      'ชำระเงินแล้ว',
  APPROVED:  'อนุมัติส่งแผน',
  COMPLETED: 'ส่งแผนเสร็จสิ้น',
  CANCELLED: 'ยกเลิก'
};

// คอลัมน์ชีต Orders (1-indexed)
const O = {
  TIMESTAMP: 1, ORDER_ID: 2, CUSTOMER: 3, ITEMS_TEXT: 4, EMAIL: 5, PHONE_FB: 6,
  ITEM_COUNT: 7, ITEMS_JSON: 8, SUBTOTAL: 9, DISCOUNT_PCT: 10,
  DISCOUNT_BAHT: 11, NET_TOTAL: 12, STATUS: 13, SLIP_URL: 14, NOTE: 15
};
const O_COUNT = 15;

// ═══ โครงชีตสินค้าแบบแยกรายชั้น: ป.1–ป.6, ม.1–ม.6, VIP & ระบบงาน, อื่นๆ ═══
const GRADE_SHEETS   = ['ป.1','ป.2','ป.3','ป.4','ป.5','ป.6','ม.1','ม.2','ม.3','ม.4','ม.5','ม.6'];
const SHEET_VIP_WORK = 'VIP & ระบบงาน';   // สินค้าประเภท "กลุ่ม VIP" และ "ระบบงาน" ทุกชั้น
const SHEET_OTHER    = 'อื่นๆ';           // แผนการสอนที่ชั้นไม่อยู่ในรายการ (เช่น อ.1 หรือเว้นว่าง)

// รายชื่อชีตสินค้าที่ระบบอ่าน (เรียงตามลำดับแสดงผล)
// ช่วงเปลี่ยนผ่าน: ถ้าชีตรวมรุ่นก่อน (แผนประถม/แผนมัธยม/แผนการสอน) ยังอยู่ ให้อ่านต่อด้วย
// จนกว่าจะกดเมนู "📗 แยกชีตรายชั้น" — ร้านจึงขายได้ไม่สะดุดระหว่างยังไม่ย้ายข้อมูล
function productSheetNames_() {
  const names = GRADE_SHEETS.concat([SHEET_VIP_WORK, SHEET_OTHER]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_PRODUCTS, SHEET_PRODUCTS_M, 'แผนการสอน', 'Products'].forEach(function (n) {
    if (ss.getSheetByName(n)) names.push(n);
  });
  return names;
}
// เลือกชีตปลายทางของสินค้า: ประเภท VIP/ระบบงาน → ชีตรวม, ชั้นตรงรายการ → ชีตชั้นนั้น, นอกนั้น → อื่นๆ
function sheetNameForProduct_(grade, category) {
  const cat = String(category || 'แผนการสอน').trim();
  if (cat === 'ระบบงาน' || cat === 'กลุ่ม VIP') return SHEET_VIP_WORK;
  const g = String(grade || '').trim();
  return GRADE_SHEETS.indexOf(g) >= 0 ? g : SHEET_OTHER;
}
// เรียงชื่อวิชาตามอักษรไทย (ทนกรณี runtime ไม่มีตัวเทียบ locale)
function thaiCompare_(a, b) {
  a = String(a || ''); b = String(b || '');
  try { return a.localeCompare(b, 'th'); }
  catch (e) { return a < b ? -1 : a > b ? 1 : 0; }
}
// เรียงแถวในชีตตามชื่อวิชา ก-ฮ — อ่านทั้งช่วงแล้วเขียนกลับ "ทั้งแถว" ข้อมูลทุกช่องจึงย้ายตามกันครบ
function sortProductSheet_(sh) {
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < 3) return;
  const vip = isVipSheet_(sh.getName());
  const width = vip ? PV_COUNT : P_COUNT;
  const cSub = vip ? PV.SUBJECT : P.SUBJECT;
  const cAdd = vip ? PV.ADDED : P.ADDED;
  const rng = sh.getRange(2, 1, last - 1, width);
  const vals = rng.getValues();
  vals.sort(function (x, y) {
    return thaiCompare_(x[cSub - 1], y[cSub - 1]) ||
           (vip ? 0 : (thaiCompare_(x[P.CURRICULUM - 1], y[P.CURRICULUM - 1]) ||
                       (Number(x[P.HOURS - 1]) || 0) - (Number(y[P.HOURS - 1]) || 0)));
  });
  rng.setValues(vals);
  rng.offset(0, cAdd - 1, vals.length, 1).setNumberFormat('@');
}

function isStatus_(v) {
  v = String(v || '').trim();
  return v === STATUSES.PENDING || v === STATUSES.PAID || v === STATUSES.APPROVED ||
         v === STATUSES.COMPLETED || v === STATUSES.CANCELLED;
}

// แปลง JSON รายการอย่างปลอดภัย — ถ้าข้อมูลเพี้ยน (เช่นแถวคอลัมน์เคลื่อน) คืน [] แทนที่จะพังทั้งระบบ
function parseItems_(raw) {
  try {
    const v = JSON.parse(String(raw == null ? '[]' : raw) || '[]');
    return Array.isArray(v) ? v : [];
  } catch (e) { return []; }
}

// แปลงรายการเป็นข้อความอ่านง่าย ใช้ในคอลัมน์ "วิชาที่สั่ง"
function itemsText_(items) {
  return (items || []).map(function (it) {
    var t = (String(it.subject || '') + ' ' + String(it.grade || '')).trim();
    var h = String(it.hours || '').replace(/\D/g, '');
    return t && h ? t + ' (' + h + 'ชม.)' : t;
  }).filter(Boolean).join('\n');
}

// 🔤 สร้างข้อความ "วิชาที่สั่ง" ใหม่ทุกแถวจาก JSON (ใช้หลังเปลี่ยนรูปแบบ เช่น เติมชม.)
function refreshItemsTextAll() {
  const sh = getSheet_(SHEET_ORDERS);
  const last = sh.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('ยังไม่มีออเดอร์'); return; }
  const jsons = sh.getRange(2, O.ITEMS_JSON, last - 1, 1).getValues();
  const texts = sh.getRange(2, O.ITEMS_TEXT, last - 1, 1).getValues();
  var changed = 0;
  for (var i = 0; i < jsons.length; i++) {
    const items = parseItems_(jsons[i][0]);
    if (!items.length) continue;
    const t = itemsText_(items);
    if (t && t !== String(texts[i][0])) { texts[i][0] = t; changed++; }
  }
  if (changed) sh.getRange(2, O.ITEMS_TEXT, last - 1, 1).setNumberFormat('@').setValues(texts);
  SpreadsheetApp.getActiveSpreadsheet().toast('อัปเดตข้อความวิชาที่สั่งแล้ว ' + changed + ' แถว', '🔤', 5);
}

// ข้อความท้ายมาตรฐาน (เหมือนกันทุกวิชา) — ใช้เติมคอลัมน์ "ข้อความท้าย (ส่งให้ลูกค้า)"
const STD_FOOTER = [
  '📱 สื่อการสอน (Interactive HTML Media)',
  '   สามารถเข้าใช้งานได้ที่: https://kruprompt27.github.io/krupromsue/',
  '(หมายเหตุ: เข้าใช้งานด้วย Gmail ที่คุณครูระบุไว้ตอนสั่งซื้อ)',
  '____________________________________________________________',
  '🧾 เอกสารเพิ่มเติม',
  'แบบฟอร์มบันทึกหลังการสอน, เครื่องมือวัดและประเมินผล (ที่สามารถเลือกใช้เพิ่มเติมจากในแผนได้), แนวทางการแก้ปัญหา และคำถามที่พบบ่อย (FAQ)',
  '📥 ดาวน์โหลดได้ที่:',
  'https://drive.google.com/drive/folders/1w8Yb7w2oUQXG4iwDiy9rZ78K7tsGJN5u?usp=sharing',
  ' ',
  '📝 ไฟล์ทั้งหมดอยู่ในรูปแบบ Word และ PDF สามารถแก้ไขและนำไปปรับใช้ได้ทันทีค่ะ',
  '✅ หากมีข้อสงสัยหรือต้องการสอบถามเพิ่มเติม ทักแชทได้นะคะ 😊',
  '🔒 หมายเหตุ: สงวนลิขสิทธิ์ ไม่อนุญาตให้เผยแพร่ ดัดแปลง หรือแจกจ่ายต่อโดยไม่ได้รับอนุญาต เพื่อเป็นการสนับสนุนผู้ผลิตผลงานค่ะ'
].join('\n');

// คอลัมน์ชีต Products (1-indexed)
// POST_NOTE = "รายละเอียดแผน" (เฉพาะวิชา) , FOOTER = "ข้อความท้าย (ส่งให้ลูกค้า)" (เหมือนกันทุกวิชา)
const P = {
  ID: 1, GRADE: 2, CURRICULUM: 3, SUBJECT: 4, HOURS: 5, PRICE: 6,
  DRIVE_LINK: 7, EXAMPLE_LINK: 8, POST_NOTE: 9, FOOTER: 10,
  ACTIVE: 11, ADDED: 12, CATEGORY: 13, GROUP: 14
};
const P_COUNT = 14;

// หัวตารางมาตรฐานชีตสินค้า (ใช้ทั้งตอนสร้างชีตใหม่และซ่อมโครง)
const PRODUCT_HEADER_STD = [
  'ID', 'ระดับชั้น', 'หลักสูตร', 'ชื่อวิชา/รายการ', 'เวลาเรียน (ชม.)', 'ราคา (บาท)',
  'ลิงก์ Drive (ส่งให้ลูกค้า)', 'ลิงก์ตัวอย่าง', 'รายละเอียดแผน', 'ข้อความท้าย (ส่งให้ลูกค้า)',
  'เปิดขาย', 'วันที่เพิ่ม', 'ประเภท', 'กลุ่มวิชา'
];

/* ═══ ชีต "VIP & ระบบงาน" ใช้โครงย่อของตัวเอง (ไม่มี ชั้น/หลักสูตร/ชม./กลุ่มวิชา) ═══
 * สินค้าประเภทนี้ไม่ผูกกับชั้นเรียน — คอลัมน์พวกนั้นจึงไม่จำเป็น ตามที่ Founder กำหนด */
const PV = { ID: 1, SUBJECT: 2, PRICE: 3, DRIVE_LINK: 4, EXAMPLE_LINK: 5, POST_NOTE: 6, FOOTER: 7, ACTIVE: 8, ADDED: 9, CATEGORY: 10 };
const PV_COUNT = 10;
const PRODUCT_HEADER_VIP = ['ID', 'ชื่อวิชา/รายการ', 'ราคา (บาท)', 'ลิงก์ Drive (ส่งให้ลูกค้า)',
  'ลิงก์ตัวอย่าง', 'รายละเอียด', 'ข้อความท้าย (ส่งให้ลูกค้า)', 'เปิดขาย', 'วันที่เพิ่ม', 'ประเภท'];
function isVipSheet_(name) { return name === SHEET_VIP_WORK; }
// แปลตำแหน่งคอลัมน์มาตรฐาน (P.*) → คอลัมน์จริงของชีตนั้น (0 = ชีตนี้ไม่มีคอลัมน์นี้)
function colFor_(sheetName, pcol) {
  if (!isVipSheet_(sheetName)) return pcol;
  return ({ 1: PV.ID, 4: PV.SUBJECT, 6: PV.PRICE, 7: PV.DRIVE_LINK, 8: PV.EXAMPLE_LINK,
    9: PV.POST_NOTE, 10: PV.FOOTER, 11: PV.ACTIVE, 12: PV.ADDED, 13: PV.CATEGORY })[pcol] || 0;
}
// สร้างแถวข้อมูลตามโครงของชีตปลายทาง (ใช้ตอนเพิ่ม/ย้ายข้ามชีต)
function buildRowFor_(sheetName, o) {
  if (isVipSheet_(sheetName)) {
    return [o.id, o.subject, o.price || '', o.driveLink || '', o.exampleLink || '',
      o.postNote || '', o.footer || STD_FOOTER, o.active ? 'TRUE' : 'FALSE', o.added, o.category];
  }
  return [o.id, o.grade || '', o.curriculum || '', o.subject, o.hours || '', o.price || '',
    o.driveLink || '', o.exampleLink || '', o.postNote || '', o.footer || STD_FOOTER,
    o.active ? 'TRUE' : 'FALSE', o.added, o.category, o.group || ''];
}
function sheetColCount_(name) { return isVipSheet_(name) ? PV_COUNT : P_COUNT; }

/* ═══ กลุ่มวิชา (ตัวกรองหน้าร้าน) — "กลุ่มคือข้อมูล ไม่ใช่โค้ด" ═══
 * ลำดับมาตรฐานต่อหลักสูตร ใช้เรียงชิปหน้าร้าน / กลุ่มใหม่ที่พิมพ์เพิ่มในชีตจะต่อท้ายเอง
 * 68 = แบ่งตามความสามารถผู้เรียน (ฐานสมรรถนะ) / 51 = แบ่งตามกลุ่มสาระ (ใช้ ป.4-6 และ ม.1-ม.6) */
const GROUPS_68 = ['การอ่านเขียน', 'การคิดคำนวณ', 'วิทยาศาสตร์และเทคโนโลยี', 'สังคมและความเป็นพลเมือง',
  'เศรษฐกิจและการเงิน', 'สุขภาพกายและจิต', 'ศิลปะและวัฒนธรรม', 'กิจกรรมพัฒนาผู้เรียน'];
const GROUPS_51 = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์และเทคโนโลยี', 'สังคมศึกษาฯ',
  'สุขศึกษาและพลศึกษา', 'ศิลปะ', 'การงานอาชีพ', 'ภาษาต่างประเทศ', 'กิจกรรมพัฒนาผู้เรียน'];

// ตาราง keyword จัดกลุ่มอัตโนมัติ — เรียงลำดับสำคัญ: กติกาเฉพาะเจาะจงอยู่บน (เช็คก่อน)
const GROUP_RULES_68 = [
  ['กิจกรรมพัฒนาผู้เรียน', ['ชุมนุม', 'ลูกเสือ', 'เนตรนารี', 'ยุวกาชาด', 'แนะแนว', 'กิจกรรมพัฒนา']],
  ['วิทยาศาสตร์และเทคโนโลยี', ['วิทยาศาสตร์', 'วิทยาการคำนวณ', 'คอมพิวเตอร์', 'เทคโนโลยี']],
  ['สังคมและความเป็นพลเมือง', ['สังคม', 'ประวัติศาสตร์', 'หน้าที่พลเมือง', 'ทุจริต', 'พระพุทธ']],
  ['เศรษฐกิจและการเงิน', ['เศรษฐกิจ', 'การเงิน', 'การงานอาชีพ']],
  ['สุขภาพกายและจิต', ['สุขศึกษา', 'พลศึกษา', 'สุขภาพ']],
  ['ศิลปะและวัฒนธรรม', ['ศิลปะ', 'ดนตรี', 'นาฏศิลป์', 'ทัศนศิลป์', 'วัฒนธรรม']],
  ['การคิดคำนวณ', ['คณิตศาสตร์', 'คณิต']],
  ['การอ่านเขียน', ['ภาษาไทย', 'ภาษาอังกฤษ', 'อังกฤษ', 'การอ่าน', 'การเขียน']]   // ✓ อังกฤษ (68) → การอ่านเขียน
];
const GROUP_RULES_51 = [
  ['กิจกรรมพัฒนาผู้เรียน', ['ชุมนุม', 'ลูกเสือ', 'เนตรนารี', 'ยุวกาชาด', 'แนะแนว', 'กิจกรรมพัฒนา']],
  ['วิทยาศาสตร์และเทคโนโลยี', ['วิทยาศาสตร์', 'ฟิสิกส์', 'เคมี', 'ชีววิทยา', 'ดาราศาสตร์', 'วิทยาการคำนวณ', 'คอมพิวเตอร์', 'เทคโนโลยี', 'การออกแบบ']],
  ['สังคมศึกษาฯ', ['สังคม', 'ประวัติศาสตร์', 'หน้าที่พลเมือง', 'ทุจริต', 'พระพุทธ']],
  ['สุขศึกษาและพลศึกษา', ['สุขศึกษา', 'พลศึกษา']],
  ['ศิลปะ', ['ศิลปะ', 'ดนตรี', 'นาฏศิลป์', 'ทัศนศิลป์']],
  ['การงานอาชีพ', ['การงานอาชีพ', 'การงาน']],
  ['ภาษาต่างประเทศ', ['อังกฤษ', 'ภาษาจีน', 'ภาษาญี่ปุ่น', 'ต่างประเทศ', 'English', 'Smile']],
  ['คณิตศาสตร์', ['คณิตศาสตร์', 'คณิต']],
  ['ภาษาไทย', ['ภาษาไทย', 'ไทย']]
];
// เดากลุ่มจากชื่อวิชา+หลักสูตร — จับไม่ได้คืนค่าว่าง (ให้เติมเองในชีต ปลอดภัยกว่าเดาผิด)
function autoGroupFor_(curriculum, subject) {
  const cur = String(curriculum || '').replace(/[^\d]/g, '');
  const s = String(subject || '');
  if (!s) return '';
  const rules = cur === '68' ? GROUP_RULES_68 : cur === '51' ? GROUP_RULES_51 : null;
  if (!rules) return '';
  for (let i = 0; i < rules.length; i++) {
    for (let k = 0; k < rules[i][1].length; k++) {
      if (s.indexOf(rules[i][1][k]) >= 0) return rules[i][0];
    }
  }
  return '';
}

// ประเภทสินค้า
const CATEGORIES = ['แผนการสอน', 'ระบบงาน', 'กลุ่ม VIP'];

// ค่าตั้งต้นชีต Settings (key | value)
const DEFAULT_SETTINGS = [
  ['pageName', PAGE_NAME],
  ['slogan', 'แผนการสอน ป.1–ม.6 พร้อมใช้ ส่งไวทาง Gmail'],
  // บัญชีแอดมิน: ชื่อ=PIN:สิทธิ์ (owner เห็นทุกเมนู / staff เห็นเฉพาะออเดอร์+ตั้งค่า)
  ['admins', 'yui=080766:owner, pinky=1234:staff'],
  ['telegramBotToken', ''],
  ['telegramChatId', ''],
  ['autoSend', 'off'],   // 🤖 ส่งแผนอัตโนมัติเมื่อยอดสลิปตรง (on/off)
  ['promptpayName', ''],
  ['promptpayNumber', ''],
  ['qrURL', ''],
  ['logoURL', ''],
  ['discountMinItems', '3'],
  ['discountPercent', '5'],
  ['priceTable', '40=200, 60=220, 80=250, 120=280, 160=300, 200=350'],
  ['announcement', ''],
  // 🖋️ ข้อความหน้าร้าน — เว้นว่าง = ใช้ข้อความมาตรฐานที่ฝังในหน้าเว็บ (แก้ที่นี่ ไม่ต้องอัปไฟล์ GitHub)
  ['shopTextDiscount', ''],
  ['shopTextVip', ''],
  ['shopTextGmail1', ''],
  ['shopTextGmail2', ''],
  ['shopTextFooter', '']
];

// แคตตาล็อกวิชาตั้งต้น — [ชั้น, หลักสูตร, วิชา, ชั่วโมง]
// ครอบคลุม ป.1–ป.6 (หลักสูตร 68 และ 51) และ ม.1–ม.6 (หลักสูตร 51)
// ราคา/ลิงก์ Drive ปล่อยว่าง รอกรอกในเมนู "จัดการแผนการสอน"
const SEED_PRODUCTS = [
  // ═══ ระดับประถม ป.1–ป.6 ═══
  ['ป.1', '68', 'ภาษาไทย', 200],
  ['ป.1', '68', 'ภาษาอังกฤษ', 160],
  ['ป.1', '68', 'ภาษาอังกฤษ', 200],
  ['ป.1', '68', 'คณิตศาสตร์', 200],
  ['ป.1', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 40],
  ['ป.1', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.1', '68', 'สังคมและความเป็นพลเมือง', 40],
  ['ป.1', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.1', '68', 'ประวัติศาสตร์', 40],
  ['ป.1', '68', 'เศรษฐกิจและการเงิน', 40],
  ['ป.1', '68', 'เศรษฐกิจและการเงิน', 80],
  ['ป.1', '68', 'สุขภาพกายและจิต', 40],
  ['ป.1', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 40],
  ['ป.1', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 80],
  ['ป.1', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.1', '68', 'วิทยาการคำนวณ', 40],
  ['ป.1', '68', 'พลศึกษา', 40],
  ['ป.1', '68', 'การงานอาชีพ', 40],
  ['ป.1', '68', 'ดนตรี-นาฏศิลป์', 40],
  ['ป.2', '68', 'ภาษาไทย', 200],
  ['ป.2', '68', 'ภาษาอังกฤษ', 160],
  ['ป.2', '68', 'ภาษาอังกฤษ', 200],
  ['ป.2', '68', 'คณิตศาสตร์', 200],
  ['ป.2', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 40],
  ['ป.2', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.2', '68', 'สังคมและความเป็นพลเมือง', 40],
  ['ป.2', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.2', '68', 'ประวัติศาสตร์', 40],
  ['ป.2', '68', 'เศรษฐกิจและการเงิน', 40],
  ['ป.2', '68', 'เศรษฐกิจและการเงิน', 80],
  ['ป.2', '68', 'สุขภาพกายและจิต', 40],
  ['ป.2', '68', 'สุขภาพกายและจิต', 80],
  ['ป.2', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 40],
  ['ป.2', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 80],
  ['ป.2', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.2', '68', 'วิทยาการคำนวณ', 40],
  ['ป.2', '68', 'พลศึกษา', 40],
  ['ป.2', '68', 'การงานอาชีพ', 40],
  ['ป.2', '68', 'ดนตรี-นาฏศิลป์', 40],
  ['ป.3', '68', 'ภาษาไทย', 160],
  ['ป.3', '68', 'ภาษาไทย', 200],
  ['ป.3', '68', 'ภาษาอังกฤษ', 160],
  ['ป.3', '68', 'ภาษาอังกฤษ', 200],
  ['ป.3', '68', 'คณิตศาสตร์', 200],
  ['ป.3', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 40],
  ['ป.3', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.3', '68', 'สังคมและความเป็นพลเมือง', 40],
  ['ป.3', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.3', '68', 'ประวัติศาสตร์', 40],
  ['ป.3', '68', 'เศรษฐกิจและการเงิน', 40],
  ['ป.3', '68', 'เศรษฐกิจและการเงิน', 80],
  ['ป.3', '68', 'สุขภาพกายและจิต', 40],
  ['ป.3', '68', 'สุขภาพกายและจิต', 80],
  ['ป.3', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 40],
  ['ป.3', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 80],
  ['ป.3', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.3', '68', 'วิทยาการคำนวณ', 40],
  ['ป.3', '68', 'พลศึกษา', 40],
  ['ป.3', '68', 'การงานอาชีพ', 40],
  ['ป.3', '68', 'ดนตรี-นาฏศิลป์', 40],
  ['ป.4', '68', 'ภาษาไทย', 160],
  ['ป.4', '68', 'ภาษาไทย (บูรณาการ)', 40],
  ['ป.4', '68', 'ภาษาอังกฤษ', 80],
  ['ป.4', '68', 'ภาษาอังกฤษ (บูรณาการ)', 40],
  ['ป.4', '68', 'คณิตศาสตร์', 160],
  ['ป.4', '68', 'คณิตศาสตร์ (บูรณาการ)', 40],
  ['ป.4', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.4', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.4', '68', 'ประวัติศาสตร์', 40],
  ['ป.4', '68', 'เศรษฐกิจและการเงิน', 40],
  ['ป.4', '68', 'สุขภาพกายและสุขภาวะจิต', 80],
  ['ป.4', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 40],
  ['ป.4', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.4', '51', 'ภาษาไทย (หลักภาษาและการใช้ภาษาไทย)', 80],
  ['ป.4', '51', 'ภาษาไทย (วรรณคดีและวรรณกรรม)', 80],
  ['ป.4', '51', 'ภาษาอังกฤษ Smile', 80],
  ['ป.4', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ป.4', '51', 'คอมพิวเตอร์', 40],
  ['ป.4', '51', 'วิทยาการคำนวณ', 40],
  ['ป.4', '51', 'การงานอาชีพ', 40],
  ['ป.5', '68', 'ภาษาไทย', 120],
  ['ป.5', '68', 'ภาษาไทย (บูรณาการ)', 40],
  ['ป.5', '68', 'ภาษาอังกฤษ', 60],
  ['ป.5', '68', 'ภาษาอังกฤษ (บูรณาการ)', 40],
  ['ป.5', '68', 'คณิตศาสตร์', 120],
  ['ป.5', '68', 'คณิตศาสตร์ (บูรณาการ)', 40],
  ['ป.5', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.5', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.5', '68', 'ประวัติศาสตร์', 40],
  ['ป.5', '68', 'เศรษฐกิจและการเงิน', 80],
  ['ป.5', '68', 'สุขภาพกายและสุขภาวะจิต', 80],
  ['ป.5', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 80],
  ['ป.5', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.5', '51', 'ภาษาไทย (หลักภาษาและการใช้ภาษาไทย)', 80],
  ['ป.5', '51', 'ภาษาไทย (วรรณคดีและวรรณกรรม)', 80],
  ['ป.5', '51', 'ภาษาอังกฤษ Smile', 80],
  ['ป.5', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ป.5', '51', 'คอมพิวเตอร์', 40],
  ['ป.5', '51', 'วิทยาการคำนวณ', 40],
  ['ป.5', '51', 'การงานอาชีพ', 40],
  ['ป.6', '68', 'ภาษาไทย', 80],
  ['ป.6', '68', 'ภาษาไทย (บูรณาการ)', 40],
  ['ป.6', '68', 'ภาษาไทย (บูรณาการ)', 80],
  ['ป.6', '68', 'ภาษาอังกฤษ', 40],
  ['ป.6', '68', 'ภาษาอังกฤษ (บูรณาการ)', 40],
  ['ป.6', '68', 'ภาษาอังกฤษ (บูรณาการ)', 80],
  ['ป.6', '68', 'คณิตศาสตร์', 80],
  ['ป.6', '68', 'คณิตศาสตร์ (บูรณาการ)', 40],
  ['ป.6', '68', 'คณิตศาสตร์ (บูรณาการ)', 80],
  ['ป.6', '68', 'วิทยาศาสตร์และสิ่งแวดล้อม', 80],
  ['ป.6', '68', 'สังคมและความเป็นพลเมือง', 80],
  ['ป.6', '68', 'ประวัติศาสตร์', 40],
  ['ป.6', '68', 'เศรษฐกิจและการเงิน', 80],
  ['ป.6', '68', 'สุขภาพกายและสุขภาวะจิต', 80],
  ['ป.6', '68', 'ศิลปะและวัฒนธรรมเพื่อสุนทรียภาพ', 80],
  ['ป.6', '68', 'การใช้เทคโนโลยีอย่างฉลาดรู้', 40],
  ['ป.6', '51', 'ภาษาไทย (หลักภาษาและการใช้ภาษาไทย)', 80],
  ['ป.6', '51', 'ภาษาไทย (วรรณคดีและวรรณกรรม)', 80],
  ['ป.6', '51', 'ภาษาอังกฤษ Smile', 80],
  ['ป.6', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ป.6', '51', 'คอมพิวเตอร์', 40],
  ['ป.6', '51', 'วิทยาการคำนวณ', 40],
  ['ป.6', '51', 'การงานอาชีพ', 40],
  // ═══ ระดับมัธยม ม.1–ม.6 (หลักสูตร 2551) ═══
  ['ม.1', '51', 'ภาษาไทย', 120],
  ['ม.1', '51', 'คณิตศาสตร์', 120],
  ['ม.1', '51', 'วิทยาศาสตร์และเทคโนโลยี', 120],
  ['ม.1', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 120],
  ['ม.1', '51', 'ประวัติศาสตร์', 40],
  ['ม.1', '51', 'สุขศึกษาและพลศึกษา', 80],
  ['ม.1', '51', 'ศิลปะ', 80],
  ['ม.1', '51', 'การงานอาชีพ', 40],
  ['ม.1', '51', 'ภาษาอังกฤษ', 120],
  ['ม.1', '51', 'วิทยาการคำนวณ', 40],
  ['ม.1', '51', 'การออกแบบและเทคโนโลยี', 40],
  ['ม.2', '51', 'ภาษาไทย', 120],
  ['ม.2', '51', 'คณิตศาสตร์', 120],
  ['ม.2', '51', 'วิทยาศาสตร์และเทคโนโลยี', 120],
  ['ม.2', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 120],
  ['ม.2', '51', 'ประวัติศาสตร์', 40],
  ['ม.2', '51', 'สุขศึกษาและพลศึกษา', 80],
  ['ม.2', '51', 'ศิลปะ', 80],
  ['ม.2', '51', 'การงานอาชีพ', 40],
  ['ม.2', '51', 'ภาษาอังกฤษ', 120],
  ['ม.2', '51', 'วิทยาการคำนวณ', 40],
  ['ม.2', '51', 'การออกแบบและเทคโนโลยี', 40],
  ['ม.3', '51', 'ภาษาไทย', 120],
  ['ม.3', '51', 'คณิตศาสตร์', 120],
  ['ม.3', '51', 'วิทยาศาสตร์และเทคโนโลยี', 120],
  ['ม.3', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 120],
  ['ม.3', '51', 'ประวัติศาสตร์', 40],
  ['ม.3', '51', 'สุขศึกษาและพลศึกษา', 80],
  ['ม.3', '51', 'ศิลปะ', 80],
  ['ม.3', '51', 'การงานอาชีพ', 40],
  ['ม.3', '51', 'ภาษาอังกฤษ', 120],
  ['ม.3', '51', 'วิทยาการคำนวณ', 40],
  ['ม.3', '51', 'การออกแบบและเทคโนโลยี', 40],
  ['ม.4', '51', 'ภาษาไทย', 80],
  ['ม.4', '51', 'คณิตศาสตร์พื้นฐาน', 80],
  ['ม.4', '51', 'คณิตศาสตร์เพิ่มเติม', 80],
  ['ม.4', '51', 'ภาษาอังกฤษ', 80],
  ['ม.4', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ม.4', '51', 'ประวัติศาสตร์', 40],
  ['ม.4', '51', 'ฟิสิกส์', 80],
  ['ม.4', '51', 'เคมี', 80],
  ['ม.4', '51', 'ชีววิทยา', 80],
  ['ม.4', '51', 'สุขศึกษาและพลศึกษา', 40],
  ['ม.4', '51', 'ศิลปะ', 40],
  ['ม.4', '51', 'การงานอาชีพ', 40],
  ['ม.4', '51', 'วิทยาการคำนวณ', 40],
  ['ม.5', '51', 'ภาษาไทย', 80],
  ['ม.5', '51', 'คณิตศาสตร์พื้นฐาน', 80],
  ['ม.5', '51', 'คณิตศาสตร์เพิ่มเติม', 80],
  ['ม.5', '51', 'ภาษาอังกฤษ', 80],
  ['ม.5', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ม.5', '51', 'ประวัติศาสตร์', 40],
  ['ม.5', '51', 'ฟิสิกส์', 80],
  ['ม.5', '51', 'เคมี', 80],
  ['ม.5', '51', 'ชีววิทยา', 80],
  ['ม.5', '51', 'สุขศึกษาและพลศึกษา', 40],
  ['ม.5', '51', 'ศิลปะ', 40],
  ['ม.5', '51', 'การงานอาชีพ', 40],
  ['ม.5', '51', 'วิทยาการคำนวณ', 40],
  ['ม.6', '51', 'ภาษาไทย', 80],
  ['ม.6', '51', 'คณิตศาสตร์พื้นฐาน', 80],
  ['ม.6', '51', 'คณิตศาสตร์เพิ่มเติม', 80],
  ['ม.6', '51', 'ภาษาอังกฤษ', 80],
  ['ม.6', '51', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80],
  ['ม.6', '51', 'ประวัติศาสตร์', 40],
  ['ม.6', '51', 'ฟิสิกส์', 80],
  ['ม.6', '51', 'เคมี', 80],
  ['ม.6', '51', 'ชีววิทยา', 80],
  ['ม.6', '51', 'สุขศึกษาและพลศึกษา', 40],
  ['ม.6', '51', 'ศิลปะ', 40],
  ['ม.6', '51', 'การงานอาชีพ', 40],
  ['ม.6', '51', 'วิทยาการคำนวณ', 40]
];


// ============================================================
// 🌐 Web App Entry
// ============================================================
function doGet(e) {
  ensureSheets_();
  // 📖 หน้าตัวอย่างสื่อสาธารณะ (Preview) — เสิร์ฟจาก Apps Script
  if (e && e.parameter && e.parameter.id)      return servePreviewMedia_(e.parameter.id);
  if (e && e.parameter && e.parameter.preview) return servePreviewPage_(e);
  // หน้าร้านย้ายไป GitHub Pages แล้ว — Apps Script ทำหน้าที่เป็น API อย่างเดียว
  // ถ้าเรียกแบบ ?fn=... ผ่าน GET ก็ให้ทำงานเป็น API (เผื่อทดสอบ), ไม่งั้นคืนข้อความสถานะ
  if (e && e.parameter && e.parameter.fn) return doPost(e);
  return ContentService
    .createTextOutput('✅ ระบบสั่งซื้อแผนการสอน ' + PAGE_NAME + ' — API พร้อมทำงาน (หน้าร้านอยู่ที่ ' + SHOP_URL + ')')
    .setMimeType(ContentService.MimeType.TEXT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 🌍 JSON API สำหรับหน้าเว็บบน GitHub Pages (เรียกผ่าน fetch)
//    รูปแบบคำขอ: POST { "fn": "ชื่อฟังก์ชัน", "args": [...] }
//    ความปลอดภัยเท่าเวอร์ชันฝัง: ฟังก์ชันแอดมินตรวจ PIN ฝั่งเซิร์ฟเวอร์ทุกครั้ง
// ============================================================
function doPost(e) {
  try {
    ensureSheets_();
    // รับคำสั่งได้ทั้งแบบ POST (body JSON) และ GET (?fn=...&args=[...])
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents || '{}');
    } else if (e && e.parameter && e.parameter.fn) {
      body = { fn: e.parameter.fn, args: e.parameter.args ? JSON.parse(e.parameter.args) : [] };
    }
    // รายชื่อฟังก์ชันที่เปิดให้เรียกได้ (whitelist เท่านั้น)
    const api = {
      // ฝั่งลูกค้า
      getBootstrap: getBootstrap,
      getPublicSettings: getPublicSettings,
      getShopProducts: getShopProducts,
      createOrder: createOrder,
      uploadSlip: uploadSlip,
      getOrderStatus: getOrderStatus,
      getPreviewData: getPreviewData,
      getPreviewMediaContent: getPreviewMediaContent,
      // ฝั่งแอดมิน (ทุกตัวตรวจ PIN เอง)
      checkAdminPin: checkAdminPin,
      getEdgeStatus: getEdgeStatus,
      publishShopSnapshotApi: publishShopSnapshotApi,
      exportOrdersMonth: exportOrdersMonth,
      listOrders: listOrders,
      updateOrderStatus: updateOrderStatus,
      updateOrderEmail: updateOrderEmail,
      approveOrder: approveOrder,
      getProducts: getProducts,
      addProduct: addProduct,
      updateProduct: updateProduct,
      deleteProduct: deleteProduct,
      bulkImportProducts: bulkImportProducts,
      autoSetPrices: autoSetPrices,
      getReportData: getReportData,
      listCustomers: listCustomers,
      getCustomerHistory: getCustomerHistory,
      getSettings: getSettings,
      saveSettings: saveSettings,
      uploadImageToDrive: uploadImageToDrive,
      testTelegram: testTelegram,
      getTelegramChatId: getTelegramChatId,
      backupAllData: backupAllData,
      importBackupData: importBackupData,
      clearAllData: clearAllData,
      clearOrders: clearOrders
    };
    const fn = api[String(body.fn || '')];
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง: ' + body.fn);
    return jsonOut_(fn.apply(null, body.args || []));
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// 🏗️ สร้างชีตอัตโนมัติเมื่อรันครั้งแรก
// ============================================================
let ENSURE_DONE_ = false;   // ตรวจโครงชีตแค่ครั้งแรกของแต่ละคำขอ (execution) พอ
function ensureSheets_() {
  if (ENSURE_DONE_) return;
  ENSURE_DONE_ = true;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Settings ---
  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.getRange(1, 1, 1, 2).setValues([['Key', 'Value']])
      .setFontWeight('bold').setBackground('#ede9fe');
    sh.getRange(2, 1, DEFAULT_SETTINGS.length, 2).setValues(DEFAULT_SETTINGS);
    sh.setColumnWidth(1, 160).setColumnWidth(2, 320);
    sh.setFrozenRows(1);
  } else {
    // อัปเกรด: เติมคีย์ใหม่ที่ยังไม่มี (เช่น admins, telegram) โดยไม่แตะค่าที่ตั้งไว้แล้ว
    const have = {};
    if (sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
        have[String(r[0]).trim()] = true;
      });
    }
    DEFAULT_SETTINGS.forEach(function (kv) {
      if (!have[kv[0]]) sh.appendRow(kv);
    });
  }

  // --- ชีตสินค้าแยกรายชั้น: ป.1–ป.6, ม.1–ม.6, VIP & ระบบงาน, อื่นๆ ---
  const PRODUCT_HEADER = PRODUCT_HEADER_STD;
  const allProductSheets = GRADE_SHEETS.concat([SHEET_VIP_WORK, SHEET_OTHER]);
  let createdProductSheet = false;
  allProductSheets.forEach(function (name) {
    if (ss.getSheetByName(name)) return;
    const s = ss.insertSheet(name);
    if (isVipSheet_(name)) {
      s.getRange(1, 1, 1, PV_COUNT).setValues([PRODUCT_HEADER_VIP]).setFontWeight('bold').setBackground('#fef3c7');
    } else {
      s.getRange(1, 1, 1, P_COUNT).setValues([PRODUCT_HEADER]).setFontWeight('bold')
        .setBackground(name.indexOf('ม.') === 0 ? '#e0f2fe' : '#dbeafe');
    }
    s.setFrozenRows(1);
    createdProductSheet = true;
  });
  // อัปเกรดชีตรายชั้นเดิม: เติมหัวคอลัมน์ 14 "กลุ่มวิชา" (non-destructive) — ชีต VIP ใช้โครงย่อ ไม่ต้อง
  allProductSheets.forEach(function (name) {
    if (isVipSheet_(name)) return;
    const s = ss.getSheetByName(name);
    if (s && String(s.getRange(1, P.GROUP).getValue()).trim() !== 'กลุ่มวิชา') {
      s.getRange(1, P.GROUP).setValue('กลุ่มวิชา').setFontWeight('bold')
        .setBackground(name.indexOf('ม.') === 0 ? '#e0f2fe' : '#dbeafe');
    }
  });

  // seed แคตตาล็อกตั้งต้น เฉพาะติดตั้งใหม่จริงๆ (ไม่มีชีตรวมรุ่นเก่า และชีตรายชั้นยังว่างทั้งหมด)
  if (createdProductSheet) {
    const legacyExists = [SHEET_PRODUCTS, SHEET_PRODUCTS_M, 'แผนการสอน', 'Products'].some(function (n) { return !!ss.getSheetByName(n); });
    const anyData = allProductSheets.some(function (n) { const s = ss.getSheetByName(n); return s && s.getLastRow() > 1; });
    if (!legacyExists && !anyData && SEED_PRODUCTS.length) {
      const now = ymd_(new Date());
      const bySheet = {};
      SEED_PRODUCTS.forEach(function (r, i) {
        const name = sheetNameForProduct_(r[0], 'แผนการสอน');
        (bySheet[name] = bySheet[name] || []).push(
          ['PD' + padNum_(i + 1, 4), r[0], r[1], r[2], r[3], '', '', '', '', STD_FOOTER, 'TRUE', now, 'แผนการสอน', autoGroupFor_(r[1], r[2])]);
      });
      Object.keys(bySheet).forEach(function (name) {
        const s = ss.getSheetByName(name);
        const rows = bySheet[name];
        s.getRange(2, 1, rows.length, P_COUNT).setValues(rows);
        s.getRange(2, P.ADDED, rows.length, 1).setNumberFormat('@'); // กัน Sheets แปลงวันที่
        sortProductSheet_(s);   // เรียงชื่อวิชา ก-ฮ ตั้งแต่แรก
      });
    }
  }

  // --- Orders ---
  // --- SlipHash (ชีตซ่อน เก็บลายนิ้วมือสลิปกันใช้ซ้ำ) ---
  let shh = ss.getSheetByName(SHEET_SLIPHASH);
  if (!shh) {
    shh = ss.insertSheet(SHEET_SLIPHASH);
    shh.getRange(1, 1, 1, 3).setValues([['ลายนิ้วมือไฟล์ (MD5)', 'เลขที่ออเดอร์', 'เวลา']])
      .setFontWeight('bold').setBackground('#fee2e2');
    try { shh.hideSheet(); } catch (eH) {}
  }

  sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ORDERS);
    sh.getRange(1, 1, 1, O_COUNT).setValues([[
      'วันที่สั่งซื้อ', 'เลขที่ออเดอร์', 'ชื่อลูกค้า/Facebook', 'วิชาที่สั่ง', 'อีเมล (Gmail)', 'เบอร์โทร',
      'จำนวนรายการ', 'รายการ (JSON)', 'ยอดก่อนลด', 'ส่วนลด %',
      'ส่วนลด (บาท)', 'ยอดสุทธิ', 'สถานะ', 'ลิงก์สลิป', 'หมายเหตุ'
    ]]).setFontWeight('bold').setBackground('#dcfce7');
    sh.getRange('A:A').setNumberFormat('@'); // วันที่เก็บเป็นข้อความ
    sh.setColumnWidth(O.ITEMS_TEXT, 240);
    sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, O.ITEMS_TEXT).getValue()) !== 'วิชาที่สั่ง') {
    // อัปเกรดชีตเดิมเป็นโครง 15 คอลัมน์ — เช็คก่อนว่า "ข้อมูล" อยู่โครงไหน จะได้ไม่แทรกคอลัมน์ซ้ำ
    const lastO = sh.getLastRow();
    let dataIs15 = false;
    if (lastO > 1) {
      const stVals = sh.getRange(2, O.STATUS, lastO - 1, 1).getValues();
      dataIs15 = stVals.some(function (r) { return isStatus_(r[0]); });
    }
    if (!dataIs15) {
      // ข้อมูลยังเป็นโครงเก่า 14 คอลัมน์ → แทรกคอลัมน์ "วิชาที่สั่ง" (หัว+ข้อมูลขยับพร้อมกัน)
      sh.insertColumnAfter(O.CUSTOMER);
      if (lastO > 1) {
        const jsons = sh.getRange(2, O.ITEMS_JSON, lastO - 1, 1).getValues();
        sh.getRange(2, O.ITEMS_TEXT, lastO - 1, 1).setNumberFormat('@')
          .setValues(jsons.map(function (r) { return [itemsText_(parseItems_(r[0]))]; }));
      }
    } else if (lastO > 1) {
      // ข้อมูลเป็นโครงใหม่อยู่แล้ว (แค่หัวตารางเก่า) → เติมวิชาที่สั่งเฉพาะช่องที่ว่าง
      const dvals = sh.getRange(2, O.ITEMS_TEXT, lastO - 1, 1).getValues();
      const jsons = sh.getRange(2, O.ITEMS_JSON, lastO - 1, 1).getValues();
      sh.getRange(2, O.ITEMS_TEXT, lastO - 1, 1).setNumberFormat('@')
        .setValues(dvals.map(function (r, i) {
          return [String(r[0] || '') || itemsText_(parseItems_(jsons[i][0]))];
        }));
    }
    // เขียนหัวตาราง 15 ช่องให้ตรงโครงปัจจุบันเสมอ
    sh.getRange(1, 1, 1, O_COUNT).setValues([[
      'วันที่สั่งซื้อ', 'เลขที่ออเดอร์', 'ชื่อลูกค้า/Facebook', 'วิชาที่สั่ง', 'อีเมล (Gmail)', 'เบอร์โทร',
      'จำนวนรายการ', 'รายการ (JSON)', 'ยอดก่อนลด', 'ส่วนลด %',
      'ส่วนลด (บาท)', 'ยอดสุทธิ', 'สถานะ', 'ลิงก์สลิป', 'หมายเหตุ'
    ]]).setFontWeight('bold').setBackground('#dcfce7');
    sh.setColumnWidth(O.ITEMS_TEXT, 240);
  }
}


// ============================================================
// 🧰 Helpers
// ============================================================
function ymd_(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}
function padNum_(n, len) {
  return ('000000' + n).slice(-len);
}
function escapeHtml_(t) {
  if (t === null || t === undefined) return '';
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function getSheet_(name) {
  ensureSheets_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function settingsMap_() {
  const sh = getSheet_(SHEET_SETTINGS);
  const last = sh.getLastRow();
  const map = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      const k = String(r[0] || '').trim();
      if (k) map[k] = String(r[1] === null || r[1] === undefined ? '' : r[1]);
    });
  }
  return map;
}

// ============================================================
// 🔐 บัญชีแอดมิน (yui = owner / pinky = staff)
// ============================================================
// รูปแบบในชีต Settings คีย์ admins: "yui=080766:owner, pinky=1234:staff"
function admins_() {
  const m = settingsMap_();
  const raw = m.admins || '';
  const list = [];
  String(raw).split(',').forEach(function (part) {
    const mt = String(part).trim().match(/^([^=]+)=([^:]+):(owner|staff)$/);
    if (mt) list.push({ name: mt[1].trim(), pin: mt[2].trim(), role: mt[3].trim() });
  });
  // รองรับระบบเก่าที่ใช้ adminPIN เดี่ยว
  if (!list.length && m.adminPIN) list.push({ name: 'admin', pin: String(m.adminPIN).trim(), role: 'owner' });
  return list;
}
function verifyAdmin_(pin) {
  const p = String(pin || '').trim();
  const found = admins_().filter(function (a) { return a.pin === p; })[0];
  if (!found) throw new Error('PIN ไม่ถูกต้อง');
  return found;
}
function requireOwner_(pin) {
  const a = verifyAdmin_(pin);
  if (a.role !== 'owner') throw new Error('เมนูนี้สำหรับเจ้าของระบบเท่านั้น (บัญชีพนักงานใช้ได้เฉพาะเมนูจัดการคำสั่งซื้อ)');
  return a;
}
function checkAdminPin(pin) {
  try {
    const a = verifyAdmin_(pin);
    return { success: true, name: a.name, role: a.role };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// ⚙️ Settings (ข้อมูลเพจ / PromptPay / ส่วนลด)
// ============================================================

// ฝั่งลูกค้า — คืนเฉพาะค่าที่ปลอดภัย (ไม่ส่ง adminPIN ออกไป)
function getPublicSettings() {
  const m = settingsMap_();
  return {
    success: true,
    settings: {
      pageName: m.pageName || PAGE_NAME,
      slogan: m.slogan || '',
      promptpayName: m.promptpayName || '',
      promptpayNumber: m.promptpayNumber || '',
      qrURL: m.qrURL || '',
      logoURL: m.logoURL || '',
      discountMinItems: Number(m.discountMinItems || 3),
      discountPercent: Number(m.discountPercent || 5),
      announcement: m.announcement || '',
      shopTextDiscount: m.shopTextDiscount || '',
      shopTextVip: m.shopTextVip || '',
      shopTextGmail1: m.shopTextGmail1 || '',
      shopTextGmail2: m.shopTextGmail2 || '',
      shopTextFooter: m.shopTextFooter || ''
    }
  };
}

// ฝั่งแอดมิน — คืนทุกค่า
function getSettings(pin) {
  try {
    requireOwner_(pin);
    return { success: true, settings: settingsMap_() };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveSettings(pin, data) {
  try {
    requireOwner_(pin); clearShopCache_();
    const sh = getSheet_(SHEET_SETTINGS);
    const last = sh.getLastRow();
    const keys = last >= 2
      ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); })
      : [];
    Object.keys(data || {}).forEach(function (k) {
      const v = String(data[k] === null || data[k] === undefined ? '' : data[k]);
      const idx = keys.indexOf(k);
      if (idx >= 0) sh.getRange(idx + 2, 2).setValue(v);
      else sh.appendRow([k, v]);
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 📚 Products (รายการแผนการสอน)
// ============================================================
function readProducts_() {
  // อ่านรวมจากทุกชีตสินค้า (ป. + ม.) — ทุกรายการจำชื่อชีตกับแถวของตัวเองไว้ให้ตัวแก้/ลบใช้
  const out = [];
  productSheetNames_().forEach(function (name) {
    const sh = getSheet_(name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last < 2) return;
    if (isVipSheet_(name)) {   // ชีต VIP & ระบบงาน ใช้โครงย่อ 10 คอลัมน์
      sh.getRange(2, 1, last - 1, PV_COUNT).getValues().forEach(function (r, i) {
        out.push({
          sheetName: name, row: i + 2,
          id: String(r[PV.ID - 1] || '').trim(),
          grade: '', curriculum: '', hours: '', group: '',
          subject: String(r[PV.SUBJECT - 1] || '').trim(),
          price: Number(r[PV.PRICE - 1] || 0),
          driveLink: String(r[PV.DRIVE_LINK - 1] || '').trim(),
          exampleLink: String(r[PV.EXAMPLE_LINK - 1] || '').trim(),
          htmlMedia: '', postNote: String(r[PV.POST_NOTE - 1] || ''),
          footer: String(r[PV.FOOTER - 1] || ''),
          active: String(r[PV.ACTIVE - 1] || '').toUpperCase() === 'TRUE',
          added: String(r[PV.ADDED - 1] || ''),
          category: String(r[PV.CATEGORY - 1] || '').trim() || 'ระบบงาน'
        });
      });
      return;
    }
    sh.getRange(2, 1, last - 1, P_COUNT).getValues().forEach(function (r, i) {
      out.push(makeProduct_(r, i + 2, name));
    });
  });
  return out.filter(function (p) { return p.id; });
}
function makeProduct_(r, row, sheetName) {
  return {
      sheetName: sheetName,
      row: row,
      id: String(r[P.ID - 1] || ''),
      grade: String(r[P.GRADE - 1] || ''),
      curriculum: String(r[P.CURRICULUM - 1] || ''),
      subject: String(r[P.SUBJECT - 1] || ''),
      hours: Number(r[P.HOURS - 1] || 0),
      price: Number(r[P.PRICE - 1] || 0),
      driveLink: String(r[P.DRIVE_LINK - 1] || ''),
      exampleLink: String(r[P.EXAMPLE_LINK - 1] || ''),
      htmlMedia: '',
      planDetail: String(r[P.POST_NOTE - 1] || ''),
      footer: String(r[P.FOOTER - 1] || ''),
      // postNote = รายละเอียดแผน + ข้อความท้าย (รวมตอนส่งให้ลูกค้า เหมือนเดิม)
      postNote: (function () {
        const d = String(r[P.POST_NOTE - 1] || '').trim();
        const f = String(r[P.FOOTER - 1] || '').trim();
        return d && f ? (d + '\n\n' + f) : (d || f);
      })(),
      active: String(r[P.ACTIVE - 1]).toUpperCase() !== 'FALSE',
      added: String(r[P.ADDED - 1] || ''),
      category: String(r[P.CATEGORY - 1] || '').trim() || 'แผนการสอน',
      group: String(r[P.GROUP - 1] || '').trim()
  };
}

// ฝั่งลูกค้า — เฉพาะวิชาเปิดขาย + ราคา > 0 (ไม่ส่งลิงก์ Drive ออกไปเด็ดขาด)
function getShopProducts() {
  try {
    const items = readProducts_()
      .filter(function (p) { return p.active && p.price > 0; })
      .map(function (p) {
        return {
          id: p.id, grade: p.grade, curriculum: p.curriculum,
          subject: p.subject, hours: p.hours, price: p.price,
          exampleLink: p.exampleLink, category: p.category, group: p.group
        };
      });
    return { success: true, products: items, groupOrder: { '68': GROUPS_68, '51': GROUPS_51 } };
  } catch (e) { return { success: false, error: e.message }; }
}

// ฝั่งแอดมิน — เห็นทุกคอลัมน์
function getProducts(pin) {
  try {
    requireOwner_(pin);
    return { success: true, products: readProducts_() };
  } catch (e) { return { success: false, error: e.message }; }
}

// ⚡ รวม settings + products ในการเรียกครั้งเดียว — ลด round-trip ตอนเปิดหน้าร้านให้เหลือครั้งเดียว
function getBootstrap() {
  try {
    // แคชฝั่งเซิร์ฟเวอร์ 30 วินาที — ลูกค้าคนถัดๆ ไปโหลดไวขึ้นมาก (ล้างอัตโนมัติเมื่อแก้สินค้า/ตั้งค่า)
    const cache = CacheService.getScriptCache();
    const hit = cache.get(SHOP_CACHE_KEY);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
    const s = getPublicSettings();
    const p = getShopProducts();
    const out = {
      success: true,
      settings: (s && s.settings) || {},
      products: (p && p.products) || [],
      groupOrder: (p && p.groupOrder) || {},
      error: (s && s.error) || (p && p.error) || ''
    };
    try { cache.put(SHOP_CACHE_KEY, JSON.stringify(out), 30); } catch (e) {}
    return out;
  } catch (e) { return { success: false, error: e.message }; }
}

const SHOP_CACHE_KEY = 'shop_bootstrap_v1';
// ล้างแคชหน้าร้าน — เรียกทุกครั้งที่มีการแก้ไขสินค้า/ราคา/ตั้งค่า เพื่อให้ลูกค้าเห็นข้อมูลล่าสุดทันที
function clearShopCache_() {
  try { CacheService.getScriptCache().remove(SHOP_CACHE_KEY); } catch (e) {}
  // ⚡ Edge: ข้อมูลร้านเปลี่ยน → ติดธงไว้ ให้ตัวเผยแพร่อัตโนมัติ (ทุก 5 นาที) อัปขึ้น GitHub
  try { PropertiesService.getScriptProperties().setProperty('EDGE_DIRTY', '1'); } catch (e) {}
}

/* ============================================================
 *  ⚡ SHEETS + EDGE — เผยแพร่ snapshot หน้าร้าน (products.json)
 *  ขึ้น GitHub Pages CDN เพื่อให้หน้าเว็บลูกค้าเปิดเร็วทันที
 *  Sheets ยังเป็นฐานข้อมูลจริง / GAS ยังรับออเดอร์-สลิปเหมือนเดิม
 *  ตั้งค่าครั้งแรก: เมนู 🛠 เครื่องมือระบบ → 🚀 ตั้งค่า Edge
 * ============================================================ */
function edgeConfig_() {
  const pr = PropertiesService.getScriptProperties();
  return {
    token: pr.getProperty('GITHUB_TOKEN') || '',
    repo: pr.getProperty('GITHUB_REPO') || 'kruprompt27/kruprompt27.github.io',
    path: pr.getProperty('GITHUB_PATH') || 'shop/products.json',
    branch: pr.getProperty('GITHUB_BRANCH') || 'main'
  };
}

// ตั้งค่าครั้งแรก: ขอ GitHub Token (สร้างที่ github.com → Settings → Developer settings →
// Personal access tokens → Fine-grained → เลือกรีโปหน้าร้าน → สิทธิ์ Contents: Read and write)
function setupEdgePublish() {
  const ui = SpreadsheetApp.getUi();
  const pr = PropertiesService.getScriptProperties();
  const cfg = edgeConfig_();
  const t = ui.prompt('🚀 ตั้งค่า Edge (1/2) — GitHub Token',
    'วาง Personal Access Token ของ GitHub\n(สิทธิ์ Contents: Read/Write เฉพาะรีโปหน้าร้าน)\n\n' +
    (cfg.token ? 'มี token เดิมอยู่แล้ว — เว้นว่างเพื่อใช้ตัวเดิม' : ''), ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;
  const token = String(t.getResponseText() || '').trim();
  if (token) pr.setProperty('GITHUB_TOKEN', token);
  else if (!cfg.token) { ui.alert('ยังไม่ได้วาง token'); return; }

  const r = ui.prompt('🚀 ตั้งค่า Edge (2/2) — ตำแหน่งไฟล์',
    'รูปแบบ: owner/repo|path|branch\nกด OK เพื่อใช้ค่าปัจจุบัน:\n' +
    cfg.repo + '|' + cfg.path + '|' + cfg.branch, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const v = String(r.getResponseText() || '').trim();
  if (v) {
    const parts = v.split('|');
    if (parts[0]) pr.setProperty('GITHUB_REPO', parts[0].trim());
    if (parts[1]) pr.setProperty('GITHUB_PATH', parts[1].trim());
    if (parts[2]) pr.setProperty('GITHUB_BRANCH', parts[2].trim());
  }
  try {
    const n = publishShopSnapshot_();
    ui.alert('✅ ตั้งค่าสำเร็จ และเผยแพร่ทดสอบแล้ว (' + n + ' วิชา)\n\nหน้าร้านจะโหลดข้อมูลจากไฟล์นี้ทันที เปิดเร็วขึ้นชัดเจนค่ะ\nต่อไประบบเผยแพร่ให้เองทุกครั้งที่ข้อมูลเปลี่ยน (ภายใน ~5 นาที)');
  } catch (e) {
    ui.alert('⚠️ บันทึกค่าแล้ว แต่เผยแพร่ทดสอบไม่สำเร็จ:\n' + e.message +
      '\n\nเช็ค: token มีสิทธิ์ Contents ของรีโปนี้ไหม / ชื่อรีโป-branch ถูกไหม');
  }
}

function publishShopSnapshotMenu() {
  const ui = SpreadsheetApp.getUi();
  try { const n = publishShopSnapshot_(); ui.alert('🚀 เผยแพร่หน้าร้านแล้ว (' + n + ' วิชา)\nลูกค้าเห็นข้อมูลใหม่ภายใน ~1 นาที (CDN)'); }
  catch (e) { ui.alert('❌ เผยแพร่ไม่สำเร็จ: ' + e.message); }
}

// ตัวเผยแพร่อัตโนมัติ (trigger ทุก 5 นาที) — ทำงานเฉพาะเมื่อมีธงข้อมูลเปลี่ยน และตั้ง token แล้ว
function edgeAutoPublish() {
  const pr = PropertiesService.getScriptProperties();
  if (pr.getProperty('EDGE_DIRTY') !== '1') return;
  if (!edgeConfig_().token) return;   // ยังไม่ตั้งค่า Edge — ข้ามเงียบๆ ระบบเดิมทำงานปกติ
  try {
    publishShopSnapshot_();
    pr.deleteProperty('EDGE_DIRTY');
    pr.deleteProperty('EDGE_FAIL_NOTIFIED');   // กลับมาปกติ → พร้อมเตือนใหม่ถ้าพังอีกในอนาคต
  } catch (e) {
    Logger.log('edgeAutoPublish: ' + e);
    // ระบบขายจริงห้ามพังเงียบ — แจ้ง Telegram ทันที (ไม่เกินวันละครั้ง) หน้าร้านยังเปิดได้ผ่าน GAS ตามปกติ
    const today = ymd_(new Date()).slice(0, 10);
    if (pr.getProperty('EDGE_FAIL_NOTIFIED') !== today) {
      pr.setProperty('EDGE_FAIL_NOTIFIED', today);
      try {
        sendTelegram_('🚨 เผยแพร่หน้าร้านขึ้น Edge ไม่สำเร็จ\nสาเหตุ: ' + String(e.message || e).slice(0, 140) +
          '\n\n👉 หน้าร้านยังขายได้ตามปกติ (ระบบถอยไปใช้ GAS อัตโนมัติ) แต่จะโหลดช้าลง' +
          '\n👉 เช็ค token ที่เมนู 🛠 → 🚀 ตั้งค่า Edge / GitHub');
      } catch (e2) {}
    }
  }
}

function publishShopSnapshot_() {
  const cfg = edgeConfig_();
  if (!cfg.token) throw new Error('ยังไม่ได้ตั้งค่า GitHub Token (เมนู 🚀 ตั้งค่า Edge)');
  const prods = getShopProducts();
  if (!prods.success) throw new Error(prods.error);
  const sets = getPublicSettings();
  const snap = {
    updated: new Date().toISOString(),
    settings: sets.success ? sets.settings : {},
    products: prods.products
  };
  const url = 'https://api.github.com/repos/' + cfg.repo + '/contents/' +
    cfg.path.split('/').map(encodeURIComponent).join('/');
  const headers = { Authorization: 'Bearer ' + cfg.token, Accept: 'application/vnd.github+json' };
  // หา sha ของไฟล์เดิม (ต้องใช้ตอนแก้ทับ)
  let sha = null;
  const g = UrlFetchApp.fetch(url + '?ref=' + encodeURIComponent(cfg.branch), { headers: headers, muteHttpExceptions: true });
  if (g.getResponseCode() === 200) { try { sha = JSON.parse(g.getContentText()).sha; } catch (e) {} }
  const body = {
    message: '🚀 shop snapshot ' + ymd_(new Date()),
    branch: cfg.branch,
    content: Utilities.base64Encode(JSON.stringify(snap), Utilities.Charset.UTF_8)
  };
  if (sha) body.sha = sha;
  const res = UrlFetchApp.fetch(url, {
    method: 'put', contentType: 'application/json', headers: headers,
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    const errMsg = 'GitHub ตอบ ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 160);
    try { PropertiesService.getScriptProperties().setProperty('EDGE_LAST_ERROR', errMsg); } catch (e) {}
    throw new Error(errMsg);
  }
  // สำเร็จ → บันทึกเวลา/จำนวน ไว้โชว์บนหน้าแอดมิน และล้าง error เดิม
  try {
    const pr = PropertiesService.getScriptProperties();
    pr.setProperty('EDGE_LAST_PUBLISH', new Date().toISOString());
    pr.setProperty('EDGE_LAST_COUNT', String(prods.products.length));
    pr.deleteProperty('EDGE_LAST_ERROR');
  } catch (e) {}
  return prods.products.length;
}

// 🚀 สถานะ Edge สำหรับหน้าแอดมิน (แอดมินทุกสิทธิ์ดูได้)
function getEdgeStatus(pin) {
  try {
    verifyAdmin_(pin);
    const pr = PropertiesService.getScriptProperties();
    return {
      success: true,
      configured: !!edgeConfig_().token,
      lastPublish: pr.getProperty('EDGE_LAST_PUBLISH') || '',
      lastCount: Number(pr.getProperty('EDGE_LAST_COUNT') || 0),
      dirty: pr.getProperty('EDGE_DIRTY') === '1',
      lastError: pr.getProperty('EDGE_LAST_ERROR') || ''
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// 🚀 กดเผยแพร่จากหน้าแอดมิน (เจ้าของระบบ)
function publishShopSnapshotApi(pin) {
  try {
    requireOwner_(pin);
    const n = publishShopSnapshot_();
    try { PropertiesService.getScriptProperties().deleteProperty('EDGE_DIRTY'); } catch (e) {}
    return { success: true, published: n };
  } catch (e) { return { success: false, error: e.message }; }
}

// 📥 ดึงออเดอร์รายเดือนสำหรับ Export Excel (เจ้าของระบบ) — month รูปแบบ 'YYYY-MM'
function exportOrdersMonth(pin, month) {
  try {
    requireOwner_(pin);
    const mo = String(month || '').slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(mo)) throw new Error('รูปแบบเดือนไม่ถูกต้อง (YYYY-MM)');
    const rows = readOrders_()
      .filter(function (o) { return String(o.timestamp).slice(0, 7) === mo; })
      .map(function (o) {
        return {
          orderId: o.orderId, timestamp: o.timestamp, customer: o.customer, email: o.email,
          items: (o.items || []).map(function (it) { return it.subject + ' ' + (it.grade || ''); }).join(', '),
          itemCount: (o.items || []).length,
          discount: Number(o.discount || 0), netTotal: Number(o.netTotal || 0), status: o.status
        };
      });
    return { success: true, rows: rows, month: mo };
  } catch (e) { return { success: false, error: e.message }; }
}

// ============================================================
// 💰 ราคาอัตโนมัติตามเวลาเรียน (ชม.)
// ============================================================
// ตารางราคาแก้ได้ที่ชีต Settings คีย์ priceTable รูปแบบ "40=200, 60=220, ..."
const DEFAULT_PRICE_TABLE = '40=200, 60=220, 80=250, 120=280, 160=300, 200=350';

function priceTable_() {
  const raw = settingsMap_().priceTable || DEFAULT_PRICE_TABLE;
  const table = {};
  String(raw).split(',').forEach(function (pair) {
    const m = pair.split('=');
    const h = Number(String(m[0] || '').trim());
    const p = Number(String(m[1] || '').trim());
    if (h > 0 && p > 0) table[h] = p;
  });
  return table;
}

// คืนราคาตามชั่วโมง (ไม่อยู่ในตาราง = 0 → ต้องตั้งเอง)
function priceForHours_(hours) {
  return priceTable_()[Number(hours || 0)] || 0;
}

// ตั้งราคาอัตโนมัติทั้งชีต Products ตามตารางชั่วโมง
// overwrite=true ทับราคาที่ตั้งไว้แล้วทั้งหมด / false เติมเฉพาะวิชาที่ราคายังว่างหรือ 0
function autoSetPrices(pin, overwrite) {
  try {
    requireOwner_(pin); clearShopCache_();
    const table = priceTable_();
    let updated = 0, skipped = 0;
    productSheetNames_().forEach(function (name) {   // ตั้งราคาทั้งชีต ป. และชีต ม.
      if (isVipSheet_(name)) return;   // ชีต VIP & ระบบงาน ไม่มีคอลัมน์ ชม. — ราคากรอกเองในชีต
      const sh = getSheet_(name);
      if (!sh) return;
      const last = sh.getLastRow();
      if (last < 2) return;
      const range = sh.getRange(2, P.HOURS, last - 1, 2); // คอลัมน์ ชม. + ราคา
      const vals = range.getValues();
      vals.forEach(function (r) {
        const hours = Number(r[0] || 0);
        const cur = Number(r[1] || 0);
        const auto = table[hours] || 0;
        if (!auto) { skipped++; return; }            // ชั่วโมงไม่อยู่ในตาราง
        if (cur > 0 && !overwrite) { skipped++; return; } // มีราคาแล้วและไม่ให้ทับ
        if (cur !== auto) { r[1] = auto; updated++; } else skipped++;
      });
      range.setValues(vals);
    });
    return { success: true, updated: updated, skipped: skipped };
  } catch (e) { return { success: false, error: e.message }; }
}


function nextProductId_() {
  const ids = readProducts_().map(function (p) {
    const m = p.id.match(/^PD(\d+)$/);
    return m ? Number(m[1]) : 0;
  });
  const max = ids.length ? Math.max.apply(null, ids) : 0;
  return 'PD' + padNum_(max + 1, 4);
}

function addProduct(pin, data) {
  try {
    requireOwner_(pin); clearShopCache_();
    if (!data || !String(data.subject || '').trim()) throw new Error('กรุณากรอกชื่อวิชา');
    const sh = getSheet_(sheetNameForProduct_(data.grade, data.category));  // ลงชีตชั้นนั้นๆ / VIP & ระบบงาน / อื่นๆ
    const id = nextProductId_();
    const targetName = sheetNameForProduct_(data.grade, data.category);
    const row = sh.getLastRow() + 1;
    const width = sheetColCount_(targetName);
    sh.getRange(row, 1, 1, width).setValues([buildRowFor_(targetName, {
      id: id, grade: String(data.grade || ''), curriculum: String(data.curriculum || ''),
      subject: String(data.subject).trim(), hours: Number(data.hours || 0) || '',
      price: Number(data.price || 0) || priceForHours_(Number(data.hours || 0)) || '',
      driveLink: String(data.driveLink || ''), exampleLink: String(data.exampleLink || ''),
      postNote: String(data.planDetail || data.postNote || ''), footer: STD_FOOTER,
      active: true, added: ymd_(new Date()), category: String(data.category || 'แผนการสอน'),
      group: String(data.group || '').trim() || autoGroupFor_(data.curriculum, data.subject)
    })]);
    sh.getRange(row, colFor_(targetName, P.ADDED)).setNumberFormat('@');
    sortProductSheet_(sh);   // แทรกแล้วเรียงชื่อวิชา ก-ฮ ให้อัตโนมัติ
    return { success: true, id: id };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateProduct(pin, id, data) {
  try {
    requireOwner_(pin); clearShopCache_();
    const p = readProducts_().filter(function (x) { return x.id === id; })[0];
    if (!p) throw new Error('ไม่พบรายการ ' + id);
    const sh = getSheet_(p.sheetName || SHEET_PRODUCTS);
    // เขียนผ่านตัวแปลตำแหน่ง — ชีต VIP & ระบบงาน (โครงย่อ) จะข้ามคอลัมน์ที่ชีตนั้นไม่มีให้เอง
    const set = function (pcol, val) {
      if (val === undefined || val === null) return;
      const c = colFor_(p.sheetName, pcol);
      if (c) sh.getRange(p.row, c).setValue(val);
    };
    set(P.GRADE, data.grade); set(P.CURRICULUM, data.curriculum);
    set(P.SUBJECT, data.subject); set(P.HOURS, Number(data.hours || 0));
    set(P.PRICE, Number(data.price || 0)); set(P.DRIVE_LINK, data.driveLink);
    set(P.EXAMPLE_LINK, data.exampleLink);
    set(P.POST_NOTE, data.planDetail !== undefined ? data.planDetail : data.postNote);
    set(P.FOOTER, data.footer); set(P.CATEGORY, data.category);
    if (data.active !== undefined) set(P.ACTIVE, data.active ? 'TRUE' : 'FALSE');
    if (data.group !== undefined) set(P.GROUP, String(data.group || '').trim());
    // ประเภท/ชั้นใหม่ทำให้ต้องอยู่คนละชีต → ย้ายโดยประกอบแถวใหม่ตามโครงชีตปลายทาง (รองรับย้ายข้ามโครง)
    const target = sheetNameForProduct_(
      data.grade !== undefined && data.grade !== null ? data.grade : p.grade,
      data.category !== undefined && data.category !== null ? data.category : p.category);
    if (target !== p.sheetName) {
      const fresh = readProducts_().filter(function (x) { return x.id === id; })[0] || p;  // อ่านค่าหลังแก้
      const tsh = getSheet_(target);
      const trow = tsh.getLastRow() + 1;
      const width = sheetColCount_(target);
      tsh.getRange(trow, 1, 1, width).setValues([buildRowFor_(target, fresh)]);
      tsh.getRange(trow, colFor_(target, P.ADDED)).setNumberFormat('@');
      sh.deleteRow(p.row);
      sortProductSheet_(tsh);
    } else {
      sortProductSheet_(sh);   // ชื่อวิชาอาจถูกแก้ → จัดเรียงใหม่ให้คงลำดับ ก-ฮ
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function deleteProduct(pin, id) {
  try {
    requireOwner_(pin); clearShopCache_();
    const p = readProducts_().filter(function (x) { return x.id === id; })[0];
    if (!p) throw new Error('ไม่พบรายการ ' + id);
    getSheet_(p.sheetName || SHEET_PRODUCTS).deleteRow(p.row);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// นำเข้าจาก Excel/CSV — array ของ {grade, curriculum, subject, hours, price}
function bulkImportProducts(pin, items) {
  try {
    requireOwner_(pin); clearShopCache_();
    if (!items || !items.length) throw new Error('ไม่พบข้อมูลนำเข้า');
    let n = Number(nextProductId_().replace('PD', ''));
    const now = ymd_(new Date());
    const bySheet = {};   // แยกแถวนำเข้าลงชีตตามระดับชั้นอัตโนมัติ (ม.* → ชีต ม.)
    items.forEach(function (it) {
      if (!String(it.subject || '').trim()) return;
      const hours = Number(it.hours || 0);
      const price = Number(it.price || 0) || priceForHours_(hours); // ไม่กรอกราคา → คิดตาม ชม.
      const row = ['PD' + padNum_(n++, 4), String(it.grade || ''), String(it.curriculum || ''),
        String(it.subject).trim(), hours, price,
        String(it.driveLink || ''), String(it.exampleLink || ''),
        String(it.planDetail || it.postNote || ''), STD_FOOTER, 'TRUE', now,
        String(it.category || 'แผนการสอน'),
        String(it.group || '').trim() || autoGroupFor_(it.curriculum, it.subject)];
      const name = sheetNameForProduct_(it.grade, it.category || 'แผนการสอน');
      (bySheet[name] = bySheet[name] || []).push(row);
    });
    let imported = 0;
    Object.keys(bySheet).forEach(function (name) {
      // แปลงแถวมาตรฐาน 14 ช่อง → โครงของชีตปลายทาง (ชีต VIP ใช้โครงย่อ)
      const rows = bySheet[name].map(function (r) {
        if (!isVipSheet_(name)) return r;
        return [r[P.ID - 1], r[P.SUBJECT - 1], r[P.PRICE - 1], r[P.DRIVE_LINK - 1], r[P.EXAMPLE_LINK - 1],
          r[P.POST_NOTE - 1], r[P.FOOTER - 1], r[P.ACTIVE - 1], r[P.ADDED - 1], r[P.CATEGORY - 1]];
      });
      const width = sheetColCount_(name);
      const sh = getSheet_(name);
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, width).setValues(rows);
      sh.getRange(startRow, colFor_(name, P.ADDED), rows.length, 1).setNumberFormat('@');
      sortProductSheet_(sh);   // นำเข้าแล้วเรียงชื่อวิชา ก-ฮ
      imported += rows.length;
    });
    if (!imported) throw new Error('ไม่มีแถวที่ใช้ได้ (ต้องมีชื่อวิชา)');
    return { success: true, imported: imported };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 🛒 Orders (ฝั่งลูกค้า)
// ============================================================

// สร้างคำสั่งซื้อ — คำนวณราคาฝั่งเซิร์ฟเวอร์เสมอ (กันแก้ราคาจากหน้าเว็บ)
function createOrder(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const customer = String(data && data.customer || '').trim();
    const email = String(data && data.email || '').trim().toLowerCase();
    const phone = String(data && data.phone || '').trim();
    const ids = (data && data.productIds) || [];
    if (!customer) throw new Error('กรุณากรอกชื่อ/Facebook');
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
    if (!ids.length) throw new Error('ตะกร้าสินค้าว่าง');

    const all = {};
    readProducts_().forEach(function (p) { all[p.id] = p; });
    const split = !!(data && data.splitEmail);
    const itemEmails = (data && data.itemEmails) || {};
    const items = [];
    ids.forEach(function (id) {
      const p = all[id];
      if (!p || !p.active || p.price <= 0) throw new Error('รายการ ' + id + ' ไม่พร้อมขาย');
      let itemEmail = email;
      if (split) {
        const e = String(itemEmails[id] || '').trim().toLowerCase();
        if (e) {
          if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Gmail ของวิชา ' + p.subject + ' ไม่ถูกต้อง');
          itemEmail = e;
        }
      }
      items.push({ id: p.id, grade: p.grade, curriculum: p.curriculum,
        subject: p.subject, hours: p.hours, price: p.price, category: p.category, email: itemEmail });
    });

    const m = settingsMap_();
    const minItems = Number(m.discountMinItems || 3);
    const pct = Number(m.discountPercent || 5);
    const subtotal = items.reduce(function (s, it) { return s + it.price; }, 0);
    // ส่วนลดคิดจากเฉพาะหมวด "แผนการสอน" (ไม่รวมระบบงาน/กลุ่ม VIP) ตามกติกาเดิมของร้าน
    const lessons = items.filter(function (it) { return (it.category || 'แผนการสอน') === 'แผนการสอน'; });
    const lessonSubtotal = lessons.reduce(function (s, it) { return s + it.price; }, 0);
    const usePct = lessons.length >= minItems ? pct : 0;
    const discount = Math.round(lessonSubtotal * usePct / 100);
    const net = subtotal - discount;

    // เลขออเดอร์: KPS + วันที่ + ลำดับ 3 หลักของวันนั้น เช่น KPS250610-003
    // เลขออเดอร์: ใช้ "เลขลำดับสูงสุดของวันนั้น + 1" (ไม่ใช่นับจำนวนแถว)
    // กันเลขซ้ำเมื่อมีออเดอร์ถูกยกเลิก/ลบ — เลขที่เคยออกไปแล้วจะไม่ถูกนำกลับมาใช้ซ้ำ
    const sh = getSheet_(SHEET_ORDERS);
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMdd');
    const prefix = 'KPS' + today + '-';
    let maxSeq = 0;
    const last = sh.getLastRow();
    if (last >= 2) {
      sh.getRange(2, O.ORDER_ID, last - 1, 1).getValues().forEach(function (r) {
        const v = String(r[0] || '');
        if (v.indexOf(prefix) === 0) {
          const n = parseInt(v.slice(prefix.length).replace(/\D/g, ''), 10);
          if (n > maxSeq) maxSeq = n;
        }
      });
    }
    const orderId = prefix + padNum_(maxSeq + 1, 3);

    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, O_COUNT).setValues([[
      ymd_(new Date()), orderId, customer, itemsText_(items), email, phone,
      items.length, JSON.stringify(items), subtotal, usePct, discount, net,
      STATUSES.PENDING, '', ''
    ]]);
    sh.getRange(row, O.TIMESTAMP).setNumberFormat('@');

    // 📥 เมลยืนยันรับออเดอร์ (ถ้าส่งไม่ได้ ไม่ให้กระทบการสั่งซื้อ)
    try {
      sendOrderConfirmEmail_({ orderId: orderId, customer: customer, email: email,
        items: items, discountPercent: usePct, discount: discount, netTotal: net });
    } catch (eMail) { Logger.log('confirm email: ' + eMail); }

    return { success: true, orderId: orderId, subtotal: subtotal,
      discountPercent: usePct, discount: discount, netTotal: net };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// อัปโหลดสลิป (base64 ถูกบีบอัดจากฝั่งหน้าเว็บแล้ว) → Drive → อัปเดตสถานะ
function uploadSlip(orderId, base64, mimeType) {
  try {
    const found = findOrder_(orderId);
    if (!found) throw new Error('ไม่พบเลขที่ออเดอร์ ' + orderId);
    if (found.status === STATUSES.CANCELLED) throw new Error('ออเดอร์นี้ถูกยกเลิกแล้ว');

    const folder = getOrCreateFolder_(SLIPS_FOLDER_NAME);
    const bytes = Utilities.base64Decode(String(base64).replace(/^data:[^;]+;base64,/, ''));
    const ext = (mimeType === 'image/png') ? 'png' : 'jpg';
    const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg',
      'slip_' + orderId + '_' + Date.now() + '.' + ext);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

    const sh = getSheet_(SHEET_ORDERS);
    sh.getRange(found.row, O.SLIP_URL).setValue(url);
    if (found.status === STATUSES.PENDING) {
      sh.getRange(found.row, O.STATUS).setValue(STATUSES.PAID);
    }

    // 🧾 กันสลิปซ้ำ: เทียบลายนิ้วมือไฟล์กับสลิปทุกใบที่เคยรับ
    const hashHex = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes));
    const dupOf = findSlipHash_(hashHex, found.orderId);
    recordSlipHash_(hashHex, found.orderId);

    // แจ้งเตือนแอดมินทาง Telegram อย่างเดียว (ไม่ส่งอีเมลแจ้งเตือนแล้ว)
    sendTelegram_('💸 สลิปใหม่ #' + found.orderId + '\nลูกค้า: ' + found.customer + ' (' + found.email + ')\nยอด: ' + found.netTotal + ' บาท\nดูสลิป: ' + url +
      (dupOf ? '\n\n🚨 สลิปใบนี้ซ้ำกับออเดอร์ #' + dupOf + ' — งดส่งอัตโนมัติ รอตรวจเอง' : ''));

    if (dupOf) {
      const cur = String(sh.getRange(found.row, O.NOTE).getValue() || '');
      sh.getRange(found.row, O.NOTE).setValue((cur ? cur + ' | ' : '') + '🚨 สลิปซ้ำกับ #' + dupOf);
    } else {
      // 🤖 ถ้าเปิดส่งอัตโนมัติ: ตรวจยอดในสลิปด้วย OCR → ตรงก็แชร์สิทธิ์+ส่งแผนทันที
      autoSendAfterSlip_(found.orderId, file.getId());
    }

    return { success: true, slipUrl: url };
  } catch (e) { return { success: false, error: e.message }; }
}

// ลูกค้าตรวจสอบสถานะ — ค้นด้วยเลขออเดอร์ หรืออีเมล
function getOrderStatus(query) {
  try {
    const q = String(query || '').trim().toLowerCase();
    if (!q) throw new Error('กรุณากรอกเลขที่ออเดอร์หรืออีเมล');
    const results = readOrders_().filter(function (o) {
      return o.orderId.toLowerCase() === q || o.email.toLowerCase() === q;
    }).map(function (o) {
      return { orderId: o.orderId, timestamp: o.timestamp, customer: o.customer,
        itemCount: o.itemCount, items: (o.items || []).map(function (it) {
          return { subject: it.subject, grade: it.grade, hours: it.hours, price: it.price };
        }),
        netTotal: o.netTotal, status: o.status, hasSlip: !!o.slipUrl };
    });
    return { success: true, orders: results.reverse().slice(0, 20) };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 📦 Orders (ฝั่งแอดมิน)
// ============================================================
function readOrders_() {
  const sh = getSheet_(SHEET_ORDERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, O_COUNT).getValues().map(function (r, i) {
    let items = [];
    items = parseItems_(r[O.ITEMS_JSON - 1]);
    return {
      row: i + 2,
      timestamp: String(r[O.TIMESTAMP - 1] || ''),
      orderId: String(r[O.ORDER_ID - 1] || ''),
      customer: String(r[O.CUSTOMER - 1] || ''),
      email: String(r[O.EMAIL - 1] || ''),
      phone: String(r[O.PHONE_FB - 1] || ''),
      itemCount: Number(r[O.ITEM_COUNT - 1] || 0),
      items: items,
      subtotal: Number(r[O.SUBTOTAL - 1] || 0),
      discountPct: Number(r[O.DISCOUNT_PCT - 1] || 0),
      discount: Number(r[O.DISCOUNT_BAHT - 1] || 0),
      netTotal: Number(r[O.NET_TOTAL - 1] || 0),
      status: String(r[O.STATUS - 1] || ''),
      slipUrl: String(r[O.SLIP_URL - 1] || ''),
      note: String(r[O.NOTE - 1] || '')
    };
  }).filter(function (o) { return o.orderId; });
}

function findOrder_(orderId) {
  const id = String(orderId || '').trim();
  return readOrders_().filter(function (o) { return o.orderId === id; })[0] || null;
}

// อ่านออเดอร์จาก "แถวที่ระบุ" ตรง ๆ — ใช้เมื่อทำงานจากในชีต (กดเมนู/แก้สถานะ)
// กันปัญหาเลขออเดอร์ซ้ำ: ถ้าใช้ findOrder_ ด้วยเลขที่ซ้ำ จะหยิบผิดแถว แต่ตัวนี้ผูกกับแถวที่เลือกจริง
function readOrderAtRow_(sh, row) {
  if (row < 2 || row > sh.getLastRow()) return null;
  const r = sh.getRange(row, 1, 1, O_COUNT).getValues()[0];
  const o = {
    row: row,
    timestamp: String(r[O.TIMESTAMP - 1] || ''),
    orderId: String(r[O.ORDER_ID - 1] || ''),
    customer: String(r[O.CUSTOMER - 1] || ''),
    email: String(r[O.EMAIL - 1] || ''),
    phone: String(r[O.PHONE_FB - 1] || ''),
    itemCount: Number(r[O.ITEM_COUNT - 1] || 0),
    items: parseItems_(r[O.ITEMS_JSON - 1]),
    subtotal: Number(r[O.SUBTOTAL - 1] || 0),
    discountPct: Number(r[O.DISCOUNT_PCT - 1] || 0),
    discount: Number(r[O.DISCOUNT_BAHT - 1] || 0),
    netTotal: Number(r[O.NET_TOTAL - 1] || 0),
    status: String(r[O.STATUS - 1] || ''),
    slipUrl: String(r[O.SLIP_URL - 1] || ''),
    note: String(r[O.NOTE - 1] || '')
  };
  return o.orderId ? o : null;
}

function listOrders(pin, statusFilter) {
  try {
    const admin = verifyAdmin_(pin);
    let orders = readOrders_().reverse();
    // บัญชีพนักงาน (staff) เห็นออเดอร์ 7 วันล่าสุด (รวมวันนี้) เท่านั้น
    if (admin.role === 'staff') {
      const cutoff = ymd_(new Date(Date.now() - 6 * 24 * 60 * 60 * 1000)).slice(0, 10);
      orders = orders.filter(function (o) { return o.timestamp.slice(0, 10) >= cutoff; });
    }
    if (statusFilter && statusFilter !== 'ทั้งหมด') {
      orders = orders.filter(function (o) { return o.status === statusFilter; });
    }
    return { success: true, orders: orders.slice(0, 500) };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateOrderStatus(pin, orderId, status, note) {
  try {
    const admin = verifyAdmin_(pin);
    const o = findOrder_(orderId);
    if (!o) throw new Error('ไม่พบออเดอร์ ' + orderId);
    const sh = getSheet_(SHEET_ORDERS);
    sh.getRange(o.row, O.STATUS).setValue(status);
    if (note !== undefined && note !== null && String(note).trim()) {
      sh.getRange(o.row, O.NOTE).setValue(String(note) + ' โดย ' + admin.name + ' เมื่อ ' + ymd_(new Date()));
    }
    // 🚫 ยกเลิกออเดอร์ → ถอนสิทธิ์ผู้อ่าน Drive คืนจากลูกค้าด้วย
    if (status === STATUSES.CANCELLED) {
      const res = revokeReadersForOrder_(o);
      const cur = String(sh.getRange(o.row, O.NOTE).getValue() || '');
      sh.getRange(o.row, O.NOTE).setValue((cur ? cur + ' | ' : '') + revokeSummary_(res));
      // (ปิดแจ้งเตือน Telegram ตอนยกเลิก+ถอนสิทธิ์ตามที่ขอ — ยกเลิก/ถอนสิทธิ์ยังทำงานปกติ)
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ✅ อนุมัติออเดอร์ → ส่งอีเมลลิงก์แผนอัตโนมัติ → สถานะ "ส่งแผนเสร็จสิ้น"
// ✏️ แก้ไขอีเมลลูกค้า (กรณีพิมพ์ผิด) — อัปเดตทั้งอีเมลหลักและอีเมลรายวิชาใน JSON
function updateOrderEmail(pin, orderId, newEmail) {
  try {
    const admin = verifyAdmin_(pin);
    const email = String(newEmail || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
    const o = findOrder_(orderId);
    if (!o) throw new Error('ไม่พบออเดอร์ ' + orderId);
    const sh = getSheet_(SHEET_ORDERS);
    sh.getRange(o.row, O.EMAIL).setValue(email);
    var items = o.items || [];
    if (items.length) {
      items = items.map(function (it) { it.email = email; return it; });
      sh.getRange(o.row, O.ITEMS_JSON).setValue(JSON.stringify(items));
    }
    return { success: true, email: email };
  } catch (e) { return { success: false, error: e.message }; }
}

function approveOrder(pin, orderId) {
  try {
    const admin = verifyAdmin_(pin);
    const o = findOrder_(orderId);
    if (!o) throw new Error('ไม่พบออเดอร์ ' + orderId);
    if (o.status === STATUSES.CANCELLED) throw new Error('ออเดอร์นี้ถูกยกเลิกแล้ว');

    const built = buildOrderBlocks_(o);
    if (!built.blocks.length) {
      throw new Error('อ่านรายการสินค้าในออเดอร์นี้ไม่ได้ (แถวอาจคอลัมน์เคลื่อน) — รันเมนู 🩹 ซ่อมแถวคอลัมน์เคลื่อนก่อน');
    }
    if (built.missing.length) {
      throw new Error('ยังไม่ได้ใส่ลิงก์ Drive/ข้อความของ: ' +
        built.missing.map(function (b) { return b.subject + ' ' + b.grade; }).join(', ') +
        ' (กรอกในเมนูจัดการแผนการสอนก่อน)');
    }

    // 🔑 แชร์สิทธิ์ "ผู้อ่าน" ใน Drive ให้ Gmail ลูกค้าทุกวิชาที่ซื้อ ก่อนส่งเมล
    const shareFailed = grantReadersForBlocks_(built.blocks, o.email);

    sendApprovalEmail_(o, built.blocks);

    const sh = getSheet_(SHEET_ORDERS);
    sh.getRange(o.row, O.STATUS).setValue(STATUSES.COMPLETED);
    sh.getRange(o.row, O.NOTE).setValue('ส่งโดย ' + admin.name + ' เมื่อ ' + ymd_(new Date()) +
      (shareFailed.length ? ' | ⚠️ แชร์สิทธิ์ไม่สำเร็จ: ' + shareFailed.join(', ') : ''));
    sendTelegram_('✅ ส่งแผนสำเร็จ #' + o.orderId + '\nลูกค้า: ' + o.customer + '\nยอด: ' + o.netTotal + ' บาท\nส่งโดย: ' + admin.name +
      (shareFailed.length ? '\n⚠️ แชร์สิทธิ์ไม่สำเร็จ ต้องแชร์เอง: ' + shareFailed.join(', ') : '\n🔑 แชร์สิทธิ์ผู้อ่านให้ลูกค้าแล้ว'));
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 📧 อีเมล
// ============================================================
// ทำคีย์ให้สม่ำเสมอ (ตัดช่องว่างทั้งหมด) สำหรับจับคู่ชื่อวิชา/ชั้น/ชั่วโมง
function normKey_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').trim();
}

// จับคู่รายการในออเดอร์กับลิงก์/ข้อความปัจจุบันในชีตแผนการสอน (ใช้ร่วมทุกช่องทางส่ง)
// ใช้ "คีย์รวมชื่อวิชา+ระดับชั้น" (ไม่มีตัวคั่น) จึงจับคู่ได้ทั้ง 2 แบบ:
//   • ออเดอร์ใหม่: subject="เศรษฐกิจและการเงิน", grade="ป.1"  → "เศรษฐกิจและการเงินป.1"
//   • ออเดอร์เก่า/ย้ายมา: subject="เศรษฐกิจและการเงิน ป.1", grade="" → "เศรษฐกิจและการเงินป.1"
// ทั้งคู่ตรงกับแผนในชีต (subject="เศรษฐกิจและการเงิน", grade="ป.1") → "เศรษฐกิจและการเงินป.1"
// เมื่อมีหลายแผนชื่อ+ชั้นเดียวกัน เลือกตัวที่ "มีลิงก์ + ชั่วโมงตรงที่สุด"
function buildOrderBlocks_(o) {
  const products = readProducts_();
  const byId = {};
  products.forEach(function (p) { byId[p.id] = p; });

  function hasLink_(p) { return !!(p && (p.driveLink || p.postNote)); }
  function combo_(subject, grade) { return normKey_(subject) + normKey_(grade); }

  function pick_(it) {
    const want = combo_(it.subject, it.grade);
    const hItem = normKey_(it.hours);
    const cItem = normKey_(it.curriculum);

    // รวมรายชื่อแผนที่ "อาจใช่" — ตัวที่ ID ตรง + ทุกตัวที่คีย์รวมชื่อวิชา+ชั้นตรงกัน
    const cands = [];
    const byid = byId[it.id];
    if (byid) cands.push(byid);
    products.forEach(function (p) {
      if (p === byid) return;
      if (want && combo_(p.subject, p.grade) === want) cands.push(p);
    });
    if (!cands.length) return null;

    // ให้คะแนน: มีลิงก์/ข้อความ (สำคัญสุด) > ชั่วโมงตรง > หลักสูตรตรง > ID ตรง
    function score_(p) {
      return (hasLink_(p) ? 1000 : 0) +
             (hItem && normKey_(p.hours) === hItem ? 100 : 0) +
             (cItem && normKey_(p.curriculum) === cItem ? 10 : 0) +
             (p === byid ? 1 : 0);
    }
    cands.sort(function (a, b) { return score_(b) - score_(a); });
    return cands[0];
  }

  const blocks = (Array.isArray(o.items) ? o.items : []).map(function (it) {
    const p = pick_(it) || {};
    return {
      subject: it.subject, grade: it.grade, curriculum: it.curriculum, hours: it.hours,
      category: (p.category || it.category || 'แผนการสอน'),
      email: it.email || '',          // อีเมลผู้รับเฉพาะวิชา (รองรับแยก Gmail ตามวิชา)
      link: p.driveLink || '', htmlMedia: p.htmlMedia || '',
      postNote: p.planDetail || '',   // เฉพาะวิชา (แสดงใต้แต่ละวิชา)
      footer: p.footer || ''          // ข้อความท้ายมาตรฐาน (แสดงครั้งเดียวท้ายเมล)
    };
  });
  // ต้องมีลิงก์ Drive หรือรายละเอียด/ข้อความสำหรับลูกค้า อย่างใดอย่างหนึ่ง (กลุ่ม VIP ใช้ข้อความลิงก์เชิญ)
  const missing = blocks.filter(function (b) { return !b.link && !b.postNote; });
  return { blocks: blocks, missing: missing };
}

// ============================================================
// 🔑 แชร์สิทธิ์ "ผู้อ่าน" ใน Google Drive ให้ Gmail ลูกค้า
// ============================================================
function extractDriveId_(url) {
  const m = String(url || '').match(/\/folders\/([a-zA-Z0-9_-]{10,})|\/file\/d\/([a-zA-Z0-9_-]{10,})|[?&]id=([a-zA-Z0-9_-]{10,})|\/d\/([a-zA-Z0-9_-]{10,})/);
  if (!m) return '';
  return m[1] || m[2] || m[3] || m[4] || '';
}
// คืนค่า '' = สำเร็จ / ข้อความ = เหตุผลที่ไม่สำเร็จ
function grantReader_(url, email) {
  const id = extractDriveId_(url);
  if (!id) return 'ไม่พบรหัสโฟลเดอร์ในลิงก์';
  try { DriveApp.getFolderById(id).addViewer(email); return ''; }
  catch (e1) {
    try { DriveApp.getFileById(id).addViewer(email); return ''; }
    catch (e2) { return String(e2.message || e2).substring(0, 80); }
  }
}
// แชร์ทุกวิชาที่มีลิงก์ — ส่งสิทธิ์ให้อีเมลเฉพาะวิชา (ถ้ามี) ไม่งั้นใช้อีเมลหลัก
function grantReadersForBlocks_(blocks, fallbackEmail) {
  const failed = [];
  blocks.forEach(function (b) {
    if (!b.link) return;
    const to = b.email || fallbackEmail;
    const err = grantReader_(b.link, to);
    if (err) failed.push(b.subject + ' ' + b.grade + ' (' + err + ')');
  });
  return failed;
}

// ============================================================
// 🤖 ส่งแผนอัตโนมัติเมื่อยอดสลิปตรง (สวิตช์ autoSend ในหน้าตั้งค่า)
// ============================================================
function autoSendAfterSlip_(orderId, slipFileId) {
  try {
    if (String(settingsMap_().autoSend || 'off') !== 'on') return;
    const o = findOrder_(orderId);
    if (!o || o.status !== STATUSES.PAID) return;

    const check = verifySlipAmount_(slipFileId, o.netTotal);
    if (!check.matched) {
      sendTelegram_('⚠️ ส่งอัตโนมัติไม่ได้ — ต้องตรวจเอง\n#' + o.orderId + ' ' + o.customer +
        '\nยอดออเดอร์: ' + o.netTotal + ' บาท\nผลตรวจสลิป: ' + check.reason +
        (check.foundNumbers.length ? '\nตัวเลขที่อ่านได้: ' + check.foundNumbers.join(', ') : ''));
      return;
    }
    const built = buildOrderBlocks_(o);
    if (!built.blocks.length) {
      sendTelegram_('⚠️ ยอดตรง (฿' + check.matchedAmount + ') แต่อ่านรายการในออเดอร์ไม่ได้ — รันเมนู 🩹 แล้วส่งเองจากหน้าต้องส่ง\n#' + o.orderId);
      return;
    }
    if (built.missing.length) {
      sendTelegram_('⚠️ ยอดตรง (฿' + check.matchedAmount + ') แต่ส่งอัตโนมัติไม่ได้\n#' + o.orderId +
        '\nยังไม่มีลิงก์/ข้อความของ: ' +
        built.missing.map(function (b) { return b.subject + ' ' + b.grade; }).join(', '));
      return;
    }
    const shareFailed = grantReadersForBlocks_(built.blocks, o.email);
    sendApprovalEmail_(o, built.blocks);
    const sh = getSheet_(SHEET_ORDERS);
    sh.getRange(o.row, O.STATUS).setValue(STATUSES.COMPLETED);
    sh.getRange(o.row, O.NOTE).setValue('🤖 ส่งอัตโนมัติ เมื่อ ' + ymd_(new Date()) +
      (shareFailed.length ? ' | ⚠️ แชร์สิทธิ์ไม่สำเร็จ: ' + shareFailed.join(', ') : ''));
    sendTelegram_('🤖✅ ส่งแผนอัตโนมัติสำเร็จ!\n#' + o.orderId + '\nลูกค้า: ' + o.customer + ' (' + o.email + ')' +
      '\nตรวจสลิป: ยอดตรง ฿' + check.matchedAmount +
      '\nส่งแผน ' + built.blocks.length + ' รายการทางอีเมล + แชร์สิทธิ์ผู้อ่านแล้ว' +
      (shareFailed.length ? '\n⚠️ แชร์สิทธิ์ไม่สำเร็จ ต้องแชร์เอง: ' + shareFailed.join(', ') : ''));
  } catch (e) {
    sendTelegram_('⚠️ ระบบส่งอัตโนมัติขัดข้อง #' + orderId + ': ' + String(e.message || e).substring(0, 120) + '\nรายการยังค้างอยู่ที่หน้า "ต้องส่ง" กดส่งเองได้เลย');
    Logger.log('autoSendAfterSlip_ error: ' + e);
  }
}

// อ่านยอดเงินในรูปสลิปด้วย Google Drive OCR แล้วเทียบกับยอดออเดอร์ (คลาดเคลื่อนได้ไม่เกิน 1 บาท)
function verifySlipAmount_(fileId, expectedAmount) {
  var tempDocId = null;
  try {
    expectedAmount = Number(expectedAmount) || 0;
    if (expectedAmount <= 0) return { matched: false, matchedAmount: null, foundNumbers: [], reason: 'ยอดออเดอร์ไม่ถูกต้อง' };

    var blob = DriveApp.getFileById(fileId).getBlob();

    // วิธีที่ 1: Advanced Drive Service (ต้องเปิดใน Services → Drive API)
    if (typeof Drive !== 'undefined' && Drive.Files) {
      try {
        if (Drive.Files.insert) {
          tempDocId = Drive.Files.insert(
            { title: 'OCR_temp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
            blob, { ocr: true, ocrLanguage: 'en' }).id;
        } else if (Drive.Files.create) {
          tempDocId = Drive.Files.create(
            { name: 'OCR_temp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
            blob, { ocrLanguage: 'en', fields: 'id' }).id;
        }
      } catch (e1) { Logger.log('Advanced Drive OCR failed, fallback REST: ' + e1); }
    }

    // วิธีที่ 2 (สำรอง): REST API อัปโหลดภาพ → copy เป็น Google Doc พร้อม OCR
    if (!tempDocId) {
      var token = ScriptApp.getOAuthToken();
      var uploadResp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files?uploadType=media', {
        method: 'post', contentType: blob.getContentType() || 'image/jpeg',
        payload: blob.getBytes(), headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
      if (uploadResp.getResponseCode() !== 200) {
        return { matched: false, matchedAmount: null, foundNumbers: [],
          reason: 'OCR upload ล้มเหลว (HTTP ' + uploadResp.getResponseCode() + ') — แนะนำเปิด Advanced Drive Service' };
      }
      var uploadedId = JSON.parse(uploadResp.getContentText()).id;
      try {
        var copyResp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/' + uploadedId + '/copy?ocr=true&ocrLanguage=en', {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify({ title: 'OCR_temp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' }),
          headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
        });
        if (copyResp.getResponseCode() === 200) tempDocId = JSON.parse(copyResp.getContentText()).id;
      } finally {
        try { DriveApp.getFileById(uploadedId).setTrashed(true); } catch (eDel) {}
      }
      if (!tempDocId) {
        return { matched: false, matchedAmount: null, foundNumbers: [],
          reason: 'OCR แปลงไฟล์ล้มเหลว — แนะนำเปิด Advanced Drive Service' };
      }
    }

    // อ่านตัวเลขทั้งหมดจากข้อความที่ OCR ได้ แล้วหายอดที่ตรง
    var text = DocumentApp.openById(tempDocId).getBody().getText();
    var matches = text.match(/\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+\.\d{2}|\d{2,}/g) || [];
    var foundNumbers = [];
    var matched = false, matchedAmount = null;
    for (var i = 0; i < matches.length; i++) {
      var num = parseFloat(matches[i].replace(/,/g, ''));
      if (isNaN(num) || num < 1) continue;
      foundNumbers.push(matches[i]);
      if (Math.abs(num - expectedAmount) <= 1) { matched = true; matchedAmount = num; break; }
    }
    foundNumbers = foundNumbers.filter(function (v, idx, arr) { return arr.indexOf(v) === idx; }).slice(0, 10);

    return matched
      ? { matched: true, matchedAmount: matchedAmount, foundNumbers: foundNumbers, reason: 'พบยอดตรง ฿' + matchedAmount }
      : { matched: false, matchedAmount: null, foundNumbers: foundNumbers,
          reason: foundNumbers.length ? 'ไม่พบยอด ฿' + expectedAmount + ' ในสลิป' : 'อ่านตัวเลขจากสลิปไม่ได้ (ภาพอาจไม่ชัด)' };
  } catch (err) {
    Logger.log('verifySlipAmount_ error: ' + err);
    return { matched: false, matchedAmount: null, foundNumbers: [], reason: 'ระบบ OCR ขัดข้อง (' + String(err).substring(0, 60) + ')' };
  } finally {
    if (tempDocId) { try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch (eC) {} }
  }
}

// 📥 เมลยืนยันรับออเดอร์ — ส่งให้ลูกค้าทันทีที่กดสั่งซื้อ (มียอด + ช่องทางโอน + ขั้นตอนถัดไป)
function sendOrderConfirmEmail_(order) {
  const m = settingsMap_();
  const pageName = m.pageName || PAGE_NAME;
  const itemRows = (order.items || []).map(function (it) {
    return '<tr><td style="padding:7px 10px; border-bottom:1px solid #f3f4f6; font-size:13px; color:#374151;">📘 ' +
      escapeHtml_(String(it.subject || '')) + ' ' + escapeHtml_(String(it.grade || '')) + '</td>' +
      '<td style="padding:7px 10px; border-bottom:1px solid #f3f4f6; font-size:13px; color:#374151; text-align:right; white-space:nowrap;">' +
      (Number(it.price) || 0).toLocaleString() + ' บาท</td></tr>';
  }).join('');
  const discountRow = (Number(order.discount) > 0)
    ? '<tr><td style="padding:7px 10px; font-size:13px; color:#059669;">ส่วนลด ' + order.discountPercent + '%</td>' +
      '<td style="padding:7px 10px; font-size:13px; color:#059669; text-align:right;">-' + Number(order.discount).toLocaleString() + ' บาท</td></tr>'
    : '';
  const qrBlock = m.qrURL
    ? '<div style="text-align:center; margin:16px 0;"><img src="' + escapeHtml_(m.qrURL) + '" alt="QR ชำระเงิน" style="max-width:240px; border-radius:10px; border:1px solid #e5e7eb;"><div style="font-size:12px; color:#6b7280; margin-top:6px;">' + escapeHtml_(m.promptpayName || '') + '</div></div>'
    : '';
  const html =
    '<!DOCTYPE html><html><body style="font-family: Sarabun, Arial, sans-serif; background:#f9fafb; padding:20px; margin:0;">' +
    '<div style="max-width:640px; margin:0 auto; background:#fff; border-radius:12px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.08);">' +
    '<div style="background:linear-gradient(135deg,#34d399,#10b981); padding:26px 24px; color:#fff; text-align:center;">' +
    '<div style="font-size:34px; margin-bottom:6px;">📥</div>' +
    '<h1 style="margin:0; font-size:21px;">รับคำสั่งซื้อเรียบร้อยแล้ว</h1>' +
    '<p style="margin:8px 0 0; opacity:0.95; font-size:13px;">ขอบคุณที่อุดหนุนเพจ' + escapeHtml_(pageName) + ' ค่ะ</p></div>' +
    '<div style="padding:24px;">' +
    '<p style="font-size:15px; color:#1f2937; margin:0 0 14px;">สวัสดีค่ะ คุณ' + escapeHtml_(order.customer) + ' 🌷</p>' +
    '<div style="background:#f0fdf4; border-radius:10px; padding:14px; margin-bottom:18px; border:1px solid #bbf7d0;">' +
    '<div style="font-size:12px; color:#6b7280;">รหัสคำสั่งซื้อ (ใช้ตรวจสอบสถานะ)</div>' +
    '<div style="font-size:18px; font-weight:bold; color:#047857;">#' + escapeHtml_(order.orderId) + '</div></div>' +
    '<table style="width:100%; border-collapse:collapse; margin-bottom:6px;">' + itemRows + discountRow +
    '<tr><td style="padding:10px; font-size:15px; font-weight:700; color:#111827; background:#fafafa;">ยอดที่ต้องชำระ</td>' +
    '<td style="padding:10px; font-size:17px; font-weight:800; color:#dc2626; text-align:right; background:#fafafa;">' + Number(order.netTotal).toLocaleString() + ' บาท</td></tr></table>' +
    qrBlock +
    '<div style="background:#fffbeb; border:1px solid #fde68a; border-radius:10px; padding:14px; font-size:13px; color:#92400e; line-height:1.8; margin-top:14px;">' +
    '<b>ขั้นตอนถัดไป:</b><br>1) โอนเงินตามยอดด้านบน<br>2) กดปุ่มด้านล่างเพื่อเปิดหน้าสถานะออเดอร์นี้ แล้วแนบสลิปได้เลย<br>3) ยอดถูกต้อง ระบบจะส่งแผนเข้าอีเมลนี้ให้อัตโนมัติทันทีค่ะ 🤖</div>' +
    '<div style="text-align:center; margin-top:18px;"><a href="' + SHOP_URL + '?order=' + encodeURIComponent(order.orderId) + '" style="display:inline-block; background:linear-gradient(135deg,#0284c7,#0ea5e9); color:#fff; text-decoration:none; padding:11px 26px; border-radius:8px; font-weight:700; font-size:14px;">🔍 ตรวจสถานะ / แนบสลิปออเดอร์นี้</a></div>' +
    '<p style="font-size:12px; color:#9ca3af; margin-top:20px; text-align:center;">หากมีข้อสงสัย ทักแชทเพจ ' + escapeHtml_(pageName) + ' ได้เลยนะคะ 💛</p>' +
    '</div></div></body></html>';
  MailApp.sendEmail({ to: order.email, subject: '📥 รับคำสั่งซื้อแล้ว #' + order.orderId + ' - ' + pageName, htmlBody: html, name: pageName });
}

function sendApprovalEmail_(order, blocks) {
  // จัดกลุ่มบล็อกตามอีเมลผู้รับ (รองรับ "แยก Gmail ตามวิชา") แล้วส่งแยกฉบับให้แต่ละคน
  const groups = {};
  blocks.forEach(function (b) {
    const to = b.email || order.email;
    (groups[to] = groups[to] || []).push(b);
  });
  Object.keys(groups).forEach(function (to) {
    sendApprovalEmailTo_(order, groups[to], to);
  });
}

function sendApprovalEmailTo_(order, blocks, toEmail) {
  const m = settingsMap_();
  const pageName = m.pageName || PAGE_NAME;
  const subject = '✅ แผนการสอนของคุณพร้อมแล้ว - ' + pageName + ' #' + order.orderId;

  const blockHtml = blocks.map(function (b) {
    const icon = b.category === 'ระบบงาน' ? '💼' : (b.category === 'กลุ่ม VIP' ? '👑' : '📘');
    return '' +
      '<div style="background:#fafafa; border-radius:10px; padding:18px 20px; margin-bottom:14px; border-left:4px solid #818cf8;">' +
      '<div style="font-weight:700; color:#3730a3; font-size:15px; margin-bottom:8px;">' + icon + ' ' +
      escapeHtml_(b.subject) + ' ' + escapeHtml_(b.grade) +
      (b.curriculum ? ' (หลักสูตร ' + escapeHtml_(b.curriculum) + ')' : '') +
      (b.hours ? ' — ' + b.hours + ' ชม.' : '') + '</div>' +
      (b.link ? '<a href="' + escapeHtml_(b.link) + '" style="display:inline-block; background:linear-gradient(135deg,#818cf8,#a78bfa); color:#fff; text-decoration:none; padding:10px 22px; border-radius:8px; font-weight:700; font-size:14px;">📂 เปิดไฟล์ใน Google Drive</a>' : '') +
      (b.postNote ? '<div style="margin-top:12px; padding-top:12px; border-top:1px dashed #d1d5db; font-size:13px; color:#4b5563; line-height:1.75;">' + linkifyHtml_(b.postNote) + '</div>' : '') +
      '</div>';
  }).join('');

  // ข้อความท้ายมาตรฐาน — แสดงครั้งเดียวท้ายเมล (ไม่ซ้ำใต้ทุกวิชา)
  let footerText = '';
  for (var fi = 0; fi < blocks.length; fi++) { if (blocks[fi].footer) { footerText = blocks[fi].footer; break; } }
  const footerHtml = footerText
    ? '<div style="background:#fafafa; border:1px solid #eee; border-radius:10px; padding:16px 20px; margin-top:6px; font-size:13px; color:#4b5563; line-height:1.75;">' + linkifyHtml_(footerText) + '</div>'
    : '';  const html =
    '<!DOCTYPE html><html><body style="font-family: Sarabun, Arial, sans-serif; background:#f9fafb; padding:20px; margin:0;">' +
    '<div style="max-width:640px; margin:0 auto; background:#fff; border-radius:12px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.08);">' +
    '<div style="background:linear-gradient(135deg,#818cf8,#a78bfa,#f0abfc); padding:28px 24px; color:#fff; text-align:center;">' +
    '<div style="font-size:36px; margin-bottom:6px;">✨</div>' +
    '<h1 style="margin:0; font-size:22px;">แอดมินได้รับยอดโอนเรียบร้อยแล้ว</h1>' +
    '<p style="margin:10px 0 0; opacity:0.95; font-size:14px;">🙏 ขอบคุณที่ใช้บริการเพจ' + escapeHtml_(pageName) + '</p>' +
    '</div>' +
    '<div style="padding:24px;">' +
    '<p style="font-size:15px; color:#1f2937; margin:0 0 14px;">สวัสดีค่ะ คุณ' + escapeHtml_(order.customer) + ' 🌷</p>' +
    '<p style="font-size:14px; color:#4b5563; margin:0 0 18px; line-height:1.7;">นี่คือลิงก์ไฟล์แผนการสอนของคุณ เปิดด้วย <strong>Gmail นี้ (' + escapeHtml_(toEmail) + ')</strong> ได้เลยค่ะ</p>' +
    '<div style="background:#f8f7ff; border-radius:10px; padding:14px; margin-bottom:24px; border:1px solid #e9d5ff;">' +
    '<div style="font-size:12px; color:#6b7280;">รหัสคำสั่งซื้อ</div>' +
    '<div style="font-size:18px; font-weight:bold; color:#5b21b6;">#' + escapeHtml_(order.orderId) + '</div>' +
    '</div>' +
    blockHtml +
    footerHtml +
    '<p style="font-size:13px; color:#6b7280; margin-top:22px; line-height:1.7;">หากเปิดลิงก์ไม่ได้ หรือมีข้อสงสัย ทักแชทเพจ <b>' + escapeHtml_(pageName) + '</b> ได้เลยนะคะ 💛</p>' +
    '</div></div></body></html>';

  MailApp.sendEmail({
    to: toEmail,
    subject: subject,
    htmlBody: html,
    name: pageName
  });
}

function linkifyHtml_(text) {
  return escapeHtml_(text)
    .replace(/(https?:\/\/[^\s]+)/g, '<a href="$1">$1</a>')
    .replace(/\n/g, '<br>');
}


// ============================================================
// 📊 รายงานสรุป (แอดมิน)
// ============================================================
function getReportData(pin) {
  try {
    requireOwner_(pin);
    const orders = readOrders_();
    const tz = 'Asia/Bangkok';
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const monthStr = todayStr.slice(0, 7);

    const done = orders.filter(function (o) {
      return o.status === STATUSES.COMPLETED || o.status === STATUSES.APPROVED;
    });
    const statusCount = {};
    Object.keys(STATUSES).forEach(function (k) { statusCount[STATUSES[k]] = 0; });
    orders.forEach(function (o) {
      statusCount[o.status] = (statusCount[o.status] || 0) + 1;
    });

    // ยอดขายรายวัน 30 วันล่าสุด (นับเฉพาะออเดอร์สำเร็จ)
    const daily = {};
    for (let i = 29; i >= 0; i--) {
      const d = new Date(); d.setDate(d.getDate() - i);
      daily[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = 0;
    }
    let todaySum = 0, monthSum = 0, allSum = 0;
    const subjectCount = {};
    done.forEach(function (o) {
      const day = o.timestamp.slice(0, 10);
      allSum += o.netTotal;
      if (day === todayStr) todaySum += o.netTotal;
      if (day.slice(0, 7) === monthStr) monthSum += o.netTotal;
      if (daily.hasOwnProperty(day)) daily[day] += o.netTotal;
      (o.items || []).forEach(function (it) {
        const key = it.subject + ' ' + it.grade;
        subjectCount[key] = (subjectCount[key] || 0) + 1;
      });
    });
    const topSubjects = Object.keys(subjectCount)
      .map(function (k) { return { name: k, count: subjectCount[k] }; })
      .sort(function (a, b) { return b.count - a.count; })
      .slice(0, 10);

    return {
      success: true,
      report: {
        totalOrders: orders.length,
        completedOrders: done.length,
        todaySum: todaySum, monthSum: monthSum, allSum: allSum,
        statusCount: statusCount,
        dailyLabels: Object.keys(daily),
        dailyValues: Object.keys(daily).map(function (k) { return daily[k]; }),
        topSubjects: topSubjects
      }
    };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 🖼️ อัปโหลดรูป (โลโก้ / QR PromptPay) → Drive → คืน URL
// ============================================================
function uploadImageToDrive(pin, base64, filename, mimeType) {
  try {
    requireOwner_(pin);
    const folder = getOrCreateFolder_(IMAGES_FOLDER_NAME);
    const bytes = Utilities.base64Decode(String(base64).replace(/^data:[^;]+;base64,/, ''));
    const blob = Utilities.newBlob(bytes, mimeType || 'image/png',
      (filename || 'image') + '_' + Date.now());
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // ใช้รูปแบบ lh3.googleusercontent.com (uc?export=view เลิกใช้แล้วตั้งแต่ปี 2025)
    return { success: true, url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
  } catch (e) { return { success: false, error: e.message }; }
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}


// ============================================================
// 💾 Backup / Import / ล้างข้อมูล (แอดมิน)
// ============================================================
function backupAllData(pin) {
  try {
    requireOwner_(pin);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dump = {};
    [SHEET_SETTINGS].concat(productSheetNames_(), [SHEET_ORDERS]).forEach(function (name) {
      const sh = ss.getSheetByName(name);
      dump[name] = sh ? sh.getDataRange().getValues() : [];
    });
    return { success: true, exportedAt: ymd_(new Date()), data: dump };
  } catch (e) { return { success: false, error: e.message }; }
}

function importBackupData(pin, jsonData) {
  try {
    requireOwner_(pin);
    const obj = (typeof jsonData === 'string') ? JSON.parse(jsonData) : jsonData;
    const dump = obj.data || obj;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [SHEET_SETTINGS].concat(productSheetNames_(), [SHEET_ORDERS]).forEach(function (name) {
      const rows = dump[name];
      if (!rows || !rows.length) return;
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      sh.clearContents();
      sh.getRange(1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
      sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
    });
    return { success: true };
  } catch (e) { return { success: false, error: 'นำเข้าไม่สำเร็จ: ' + e.message }; }
}

// ล้างเฉพาะประวัติคำสั่งซื้อ (ไม่แตะแผนการสอน/ตั้งค่า) — ใช้กับปุ่มในหน้า GitHub
function clearOrders(pin) {
  try {
    requireOwner_(pin);
    const sh = getSheet_(SHEET_ORDERS);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ล้างข้อมูล Orders + Products (คงหัวตารางและ Settings ไว้)
function clearAllData(pin) {
  try {
    requireOwner_(pin);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [SHEET_ORDERS].concat(productSheetNames_()).forEach(function (name) {
      const sh = ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 1) {
        sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
      }
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 💬 Telegram (แจ้งเตือน 2 จังหวะ: สลิปใหม่ / ส่งแผนสำเร็จ)
// ============================================================
function sendTelegram_(text) {
  try {
    const m = settingsMap_();
    const token = String(m.telegramBotToken || '').trim();
    const chatId = String(m.telegramChatId || '').trim();
    if (!token || !chatId) return; // ยังไม่ได้ตั้งค่า = ข้ามเงียบๆ
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      payload: { chat_id: chatId, text: text },
      muteHttpExceptions: true
    });
  } catch (e) { Logger.log('sendTelegram_ error: ' + e); }
}

function testTelegram(pin) {
  try {
    const admin = requireOwner_(pin);
    const m = settingsMap_();
    if (!String(m.telegramBotToken || '').trim()) throw new Error('ยังไม่ได้ใส่ Bot Token ในตั้งค่า');
    if (!String(m.telegramChatId || '').trim()) throw new Error('ยังไม่ได้ใส่ Chat ID — กดปุ่ม "หา Chat ID" ก่อน');
    sendTelegram_('🔔 ทดสอบแจ้งเตือนจากระบบสั่งซื้อแผนการสอน ' + PAGE_NAME + '\nผู้ทดสอบ: ' + admin.name + '\nเวลา: ' + ymd_(new Date()));
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// หา Chat ID: ให้ทักข้อความหาบอทใน Telegram ก่อน แล้วกดปุ่มนี้
function getTelegramChatId(pin) {
  try {
    requireOwner_(pin);
    const token = String(settingsMap_().telegramBotToken || '').trim();
    if (!token) throw new Error('ใส่และบันทึก Bot Token ก่อน');
    const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates',
      { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    if (!data.ok) throw new Error('Token ไม่ถูกต้อง: ' + (data.description || ''));
    const ids = {};
    (data.result || []).forEach(function (u) {
      const chat = u.message && u.message.chat;
      if (chat) ids[chat.id] = (chat.first_name || '') + ' ' + (chat.username ? '@' + chat.username : '');
    });
    const keys = Object.keys(ids);
    if (!keys.length) throw new Error('ไม่พบข้อความ — เปิด Telegram ทักหาบอทของคุณก่อน (พิมพ์อะไรก็ได้) แล้วกดปุ่มนี้อีกครั้ง');
    return { success: true, chats: keys.map(function (k) { return { id: k, name: ids[k].trim() }; }) };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 👥 ประวัติลูกค้า (คำนวณสดจากชีต Orders — ไม่มีชีตซ้ำซ้อน)
// ============================================================
function listCustomers(pin) {
  try {
    requireOwner_(pin);
    const map = {};
    readOrders_().forEach(function (o) {
      if (!o.email) return;
      if (o.status === STATUSES.CANCELLED) return;
      const c = map[o.email] || { email: o.email, name: o.customer, phone: '',
        orderCount: 0, totalSpent: 0, lastDate: '', subjects: {} };
      c.name = o.customer || c.name;
      if (o.phone) c.phone = o.phone;
      c.orderCount++;
      if (o.status === STATUSES.COMPLETED || o.status === STATUSES.APPROVED) c.totalSpent += o.netTotal;
      if (o.timestamp > c.lastDate) c.lastDate = o.timestamp;
      (o.items || []).forEach(function (it) { c.subjects[it.subject + ' ' + it.grade] = true; });
      map[o.email] = c;
    });
    const customers = Object.keys(map).map(function (k) {
      const c = map[k];
      return { email: c.email, name: c.name, phone: c.phone, orderCount: c.orderCount,
        totalSpent: c.totalSpent, lastDate: c.lastDate,
        subjectCount: Object.keys(c.subjects).length };
    }).sort(function (a, b) { return a.lastDate < b.lastDate ? 1 : -1; });
    return { success: true, customers: customers };
  } catch (e) { return { success: false, error: e.message }; }
}

function getCustomerHistory(pin, email) {
  try {
    requireOwner_(pin);
    const q = String(email || '').trim().toLowerCase();
    const orders = readOrders_().filter(function (o) { return o.email.toLowerCase() === q; })
      .reverse().map(function (o) {
        return { orderId: o.orderId, timestamp: o.timestamp, items: o.items,
          netTotal: o.netTotal, status: o.status, note: o.note, slipUrl: o.slipUrl };
      });
    return { success: true, orders: orders };
  } catch (e) { return { success: false, error: e.message }; }
}


// ============================================================
// 🎓 เมนู "ครูพร้อมสอน" ใน Google Sheets + สีรายวัน + กรองเดือน
// ============================================================
/* ============================================================
 *  ➕ เติมรายวิชา ม.1–ม.6 (หลักสูตร 51) ลงชีต "แผนการสอน" ที่มีอยู่
 *  ใช้กับชีตเดิมที่มีวิชา ป.1–6 อยู่แล้ว — ต่อท้ายอย่างเดียว ไม่แตะแถวเดิม
 *  กดซ้ำได้ปลอดภัย: รายการที่มีอยู่แล้ว (ชั้น+หลักสูตร+วิชา+ชม. ตรงกัน) จะถูกข้าม
 * ============================================================ */
function seedMattayomProducts() {
  const ui = SpreadsheetApp.getUi();
  const rows = SEED_PRODUCTS.filter(function (r) { return String(r[0]).indexOf('ม.') === 0; });
  if (!rows.length) { ui.alert('ไม่พบรายวิชา ม. ในแคตตาล็อกตั้งต้น'); return; }

  // คีย์กันซ้ำ: ชั้น|หลักสูตร|วิชา|ชม. (ตัดช่องว่างเพื่อเทียบแบบทน)
  const norm = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); };
  const key = function (g, c, s, h) { return [norm(g), norm(c), norm(s), String(Number(h) || '')].join('|'); };
  const existing = {};
  readProducts_().forEach(function (p) { existing[key(p.grade, p.curriculum, p.subject, p.hours)] = 1; });

  const toAdd = rows.filter(function (r) { return !existing[key(r[0], r[1], r[2], r[3])]; });
  if (!toAdd.length) { ui.alert('รายวิชา ม.1–ม.6 (หลักสูตร 51) มีครบในชีตอยู่แล้ว ' + rows.length + ' รายการ — ไม่ต้องเติมเพิ่มค่ะ'); return; }

  const res = ui.alert('➕ เติมรายวิชา ม.1–ม.6 (หลักสูตร 51)',
    'จะเพิ่ม ' + toAdd.length + ' วิชาลงชีต "' + SHEET_PRODUCTS_M + '" (แยกจากชีต ป.)' +
    (toAdd.length < rows.length ? ' (ข้ามที่มีแล้ว ' + (rows.length - toAdd.length) + ' รายการ)' : '') +
    '\n\nราคา = คิดอัตโนมัติตามชั่วโมงจากตารางราคา / ลิงก์ Drive ปล่อยว่างไว้กรอกภายหลัง\n\nดำเนินการต่อไหมคะ?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  // เลข PD ถัดไป นับต่อจากเลขสูงสุดทุกชีต (รูปแบบเดียวกับ nextProductId_)
  let maxId = 0;
  readProducts_().forEach(function (p) {
    const m = String(p.id).match(/^PD(\d+)$/);
    if (m && Number(m[1]) > maxId) maxId = Number(m[1]);
  });
  const now = ymd_(new Date());
  const bySheet = {};   // กระจายลงชีต ม.1–ม.6 ตามชั้นของแต่ละวิชา
  toAdd.forEach(function (r, i) {
    const hours = Number(r[3]) || 0;
    const name = sheetNameForProduct_(r[0], 'แผนการสอน');
    (bySheet[name] = bySheet[name] || []).push(
      ['PD' + padNum_(maxId + 1 + i, 4), r[0], r[1], r[2], hours,
        priceForHours_(hours) || '', '', '', '', STD_FOOTER, 'TRUE', now, 'แผนการสอน', autoGroupFor_(r[1], r[2])]);
  });
  const parts = [];
  Object.keys(bySheet).forEach(function (name) {
    const s = getSheet_(name);
    const rows = bySheet[name];
    const start = s.getLastRow() + 1;
    s.getRange(start, 1, rows.length, P_COUNT).setValues(rows);
    s.getRange(start, P.ADDED, rows.length, 1).setNumberFormat('@');
    sortProductSheet_(s);
    parts.push(name + ' ' + rows.length + ' วิชา');
  });
  clearShopCache_();
  ui.alert('✅ เพิ่มรายวิชามัธยมแล้ว: ' + parts.join(' • ') + '\n(เรียงชื่อวิชา ก-ฮ ให้เรียบร้อย)\n\nขั้นถัดไป:\n1) กรอกลิงก์ Drive ของแต่ละวิชา\n2) เมนู 🛠 เครื่องมือระบบ → 🔗 เติมลิงก์ตัวอย่าง (Preview) ทุกวิชา');
}

/* ============================================================
 *  📗 แยกชีตรายชั้น (ครั้งเดียว): ย้ายข้อมูลจากชีตรวมเดิม
 *  (แผนประถม / แผนมัธยม / แผนการสอน / Products) ไปยังชีตรายชั้น
 *  ป.1–ป.6, ม.1–ม.6, VIP & ระบบงาน, อื่นๆ แบบ "ยกทั้งแถว"
 *  ข้อมูลทุกช่องจึงตรงกันเหมือนเดิม แล้วเรียงชื่อวิชา ก-ฮ ให้ทุกชีต
 *  ชีตเดิมไม่ถูกลบ — เปลี่ยนชื่อเป็น "สำรอง_..." เก็บไว้ตรวจย้อนหลัง
 * ============================================================ */
function migrateToGradeSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets_();
  const legacy = [SHEET_PRODUCTS, SHEET_PRODUCTS_M, 'แผนการสอน', 'Products']
    .map(function (n) { return ss.getSheetByName(n); })
    .filter(function (s) { return s && s.getLastRow() > 1; });
  if (!legacy.length) { ui.alert('ไม่พบชีตรวมรุ่นเก่าที่มีข้อมูล — ระบบใช้ชีตรายชั้นอยู่แล้วค่ะ'); return; }

  const total = legacy.reduce(function (n, s) { return n + (s.getLastRow() - 1); }, 0);
  const res = ui.alert('📗 แยกชีตรายชั้น',
    'จะย้ายข้อมูล ' + total + ' แถว จากชีต: ' + legacy.map(function (s) { return s.getName(); }).join(', ') +
    '\nไปยังชีตรายชั้น (ป.1–ป.6, ม.1–ม.6, VIP & ระบบงาน, อื่นๆ) แบบยกทั้งแถว' +
    '\nแล้วเรียงชื่อวิชา ก-ฮ ให้ทุกชีต' +
    '\n\nชีตเดิมจะถูกเปลี่ยนชื่อเป็น "สำรอง_..." เก็บไว้ (ไม่ลบทิ้ง)\n\nดำเนินการต่อไหมคะ?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  let moved = 0;
  const bySheet = {};
  legacy.forEach(function (s) {
    const last = s.getLastRow();
    s.getRange(2, 1, last - 1, P_COUNT).getValues().forEach(function (r) {
      if (!String(r[P.ID - 1] || '').trim()) return;   // ข้ามแถวว่าง
      const name = sheetNameForProduct_(r[P.GRADE - 1], r[P.CATEGORY - 1]);
      (bySheet[name] = bySheet[name] || []).push(r);
      moved++;
    });
  });
  Object.keys(bySheet).forEach(function (name) {
    const t = getSheet_(name);
    const rows = bySheet[name].map(function (r) {
      if (!isVipSheet_(name)) return r.slice(0, P_COUNT);
      return [r[P.ID - 1], r[P.SUBJECT - 1], r[P.PRICE - 1], r[P.DRIVE_LINK - 1], r[P.EXAMPLE_LINK - 1],
        r[P.POST_NOTE - 1], r[P.FOOTER - 1], r[P.ACTIVE - 1], r[P.ADDED - 1], r[P.CATEGORY - 1]];
    });
    const width = sheetColCount_(name);
    const start = t.getLastRow() + 1;
    t.getRange(start, 1, rows.length, width).setValues(rows);
    t.getRange(start, colFor_(name, P.ADDED), rows.length, 1).setNumberFormat('@');
    sortProductSheet_(t);
  });
  // เก็บชีตเดิมเป็นตัวสำรอง (พ้นสายตาระบบ เพราะชื่อไม่อยู่ใน productSheetNames_ แล้ว)
  legacy.forEach(function (s) {
    let base = 'สำรอง_' + s.getName(), name = base, k = 2;
    while (ss.getSheetByName(name)) { name = base + '_' + (k++); }
    s.setName(name);
    try { s.hideSheet(); } catch (eH) {}
  });
  clearShopCache_();
  const parts = Object.keys(bySheet).map(function (n) { return n + ' ' + bySheet[n].length; });
  ui.alert('✅ แยกชีตรายชั้นเสร็จ ย้าย ' + moved + '/' + total + ' แถว\n' + parts.join(' • ') +
    '\n\nชีตเดิมถูกซ่อนไว้ในชื่อ "สำรอง_..." — ตรวจความถูกต้องสัก 2-3 วันแล้วค่อยลบได้ค่ะ');
}

// 🔤 เรียงชื่อวิชา ก-ฮ ทุกชีตสินค้า (กดซ้ำได้เสมอ เช่น หลังแก้ข้อมูลในชีตเองโดยตรง)
function sortAllProductSheets() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let n = 0;
  GRADE_SHEETS.concat([SHEET_VIP_WORK, SHEET_OTHER]).forEach(function (name) {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() > 2) { sortProductSheet_(s); n++; }
  });
  clearShopCache_();
  try { SpreadsheetApp.getUi().alert('🔤 เรียงชื่อวิชา ก-ฮ แล้ว ' + n + ' ชีต (ย้ายทั้งแถว ข้อมูลตรงกันครบทุกช่อง)'); } catch (e) {}
}


/* ============================================================
 *  📢 แจ้งอัปเดตไฟล์ถึงลูกค้าที่เคยซื้อ (สถานะ "ส่งแผนเสร็จสิ้น")
 *  - คลิกแถววิชาในชีตสินค้า → กดเมนู → พิมพ์ข้อความอัปเดต
 *  - กันโควตา Gmail ไว้ให้งานขาย RESERVE_MAIL_QUOTA ฉบับ/วันเสมอ
 *  - เกินโควตา → เข้าชีตคิว ทยอยส่งอัตโนมัติทุกคืน ~03:00 จนครบ
 * ============================================================ */
const SHEET_UPDATE_QUEUE = '📢 คิวแจ้งอัปเดต';
const RESERVE_MAIL_QUOTA = 30;   // โควตาที่กันไว้ให้อีเมลออเดอร์/ส่งแผนเสมอ

function updateQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_UPDATE_QUEUE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_UPDATE_QUEUE);
    sh.getRange(1, 1, 1, 9).setValues([[
      'เวลาเข้าคิว', 'อีเมล', 'ชื่อลูกค้า', 'วิชา', 'ชั้น', 'ข้อความอัปเดต', 'ลิงก์ Drive', 'สถานะ', 'เวลาส่ง'
    ]]).setFontWeight('bold').setBackground('#fef9c3');
    sh.setFrozenRows(1);
  }
  return sh;
}

// เมนู: แจ้งอัปเดตวิชาในแถวที่เลือกอยู่
function notifyUpdateCurrentRow() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const okSheets = GRADE_SHEETS.concat([SHEET_VIP_WORK, SHEET_OTHER]);
  if (okSheets.indexOf(sh.getName()) < 0) {
    ui.alert('เมนูนี้ใช้ในชีตสินค้า (ป.1–ม.6 / VIP & ระบบงาน / อื่นๆ)\nคลิกแถววิชาที่อัปเดตไฟล์ก่อน แล้วกดเมนูอีกครั้งค่ะ');
    return;
  }
  const row = sh.getActiveCell().getRow();
  if (row < 2 || row > sh.getLastRow()) { ui.alert('คลิกแถววิชาก่อนค่ะ'); return; }
  const vip = isVipSheet_(sh.getName());
  const r = sh.getRange(row, 1, 1, vip ? PV_COUNT : P_COUNT).getValues()[0];
  const id = String(r[(vip ? PV.ID : P.ID) - 1] || '').trim();
  const subject = String(r[(vip ? PV.SUBJECT : P.SUBJECT) - 1] || '').trim();
  const grade = vip ? '' : String(r[P.GRADE - 1] || '').trim();
  const driveLink = String(r[(vip ? PV.DRIVE_LINK : P.DRIVE_LINK) - 1] || '').trim();
  if (!id || !subject) { ui.alert('แถวนี้ไม่มีข้อมูลวิชา'); return; }

  const ask = ui.prompt('📢 แจ้งอัปเดต: ' + subject + ' ' + grade,
    'พิมพ์ข้อความอัปเดตที่จะแจ้งลูกค้า เช่น\n"เพิ่มแผนเทอม 2 อีก 40 แผน และแก้ไขใบงานหน่วยที่ 3"',
    ui.ButtonSet.OK_CANCEL);
  if (ask.getSelectedButton() !== ui.Button.OK) return;
  const message = String(ask.getResponseText() || '').trim();
  if (!message) { ui.alert('ยังไม่ได้พิมพ์ข้อความอัปเดต'); return; }

  // หาลูกค้าที่เคยซื้อวิชานี้ (เฉพาะส่งแผนเสร็จสิ้น) — จับคู่ด้วยรหัสวิชา, ออเดอร์เก่าใช้ชื่อวิชา+ชั้น
  const seen = {};
  const recipients = [];
  readOrders_().forEach(function (o) {
    if (o.status !== 'ส่งแผนเสร็จสิ้น') return;
    const hit = (o.items || []).some(function (it) {
      return String(it.id || '') === id ||
        (String(it.subject || '').trim() === subject && String(it.grade || '').trim() === grade);
    });
    if (!hit) return;
    const em = String(o.email || '').trim().toLowerCase();
    if (!em || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em) || seen[em]) return;
    seen[em] = 1;
    recipients.push({ email: em, customer: o.customer || '' });
  });
  if (!recipients.length) { ui.alert('ไม่พบลูกค้าที่เคยซื้อ ' + subject + ' ' + grade + ' (สถานะส่งแผนเสร็จสิ้น)'); return; }

  const quota = MailApp.getRemainingDailyQuota();
  const sendableToday = Math.max(0, quota - RESERVE_MAIL_QUOTA);
  const nowCount = Math.min(recipients.length, sendableToday);
  const res = ui.alert('📢 ยืนยันการแจ้งอัปเดต',
    'วิชา: ' + subject + ' ' + grade +
    '\nลูกค้าที่จะแจ้ง: ' + recipients.length + ' คน (ตัดอีเมลซ้ำแล้ว)' +
    '\nโควตา Gmail วันนี้เหลือ: ' + quota + ' (กันไว้ให้งานขาย ' + RESERVE_MAIL_QUOTA + ')' +
    '\n→ ส่งทันที ' + nowCount + ' คน' +
    (recipients.length > nowCount ? ' / อีก ' + (recipients.length - nowCount) + ' คนเข้าคิว ส่งอัตโนมัติคืนละรอบ (~03:00) จนครบ' : '') +
    '\n\nดำเนินการต่อไหมคะ?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  // ลงคิวทั้งหมดก่อน แล้วให้ตัวประมวลผลส่งเท่าที่โควตาอนุญาต — ที่เหลือรอ trigger กลางคืน
  const qsh = updateQueueSheet_();
  const now = ymd_(new Date());
  const rows = recipients.map(function (rc) {
    return [now, rc.email, rc.customer, subject, grade, message, driveLink, 'รอส่ง', ''];
  });
  qsh.getRange(qsh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);

  const sent = processUpdateQueue_();
  ui.alert('✅ ส่งแจ้งอัปเดตแล้ว ' + sent + ' คน' +
    (recipients.length > sent ? '\n⏳ อีก ' + (recipients.length - sent) + ' คนอยู่ในชีต "' + SHEET_UPDATE_QUEUE + '" ระบบจะทยอยส่งให้ทุกคืนจนครบค่ะ' : '') +
    '\n\n(ถ้ายังไม่เคยเปิด trigger กลางคืน: เมนู 🛠 เครื่องมือระบบ → ⏰ เปิดใช้รายงานเช้า 8 โมง)');
}

// ตัวประมวลผลคิว — เรียกจากเมนูตอนสร้าง และจาก trigger กลางคืน
function processUpdateQueue() { try { processUpdateQueue_(); } catch (e) { Logger.log('processUpdateQueue: ' + e); } }
function processUpdateQueue_() {
  const sh = updateQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  let budget = Math.max(0, MailApp.getRemainingDailyQuota() - RESERVE_MAIL_QUOTA);
  if (!budget) return 0;
  const m = settingsMap_();
  const pageName = m.pageName || PAGE_NAME;
  const vals = sh.getRange(2, 1, last - 1, 9).getValues();
  let sent = 0;
  for (let i = 0; i < vals.length && budget > 0; i++) {
    if (String(vals[i][7]) !== 'รอส่ง') continue;
    const email = String(vals[i][1] || '');
    const customer = String(vals[i][2] || '');
    const subject = String(vals[i][3] || '');
    const grade = String(vals[i][4] || '');
    const message = String(vals[i][5] || '');
    const driveLink = String(vals[i][6] || '');
    try {
      MailApp.sendEmail({
        to: email,
        subject: '📢 อัปเดตไฟล์ ' + subject + ' ' + grade + ' - ' + pageName,
        htmlBody: buildUpdateEmail_(customer, subject, grade, message, driveLink, pageName),
        name: pageName
      });
      sh.getRange(i + 2, 8, 1, 2).setValues([['ส่งแล้ว', ymd_(new Date())]]);
      sent++; budget--;
    } catch (e) {
      sh.getRange(i + 2, 8).setValue('ผิดพลาด: ' + String(e.message || e).slice(0, 80));
    }
  }
  return sent;
}

function buildUpdateEmail_(customer, subject, grade, message, driveLink, pageName) {
  return '<div style="font-family:Arial,sans-serif; max-width:560px; margin:0 auto; border:1px solid #e0f2fe; border-radius:12px; overflow:hidden">' +
    '<div style="background:linear-gradient(135deg,#0284c7,#38bdf8); color:#fff; padding:18px 22px;">' +
      '<div style="font-size:18px; font-weight:700;">📢 มีการอัปเดตไฟล์ที่คุณเคยสั่งซื้อ</div>' +
      '<div style="font-size:13px; opacity:.9; margin-top:2px;">' + pageName + '</div></div>' +
    '<div style="padding:20px 22px; color:#334155; font-size:14px; line-height:1.7;">' +
      'สวัสดีค่ะ' + (customer ? ' คุณ' + customer : '') + ' 🙏<br><br>' +
      'ไฟล์ <b>' + subject + ' ' + grade + '</b> ที่คุณเคยสั่งซื้อ มีการอัปเดตดังนี้ค่ะ' +
      '<div style="background:#f0f9ff; border-left:4px solid #38bdf8; border-radius:8px; padding:12px 14px; margin:12px 0;">' + message + '</div>' +
      '✅ เข้าดูไฟล์ล่าสุดได้จาก <b>ลิงก์เดิม</b> ที่เคยได้รับ — สิทธิ์การเข้าถึงของคุณยังอยู่ครบ ไม่มีค่าใช้จ่ายเพิ่มค่ะ' +
      (driveLink ? '<div style="text-align:center; margin-top:16px;"><a href="' + driveLink + '" style="display:inline-block; background:linear-gradient(135deg,#0284c7,#0ea5e9); color:#fff; text-decoration:none; padding:11px 26px; border-radius:8px; font-weight:700; font-size:14px;">📂 เปิดไฟล์งาน</a></div>' : '') +
      '<div style="color:#94a3b8; font-size:12px; margin-top:18px;">ขอบคุณที่ไว้วางใจ ' + pageName + ' ค่ะ 💛</div>' +
    '</div></div>';
}

/* ============================================================
 *  🏷️ จัดกลุ่มวิชาอัตโนมัติ — เติมเฉพาะช่องที่ว่าง กดซ้ำได้ ไม่ทับที่แก้มือ
 * ============================================================ */
function autoGroupProducts() {
  const ui = SpreadsheetApp.getUi();
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let filled = 0, blank = [];
  GRADE_SHEETS.concat([SHEET_OTHER]).forEach(function (name) {   // ชีต VIP & ระบบงาน ไม่ใช้กลุ่มวิชา
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last < 2) return;
    const rng = sh.getRange(2, 1, last - 1, P_COUNT);
    const vals = rng.getValues();
    let changed = false;
    vals.forEach(function (r) {
      if (!String(r[P.ID - 1] || '').trim()) return;
      if (String(r[P.GROUP - 1] || '').trim()) return;               // มีกลุ่มแล้ว — ไม่แตะ
      const g = autoGroupFor_(r[P.CURRICULUM - 1], r[P.SUBJECT - 1]);
      if (g) { r[P.GROUP - 1] = g; filled++; changed = true; }
      else blank.push(String(r[P.SUBJECT - 1]) + ' (' + name + ')');
    });
    if (changed) {
      rng.setValues(vals);
      rng.offset(0, P.ADDED - 1, vals.length, 1).setNumberFormat('@');
    }
  });
  clearShopCache_();
  ui.alert('🏷️ จัดกลุ่มวิชาแล้ว ' + filled + ' วิชา' +
    (blank.length ? '\n\n⚠️ จับกลุ่มไม่ได้ ' + blank.length + ' วิชา (เติมเองที่คอลัมน์ "กลุ่มวิชา"):\n• ' +
      blank.slice(0, 12).join('\n• ') + (blank.length > 12 ? '\n…และอีก ' + (blank.length - 12) + ' วิชา' : '') : '\nครบทุกวิชาแล้วค่ะ ✅'));
}

/* ============================================================
 *  🧱 ปรับโครงชีต VIP & ระบบงาน (ครั้งเดียว)
 *  ชีตนี้ยกมาจากระบบรุ่นเก่า (10 คอลัมน์ ไม่มีชั้น/หลักสูตร/ชม.)
 *  → อ่านแบบอิงชื่อหัวคอลัมน์ + ซ่อมแถวเหลื่อมด้วยรูปแบบข้อมูล
 *  → เขียนกลับเป็นโครงมาตรฐาน 14 คอลัมน์ / ชีตเดิมเก็บเป็น สำรอง_
 * ============================================================ */
function fixVipWorkSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VIP_WORK);
  if (!sh || sh.getLastRow() < 2) { ui.alert('ไม่พบข้อมูลในชีต "' + SHEET_VIP_WORK + '"'); return; }

  // โครงย่อที่ต้องการ: ช่อง 2 = ชื่อวิชา/รายการ, ช่อง 3 = ราคา → ถ้าเป็นแบบนี้แล้วไม่ต้องทำซ้ำ
  const h2 = String(sh.getRange(1, 2).getValue()).trim();
  const h3 = String(sh.getRange(1, 3).getValue()).trim();
  if (h2.indexOf('ชื่อวิชา') === 0 && h3.indexOf('ราคา') === 0) {
    ui.alert('ชีต "' + SHEET_VIP_WORK + '" เป็นโครงย่อ 10 คอลัมน์อยู่แล้วค่ะ'); return;
  }

  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const col = function (names) {
    for (let i = 0; i < header.length; i++) {
      for (let k = 0; k < names.length; k++) {
        if (header[i].indexOf(names[k]) === 0) return i;
      }
    }
    return -1;
  };
  const ix = {
    id: col(['ID']), subject: col(['ชื่อวิชา']), price: col(['ราคา']),
    drive: col(['ลิงก์ Drive', 'ลิงก์Drive']), example: col(['ลิงก์ตัวอย่าง']),
    detail: col(['รายละเอียด']), footer: col(['ข้อความท้าย']),
    active: col(['เปิดขาย']), added: col(['วันที่เพิ่ม']), category: col(['ประเภท'])
  };
  if (ix.id < 0 || ix.subject < 0) { ui.alert('อ่านหัวตารางไม่สำเร็จ — หัวคอลัมน์ ID/ชื่อวิชา ไม่พบ'); return; }

  const raw = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  const isUrl = function (v) { return /^https?:\/\//.test(String(v).trim()); };
  const isBool = function (v) { return /^(TRUE|FALSE)$/i.test(String(v).trim()); };
  const pick = function (r, i) { return i >= 0 ? String(r[i] == null ? '' : r[i]).trim() : ''; };
  const rows = [], notes = [];
  raw.forEach(function (r, i) {
    const id = pick(r, ix.id);
    if (!id) return;
    let subject = pick(r, ix.subject), price = pick(r, ix.price);
    let drive = pick(r, ix.drive), example = pick(r, ix.example);
    let detail = pick(r, ix.detail), footer = pick(r, ix.footer);
    let active = pick(r, ix.active), added = pick(r, ix.added), category = pick(r, ix.category);
    if (isUrl(active)) {   // ซ่อมแถวเหลื่อม: ลิงก์ไปโผล่ช่องเปิดขาย
      if (!example) example = active; else if (!drive) drive = active;
      active = 'TRUE';
      notes.push('แถว ' + (i + 2) + ': พบลิงก์ในช่องเปิดขาย — จัดเข้าช่องลิงก์ให้แล้ว');
    }
    if (!isBool(active)) { active = 'TRUE'; notes.push('แถว ' + (i + 2) + ' (' + subject + '): ช่องเปิดขายว่าง — ตั้งเป็น TRUE ให้'); }
    if (!category) category = /vip/i.test(subject) ? 'กลุ่ม VIP' : 'ระบบงาน';
    if (!Number(price)) notes.push('แถว ' + (i + 2) + ' (' + subject + '): ไม่มีราคา — จะยังไม่ขึ้นหน้าร้าน');
    rows.push([id, subject, Number(price) || '', drive, example, detail, footer || STD_FOOTER,
      active.toUpperCase(), added || ymd_(new Date()), category]);
  });
  if (!rows.length) { ui.alert('ไม่พบแถวข้อมูลที่ใช้ได้'); return; }

  const res = ui.alert('🧱 ปรับโครงชีต VIP & ระบบงาน (โครงย่อ 10 คอลัมน์)',
    'จะจัดข้อมูล ' + rows.length + ' รายการเข้าโครงย่อ: ID • ชื่อวิชา/รายการ • ราคา • ลิงก์ Drive • ลิงก์ตัวอย่าง • รายละเอียด • ข้อความท้าย • เปิดขาย • วันที่เพิ่ม • ประเภท' +
    '\n(ตัดคอลัมน์ ระดับชั้น/หลักสูตร/ชม./กลุ่มวิชา ออกตามที่กำหนด)' +
    '\nชีตเดิมเก็บเป็น "สำรอง_' + SHEET_VIP_WORK + '" (ไม่ลบ)' +
    (notes.length ? '\n\nจุดที่ระบบซ่อม/ควรตรวจ ' + notes.length + ' จุด (ดูสรุปหลังเสร็จ)' : '') +
    '\n\nดำเนินการต่อไหมคะ?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  let base = 'สำรอง_' + SHEET_VIP_WORK, bak = base, k = 2;
  while (ss.getSheetByName(bak)) { bak = base + '_' + (k++); }
  sh.setName(bak);
  try { sh.hideSheet(); } catch (e) {}

  const ns = ss.insertSheet(SHEET_VIP_WORK);
  ns.getRange(1, 1, 1, PV_COUNT).setValues([PRODUCT_HEADER_VIP]).setFontWeight('bold').setBackground('#fef3c7');
  ns.setFrozenRows(1);
  ns.getRange(2, 1, rows.length, PV_COUNT).setValues(rows);
  ns.getRange(2, PV.ADDED, rows.length, 1).setNumberFormat('@');
  sortProductSheet_(ns);
  clearShopCache_();
  ui.alert('✅ ปรับโครงเสร็จ ' + rows.length + ' รายการ (เปิดขาย TRUE ครบ)' +
    '\n\n👉 ขั้นสุดท้าย: กดเมนู 🚀 เผยแพร่หน้าร้านตอนนี้ (Edge) แล้ว hard refresh หน้าเว็บ — ชิป 💼/👑 จะแสดงทันที' +
    (notes.length ? '\n\n📌 จุดที่ควรตรวจ:\n• ' + notes.slice(0, 10).join('\n• ') + (notes.length > 10 ? '\n…อีก ' + (notes.length - 10) + ' จุด' : '') : ''));
}

/* ============================================================
 *  🩺 ตรวจสุขภาพตัวอย่างทุกวิชา — รัน logic เดียวกับที่ลูกค้าเจอจริง
 *  ผล: ✅ ปกติ / 🔍 หาไม่เจอ / 📂 โฟลเดอร์ว่าง — เขียนรายงานลงชีต
 *  แถมซ่อมสิทธิ์ PDF เป็นสาธารณะให้อัตโนมัติระหว่างตรวจ (ในตัว collect)
 *  รันต่อจากเดิมได้ถ้าครั้งก่อนหมดเวลา (ข้ามรายการที่ตรวจแล้ววันนี้)
 * ============================================================ */
const SHEET_PREVIEW_HEALTH = '🩺 สุขภาพตัวอย่าง';
function previewHealthCheck() {
  const ui = SpreadsheetApp.getUi();
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = readProducts_().filter(function (p) { return p.active && p.price > 0; });
  if (!products.length) { ui.alert('ไม่มีวิชาที่เปิดขาย'); return; }

  let sh = ss.getSheetByName(SHEET_PREVIEW_HEALTH);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PREVIEW_HEALTH);
    sh.getRange(1, 1, 1, 7).setValues([['วันที่ตรวจ', 'ID', 'วิชา', 'ชั้น', 'หลักสูตร', 'ผลตรวจ', 'หมายเหตุ']])
      .setFontWeight('bold').setBackground('#dcfce7');
    sh.setFrozenRows(1);
  }
  const today = ymd_(new Date()).slice(0, 10);
  const done = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (String(r[0]).slice(0, 10) === today) done[String(r[1])] = 1;
    });
  }
  let todo = products.filter(function (p) { return !done[p.id]; });
  if (!todo.length) {   // ตรวจครบรอบวันนี้แล้ว → เริ่มรอบใหม่ (ล้างรายงานเก่า)
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    todo = products;
  }
  const res = ui.alert('🩺 ตรวจสุขภาพตัวอย่าง',
    'จะตรวจ ' + todo.length + ' วิชา (จากทั้งหมด ' + products.length + ')' +
    '\nใช้เวลาประมาณ 1-2 วินาที/วิชา — ถ้าหมดเวลากลางทาง กดเมนูซ้ำเพื่อตรวจต่อจากเดิมได้' +
    '\nระหว่างตรวจ ระบบเปิดสิทธิ์ PDF เป็นสาธารณะให้อัตโนมัติ\n\nเริ่มเลยไหมคะ?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const t0 = Date.now();
  let ok = 0, notFound = 0, empty = 0, checked = 0;
  const out = [];
  for (let i = 0; i < todo.length; i++) {
    if (Date.now() - t0 > 250000) break;   // กันชนเพดาน 6 นาที — เหลือเวลาเขียนรายงาน
    const p = todo[i];
    let status = '', note = '';
    try {
      const r = getPreviewData_(p.subject, p.grade, p.curriculum, String(p.hours || ''));
      if (!r.ok) { status = '🔍 หาไม่เจอ'; note = r.error || ''; notFound++; }
      else if (!(r.pdfs || []).length && !(r.htmls || []).length) { status = '📂 โฟลเดอร์ว่าง'; note = 'เจอโฟลเดอร์แต่ไม่มีไฟล์ PDF/HTML'; empty++; }
      else { status = '✅ ปกติ'; note = (r.pdfs || []).length + ' PDF, ' + (r.htmls || []).length + ' สื่อ HTML'; ok++; }
    } catch (e) { status = '⚠️ ตรวจไม่สำเร็จ'; note = String(e.message || e).slice(0, 100); }
    out.push([ymd_(new Date()), p.id, p.subject, p.grade, p.curriculum, status, note]);
    checked++;
  }
  if (out.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, out.length, 7).setValues(out);
  }
  const remain = todo.length - checked;
  ui.alert('🩺 ตรวจแล้ว ' + checked + ' วิชา: ✅ ' + ok + ' • 🔍 หาไม่เจอ ' + notFound + ' • 📂 ว่าง ' + empty +
    (remain > 0 ? '\n\n⏳ เหลืออีก ' + remain + ' วิชา — กดเมนูนี้ซ้ำเพื่อตรวจต่อค่ะ' : '\n\nดูรายละเอียดรายวิชาในชีต "' + SHEET_PREVIEW_HEALTH + '"' +
      '\n🔍 = สร้าง/แก้ชื่อโฟลเดอร์ใน Drive ให้มีชั้น+หลักสูตร+ชื่อวิชาตรงกับชีต\n📂 = เติมไฟล์ตัวอย่างในโฟลเดอร์'));
}


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // เมนูฉบับกระชับ — เหลือเฉพาะงานที่ใช้จริงประจำวัน ที่เหลือเก็บใน "เครื่องมือระบบ"
  // (ฟังก์ชันเก่าที่ถอดออกจากเมนูยังอยู่ในโค้ดครบ สั่งรันได้จากหน้า Apps Script หากจำเป็น)
  // เมนูจัดหมวดหมู่: [งานแถวที่เลือก] → [ดูออเดอร์] → [เครื่องมือระบบ]
  // งานติดตั้งครั้งเดียวที่ทำจบแล้ว (เติมวิชา ม. / แยกชีตรายชั้น) ถอดออกจากเมนู
  // แต่ฟังก์ชันยังอยู่ในโค้ด รันได้จากหน้า Apps Script หากต้องใช้อีก
  ui.createMenu('🎓 ครูพร้อมสอน')
    .addItem('🔁 ส่งแผนแถวนี้อีกครั้ง', 'resendCurrentRow')
    .addItem('🚫 ยกเลิก + ถอนสิทธิ์ Drive แถวนี้', 'cancelAndRevokeCurrentRow')
    .addItem('📢 แจ้งอัปเดตวิชานี้ถึงลูกค้า', 'notifyUpdateCurrentRow')
    .addSeparator()
    .addItem('⏳ ค้างส่ง (ชำระแล้ว รอส่งแผน)', 'viewOrdersPendingSend')
    .addSubMenu(ui.createMenu('📦 ดูออเดอร์')
      .addItem('📆 วันนี้', 'viewOrdersToday')
      .addItem('🗓️ เดือนนี้', 'viewOrdersThisMonth')
      .addItem('🔄 แสดงทั้งหมด (ล้างตัวกรอง)', 'clearOrdersFilter'))
    .addItem('📊 รีเฟรชรายงาน (Dashboard + ประวัติลูกค้า)', 'refreshReports')
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠 เครื่องมือระบบ (นานๆ ใช้ที)')
      // — จัดการวิชา —
      .addItem('🏷️ จัดกลุ่มวิชาอัตโนมัติ (68/51)', 'autoGroupProducts')
      .addItem('🩺 ตรวจสุขภาพตัวอย่างทุกวิชา (+ซ่อมสิทธิ์)', 'previewHealthCheck')
      .addItem('🔗 เติมลิงก์ตัวอย่าง (Preview) ทุกวิชา', 'setupPreviewLinks')
      .addItem('🔤 เรียงชื่อวิชา ก-ฮ ทุกชีต', 'sortAllProductSheets')
      .addSeparator()
      // — โครงสร้าง/ข้อมูล —
      .addItem('🧱 ปรับโครงชีต VIP & ระบบงาน (ครั้งเดียว)', 'fixVipWorkSheet')
      .addItem('🧰 ตรวจ/สร้างชีตทั้งหมด (ซ่อมโครงสร้าง)', 'repairAllSheets')
      .addItem('💾 สำรองข้อมูลเดี๋ยวนี้', 'backupNowMenu')
      .addSeparator()
      // — Edge / อัตโนมัติ —
      .addItem('🚀 เผยแพร่หน้าร้านตอนนี้ (Edge)', 'publishShopSnapshotMenu')
      .addItem('🚀 ตั้งค่า Edge / GitHub', 'setupEdgePublish')
      .addItem('⚡ เปิดระบบอัตโนมัติทั้งหมด (ครั้งแรก/ย้ายไฟล์)', 'setupAllAutomation'))
    .addToUi();
  try { applyOrdersSheetSetup_(); refreshOrderColors(); } catch (e) { Logger.log('onOpen: ' + e); }
}

// dropdown สถานะ + สีสถานะ + เปิด Filter ที่หัวตาราง
function applyOrdersSheetSetup_() {
  const sh = getSheet_(SHEET_ORDERS);
  // ซ่อนคอลัมน์ที่ระบบใช้ภายในแต่ไม่ต้องเห็น: เบอร์โทร + รายการ (JSON)
  try { sh.hideColumns(O.PHONE_FB); sh.hideColumns(O.ITEMS_JSON); } catch (eHide) {}
  const maxRows = Math.max(sh.getLastRow(), 50);
  const statusList = [STATUSES.PENDING, STATUSES.PAID, STATUSES.APPROVED, STATUSES.COMPLETED, STATUSES.CANCELLED];
  // ล้าง dropdown เก่าที่อาจค้างผิดคอลัมน์ (เช่นไปโผล่ช่องลิงก์สลิปหลังแทรกคอลัมน์ใหม่)
  sh.getRange(2, 1, maxRows - 1, O_COUNT).clearDataValidations();
  sh.getRange(2, O.STATUS, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(statusList, true).setAllowInvalid(true).build());
  // สีตัวอักษรคอลัมน์สถานะตามค่า (ใช้ conditional formatting)
  const colors = {};
  colors[STATUSES.PENDING]   = '#a16207';
  colors[STATUSES.PAID]      = '#1e40af';
  colors[STATUSES.APPROVED]  = '#6b21a8';
  colors[STATUSES.COMPLETED] = '#15803d';
  colors[STATUSES.CANCELLED] = '#b91c1c';
  const rules = statusList.map(function (st) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(st).setFontColor(colors[st]).setBold(true)
      .setRanges([sh.getRange(2, O.STATUS, maxRows - 1, 1)]).build();
  });
  sh.setConditionalFormatRules(rules);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), O_COUNT).createFilter();
}

// สีพื้นหลังสลับโทนพาสเทลตามวันที่ — วันเดียวกันสีเดียวกัน
function refreshOrderColors() {
  const sh = getSheet_(SHEET_ORDERS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const PALETTE = ['#fff7ed', '#eff6ff', '#f0fdf4', '#fdf2f8', '#f5f3ff', '#fefce8', '#ecfeff', '#fef2f2'];
  const dates = sh.getRange(2, O.TIMESTAMP, last - 1, 1).getValues();
  const dayIndex = {};
  let n = 0;
  const bgs = dates.map(function (r) {
    const day = String(r[0] || '').slice(0, 10);
    if (!(day in dayIndex)) dayIndex[day] = n++;
    return new Array(O_COUNT).fill(PALETTE[dayIndex[day] % PALETTE.length]);
  });
  sh.getRange(2, 1, last - 1, O_COUNT).setBackgrounds(bgs);
}

function filterCurrentMonth() {
  filterOrdersByMonth_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM'));
}
function promptFilterMonth() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('📅 เลือกเดือน',
    'พิมพ์เดือนรูปแบบ ปี-เดือน เช่น 2026-05\nหรือเว้นว่างแล้วกดตกลง = ล้างตัวกรอง แสดงทั้งหมด',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const m = String(res.getResponseText() || '').trim();
  if (!m) { clearOrdersFilter(); return; }
  if (!/^\d{4}-\d{2}$/.test(m)) { ui.alert('รูปแบบไม่ถูกต้อง ตัวอย่าง: 2026-05'); return; }
  filterOrdersByMonth_(m);
}
function filterOrdersByMonth_(yyyymm) {
  const sh = getSheet_(SHEET_ORDERS);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), O_COUNT).createFilter();
  sh.getFilter().setColumnFilterCriteria(O.TIMESTAMP,
    SpreadsheetApp.newFilterCriteria().whenTextStartsWith(yyyymm).build());
  SpreadsheetApp.getActiveSpreadsheet().toast('กำลังแสดงเฉพาะเดือน ' + yyyymm, '📅 ตัวกรอง', 4);
}
function clearOrdersFilter() {
  const sh = getSheet_(SHEET_ORDERS);
  const f = sh.getFilter();
  if (f) { f.removeColumnFilterCriteria(O.TIMESTAMP); f.removeColumnFilterCriteria(O.STATUS); }
  SpreadsheetApp.getActiveSpreadsheet().toast('แสดงออเดอร์ทั้งหมดแล้ว', '📅 ตัวกรอง', 3);
}

// ── เมนูกดครั้งเดียว: วันนี้ / เดือนนี้ / ปีนี้ (ไม่ต้องกรอกอะไร) ──
function viewOrdersToday()     { filterOrdersByPrefix_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'), 'วันนี้'); }
function viewOrdersThisMonth() { filterOrdersByPrefix_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM'), 'เดือนนี้'); }
function viewOrdersThisYear()  { filterOrdersByPrefix_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy'), 'ปีนี้'); }

// ⏳ แสดงเฉพาะออเดอร์ค้างส่ง (ชำระเงินแล้ว/อนุมัติ แต่ยังไม่ส่งเสร็จ)
function viewOrdersPendingSend() {
  const sh = getSheet_(SHEET_ORDERS);
  const orders = readOrders_();
  const cnt = orders.filter(function (o) { return o.status === STATUSES.PAID || o.status === STATUSES.APPROVED; }).length;

  // ไม่มีงานค้างส่ง → แจ้งเฉยๆ ไม่ต้องซ่อนตาราง (กันเข้าใจผิดว่าข้อมูลหาย)
  if (!cnt) {
    SpreadsheetApp.getUi().alert('⏳ ค้างส่ง', '🎉 ไม่มีออเดอร์ค้างส่งตอนนี้\n(ทุกรายการที่ชำระแล้วถูกส่งเสร็จสิ้นหมดแล้ว)', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), O_COUNT).createFilter();
  sh.getFilter().removeColumnFilterCriteria(O.TIMESTAMP);
  // ใช้วิธี "ซ่อนค่าอื่นทั้งหมด" (กลไกเดียวกับติ๊กกรองด้วยมือ — ชัวร์สุดสำหรับตัวกรองชีตปกติ)
  const keep = {}; keep[STATUSES.PAID] = 1; keep[STATUSES.APPROVED] = 1;
  const hidden = { '': 1 };
  orders.forEach(function (o) { const s = String(o.status || ''); if (!keep[s]) hidden[s] = 1; });
  sh.getFilter().setColumnFilterCriteria(O.STATUS,
    SpreadsheetApp.newFilterCriteria().setHiddenValues(Object.keys(hidden)).build());
  SpreadsheetApp.getActiveSpreadsheet().toast('ค้างส่ง ' + cnt + ' รายการ (กด 🔄 แสดงทั้งหมด เพื่อล้างตัวกรอง)', '⏳ ชำระแล้ว รอส่งแผน', 5);
}

// ⬇️ กระโดดไปออเดอร์ล่าสุด (แถวล่างสุดของชีต)
function gotoLastOrder() {
  const sh = getSheet_(SHEET_ORDERS);
  const r = Math.max(sh.getLastRow(), 2);
  sh.activate();
  sh.setActiveRange(sh.getRange(r, 1));
  SpreadsheetApp.getActiveSpreadsheet().toast('ไปที่ออเดอร์ล่าสุด (แถว ' + r + ')', '⬇️', 3);
}

// 📊 รีเฟรชรายงานทั้งหมดในคลิกเดียว (Dashboard + ประวัติลูกค้า)
function refreshReports() {
  generateDashboard();
  generateCustomerSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('รีเฟรช Dashboard และประวัติลูกค้าแล้ว', '📊 รายงาน', 4);
}

// ───────── ดูออเดอร์ตามช่วง: รายวัน / รายเดือน / รายปี ─────────
// ชีต Orders เก็บวันที่เป็นข้อความขึ้นต้น yyyy-MM-dd จึงกรองด้วย "ขึ้นต้นด้วย" ได้ตรงๆ
function viewOrdersDaily() {
  const ui = SpreadsheetApp.getUi();
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const res = ui.prompt('📆 ดูรายวัน', 'พิมพ์วันที่รูปแบบ ปี-เดือน-วัน เช่น ' + today + '\n(เว้นว่าง = วันนี้)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const v = String(res.getResponseText() || '').trim() || today;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) { ui.alert('รูปแบบไม่ถูกต้อง ตัวอย่าง: ' + today); return; }
  filterOrdersByPrefix_(v, 'รายวัน');
}
function viewOrdersMonthly() {
  const ui = SpreadsheetApp.getUi();
  const thisMonth = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
  const res = ui.prompt('🗓️ ดูรายเดือน', 'พิมพ์เดือนรูปแบบ ปี-เดือน เช่น ' + thisMonth + '\n(เว้นว่าง = เดือนนี้)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const v = String(res.getResponseText() || '').trim() || thisMonth;
  if (!/^\d{4}-\d{2}$/.test(v)) { ui.alert('รูปแบบไม่ถูกต้อง ตัวอย่าง: ' + thisMonth); return; }
  filterOrdersByPrefix_(v, 'รายเดือน');
}
function viewOrdersYearly() {
  const ui = SpreadsheetApp.getUi();
  const thisYear = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy');
  const res = ui.prompt('📅 ดูรายปี', 'พิมพ์ปี (ค.ศ.) เช่น ' + thisYear + '\n(เว้นว่าง = ปีนี้)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const v = String(res.getResponseText() || '').trim() || thisYear;
  if (!/^\d{4}$/.test(v)) { ui.alert('รูปแบบไม่ถูกต้อง ตัวอย่าง: ' + thisYear); return; }
  filterOrdersByPrefix_(v, 'รายปี');
}

// กรองตารางตาม prefix วันที่ + เด้งสรุปยอดของช่วงนั้น
function filterOrdersByPrefix_(prefix, label) {
  const sh = getSheet_(SHEET_ORDERS);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), O_COUNT).createFilter();
  sh.getFilter().setColumnFilterCriteria(O.TIMESTAMP,
    SpreadsheetApp.newFilterCriteria().whenTextStartsWith(prefix).build());

  const orders = readOrders_().filter(function (o) { return String(o.timestamp || '').indexOf(prefix) === 0; });
  let net = 0; const byStatus = {};
  orders.forEach(function (o) {
    net += Number(o.netTotal || 0);
    const s = String(o.status || '(ไม่ระบุสถานะ)');
    byStatus[s] = (byStatus[s] || 0) + 1;
  });
  const statusLines = Object.keys(byStatus).map(function (k) { return '   • ' + k + ': ' + byStatus[k] + ' รายการ'; });

  SpreadsheetApp.getUi().alert('📅 สรุป' + label + ' — ' + prefix,
    'จำนวนออเดอร์: ' + orders.length + ' รายการ\n' +
    'ยอดสุทธิรวม: ' + Number(net).toLocaleString() + ' บาท\n\n' +
    'แยกตามสถานะ:\n' + (statusLines.join('\n') || '   — ไม่มีออเดอร์ในช่วงนี้') +
    '\n\n👇 ตารางด้านล่างกรองแสดงเฉพาะ' + label + 'นี้แล้ว (กดเมนู 🔄 แสดงทั้งหมด เพื่อล้างตัวกรอง)',
    SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActiveSpreadsheet().toast('แสดงเฉพาะ ' + prefix + ' (' + orders.length + ' รายการ)', '📅 ' + label, 4);
}

// 🔎 ตรวจการจับคู่ของแถวที่เลือก — บอกว่าแต่ละวิชาในออเดอร์จับคู่กับแผนไหน มีลิงก์ไหม (อ่านอย่างเดียว ไม่ส่ง)
function diagnoseCurrentRow() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_ORDERS) { ui.alert('กรุณาเลือกแถวในชีต Orders ก่อน'); return; }
  const row = sh.getActiveCell().getRow();
  if (row < 2) { ui.alert('กรุณาเลือกแถวข้อมูล (ไม่ใช่หัวตาราง)'); return; }
  const o = readOrderAtRow_(sh, row);
  if (!o) { ui.alert('ไม่พบออเดอร์ในแถวนี้'); return; }
  const built = buildOrderBlocks_(o);
  const lines = built.blocks.map(function (b, i) {
    const it = (o.items || [])[i] || {};
    return (i + 1) + ') ' + b.subject + ' ' + b.grade +
      '\n    • รหัสในออเดอร์: ' + (it.id || '(ว่าง)') +
      '\n    • ค่าดิบ → วิชา="' + (it.subject || '') + '" ชั้น="' + (it.grade || '') + '" ชม.="' + (it.hours || '') + '"' +
      '\n    • ลิงก์ Drive: ' + (b.link ? '✅ มี' : '❌ ไม่มี') +
      '\n    • รายละเอียดแผน: ' + (b.postNote ? '✅ มี' : '— ไม่มี') +
      '\n    • ข้อความท้าย: ' + (b.footer ? '✅ มี' : '— ไม่มี');
  });
  ui.alert('🔎 ตรวจการจับคู่ #' + o.orderId + '\nลูกค้า: ' + o.customer + ' (' + o.email + ')\nสถานะ: ' + o.status +
    '\nจำนวนรายการที่อ่านได้: ' + built.blocks.length +
    '\n\n' + (lines.join('\n\n') || '(อ่านรายการในออเดอร์ไม่ได้ — รันเมนู 🩹 ก่อน)') +
    (built.missing.length
      ? '\n\n⚠️ ส่งไม่ได้เพราะวิชาเหล่านี้ยังไม่มีลิงก์/ข้อความ: ' + built.missing.map(function (b) { return b.subject + ' ' + b.grade; }).join(', ')
      : '\n\n✅ ทุกวิชามีลิงก์ครบ — ส่งได้เลย'));
}

// ส่งแผนของแถวที่เลือกอยู่อีกครั้ง (ใช้จากในชีตโดยตรง ไม่เปลี่ยนสถานะ)
function resendCurrentRow() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_ORDERS) { ui.alert('กรุณาเลือกแถวในชีต Orders ก่อน'); return; }
  const row = sh.getActiveCell().getRow();
  if (row < 2) { ui.alert('กรุณาเลือกแถวข้อมูล (ไม่ใช่หัวตาราง)'); return; }
  const o = readOrderAtRow_(sh, row);
  if (!o) { ui.alert('ไม่พบออเดอร์ในแถวนี้'); return; }
  const orderId = o.orderId;
  if (ui.alert('🔁 ส่งแผนอีกครั้ง', 'ระบบจะแชร์สิทธิ์ผู้อ่าน Drive ให้ ' + o.email + ' (ถ้าแชร์อยู่แล้วก็ไม่เป็นไร)\nแล้วส่งอีเมลแผนของ #' + orderId + ' อีกครั้ง\n\nยืนยัน?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const built = buildOrderBlocks_(o);
  if (!built.blocks.length) {
    ui.alert('อ่านรายการสินค้าในออเดอร์นี้ไม่ได้ (แถวอาจคอลัมน์เคลื่อน) — รันเมนู 🩹 ก่อน');
    return;
  }
  if (built.missing.length) {
    ui.alert('ยังไม่มีลิงก์ Drive/ข้อความของ: ' + built.missing.map(function (b) { return b.subject; }).join(', '));
    return;
  }
  const shareFailed = grantReadersForBlocks_(built.blocks, o.email);
  sendApprovalEmail_(o, built.blocks);
  sh.getRange(o.row, O.NOTE).setValue(String(o.note || '') + ' | ส่งซ้ำเมื่อ ' + ymd_(new Date()));
  ui.alert(shareFailed.length
    ? '✅ ส่งอีเมลแล้ว แต่แชร์สิทธิ์ไม่สำเร็จบางวิชา (ต้องแชร์เองใน Drive): ' + shareFailed.join(', ')
    : '✅ เรียบร้อย — แชร์สิทธิ์ผู้อ่าน Drive + ส่งอีเมลให้ ' + o.email + ' แล้ว');
}

// 🔑 ไล่แชร์สิทธิ์ย้อนหลังให้ทุกออเดอร์ "ส่งแผนเสร็จสิ้น" ที่ยังไม่เคยแชร์
// (ออเดอร์เยอะอาจไม่จบในรอบเดียว — กดเมนูซ้ำ ระบบทำต่อจากเดิมเอง ไม่ซ้ำของเก่า)
function shareDriveBackfill() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('🔑 แชร์สิทธิ์ย้อนหลัง',
      'จะแชร์สิทธิ์ผู้อ่านใน Drive ให้ลูกค้าทุกออเดอร์สถานะ "ส่งแผนเสร็จสิ้น" ที่ยังไม่เคยแชร์ (ไม่มีการส่งอีเมล)\n' +
      'ออเดอร์เยอะอาจใช้หลายนาที ถ้าระบบหยุดกลางทางให้กดเมนูนี้ซ้ำ จะทำต่อจากเดิมอัตโนมัติ\n\nเริ่มเลย?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sh = getSheet_(SHEET_ORDERS);
  const startT = Date.now();
  const MARK = '🔑แชร์แล้ว';
  const orders = readOrders_().reverse(); // ใหม่ → เก่า
  let done = 0, skip = 0;
  for (var i = 0; i < orders.length; i++) {
    const o = orders[i];
    if (o.status !== STATUSES.COMPLETED) continue;
    if (String(o.note || '').indexOf(MARK) >= 0) { skip++; continue; }
    const built = buildOrderBlocks_(o);
    const failed = grantReadersForBlocks_(built.blocks, o.email);
    sh.getRange(o.row, O.NOTE).setValue(String(o.note || '') + ' | ' + MARK +
      (failed.length ? ' (พลาด: ' + failed.join('; ') + ')' : ''));
    done++;
    if (Date.now() - startT > 4.5 * 60 * 1000) {
      ui.alert('⏳ รอบนี้แชร์ไป ' + done + ' ออเดอร์ ยังไม่หมด — กดเมนูนี้ซ้ำอีกครั้งเพื่อทำต่อ');
      return;
    }
  }
  ui.alert('✅ แชร์สิทธิ์ย้อนหลังครบทุกออเดอร์แล้ว (รอบนี้ ' + done + ' / ข้ามที่เคยแชร์แล้ว ' + skip + ')');
}

// ✂️ ตัดคอลัมน์ "สื่อ HTML" ออกจากชีตแผนการสอน (ลบคอลัมน์ที่หัวตารางชื่อ "สื่อ HTML")
// ปลอดภัย: ค้นหาด้วยชื่อหัวคอลัมน์ ไม่อิงตำแหน่งตายตัว และลบเฉพาะคอลัมน์นั้น
function removeHtmlMediaColumn() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet_(SHEET_PRODUCTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || '').trim(); });
  const idx = header.indexOf('สื่อ HTML');
  if (idx === -1) {
    ui.alert('ไม่พบคอลัมน์ "สื่อ HTML" ในชีตแผนการสอน — อาจถูกลบไปแล้ว ✅\n\nหัวคอลัมน์ปัจจุบัน: ' + header.join(' | '));
    return;
  }
  if (ui.alert('✂️ ตัดคอลัมน์ "สื่อ HTML"',
      'จะลบคอลัมน์ที่ ' + columnLetter_(idx + 1) + ' (สื่อ HTML) ออกจากชีตแผนการสอน\nคอลัมน์ที่อยู่ขวามือจะเลื่อนซ้ายขึ้นมา 1 ช่อง\n\nแนะนำให้สำรองข้อมูลก่อน — ดำเนินการต่อ?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.deleteColumn(idx + 1);
  // จัดหัวคอลัมน์ "ประเภท" ให้ครบหลังเลื่อน
  if (String(sh.getRange(1, P.CATEGORY).getValue()).trim() !== 'ประเภท') {
    sh.getRange(1, P.CATEGORY).setValue('ประเภท').setFontWeight('bold').setBackground('#dbeafe');
  }
  ui.alert('✅ ตัดคอลัมน์ "สื่อ HTML" เรียบร้อยแล้ว\n\nหัวคอลัมน์ใหม่:\n' +
    sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join(' | '));
}

// แปลงเลขคอลัมน์ → ตัวอักษร (1→A, 27→AA)
function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

// ✂️ แยก "หมายเหตุท้ายเมล" เดิม เป็น 2 คอลัมน์: "รายละเอียดแผน" (เฉพาะวิชา) + "ข้อความท้าย (ส่งให้ลูกค้า)" (มาตรฐาน)
// ตัดข้อความให้อัตโนมัติ: ส่วนบน (ก่อน "สื่อการสอน/เอกสารเพิ่มเติม") = รายละเอียดแผน, ส่วนล่าง = แทนด้วยข้อความท้ายมาตรฐาน
function splitPlanDetailFooter() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet_(SHEET_PRODUCTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || '').trim(); });
  let noteCol = header.indexOf('รายละเอียดแผน');
  if (noteCol === -1) noteCol = header.indexOf('หมายเหตุท้ายเมล');
  if (noteCol === -1) { ui.alert('ไม่พบคอลัมน์ "หมายเหตุท้ายเมล" หรือ "รายละเอียดแผน" ในชีตแผนการสอน'); return; }
  noteCol += 1; // 1-indexed

  if (ui.alert('✂️ แยกหมายเหตุเป็น 2 คอลัมน์',
      'จะแยกข้อความในคอลัมน์ "' + header[noteCol - 1] + '" ออกเป็น:\n' +
      '• รายละเอียดแผน (เฉพาะวิชา — ส่วนบน)\n' +
      '• ข้อความท้าย (ส่งให้ลูกค้า) — มาตรฐานเหมือนกันทุกวิชา\n\n' +
      'แนะนำให้สำรองข้อมูลก่อน — ดำเนินการต่อ?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let footerCol = header.indexOf('ข้อความท้าย (ส่งให้ลูกค้า)');
  if (footerCol === -1) {
    sh.insertColumnAfter(noteCol);
    sh.getRange(1, noteCol + 1).setValue('ข้อความท้าย (ส่งให้ลูกค้า)').setFontWeight('bold').setBackground('#dbeafe');
    footerCol = noteCol + 1;
  } else {
    footerCol += 1;
  }
  sh.getRange(1, noteCol).setValue('รายละเอียดแผน').setFontWeight('bold').setBackground('#dbeafe');

  const last = sh.getLastRow();
  let changed = 0;
  if (last >= 2) {
    const notes = sh.getRange(2, noteCol, last - 1, 1).getValues();
    const details = [], footers = [];
    notes.forEach(function (r) {
      const full = String(r[0] || '');
      const d = splitDetail_(full);
      if (d !== full.trim()) changed++;
      details.push([d]);
      footers.push([STD_FOOTER]);
    });
    sh.getRange(2, noteCol, details.length, 1).setValues(details);
    sh.getRange(2, footerCol, footers.length, 1).setValues(footers);
  }
  ui.alert('✅ แยกเรียบร้อยแล้ว\n' +
    '• "รายละเอียดแผน" = ส่วนบนเฉพาะวิชา (ตัดให้ ' + changed + ' แถว)\n' +
    '• "ข้อความท้าย (ส่งให้ลูกค้า)" = ข้อความมาตรฐาน เติมให้ครบทุกแถวแล้ว\n\n' +
    'อีเมลที่ส่งลูกค้าจะรวม 2 คอลัมน์นี้ให้อัตโนมัติ (หน้าตาเหมือนเดิม)');
}

// ตัดเอาเฉพาะ "ส่วนบน (รายละเอียดแผน)" จากข้อความหมายเหตุเดิม — ตัดก่อนบรรทัด "สื่อการสอน/เอกสารเพิ่มเติม"
function splitDetail_(full) {
  const markers = ['สื่อการสอน', 'Interactive HTML', 'เอกสารเพิ่มเติม', '📱'];
  let idx = -1;
  markers.forEach(function (k) {
    const p = full.indexOf(k);
    if (p !== -1 && (idx === -1 || p < idx)) idx = p;
  });
  if (idx === -1) return full.trim();
  const nl = full.lastIndexOf('\n', idx);
  return full.substring(0, nl === -1 ? idx : nl).trim();
}

// 🔍 ตรวจชื่อวิชาในชีตแผนการสอนให้ตรงกันทุกช่อง
// เทียบว่า "ชื่อวิชา + ระดับชั้น" ตรงกับที่ปรากฏใน หมายเหตุท้ายเมล หรือไม่
// และเช็กว่ามีลิงก์ Drive (ส่งให้ลูกค้า) และลิงก์ตัวอย่างครบไหม → เขียนผลลงชีต "🔍 ตรวจชื่อวิชา"
function auditProductSubjectMatch() {
  const ui = SpreadsheetApp.getUi();
  const products = readProducts_();
  if (!products.length) { ui.alert('ไม่มีข้อมูลแผนการสอนให้ตรวจ'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const NAME = '🔍 ตรวจชื่อวิชา';
  let out = ss.getSheetByName(NAME);
  if (out) out.clear(); else out = ss.insertSheet(NAME);

  const headers = ['ID', 'ชื่อวิชา', 'ระดับชั้น', 'ลิงก์ Drive', 'ลิงก์ตัวอย่าง', 'ชื่อวิชาในหมายเหตุ', 'ระดับชั้นในหมายเหตุ', 'ผลตรวจ'];
  out.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dbeafe');

  let problems = 0;
  const rows = products.map(function (p) {
    const note = normKey_(p.planDetail);
    const subjKey = normKey_(p.subject);
    const gradeKey = normKey_(p.grade);
    const hasDrive = !!p.driveLink;
    const hasExample = !!p.exampleLink;
    const noteHasSubj = subjKey && note.indexOf(subjKey) !== -1;
    const noteHasGrade = gradeKey && note.indexOf(gradeKey) !== -1;

    const issues = [];
    if (!hasDrive) issues.push('❌ ไม่มีลิงก์ Drive');
    if (!hasExample) issues.push('⚠️ ไม่มีลิงก์ตัวอย่าง');
    if (p.planDetail && !noteHasSubj) issues.push('⚠️ ชื่อวิชาในรายละเอียดไม่ตรง');
    if (p.planDetail && gradeKey && !noteHasGrade) issues.push('⚠️ ระดับชั้นในรายละเอียดไม่ตรง');
    if (!p.planDetail) issues.push('⚠️ ไม่มีรายละเอียดแผน');
    if (issues.length) problems++;

    return [
      p.id, p.subject, p.grade,
      hasDrive ? '✅ มี' : '❌ ไม่มี',
      hasExample ? '✅ มี' : '— ไม่มี',
      p.planDetail ? (noteHasSubj ? '✅ ตรง' : '⚠️ ไม่ตรง') : '—',
      p.planDetail ? (gradeKey ? (noteHasGrade ? '✅ ตรง' : '⚠️ ไม่ตรง') : '—') : '—',
      issues.length ? issues.join(' · ') : '✅ ตรงกันครบ'
    ];
  });

  out.getRange(2, 1, rows.length, headers.length).setValues(rows);
  // ระบายสีแถวที่มีปัญหาให้เห็นง่าย
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][7] !== '✅ ตรงกันครบ') out.getRange(i + 2, 1, 1, headers.length).setBackground('#fff7ed');
  }
  out.setFrozenRows(1);
  out.autoResizeColumns(1, headers.length);
  ss.setActiveSheet(out);

  ui.alert('🔍 ตรวจชื่อวิชาเสร็จแล้ว\n\nตรวจทั้งหมด ' + products.length + ' รายการ\n' +
    (problems ? '⚠️ พบที่ต้องดู ' + problems + ' รายการ (ระบายสีส้มในชีต "' + NAME + '")'
              : '✅ ทุกรายการชื่อวิชา/ลิงก์ตรงกันครบ') +
    '\n\nหมายเหตุ: ระบบเทียบ "ข้อความ" ของชื่อวิชา/ชั้นในหมายเหตุท้ายเมลได้ แต่ไม่สามารถเปิดดูว่าลิงก์ Drive/ตัวอย่างชี้ไปไฟล์วิชาถูกตัวจริงไหม จุดนั้นต้องสุ่มกดเปิดเช็กเองนะคะ');
}

// 🔢 ซ่อมเลขออเดอร์ซ้ำ — ออกเลขใหม่ให้แถวที่ซ้ำ (คงแถวแรกไว้) กันส่งผิดออเดอร์
// ใช้กับรูปแบบเลขใหม่ KPSyyMMdd-NNN เท่านั้น (รูปแบบเก่าแบบสุ่มจะไม่ยุ่ง)
function fixDuplicateOrderIds() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet_(SHEET_ORDERS);
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('ไม่มีข้อมูล'); return; }
  const ids = sh.getRange(2, O.ORDER_ID, last - 1, 1).getValues()
    .map(function (r) { return String(r[0] || ''); });

  // หาเลขลำดับสูงสุดที่ใช้ไปแล้วของแต่ละวัน (prefix = KPSyyMMdd-)
  const maxSeq = {};
  ids.forEach(function (v) {
    const m = v.match(/^(KPS\d{6}-)(\d+)$/);
    if (m) {
      const p = m[1], n = parseInt(m[2], 10);
      if (!maxSeq[p] || n > maxSeq[p]) maxSeq[p] = n;
    }
  });

  const seen = {};
  const report = [];
  let fixed = 0;
  for (var i = 0; i < ids.length; i++) {
    const v = ids[i];
    if (!v) continue;
    if (!seen[v]) { seen[v] = true; continue; }   // เจอครั้งแรก — เก็บเลขเดิมไว้
    const m = v.match(/^(KPS\d{6}-)(\d+)$/);
    if (!m) { seen[v] = true; continue; }          // รูปแบบเก่า ไม่แตะ
    const p = m[1];
    maxSeq[p] = (maxSeq[p] || 0) + 1;
    let newId = p + padNum_(maxSeq[p], 3);
    while (seen[newId]) { maxSeq[p]++; newId = p + padNum_(maxSeq[p], 3); }
    sh.getRange(i + 2, O.ORDER_ID).setValue(newId);
    sh.getRange(i + 2, O.ORDER_ID).setNumberFormat('@');
    seen[newId] = true;
    report.push(v + '  →  ' + newId + '  (แถว ' + (i + 2) + ')');
    fixed++;
  }
  ui.alert(fixed
    ? '🔢 ซ่อมเลขออเดอร์ซ้ำแล้ว ' + fixed + ' แถว:\n\n' + report.join('\n') +
      '\n\nคงเลขของแถวแรกไว้ ออกเลขใหม่ให้แถวที่ซ้ำ — ตอนนี้ส่งแผนได้ถูกออเดอร์แล้ว'
    : 'ไม่พบเลขออเดอร์ซ้ำ — เรียบร้อยดีอยู่แล้ว ✅');
}

// 🩹 ซ่อมแถวออเดอร์ที่คอลัมน์เคลื่อนซ้าย 1 ช่อง (ถูกเขียนโดย deployment เวอร์ชันเก่าที่ยังไม่มีคอลัมน์ "วิชาที่สั่ง")
// วิธีตรวจ: ช่องสถานะไม่ใช่สถานะที่รู้จัก แต่ช่องก่อนหน้าเป็นสถานะ → เลื่อนข้อมูลกลับ + เติมวิชาที่สั่งจาก JSON
function fixShiftedOrderRows() {
  const ui = SpreadsheetApp.getUi();
  const sh = getSheet_(SHEET_ORDERS);
  const last = sh.getLastRow();
  if (last < 2) { ui.alert('ไม่มีข้อมูล'); return; }
  const data = sh.getRange(2, 1, last - 1, O_COUNT).getValues();
  let fixed = 0;
  for (var i = 0; i < data.length; i++) {
    const r = data[i];
    const stOk = isStatus_(r[O.STATUS - 1]);
    const stShifted = isStatus_(r[O.STATUS - 2]);
    const tailEmpty = String(r[O_COUNT - 1] || '') === '';
    if (stOk || !stShifted || !tailEmpty) continue;
    // เลื่อนขวา 1 ช่องตั้งแต่คอลัมน์ "วิชาที่สั่ง": [A,B,C] + [ว่าง] + ของเดิม D..N
    const newRow = [r[0], r[1], r[2], ''].concat(r.slice(3, O_COUNT - 1));
    newRow[O.ITEMS_TEXT - 1] = itemsText_(parseItems_(newRow[O.ITEMS_JSON - 1]));
    sh.getRange(i + 2, 1, 1, O_COUNT).setValues([newRow]);
    sh.getRange(i + 2, O.TIMESTAMP).setNumberFormat('@');
    fixed++;
  }
  ui.alert(fixed ? '🩹 ซ่อมแถวที่คอลัมน์เคลื่อนแล้ว ' + fixed + ' แถว' : 'ไม่พบแถวที่คอลัมน์เคลื่อน — เรียบร้อยดีอยู่แล้ว');
}

// ============================================================
// 🧾 กันสลิปซ้ำ — ลายนิ้วมือไฟล์ (MD5)
// ============================================================
function bytesToHex_(bytes) {
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}
// คืนเลขออเดอร์อื่นที่เคยใช้สลิปใบเดียวกัน ('' = ไม่ซ้ำ) — อัปสลิปซ้ำในออเดอร์เดิมไม่นับ
function findSlipHash_(hashHex, orderId) {
  const sh = getSheet_(SHEET_SLIPHASH);
  if (sh.getLastRow() < 2) return '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === hashHex && String(rows[i][1]) !== String(orderId)) {
      return String(rows[i][1]);
    }
  }
  return '';
}
function recordSlipHash_(hashHex, orderId) {
  try { getSheet_(SHEET_SLIPHASH).appendRow([hashHex, orderId, ymd_(new Date())]); }
  catch (e) { Logger.log('recordSlipHash_: ' + e); }
}

// 🚫 ถอนสิทธิ์ผู้อ่าน Drive คืนจากลูกค้า เมื่อออเดอร์ถูกยกเลิก
// ตัวกันสำคัญ: ถ้าลูกค้ามีออเดอร์อื่น (ที่ไม่ยกเลิก และส่งแล้ว/อนุมัติแล้ว) ซื้อวิชาเดียวกัน → ไม่ถอนวิชานั้น
function revokeReadersForOrder_(o) {
  const built = buildOrderBlocks_(o);
  // โฟลเดอร์ที่ลูกค้าคนนี้ยังควรมีสิทธิ์จากออเดอร์อื่น
  const keep = {};
  readOrders_().forEach(function (other) {
    if (other.orderId === o.orderId) return;
    if (String(other.email || '').toLowerCase() !== String(o.email || '').toLowerCase()) return;
    if (other.status !== STATUSES.COMPLETED && other.status !== STATUSES.APPROVED) return;
    buildOrderBlocks_(other).blocks.forEach(function (b) {
      const id = extractDriveId_(b.link);
      if (id) keep[id] = true;
    });
  });
  const revoked = [], kept = [], failed = [];
  built.blocks.forEach(function (b) {
    const id = extractDriveId_(b.link);
    if (!id) return;
    const label = (b.subject + ' ' + b.grade).trim();
    if (keep[id]) { kept.push(label); return; }
    try { DriveApp.getFolderById(id).removeViewer(o.email); revoked.push(label); }
    catch (e1) {
      try { DriveApp.getFileById(id).removeViewer(o.email); revoked.push(label); }
      catch (e2) { failed.push(label); } // ส่วนใหญ่คือไม่ได้ถูกแชร์อยู่แล้ว
    }
  });
  return { revoked: revoked, kept: kept, failed: failed };
}

// สรุปผลถอนสิทธิ์เป็นข้อความสั้นๆ ใช้ทั้งใน note / Telegram / alert
function revokeSummary_(res) {
  const parts = [];
  if (res.revoked.length) parts.push('🚫 ถอนสิทธิ์แล้ว: ' + res.revoked.join(', '));
  if (res.kept.length) parts.push('✋ คงสิทธิ์ไว้ (มีออเดอร์อื่น): ' + res.kept.join(', '));
  if (res.failed.length) parts.push('ℹ️ ไม่ได้ถูกแชร์/ถอนไม่ได้: ' + res.failed.join(', '));
  return parts.join(' | ') || 'ไม่มีลิงก์ Drive ให้ถอน';
}

// 🚫 เมนูในชีต: ยกเลิกออเดอร์แถวที่เลือก + ถอนสิทธิ์ Drive คืนจากลูกค้า
function cancelAndRevokeCurrentRow() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_ORDERS) { ui.alert('กรุณาเลือกแถวในชีต Orders ก่อน'); return; }
  const row = sh.getActiveCell().getRow();
  if (row < 2) { ui.alert('กรุณาเลือกแถวข้อมูล'); return; }
  const o = readOrderAtRow_(sh, row);
  if (!o) { ui.alert('ไม่พบออเดอร์ในแถวนี้'); return; }
  if (ui.alert('🚫 ยกเลิกออเดอร์ #' + o.orderId,
      'ลูกค้า: ' + o.customer + ' (' + o.email + ')\nระบบจะเปลี่ยนสถานะเป็น "ยกเลิก" และถอนสิทธิ์ผู้อ่าน Drive ของวิชาในออเดอร์นี้คืน\n(วิชาที่ลูกค้ายังมีจากออเดอร์อื่นจะไม่ถูกถอน)\n\nยืนยัน?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.getRange(o.row, O.STATUS).setValue(STATUSES.CANCELLED);
  const res = revokeReadersForOrder_(o);
  const cur = String(sh.getRange(o.row, O.NOTE).getValue() || '');
  sh.getRange(o.row, O.NOTE).setValue((cur ? cur + ' | ' : '') + revokeSummary_(res) + ' (ยกเลิกโดยเมนูชีต ' + ymd_(new Date()) + ')');
  // (ปิดแจ้งเตือน Telegram ตอนยกเลิก+ถอนสิทธิ์ตามที่ขอ)
  ui.alert('เรียบร้อย\n' + revokeSummary_(res));
}

// ⚡ เปิดใช้: เลือกสถานะ "อนุมัติส่งแผน" ในชีตแล้วระบบส่งแผนทันที (ติดตั้งทริกเกอร์ครั้งเดียว)
// ⚡ เปิดระบบอัตโนมัติทุกตัวในคลิกเดียว: ส่งเมื่ออนุมัติ + รายงานเช้า + สำรองกลางคืน + คิวแจ้งอัปเดต
// กดซ้ำได้ปลอดภัย — ตัวไหนเปิดอยู่แล้วจะข้าม ไม่สร้าง trigger ซ้ำ
function setupAllAutomation() {
  const have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have['onOrderStatusEdit']) {
    ScriptApp.newTrigger('onOrderStatusEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  }
  if (!have['dailySummary']) ScriptApp.newTrigger('dailySummary').timeBased().atHour(8).everyDays(1).create();
  if (!have['nightlyBackup']) ScriptApp.newTrigger('nightlyBackup').timeBased().atHour(2).everyDays(1).create();
  if (!have['processUpdateQueue']) ScriptApp.newTrigger('processUpdateQueue').timeBased().atHour(3).everyDays(1).create();
  SpreadsheetApp.getUi().alert('⚡ เปิดระบบอัตโนมัติครบแล้ว!\n\n' +
    '1) เลือกสถานะ "อนุมัติส่งแผน" ในชีต Orders → แชร์สิทธิ์ Drive + ส่งแผนทางอีเมลทันที (ไม่มีหน้าต่างยืนยัน)\n' +
    '2) ทุกเช้า ~08:00 ส่งรายงานสรุปเข้า Telegram\n' +
    '3) ทุกคืน ~02:00 สำรองสเปรดชีตลง Drive (' + BACKUP_FOLDER_NAME + ' เก็บ ' + BACKUP_KEEP + ' ชุด)\n' +
    '4) ทุกคืน ~03:00 ทยอยส่งคิวแจ้งอัปเดตลูกค้า (ถ้ามี)\n' +
    '5) ทุก ~5 นาที เผยแพร่หน้าร้านขึ้น Edge เมื่อข้อมูลเปลี่ยน (ถ้าตั้งค่า 🚀 แล้ว)');
}

function setupStatusEditTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onOrderStatusEdit';
  });
  if (!exists) {
    ScriptApp.newTrigger('onOrderStatusEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit().create();
  }
  SpreadsheetApp.getUi().alert('⚡ เปิดใช้แล้ว!\n\nต่อไปเมื่อเลือกสถานะ "อนุมัติส่งแผน" ในชีต Orders ระบบจะ:\n1) แชร์สิทธิ์ผู้อ่าน Drive ให้ Gmail ลูกค้า\n2) ส่งแผนทางอีเมลทันที\n3) เปลี่ยนสถานะเป็น "ส่งแผนเสร็จสิ้น" ให้เอง + แจ้ง Telegram\n\n⚠️ ระวัง: เลือกแล้วส่งจริงเลย ไม่มีหน้าต่างยืนยัน');
}

// ทำงานอัตโนมัติเมื่อมีการแก้ไขชีต (เฉพาะคอลัมน์สถานะ → ค่า "อนุมัติส่งแผน")
function onOrderStatusEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_ORDERS) return;
    if (e.range.getColumn() !== O.STATUS || e.range.getNumRows() > 1) return;
    const edited = String(e.value || '');
    if (edited !== STATUSES.APPROVED && edited !== STATUSES.CANCELLED) return;
    const row = e.range.getRow();
    if (row < 2) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const o = readOrderAtRow_(sh, row);
    if (!o) { ss.toast('ไม่พบเลขออเดอร์ในแถวนี้', '⚠️', 6); return; }

    // เลือก "ยกเลิก" → ถอนสิทธิ์ Drive คืน
    if (edited === STATUSES.CANCELLED) {
      ss.toast('กำลังถอนสิทธิ์ Drive ของ ' + o.email + ' ...', '🚫 #' + o.orderId, 5);
      const res = revokeReadersForOrder_(o);
      const cur = String(sh.getRange(o.row, O.NOTE).getValue() || '');
      sh.getRange(o.row, O.NOTE).setValue((cur ? cur + ' | ' : '') + revokeSummary_(res));
      // (ปิดแจ้งเตือน Telegram ตอนยกเลิก+ถอนสิทธิ์ตามที่ขอ)
      ss.toast(revokeSummary_(res), '🚫 #' + o.orderId, 8);
      return;
    }

    ss.toast('กำลังแชร์สิทธิ์ + ส่งแผนให้ ' + o.email + ' ...', '⚡ #' + o.orderId, 5);
    const built = buildOrderBlocks_(o);
    if (built.missing.length) {
      ss.toast('ยังไม่มีลิงก์/ข้อความของ: ' +
        built.missing.map(function (b) { return b.subject + ' ' + b.grade; }).join(', '), '⚠️ ส่งไม่ได้', 8);
      return;
    }
    const shareFailed = grantReadersForBlocks_(built.blocks, o.email);
    sendApprovalEmail_(o, built.blocks);
    sh.getRange(o.row, O.STATUS).setValue(STATUSES.COMPLETED);
    sh.getRange(o.row, O.NOTE).setValue('ส่งจากชีต เมื่อ ' + ymd_(new Date()) +
      (shareFailed.length ? ' | ⚠️ แชร์สิทธิ์ไม่สำเร็จ: ' + shareFailed.join(', ') : ''));
    sendTelegram_('✅ ส่งแผนสำเร็จ (สั่งจากชีต) #' + o.orderId + '\nลูกค้า: ' + o.customer + '\nยอด: ' + o.netTotal + ' บาท' +
      (shareFailed.length ? '\n⚠️ แชร์สิทธิ์ไม่สำเร็จ ต้องแชร์เอง: ' + shareFailed.join(', ') : '\n🔑 แชร์สิทธิ์ผู้อ่านให้ลูกค้าแล้ว'));
    ss.toast('ส่งแผนเรียบร้อย 💌 สถานะเปลี่ยนเป็นส่งเสร็จสิ้นแล้ว', '✅ #' + o.orderId, 6);
  } catch (err) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(String(err.message || err).substring(0, 100), '⚠️ ส่งไม่สำเร็จ', 8);
    } catch (e2) {}
    Logger.log('onOrderStatusEdit error: ' + err);
  }
}

// ============================================================
// ⏰ งานอัตโนมัติรายวัน: รายงานเช้า 8 โมง + สำรองข้อมูลตี 2
// ============================================================
function setupDailyJobs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have['dailySummary']) {
    ScriptApp.newTrigger('dailySummary').timeBased().atHour(8).everyDays(1).create();
  }
  if (!have['nightlyBackup']) {
    ScriptApp.newTrigger('nightlyBackup').timeBased().atHour(2).everyDays(1).create();
  }
  if (!have['processUpdateQueue']) {
    ScriptApp.newTrigger('processUpdateQueue').timeBased().atHour(3).everyDays(1).create();
  }
  if (!have['edgeAutoPublish']) {
    ScriptApp.newTrigger('edgeAutoPublish').timeBased().everyMinutes(5).create();
  }
  SpreadsheetApp.getUi().alert('⏰ เปิดใช้แล้ว!\n\n• ทุกเช้า ~08:00 น. ส่งรายงานสรุปเข้า Telegram\n• ทุกคืน ~03:00 น. ทยอยส่งคิวแจ้งอัปเดตลูกค้า (ถ้ามี)\n• ทุกคืน ~02:00 น. สำรองสเปรดชีตลง Drive (โฟลเดอร์ ' + BACKUP_FOLDER_NAME + ' เก็บย้อนหลัง ' + BACKUP_KEEP + ' ชุด)\n\nทดสอบได้เลยจากเมนู "ส่งรายงานสรุปตอนนี้" และ "สำรองข้อมูลเดี๋ยวนี้"');
}

// 📊 รายงานสรุปประจำวัน ส่งเข้า Telegram (สรุปของเมื่อวาน + งานค้างตอนนี้)
function dailySummary() {
  try {
    const y = ymd_(new Date(Date.now() - 24 * 60 * 60 * 1000)).slice(0, 10); // เมื่อวาน yyyy-MM-dd
    const orders = readOrders_();
    let newCount = 0, sales = 0, autoSent = 0, adminSent = 0;
    let pendingPay = 0, toSend = 0;
    orders.forEach(function (o) {
      if (String(o.timestamp).indexOf(y) === 0) {
        newCount++;
        if (o.status !== STATUSES.CANCELLED) sales += Number(o.netTotal) || 0;
      }
      const note = String(o.note || '');
      if (note.indexOf(y) >= 0) {
        if (note.indexOf('ส่งอัตโนมัติ') >= 0) autoSent++;
        else if (note.indexOf('ส่งโดย') >= 0 || note.indexOf('ส่งจากชีต') >= 0) adminSent++;
      }
      if (o.status === STATUSES.PENDING) pendingPay++;
      if (o.status === STATUSES.PAID || o.status === STATUSES.APPROVED) toSend++;
    });
    const quota = MailApp.getRemainingDailyQuota();
    sendTelegram_(
      '🌅 สรุปประจำวัน — ครูพร้อมสอน\n' +
      '(ข้อมูลของเมื่อวาน ' + y + ')\n' +
      '━━━━━━━━━━━━━━\n' +
      '🛒 ออเดอร์ใหม่: ' + newCount + ' รายการ\n' +
      '💰 ยอดขาย: ' + sales.toLocaleString() + ' บาท\n' +
      '🤖 ส่งอัตโนมัติ: ' + autoSent + ' | 👩‍💻 ส่งโดยแอดมิน: ' + adminSent + '\n' +
      '━━━━━━━━━━━━━━\n' +
      '⏳ ค้างตอนนี้: รอชำระ ' + pendingPay + ' | รอส่งแผน ' + toSend + (toSend > 0 ? ' ⚠️' : ' ✅') + '\n' +
      '📧 โควต้าเมลวันนี้เหลือ: ' + quota + ' ฉบับ' + (quota < 30 ? ' ⚠️ ใกล้หมด!' : ''));
  } catch (e) {
    Logger.log('dailySummary: ' + e);
    sendTelegram_('⚠️ รายงานสรุปประจำวันขัดข้อง: ' + String(e.message || e).substring(0, 100));
  }
}

// 💾 สำรองสเปรดชีตทั้งไฟล์ลง Drive ทุกคืน + เก็บย้อนหลังตามจำนวนที่ตั้ง
function nightlyBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folder = getOrCreateFolder_(BACKUP_FOLDER_NAME);
    const stamp = ymd_(new Date()).slice(0, 16).replace(':', '.');
    DriveApp.getFileById(ss.getId()).makeCopy('💾 สำรอง ' + ss.getName() + ' ' + stamp, folder);
    // ลบชุดสำรองที่เกินจำนวนเก็บ (เก่าสุดออกก่อน)
    const files = [];
    const it = folder.getFiles();
    while (it.hasNext()) { const f = it.next(); files.push({ f: f, t: f.getDateCreated().getTime() }); }
    files.sort(function (a, b) { return b.t - a.t; });
    for (var i = BACKUP_KEEP; i < files.length; i++) files[i].f.setTrashed(true);
  } catch (e) {
    Logger.log('nightlyBackup: ' + e);
    sendTelegram_('⚠️ สำรองข้อมูลรายคืนล้มเหลว: ' + String(e.message || e).substring(0, 100));
  }
}
function backupNowMenu() {
  nightlyBackup();
  SpreadsheetApp.getUi().alert('💾 สำรองเรียบร้อย — ดูได้ใน Drive โฟลเดอร์ "' + BACKUP_FOLDER_NAME + '" (เก็บย้อนหลัง ' + BACKUP_KEEP + ' ชุดล่าสุด)');
}

function menuTestTelegram() {
  const m = settingsMap_();
  if (!String(m.telegramBotToken || '').trim() || !String(m.telegramChatId || '').trim()) {
    SpreadsheetApp.getUi().alert('ยังไม่ได้ตั้งค่า Telegram — เข้า Web App เมนู ⚙️ ตั้งค่าระบบ ใส่ Bot Token แล้วกด "หา Chat ID"');
    return;
  }
  sendTelegram_('🔔 ทดสอบแจ้งเตือนจาก Google Sheets — ' + ymd_(new Date()));
  SpreadsheetApp.getUi().alert('ส่งข้อความทดสอบแล้ว เช็คใน Telegram ได้เลย');
}

// 🏗️ สร้างชีต Orders ใหม่: อ่าน "เนื้อหา" ของแต่ละเซลล์ (อีเมล/JSON/สถานะ/ลิงก์/ตัวเลข)
// แล้วจัดเข้าโครง 15 คอลัมน์ — ไม่เดาจากตำแหน่ง โครงเพี้ยนแบบไหนก็จัดได้
// ดึงข้อมูลจากชีต "Orders_เดิม(สำรอง)" อัตโนมัติถ้ามี (ต้นฉบับก่อนการจัดรอบที่แล้ว)
function rebuildOrdersSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bak = ss.getSheetByName('Orders_เดิม(สำรอง)');
  const cur = ss.getSheetByName(SHEET_ORDERS);
  const src = bak || cur;
  if (!src) { ui.alert('ไม่พบชีต Orders'); return; }
  if (ui.alert('🏗️ สร้างชีต Orders ใหม่',
      'แหล่งข้อมูล: ' + src.getName() + ' (' + Math.max(src.getLastRow() - 1, 0) + ' แถว)\n' +
      'ระบบจะอ่านเนื้อหาแต่ละช่องแล้วจัดเข้าหัวตารางให้ถูกต้อง\n\nเริ่มเลย?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const lastRow = src.getLastRow();
  const lastCol = Math.max(src.getLastColumn(), O_COUNT);
  const data = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const out = [], noEmail = [], noItems = [];
  const EMAIL_RE = /^[^@\s,;:"']+@[^@\s,;:"']+\.[A-Za-z]{2,}$/;

  data.forEach(function (r) {
    if (!String(r[1] || '').trim()) return; // ไม่มีเลขออเดอร์ = ข้าม
    let email = '', phone = '', status = '', slip = '';
    let items = [], jsonIdx = -1, statusIdx = -1, subjTextIdx = -1;
    const noteParts = [];

    // รอบที่ 1: ระบุชนิดของแต่ละช่องจากเนื้อหา
    for (var c = 3; c < r.length; c++) {
      const sv = String(r[c] == null ? '' : r[c]).trim();
      if (!sv) continue;
      if (!email && EMAIL_RE.test(sv)) { email = sv; continue; }
      if (!items.length && sv.charAt(0) === '[' && sv.indexOf('{') >= 0) {
        // เจอ JSON — ถ้าแปลงไม่ผ่าน แปลว่าถูกหั่นข้ามช่อง → ต่อช่องถัดไปเรื่อยๆ จนแปลงผ่าน
        let cand = sv, cend = c;
        let got = parseItems_(cand);
        while (!got.length && cend + 1 < r.length && cend - c < 6) {
          cend++;
          cand += String(r[cend] == null ? '' : r[cend]);
          got = parseItems_(cand);
        }
        if (got.length) { items = got; jsonIdx = c; }
        continue;
      }
      if (!status && isStatus_(sv)) { status = sv; statusIdx = c; continue; }
      if (!slip && sv.indexOf('http') === 0) { slip = sv; continue; }
      if (!phone && /^0\d{8,9}$/.test(sv)) { phone = sv; continue; }
      // จดตำแหน่งช่องที่หน้าตาเหมือนรายชื่อวิชา ไว้เป็นแผนสำรอง
      if (subjTextIdx < 0 && /(ป\.|ม\.|VIP|อ\.)\s*\d/.test(sv)) subjTextIdx = c;
    }
    // แผนสำรอง: JSON หายสนิท → กู้รายการจากข้อความวิชา (ได้ชื่อวิชา+จำนวน แต่ส่งซ้ำต้องเติมลิงก์เอง)
    if (!items.length && subjTextIdx >= 0) {
      items = String(r[subjTextIdx]).split('\n').map(function (line) {
        return { id: '', subject: line.trim(), grade: '', curriculum: '', hours: '' };
      }).filter(function (it) { return it.subject; });
    }

    // รอบที่ 2: ตัวเลขเงินที่อยู่ระหว่าง JSON กับสถานะ = [ยอดก่อนลด, %, ลด(บาท), ยอดสุทธิ]
    const anchorL = jsonIdx >= 0 ? jsonIdx : 2;
    const anchorR = statusIdx >= 0 ? statusIdx : r.length;
    const nums = [];
    for (var c2 = anchorL + 1; c2 < anchorR; c2++) {
      const sv2 = String(r[c2] == null ? '' : r[c2]).trim();
      if (/^-?\d+(\.\d+)?$/.test(sv2)) nums.push(Number(sv2));
    }
    let sub = 0, pct = 0, baht = 0, net = 0;
    if (nums.length >= 4) { const q = nums.slice(-4); sub = q[0]; pct = q[1]; baht = q[2]; net = q[3]; }
    else if (nums.length) { sub = nums[0]; net = nums[nums.length - 1]; }

    // รอบที่ 3: ข้อความหลังสถานะที่ไม่ใช่ลิงก์ = หมายเหตุ
    if (statusIdx >= 0) {
      for (var c3 = statusIdx + 1; c3 < r.length; c3++) {
        const sv3 = String(r[c3] == null ? '' : r[c3]).trim();
        if (!sv3 || sv3 === slip || sv3.indexOf('http') === 0) continue;
        if (/^-?\d+(\.\d+)?$/.test(sv3)) continue;
        noteParts.push(sv3);
      }
    }
    if (!email) noEmail.push(String(r[1]));
    if (!items.length) noItems.push(String(r[1]));

    out.push([String(r[0] || ''), String(r[1] || ''), String(r[2] || ''), itemsText_(items),
      email, phone, items.length, JSON.stringify(items),
      sub, pct, baht, net, status, slip, noteParts.join(' | ')]);
  });

  // สร้างชีตใหม่ แล้วตั้งเป็น Orders
  const tmp = ss.getSheetByName('Orders_ใหม่ชั่วคราว');
  if (tmp) ss.deleteSheet(tmp);
  const nu = ss.insertSheet('Orders_ใหม่ชั่วคราว');
  nu.getRange(1, 1, 1, O_COUNT).setValues([[
    'วันที่สั่งซื้อ', 'เลขที่ออเดอร์', 'ชื่อลูกค้า/Facebook', 'วิชาที่สั่ง', 'อีเมล (Gmail)', 'เบอร์โทร',
    'จำนวนรายการ', 'รายการ (JSON)', 'ยอดก่อนลด', 'ส่วนลด %',
    'ส่วนลด (บาท)', 'ยอดสุทธิ', 'สถานะ', 'ลิงก์สลิป', 'หมายเหตุ'
  ]]).setFontWeight('bold').setBackground('#dcfce7');
  if (out.length) {
    nu.getRange(2, 1, out.length, O_COUNT).setValues(out);
    nu.getRange(2, O.TIMESTAMP, out.length, 1).setNumberFormat('@');
  }
  nu.setColumnWidth(O.ITEMS_TEXT, 240);
  nu.setFrozenRows(1);

  if (bak) {
    if (cur) ss.deleteSheet(cur);          // ลบตัวที่จัดพลาดรอบก่อน
  } else {
    cur.setName('Orders_เดิม(สำรอง)');     // เก็บต้นฉบับไว้
    try { cur.hideSheet(); } catch (eH) {}
  }
  nu.setName(SHEET_ORDERS);
  applyOrdersSheetSetup_();
  try { refreshOrderColors(); } catch (eC) {}

  const sample = out.length ? ('\n\nตัวอย่างแถวแรก: #' + out[0][1] + ' → ' + (out[0][4] || '(ไม่มีอีเมล)') + ' • ' + out[0][11] + ' บาท • ' + out[0][12]) : '';
  ui.alert('🏗️ เสร็จแล้ว! จัดใหม่ ' + out.length + ' ออเดอร์ จากชีต "' + src.getName() + '"' +
    (noEmail.length ? '\n⚠️ หาอีเมลไม่เจอ ' + noEmail.length + ' ออเดอร์: ' + noEmail.slice(0, 5).join(', ') + (noEmail.length > 5 ? ' ...' : '') : '') +
    (noItems.length ? '\n⚠️ หารายการวิชาไม่เจอ ' + noItems.length + ' ออเดอร์: ' + noItems.slice(0, 5).join(', ') + (noItems.length > 5 ? ' ...' : '') : '') +
    sample + '\n\nต้นฉบับยังอยู่ในชีต "Orders_เดิม(สำรอง)" (ซ่อน) — ตรวจแล้วค่อยลบ');
}

// 🧰 ตรวจ/สร้าง/ซ่อมโครงสร้างชีตทั้งหมดในคลิกเดียว
// ชีตไหนหาย → สร้างใหม่ / โครงไม่ครบ (เช่นขาดคอลัมน์ใหม่) → เติมให้ / ข้อมูลที่มีอยู่ไม่ถูกแตะ
function repairAllSheets() {
  ensureSheets_();             // Settings + แผนการสอน + Orders (พร้อมอัปเกรดคอลัมน์)
  applyOrdersSheetSetup_();    // dropdown สถานะ + สีสถานะ + ตัวกรอง
  try { refreshOrderColors(); } catch (e) { Logger.log(e); }
  generateDashboard();         // 📊 Dashboard
  generateCustomerSheet();     // 👤 ประวัติลูกค้า
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ครบแล้ว: Settings • แผนประถม • แผนมัธยม • Orders • 📊 Dashboard • 👤 ประวัติลูกค้า',
    '🧰 ตรวจ/สร้างชีตเรียบร้อย', 6);
}

function checkEmailQuota() {
  SpreadsheetApp.getUi().alert('📧 วันนี้ส่งอีเมลได้อีก ' + MailApp.getRemainingDailyQuota() + ' ฉบับ');
}


// ============================================================
// 📊 ชีต Dashboard + 👤 ชีตประวัติลูกค้า (สร้างจากข้อมูล Orders
//    กดรีเฟรชจากเมนู 🎓 ครูพร้อมสอน — ไม่ต้องกรอกมือ ไม่ซ้ำซ้อน)
// ============================================================
function generateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '📊 Dashboard';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, 0); else sh.clear();

  const orders = readOrders_();
  const done = orders.filter(function (o) {
    return o.status === STATUSES.COMPLETED || o.status === STATUSES.APPROVED;
  });
  const tz = 'Asia/Bangkok';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const monthStr = todayStr.slice(0, 7);
  const yearStr = todayStr.slice(0, 4);

  function sumCount(prefix) {
    let s = 0, c = 0;
    done.forEach(function (o) {
      if (o.timestamp.indexOf(prefix) === 0) { s += o.netTotal; c++; }
    });
    return [s, c];
  }
  const dayR = sumCount(todayStr), monR = sumCount(monthStr), yrR = sumCount(yearStr);
  const allSum = done.reduce(function (s, o) { return s + o.netTotal; }, 0);
  const statusCount = {};
  orders.forEach(function (o) { statusCount[o.status] = (statusCount[o.status] || 0) + 1; });

  // Top 10 รายการขายดี
  const subj = {};
  orders.forEach(function (o) {
    if (o.status === STATUSES.CANCELLED) return;
    (o.items || []).forEach(function (it) {
      const key = it.subject + (it.grade ? ' ' + it.grade : '');
      subj[key] = (subj[key] || 0) + 1;
    });
  });
  const top = Object.keys(subj).map(function (k) { return [k, subj[k]]; })
    .sort(function (a, b) { return b[1] - a[1]; }).slice(0, 10);

  // ---- เขียนลงชีต ----
  sh.getRange('A1:D1').merge().setValue('📊 Dashboard – ' + PAGE_NAME + '™')
    .setBackground('#818cf8').setFontColor('#fff').setFontWeight('bold')
    .setFontSize(16).setHorizontalAlignment('center');
  sh.getRange('A2:D2').merge().setValue('อัปเดตล่าสุด: ' + ymd_(new Date()))
    .setFontStyle('italic').setFontColor('#6b7280').setFontSize(9).setHorizontalAlignment('center');

  const rows = [
    ['💰 สรุปยอดขาย (เฉพาะออเดอร์ที่ส่งแผนแล้ว)', '', '', ''],
    ['ช่วงเวลา', 'ยอดขาย (บาท)', 'จำนวนออเดอร์', ''],
    ['วันนี้ (' + todayStr + ')', dayR[0], dayR[1], ''],
    ['เดือนนี้ (' + monthStr + ')', monR[0], monR[1], ''],
    ['ปีนี้ (' + yearStr + ')', yrR[0], yrR[1], ''],
    ['ยอดสะสมทั้งหมด', allSum, done.length, ''],
    ['', '', '', ''],
    ['📦 สถานะคำสั่งซื้อ (' + orders.length + ' รายการ)', '', '', ''],
    [STATUSES.PENDING, statusCount[STATUSES.PENDING] || 0,
     STATUSES.PAID, statusCount[STATUSES.PAID] || 0],
    [STATUSES.COMPLETED, statusCount[STATUSES.COMPLETED] || 0,
     STATUSES.CANCELLED, statusCount[STATUSES.CANCELLED] || 0],
    ['', '', '', ''],
    ['🏆 Top 10 รายการขายดี', '', '', ''],
    ['อันดับ', 'รายการ', 'ขายได้ (ครั้ง)', '']
  ];
  top.forEach(function (t, i) { rows.push([i + 1, t[0], t[1], '']); });
  sh.getRange(4, 1, rows.length, 4).setValues(rows);

  // ตกแต่งหัวข้อ
  [4, 11, 15].forEach(function (r) {
    sh.getRange(r, 1, 1, 4).merge().setFontWeight('bold').setBackground('#ede9fe').setFontColor('#5b21b6');
  });
  sh.getRange(5, 1, 1, 3).setFontWeight('bold').setBackground('#f5f3ff');
  sh.getRange(16, 1, 1, 3).setFontWeight('bold').setBackground('#fff7ed').setFontColor('#9a3412');
  sh.getRange(6, 2, 4, 1).setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#15803d');
  sh.setColumnWidth(1, 200).setColumnWidth(2, 150).setColumnWidth(3, 150).setColumnWidth(4, 150);
  sh.setHiddenGridlines(true);
}

function generateCustomerSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = '👤 ประวัติลูกค้า';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clear();

  // รวมข้อมูลจาก Orders (ไม่นับออเดอร์ยกเลิก)
  const map = {};
  readOrders_().forEach(function (o) {
    if (!o.email || o.status === STATUSES.CANCELLED) return;
    const c = map[o.email] || { name: o.customer, email: o.email, count: 0, total: 0,
      last: '', orderIds: [], subjects: {} };
    c.name = o.customer || c.name;
    c.count++;
    if (o.status === STATUSES.COMPLETED || o.status === STATUSES.APPROVED) c.total += o.netTotal;
    if (o.timestamp > c.last) c.last = o.timestamp;
    c.orderIds.push(o.orderId);
    (o.items || []).forEach(function (it) {
      const k = it.subject + (it.grade ? ' ' + it.grade : '');
      c.subjects[k] = (c.subjects[k] || 0) + 1;
    });
    map[o.email] = c;
  });
  const customers = Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) {
      // เรียงชื่อภาษาอังกฤษ A-Z ก่อน แล้วต่อด้วยภาษาไทย ก-ฮ
      const an = String(a.name || '').trim(), bn = String(b.name || '').trim();
      const aThai = /^[\u0E00-\u0E7F]/.test(an) ? 1 : 0;
      const bThai = /^[\u0E00-\u0E7F]/.test(bn) ? 1 : 0;
      if (aThai !== bThai) return aThai - bThai;
      return an.localeCompare(bn, 'th');
    });
  const repeat = customers.filter(function (c) { return c.count > 1; }).length;

  sh.getRange('A1:H1').merge().setValue('👤 ประวัติลูกค้า — ' + PAGE_NAME + '™')
    .setBackground('#818cf8').setFontColor('#fff').setFontWeight('bold')
    .setFontSize(15).setHorizontalAlignment('center');
  sh.getRange('A2:H2').merge().setValue('ลูกค้าทั้งหมด ' + customers.length + ' คน  |  ซื้อซ้ำ ' +
    repeat + ' คน  |  อัปเดต ' + ymd_(new Date()))
    .setFontStyle('italic').setFontColor('#6b7280').setFontSize(9).setHorizontalAlignment('center');

  sh.getRange(4, 1, 1, 8).setValues([[
    '#', 'ชื่อ Facebook', 'Gmail', 'จำนวนครั้ง', 'ยอดซื้อรวม (บาท)', 'ครั้งล่าสุด', 'ออเดอร์ทั้งหมด', 'รายการที่เคยซื้อ'
  ]]).setFontWeight('bold').setBackground('#ede9fe').setFontColor('#5b21b6');

  if (customers.length) {
    const rows = customers.map(function (c, i) {
      const subjList = Object.keys(c.subjects).map(function (k) {
        return c.subjects[k] > 1 ? k + ' (x' + c.subjects[k] + ')' : k;
      }).join(', ');
      return [i + 1, c.name, c.email, c.count, c.total, c.last.slice(0, 10),
        c.orderIds.join(', '), subjList];
    });
    const range = sh.getRange(5, 1, rows.length, 8);
    range.setValues(rows);
    sh.getRange(5, 5, rows.length, 1).setNumberFormat('#,##0').setFontWeight('bold').setFontColor('#15803d');
    sh.getRange(5, 6, rows.length, 1).setNumberFormat('@');
    // ไฮไลต์ลูกค้าซื้อซ้ำ
    const bg = rows.map(function (r) { return [r[3] > 1 ? '#fef9c3' : '#ffffff']; });
    sh.getRange(5, 4, rows.length, 1).setBackgrounds(bg);
  }
  sh.setFrozenRows(4);
  sh.setColumnWidth(2, 170).setColumnWidth(3, 200).setColumnWidth(7, 230).setColumnWidth(8, 420);
  sh.getRange(5, 8, Math.max(customers.length, 1), 1).setWrap(true);
}


/**
 * ============================================================
 *  📖 หน้าตัวอย่างสื่อสาธารณะ (Preview) — รวมในระบบขาย v9
 * ============================================================
 *  ลูกค้ากดปุ่ม "ตัวอย่าง" ในหน้าร้าน → เปิดหน้านี้ (ไม่ต้อง login)
 *  แสดง: ตัวอย่างสื่อ HTML + ตัวอย่างแผน PDF (ฝังดูในหน้า)
 *  ติดตั้ง: สร้างไฟล์ HTML ชื่อ "preview" แล้ววางโค้ด preview.html
 *          → Deploy เวอร์ชันใหม่ → รัน setupPreviewLinks ครั้งเดียว
 * ============================================================
 */

// ⚙️ ตั้งค่า
const PREVIEW_PARENT_ID = '1x06_YvNy7A3oPlsDwldy6hx7ht7QU1kX';  // โฟลเดอร์ "Ex.แผน/สื่อ เพจครูพร้อมสอน"
// ⭐ ลิงก์ /exec ของเว็บแอประบบขายนี้ (ต้องตรงกับ deployment ปัจจุบัน)
const PREVIEW_SELF_URL  = 'https://script.google.com/macros/s/AKfycby6eRRZw17H8fNCTczlMsPErOLUn1pBK4qU_6lNJhzHbj-DstyX_wT5_lz0elcQ8TMI1w/exec';

// เสิร์ฟไฟล์สื่อ HTML จาก Drive (อ่านในฐานะเจ้าของ ลูกค้าไม่ต้องมีสิทธิ์)
function servePreviewMedia_(id) {
  try {
    var html = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
    return HtmlService.createHtmlOutput(html)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput('<h2 style="text-align:center;margin-top:60px;font-family:sans-serif">เปิดไฟล์สื่อไม่ได้</h2>');
  }
}

// เสิร์ฟหน้า preview (ฝังข้อมูลไฟล์ที่สแกนได้ลงไปเลย ไม่ต้องเรียกซ้ำ)
function servePreviewPage_(e) {
  var p = e.parameter || {};
  var data = getPreviewData_(p.subject, p.level, p.cur, p.hours);
  var t = HtmlService.createTemplateFromFile('preview');
  t.data = JSON.stringify(data);
  t.selfUrl = PREVIEW_SELF_URL;
  return t.evaluate()
    .setTitle('ตัวอย่างสื่อ ' + (data.title || ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- ตัวช่วยจับคู่โฟลเดอร์/ไฟล์ ---------- */
function pv_normName_(s) { return String(s || '').replace(/\s+/g, '').trim(); }
function pv_normLvl_(s) { return String(s || '').replace(/\s+/g, '').trim(); }
function pv_lvlFromName_(n) { var m = String(n).match(/(ป|ม|อ)\.?\s*(\d+)/); return m ? (m[1] + '.' + m[2]) : ''; }
// คืน true ถ้าชื่อโฟลเดอร์ครอบคลุมระดับชั้นที่ต้องการ (รองรับช่วง เช่น ป.1-3 / ป.4-6)
function pv_lvlMatches_(folderName, wantLevel) {
  var wm = pv_normLvl_(wantLevel).match(/(ป|ม|อ)\.?(\d+)/);
  if (!wm) return false;
  var prefix = wm[1], wnum = Number(wm[2]), s = String(folderName);
  var rangeRe = new RegExp(prefix + '\\.?\\s*(\\d+)\\s*-\\s*(\\d+)', 'g'), rm;
  while ((rm = rangeRe.exec(s))) { if (wnum >= Number(rm[1]) && wnum <= Number(rm[2])) return true; }
  return new RegExp(prefix + '\\.?\\s*' + wnum + '(?!\\d)').test(s);
}
function pv_curFromName_(n) {
  var s = String(n), m = s.match(/หลักสูตร\s*(\d+)/);
  if (m) return m[1];
  var m2 = s.match(/\((\d{2,4})\)\s*$/);   // เช่น "... (68)" ท้ายชื่อ (โฟลเดอร์กลุ่ม VIP)
  return m2 ? m2[1] : '';
}
function pv_parseSubj_(name) {
  var s = String(name);
  var h = '';
  var mh = s.match(/(\d+)\s*[ซช]ม\.?/); if (mh) h = mh[1];      // ดึงชั่วโมง
  s = s
    .replace(/\([^)]*\)/g, ' ')               // ตัดวงเล็บ (...)
    .replace(/(\d+)\s*[ซช]ม\.?/g, ' ')        // ตัด "200 ชม."
    .replace(/หลักสูตร\s*\d+/g, ' ')          // ตัด "หลักสูตร 68"
    .replace(/(ป|ม|อ)\s*\.?\s*\d+/g, ' ')     // ตัดระดับชั้น ป.1 / ม.3 / อ.2
    .replace(/\s+/g, ' ')
    .replace(/\s*\d+\s*$/, '')                 // ตัดเลขท้ายที่หลงเหลือ
    .trim();
  return { name: s, hours: h };
}

// สแกนโฟลเดอร์ตัวอย่างของวิชาที่ตรง → คืนรายการ PDF + HTML
// เวอร์ชัน public (เรียกผ่าน API จากหน้า preview บน GitHub ได้)
function getPreviewData(subject, level, cur, hours) {
  return getPreviewData_(subject, level, cur, hours);
}

// อ่านเนื้อหาไฟล์สื่อ HTML ในฐานะเจ้าของ → ส่งให้หน้า preview เปิดเป็นแท็บใหม่บนเครื่องผู้ใช้
// (เสถียรกว่าเสิร์ฟผ่าน HtmlService — ไม่ติดปัญหาสิทธิ์/หลายบัญชี)
function getPreviewMediaContent(id) {
  try {
    var file = DriveApp.getFileById(id);
    var nm = file.getName();
    if (!/\.html?$/i.test(nm) && !/html/i.test(nm)) return { ok: false, error: 'ไม่ใช่ไฟล์สื่อ HTML' };
    if (!pv_isUnderPreview_(file)) return { ok: false, error: 'ไฟล์อยู่นอกโฟลเดอร์ตัวอย่าง' };
    var html = file.getBlob().getDataAsString('UTF-8');
    return { ok: true, html: html, name: nm };
  } catch (e) {
    return { ok: false, error: 'เปิดสื่อไม่ได้: ' + (e && e.message || e) };
  }
}

// ตรวจว่าไฟล์อยู่ภายใต้โฟลเดอร์ตัวอย่าง (PREVIEW_PARENT_ID) หรือไม่ — กันการอ่านไฟล์อื่นของเจ้าของ
function pv_isUnderPreview_(file) {
  try {
    var seen = {}, queue = [file];
    for (var depth = 0; depth < 6 && queue.length; depth++) {
      var next = [];
      for (var i = 0; i < queue.length; i++) {
        var ps = queue[i].getParents();
        while (ps.hasNext()) {
          var par = ps.next(), pid = par.getId();
          if (pid === PREVIEW_PARENT_ID) return true;
          if (!seen[pid]) { seen[pid] = 1; next.push(par); }
        }
      }
      queue = next;
    }
  } catch (e) {}
  return false;
}

// เก็บไฟล์ PDF/HTML จากโฟลเดอร์ (และโฟลเดอร์ย่อย 1 ชั้น ถ้า deep=true)
function pv_collectFiles_(folder, deep, pdfs, htmls) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next(), nm = f.getName(), mt = f.getMimeType();
    if (mt === 'application/pdf' || /\.pdf$/i.test(nm)) {
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}  // ให้ฝัง/เปิด PDF ได้แบบสาธารณะ
      pdfs.push({ name: nm.replace(/\.pdf$/i, ''), id: f.getId() });
    } else if (/\.html?$/i.test(nm) || /html/i.test(nm)) {
      htmls.push({ name: nm.replace(/^สื่อ\s*HTML\s*/i, '').replace(/\.html?$/i, ''), id: f.getId(), no: Number((nm.match(/plan\s*(\d+)/i) || [])[1] || 0) });
    }
  }
  if (deep) {
    var subs = folder.getFolders();
    while (subs.hasNext()) pv_collectFiles_(subs.next(), false, pdfs, htmls);
  }
}

function getPreviewData_(subject, level, cur, hours) {
  var want = {
    name: pv_normName_(subject), level: pv_normLvl_(level),
    cur: String(cur || '').replace(/[^\d]/g, ''), hours: String(hours || '').replace(/[^\d]/g, '')
  };
  var title = String(subject || ''), subtitle = (level || '') + (cur ? ' · หลักสูตร ' + cur : '') + (hours ? ' · ' + hours + ' ชม.' : '');

  var parent;
  try { parent = DriveApp.getFolderById(PREVIEW_PARENT_ID); }
  catch (e) { return { ok: false, error: 'เปิดโฟลเดอร์ตัวอย่างไม่ได้', title: title, subtitle: subtitle, pdfs: [], htmls: [] }; }

  // ── กรณีกลุ่ม VIP: จับคู่ด้วยคำว่า "VIP" + ชั้น (ไม่สนหลักสูตร) แล้วรวมไฟล์จากโฟลเดอร์ + โฟลเดอร์ย่อย ──
  if (/vip/i.test(String(subject))) {
    var vipFolder = null, vf = parent.getFolders();
    while (vf.hasNext()) {
      var vg = vf.next(), vgn = vg.getName();
      if (/vip/i.test(vgn) && pv_lvlMatches_(vgn, level || subject)) { vipFolder = vg; break; }
    }
    if (!vipFolder) return { ok: false, error: 'ยังไม่มีโฟลเดอร์ตัวอย่างกลุ่ม VIP ของชั้นนี้', title: title, subtitle: subtitle, pdfs: [], htmls: [] };
    var vpdfs = [], vhtmls = [];
    pv_collectFiles_(vipFolder, true, vpdfs, vhtmls);
    vpdfs.sort(function (a, b) { return a.name.localeCompare(b.name, 'th'); });
    vhtmls.sort(function (a, b) { return a.no - b.no; });
    return { ok: true, title: subject, subtitle: subtitle, pdfs: vpdfs, htmls: vhtmls };
  }

  // 1) รวมโฟลเดอร์ชั้น+หลักสูตรที่ตรง "ทั้งหมด" (อาจมีหลายอัน เช่น โฟลเดอร์เอกสาร + โฟลเดอร์แผน — รองรับช่วงชั้น เช่น ป.1-3)
  //    ถ้าสินค้าไม่ได้ระบุหลักสูตร (want.cur ว่าง เช่น ระบบงาน) → จับคู่ด้วยชั้นอย่างเดียว
  var gradeFolders = [], grades = parent.getFolders();
  while (grades.hasNext()) {
    var g = grades.next(), gn = g.getName();
    if ((!want.cur || pv_curFromName_(gn) === want.cur) && pv_lvlMatches_(gn, level)) gradeFolders.push(g);
  }
  if (!gradeFolders.length) return { ok: false, error: 'ยังไม่มีโฟลเดอร์ตัวอย่างของชั้น/หลักสูตรนี้', title: title, subtitle: subtitle, pdfs: [], htmls: [] };

  // 2) หาโฟลเดอร์วิชาจาก "ทุกโฟลเดอร์ชั้นที่ตรง" — เทียบแบบชื่อขึ้นต้นด้วยชื่อวิชา + ให้คะแนน
  //    คะแนน: ชั่วโมงตรง (+1000) > ชื่อตรงพอดี ไม่ใช่ตัวแปรในวงเล็บ เช่น (บูรณาการ) (+100)
  var subjFolder = null, bestScore = -1;
  for (var gi = 0; gi < gradeFolders.length; gi++) {
    var subs = gradeFolders[gi].getFolders();
    while (subs.hasNext()) {
      var s = subs.next(), folderNorm = pv_normName_(s.getName());
      if (!want.name || folderNorm.indexOf(want.name) !== 0) continue;
      var sp = pv_parseSubj_(s.getName());
      var rest = folderNorm.slice(want.name.length).replace(/^\s+/, '');
      var score = 0;
      if (want.hours && sp.hours === want.hours) score += 1000;
      if (rest.charAt(0) !== '(') score += 100;
      if (score > bestScore) { bestScore = score; subjFolder = s; }
    }
  }
  if (!subjFolder) return { ok: false, error: 'ยังไม่มีโฟลเดอร์ตัวอย่างของวิชานี้', title: title, subtitle: subtitle, pdfs: [], htmls: [] };

  // 3) แยกไฟล์ PDF / HTML
  var pdfs = [], htmls = [];
  pv_collectFiles_(subjFolder, false, pdfs, htmls);
  pdfs.sort(function (a, b) { return a.name.localeCompare(b.name, 'th'); });
  htmls.sort(function (a, b) { return a.no - b.no; });
  return { ok: true, title: want.name ? subject : title, subtitle: subtitle, pdfs: pdfs, htmls: htmls };
}

/* ============================================================
 *  เครื่องมือ: เติมลิงก์ "ตัวอย่าง" (คอลัมน์ H) ของทุกวิชาอัตโนมัติ
 *  รันครั้งเดียวหลัง Deploy — ปุ่ม "ตัวอย่าง" ในหน้าร้านจะเปิดหน้า preview เอง
 * ============================================================ */
function setupPreviewLinks() {
  var n = 0;
  productSheetNames_().forEach(function (name) {   // เติมทุกชีตสินค้า (รองรับทั้งสองโครง)
    var sh = getSheet_(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;
    if (isVipSheet_(name)) {   // โครงย่อ: ลิงก์ตัวอย่างจากชื่อรายการ (ตัวจับคู่ฝั่ง preview รองรับอยู่แล้ว)
      var vr = sh.getRange(2, 1, last - 1, PV_COUNT);
      var vv = vr.getValues();
      for (var vi = 0; vi < vv.length; vi++) {
        var vsub = String(vv[vi][PV.SUBJECT - 1] || '').trim();
        if (!vsub) continue;
        vv[vi][PV.EXAMPLE_LINK - 1] = PREVIEW_PAGE_URL + '?subject=' + encodeURIComponent(vsub) + '&level=&cur=&hours=';
        n++;
      }
      vr.setValues(vv);
      vr.offset(0, PV.ADDED - 1, vv.length, 1).setNumberFormat('@');
      return;
    }
    var rng = sh.getRange(2, 1, last - 1, P.CATEGORY);
    var vals = rng.getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var subject = String(r[P.SUBJECT - 1] || '').trim();
      if (!subject) continue;
      var level = String(r[P.GRADE - 1] || '').trim();
      var cur = String(r[P.CURRICULUM - 1] || '').trim();
      var hours = String(r[P.HOURS - 1] || '').replace(/[^\d]/g, '');
      var url = PREVIEW_PAGE_URL + '?subject=' + encodeURIComponent(subject)
        + '&level=' + encodeURIComponent(level)
        + '&cur=' + encodeURIComponent(cur)
        + '&hours=' + encodeURIComponent(hours);
      r[P.EXAMPLE_LINK - 1] = url;
      n++;
    }
    rng.setValues(vals);
  });
  if (!n) { try { SpreadsheetApp.getUi().alert('ยังไม่มีวิชาในชีต'); } catch (e) {} return; }
  clearShopCache_();
  try { SpreadsheetApp.getUi().alert('เติมลิงก์ตัวอย่างแล้ว ' + n + ' วิชา\n\nปุ่ม "ตัวอย่าง" ในหน้าร้านจะเปิดหน้า preview ให้ลูกค้าอัตโนมัติ'); } catch (e) {}
}


// 🔎 เครื่องมือตรวจโครงโฟลเดอร์ตัวอย่าง — ดูชื่อจริงเทียบกับค่าที่ระบบแปลง (level/cur/subject/hours)
function previewDebug() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('🔎 Preview Debug') || ss.insertSheet('🔎 Preview Debug');
  sh.clear();
  var rows = [['ประเภท', 'ชื่อโฟลเดอร์จริง', 'level (แปลง)', 'cur (แปลง)', 'subject (แปลง)', 'hours (แปลง)', 'ไฟล์ข้างใน']];
  var parent;
  try { parent = DriveApp.getFolderById(PREVIEW_PARENT_ID); }
  catch (e) { sh.getRange(1, 1).setValue('❌ เปิดโฟลเดอร์แม่ไม่ได้: ' + e); return; }
  rows.push(['📁 แม่', parent.getName(), '', '', '', '', '']);
  var grades = parent.getFolders();
  while (grades.hasNext()) {
    var g = grades.next(), gn = g.getName();
    var lvl = pv_normLvl_(pv_lvlFromName_(gn)), cur = pv_curFromName_(gn);
    // แสดงไฟล์ที่อยู่ในโฟลเดอร์ชั้น/VIP โดยตรงด้วย (เช่น สื่อกลุ่ม VIP มักวางไฟล์ตรงๆ ไม่มีโฟลเดอร์วิชา)
    var gFileNames = [], gfs = g.getFiles(), gc = 0;
    while (gfs.hasNext() && gc < 25) { gFileNames.push(gfs.next().getName()); gc++; }
    rows.push(['📂 ชั้น/หลักสูตร', gn, lvl, cur, (/vip/i.test(gn) ? '⭐VIP' : ''), '', gFileNames.join('  |  ')]);
    var subs = g.getFolders();
    while (subs.hasNext()) {
      var s = subs.next(), sp = pv_parseSubj_(s.getName());
      var fileNames = [], fs = s.getFiles(), c = 0;
      while (fs.hasNext() && c < 25) { fileNames.push(fs.next().getName()); c++; }
      rows.push(['📄 วิชา', s.getName(), lvl, cur, sp.name, sp.hours, fileNames.join('  |  ')]);
    }
  }
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, rows[0].length); } catch (e) {}
  ss.setActiveSheet(sh);
  try { SpreadsheetApp.getUi().alert('สร้างชีต "🔎 Preview Debug" แล้ว\n\nดูคอลัมน์ "ชื่อโฟลเดอร์จริง" เทียบกับค่าที่ระบบแปลง (level/cur/subject/hours)\nถ้า subject/level/cur ไม่ตรงกับวิชาในร้าน = สาเหตุที่หาไม่เจอ'); } catch (e) {}
}
