// =====================================================================
// 🟢 การตั้งค่า LINE OA (Messaging API)
// =====================================================================
const LINE_ACCESS_TOKEN = 'xxxx'; 
const LINE_CHANNEL_TOKEN = 'xxxx'; 
const ADMIN_USER_ID = 'xxxx';

// =====================================================================
// ⚙️ ส่วนตั้งค่า ID ของเอกสารและโฟลเดอร์
// =====================================================================
const CONFIG = {
  'form_fuel_provincial': { templateId: '1g-Satcf-dBWRr9CT2OVo-bXKBpMrdKVV3cf9Ip8plSs', folderId: '1IAypieU1VqOfCIKS5ljffnLyAfD9OqYD', sheetName: 'Records_Fuel_Prov' },
  'form_advance_provincial': { templateId: '14GzQW6njQayPFyTBKlRmeZ9JBAyS5BsY6sPKM3-Cj5s', folderId: '18XtxNRZfPRm9BNXQPMA6PEUMSkLsRl5P', sheetName: 'Records_Expense_Prov' },
  'form_advance_general': { templateId: '1yET6zXzKLq3UJjjgD7hsw6yZ5b22zdNTh-IuSHSI72w', folderId: '1PoVceNfd9euC_poJWcB4FXSCbctU4f_j', sheetName: 'Records_Advance' },
  'form_fuel_bkk': { templateId: '14SuLwjH_az3Saw5m-c5D99U3HKI_2xALs_cAcoZORLg', folderId: '11NAGifRdAVahqhsS3u-c2n4ef2vMQiDW', sheetName: 'Records_Fuel_BKK' },
  'form_tollway': { templateId: '1mZGgZ1JxWZZ0ksB2W04p0Hb6eckVxGcLtkaL4xkQzo8', folderId: '1ldUUjrh4E1Vbm7jKI5BZ7e5i-NOVQRNi', sheetName: 'Records_Tollway' },
  'form_ot': { templateId: '1oX6C8rFxkSr5CK13ys2n2vvadv16YT-Ho-FQRgK6sLw', folderId: '1ANZete9SrUdfZCTsVrNSl7TPAL7Jnsry', sheetName: 'Records_OT' },
  'form_req_ot': { templateId: '1FwswCyPCO9NBFt3Rcj_0r9LnojNWy7ySHozTHyz_2NM', folderId: '1ypcPR7iYMpMa0e93yf8d2lXZddjD447D', sheetName: 'Records_Req_OT' }
};

const ATTACHMENT_FOLDER_ID = '1xG4TKi_WCCUfjpPWnj3VJnFs_WIGBdCa';

// จำนวนวันที่เอกสารค้างรออนุมัติ ก่อนจะแจ้งเตือนแอดมินผ่าน LINE
const OVERDUE_DAYS_THRESHOLD = 3;
// จำนวนวันที่แบบร่างค้างไว้โดยยังไม่ส่ง ก่อนจะเตือนเจ้าของแบบร่างผ่าน LINE
const STALE_DRAFT_DAYS_THRESHOLD = 5;
// เตือนล่วงหน้ากี่วันก่อนถึง "วันที่เคลียร์เงิน" ของเอกสารเบิกเงินสำรองที่ยังรออนุมัติ
const CLEAR_DEADLINE_REMINDER_DAYS = 2;

function doGet() { 
  return HtmlService.createHtmlOutputFromFile('index').setTitle('E-Document System').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

// =====================================================================
// 🤖 Webhook: รับข้อความจาก LINE OA (สำหรับผูกบัญชี)
// =====================================================================
function doPost(e) {
  try {
    let event = JSON.parse(e.postData.contents).events[0];
    if (event.type === 'message' && event.message.type === 'text') {
      let userMessage = event.message.text.trim();
      let userId = event.source.userId;
      let replyToken = event.replyToken;
      
      if (userMessage.startsWith("ผูกบัญชี")) {
        let username = userMessage.replace("ผูกบัญชี", "").trim();
        if (username === "") {
          replyLineMessage(replyToken, "⚠️ กรุณาพิมพ์คำว่า 'ผูกบัญชี' ตามด้วย Username ของคุณในระบบครับ\n\nตัวอย่าง: ผูกบัญชี natthapong");
        } else {
          let result = saveLineUID(username, userId);
          if (result.success) {
            replyLineMessage(replyToken, `✅ ผูกบัญชีสำเร็จ!\nยินดีต้อนรับคุณ ${result.name} เข้าสู่ระบบแจ้งเตือนของ ATE ครับ 🎉`);
          } else {
            replyLineMessage(replyToken, `❌ ผูกบัญชีไม่สำเร็จ\nเหตุผล: ${result.message}`);
          }
        }
      } else {
        replyLineMessage(replyToken, "👋 สวัสดีครับ! ระบบเอกสาร ATE ยินดีให้บริการ\n\nหากต้องการรับการแจ้งเตือนผลเอกสาร กรุณาพิมพ์คำว่า\n\n👉 'ผูกบัญชี [Username ของคุณ]'\n\nเพื่อลงทะเบียนครับผม");
      }
    }
  } catch (error) { console.log("Webhook Error: " + error); }
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

function saveLineUID(username, uid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users'); 
  if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลพนักงาน" };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let userCol = headers.findIndex(h => h.toString().toLowerCase().trim() === 'username');
  let nameCol = headers.findIndex(h => h.toString().toLowerCase().trim() === 'fullname');
  let uidCol = headers.findIndex(h => h.toString().toLowerCase().trim() === 'line uid' || h.toString().toLowerCase().trim() === 'lineuid');
  
  if (userCol === -1 || uidCol === -1) return { success: false, message: "ตั้งค่าคอลัมน์ใน Sheet ไม่ครบ (ขาด Username หรือ LINE UID)" };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userCol].toString().toLowerCase().trim() === username.toLowerCase().trim()) {
      sheet.getRange(i + 1, uidCol + 1).setValue(uid); 
      let employeeName = (nameCol !== -1) ? data[i][nameCol] : "";
      if (!employeeName || employeeName.toString().trim() === "") employeeName = username; 
      return { success: true, name: employeeName };
    }
  }
  return { success: false, message: "ไม่พบ Username นี้ในระบบครับ โปรดตรวจสอบตัวสะกด" };
}

function replyLineMessage(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    'method': 'post',
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': [{'type': 'text', 'text': text}] })
  };
  UrlFetchApp.fetch(url, options);
}

function sendLineBot(message, userId = ADMIN_USER_ID) {
  if (!LINE_CHANNEL_TOKEN || LINE_CHANNEL_TOKEN.includes('ใส่_')) return;
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = { to: userId, messages: [{ type: 'text', text: message }] };
  const options = {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + LINE_CHANNEL_TOKEN },
    payload: JSON.stringify(payload)
  };
  try { UrlFetchApp.fetch(url, options); } catch (e) { console.error('LINE Bot Error: ' + e.message); }
}

function testLineBot() {
  sendLineBot("🎉 ฮัลโหล! นี่คือข้อความทดสอบจากระบบเอกสาร ATE ครับ บอทเชื่อมต่อสำเร็จแล้ว!");
}

// =====================================================================
// 🔐 ระบบ Login & Profile
// =====================================================================
function checkLogin(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return { success: false, message: 'ไม่พบแท็บข้อมูล Users' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == username && data[i][1] == password) {
        let role = data[i][4] ? data[i][4].toString().toLowerCase() : 'user';
        return { success: true, username: data[i][0], password: data[i][1], fullName: data[i][2], department: data[i][3], role: role, position: data[i][5] || '', plate: data[i][6] || '', vType: data[i][7] || '', vBrand: data[i][8] || '' };
      }
    } 
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (error) { return { success: false, message: error.message }; }
}

function updateUserProfile(userData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == userData.originalUsername) { 
        let currentRole = data[i][4]; 
        sheet.getRange(i + 1, 1, 1, 9).setValues([[ userData.username, userData.password, userData.fullName, userData.department, currentRole, userData.position, userData.plate, userData.vType, userData.vBrand ]]);
        return { success: true, message: 'บันทึกข้อมูลส่วนตัวเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลผู้ใช้งานในระบบ' };
  } catch(e) { return { success: false, message: e.message }; }
}

// =====================================================================
// 📊 ฟังก์ชันดึงสถิติหน้า Dashboard
// =====================================================================
function getDashboardStats(username, role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let stats = { total: 0, advance: 0, fuelProv: 0, expProv: 0, fuelBkk: 0, tollway: 0, ot: 0, reqOt: 0 };
  const sheetsConfig = [
    { name: 'Records_Advance', key: 'advance' }, { name: 'Records_Fuel_Prov', key: 'fuelProv' },
    { name: 'Records_Expense_Prov', key: 'expProv' }, { name: 'Records_Fuel_BKK', key: 'fuelBkk' },
    { name: 'Records_Tollway', key: 'tollway' }, { name: 'Records_OT', key: 'ot' },
    { name: 'Records_Req_OT', key: 'reqOt' }
  ];
  sheetsConfig.forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    if (sheet) {
      let data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (role === 'admin' || data[i][2] === username) {
          stats[config.key]++; stats.total++;
        }
      }
    }
  });
  return stats;
}

// =====================================================================
// 📈 แนวโน้มรายเดือน (สำหรับกราฟเส้นในแดชบอร์ด)
// =====================================================================
function parseRecordDate(str) {
  if (!str) return null;
  str = str.toString().trim();
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); // YYYY-MM-DD (จาก input type=date)
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // DD/MM/YYYY (จาก Utilities.formatDate)
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  let d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function getMonthlyTrend(username, role, monthsBack) {
  monthsBack = monthsBack || 6;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Records_Advance', 'Records_Fuel_Prov', 'Records_Expense_Prov', 'Records_Fuel_BKK', 'Records_Tollway', 'Records_OT', 'Records_Req_OT'];
  const thaiMonths = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

  const now = new Date();
  let buckets = []; // [{key:'2026-01', label:'ม.ค. 69', count:0}, ...] เรียงเก่า -> ใหม่
  for (let i = monthsBack - 1; i >= 0; i--) {
    let d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    let key = d.getFullYear() + '-' + (d.getMonth() + 1);
    let label = thaiMonths[d.getMonth()] + ' ' + ((d.getFullYear() + 543).toString().slice(-2));
    buckets.push({ key: key, label: label, count: 0 });
  }
  let bucketMap = {};
  buckets.forEach(b => bucketMap[b.key] = b);

  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) return;
    let data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (role !== 'admin' && data[i][2] !== username) continue;
      let d = parseRecordDate(data[i][1]);
      if (!d) continue;
      let key = d.getFullYear() + '-' + (d.getMonth() + 1);
      if (bucketMap[key]) bucketMap[key].count++;
    }
  });

  return { labels: buckets.map(b => b.label), counts: buckets.map(b => b.count) };
}

// =====================================================================
// ⏰ แจ้งเตือนเอกสารค้างอนุมัตินานเกินไป (ทำงานผ่าน Time-driven Trigger)
// =====================================================================
function checkOverdueDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = [
    { name: 'Records_Advance', type: 'เบิกเงินสำรอง' }, { name: 'Records_Fuel_Prov', type: 'ค่าน้ำมัน ตจว.' },
    { name: 'Records_Expense_Prov', type: 'สรุปค่าใช้จ่าย ตจว.' }, { name: 'Records_Fuel_BKK', type: 'เบิกน้ำมัน กทม.' },
    { name: 'Records_Tollway', type: 'เบิกทางด่วน' }, { name: 'Records_OT', type: 'สรุปทำงานล่วงเวลา' },
    { name: 'Records_Req_OT', type: 'ขออนุมัติทำ OT' }
  ];
  const now = new Date();
  let overdue = [];

  sheetsToCheck.forEach(s => {
    const sheet = ss.getSheetByName(s.name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('สถานะ');
    for (let i = 1; i < data.length; i++) {
      const status = (statusCol > -1 && data[i][statusCol]) ? data[i][statusCol].toString() : 'รออนุมัติ';
      if (status !== 'รออนุมัติ') continue;
      const d = parseRecordDate(data[i][1]);
      if (!d) continue;
      const daysPending = Math.floor((now - d) / (1000 * 60 * 60 * 24));
      if (daysPending >= OVERDUE_DAYS_THRESHOLD) {
        overdue.push({ docId: data[i][0], type: s.type, name: data[i][2], days: daysPending });
      }
    }
  });

  if (overdue.length === 0) return { success: true, count: 0 };

  let msg = `⏰ แจ้งเตือนเอกสารค้างอนุมัติ!\n\nมีเอกสารรออนุมัตินานเกิน ${OVERDUE_DAYS_THRESHOLD} วัน จำนวน ${overdue.length} รายการ:\n\n`;
  overdue.slice(0, 15).forEach(o => {
    msg += `📄 ${o.docId} (${o.type})\n👤 ${o.name} — ค้าง ${o.days} วัน\n\n`;
  });
  if (overdue.length > 15) msg += `...และอีก ${overdue.length - 15} รายการ\n\n`;
  msg += `รบกวนแอดมินช่วยตรวจสอบด้วยนะครับ 🙏`;

  let adminUIDs = getAdminUIDs();
  if (adminUIDs.length > 0) {
    adminUIDs.forEach(uid => sendLinePushMessage(uid, msg));
  } else if (typeof sendLineBot === "function") {
    sendLineBot(msg + "\n(ยังไม่มีแอดมินผูกบัญชีส่วนตัว)");
  }

  // ทำงานต่อเนื่องในทริกเกอร์เดียวกัน ไม่ต้องตั้งทริกเกอร์แยก
  checkStaleDrafts();
  checkUpcomingClearDeadlines();

  return { success: true, count: overdue.length };
}

// เรียกใช้ครั้งเดียวจาก Apps Script editor (เลือกฟังก์ชันนี้แล้วกด Run) เพื่อตั้งเวลาให้เช็กทุกวันอัตโนมัติ
// (ครอบคลุมทั้งเอกสารค้างอนุมัติ, แบบร่างค้างไว้นาน, และกำหนดเคลียร์เงินที่ใกล้ถึง)
function setupOverdueTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkOverdueDocuments') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkOverdueDocuments').timeBased().everyDays(1).atHour(9).create();
  return { success: true, message: 'ตั้งเวลาแจ้งเตือนเรียบร้อยแล้ว จะเช็กทุกวันเวลาประมาณ 9 โมงเช้า' };
}

function parseDraftDate(str) {
  if (!str) return null;
  let m = str.toString().trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]), parseInt(m[4]), parseInt(m[5]), m[6] ? parseInt(m[6]) : 0);
  let d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// เตือนพนักงานเรื่องแบบร่างที่ค้างไว้นานโดยยังไม่ได้ส่ง
function checkStaleDrafts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Drafts');
  if (!sheet) return { success: true, notifiedUsers: 0 };
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  let staleByUser = {};

  for (let i = 1; i < data.length; i++) {
    const dateStr = data[i][1], username = data[i][2], formType = data[i][3];
    const d = parseDraftDate(dateStr);
    if (!d) continue;
    const days = Math.floor((now - d) / (1000 * 60 * 60 * 24));
    if (days >= STALE_DRAFT_DAYS_THRESHOLD) {
      if (!staleByUser[username]) staleByUser[username] = [];
      staleByUser[username].push({ formType: formType, days: days });
    }
  }

  let notifiedUsers = 0;
  Object.keys(staleByUser).forEach(username => {
    const uid = getUserUIDByName(username);
    if (!uid) return;
    const items = staleByUser[username];
    let msg = `📝 แจ้งเตือนแบบร่างค้างไว้นาน!\n\nคุณมีแบบร่างที่ยังไม่ได้ส่งค้างไว้ ${items.length} รายการ:\n\n`;
    items.slice(0, 10).forEach(it => { msg += `• ${it.formType} (ค้างไว้ ${it.days} วัน)\n`; });
    msg += `\nเข้าไปที่เมนู "แบบร่างของฉัน" เพื่อกรอกให้เสร็จแล้วส่งได้เลยครับ 😊`;
    sendLinePushMessage(uid, msg);
    notifiedUsers++;
  });
  return { success: true, notifiedUsers: notifiedUsers };
}

// เตือนพนักงานล่วงหน้าก่อนถึงกำหนด "วันที่เคลียร์เงิน" ของเอกสารเบิกเงินสำรองที่ยังรออนุมัติ
function checkUpcomingClearDeadlines() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Records_Advance');
  if (!sheet) return { success: true, count: 0 };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('สถานะ');
  const now = new Date();
  let upcoming = [];

  for (let i = 1; i < data.length; i++) {
    const status = (statusCol > -1 && data[i][statusCol]) ? data[i][statusCol].toString() : 'รออนุมัติ';
    if (status !== 'รออนุมัติ') continue;
    const clearDate = parseRecordDate(data[i][1]);
    if (!clearDate) continue;
    const daysLeft = Math.floor((clearDate - now) / (1000 * 60 * 60 * 24));
    if (daysLeft >= 0 && daysLeft <= CLEAR_DEADLINE_REMINDER_DAYS) {
      upcoming.push({ docId: data[i][0], name: data[i][2], daysLeft: daysLeft });
    }
  }

  upcoming.forEach(u => {
    const uid = getUserUIDByName(u.name);
    if (!uid) return;
    const msg = u.daysLeft === 0
      ? `⏳ วันนี้ถึงกำหนดเคลียร์เงินสำรองแล้วนะครับ!\n\n📄 รหัส: ${u.docId}\nอย่าลืมติดตามเรื่องเคลียร์เงินให้เรียบร้อยครับ 🙏`
      : `⏳ ใกล้ถึงกำหนดเคลียร์เงินสำรองแล้ว!\n\n📄 รหัส: ${u.docId}\nเหลืออีก ${u.daysLeft} วัน\nอย่าลืมติดตามเรื่องเคลียร์เงินด้วยนะครับ 🙏`;
    sendLinePushMessage(uid, msg);
  });
  return { success: true, count: upcoming.length };
}

// =====================================================================
// 📄 ดึงประวัติเอกสาร
// =====================================================================
function getUserDocuments(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let userDocs = [];
  const sheetsToRead = [ { name: 'Records_Advance', type: 'เบิกเงินสำรอง' }, { name: 'Records_Fuel_Prov', type: 'ค่าน้ำมัน ตจว.' }, { name: 'Records_Expense_Prov', type: 'สรุปค่าใช้จ่าย ตจว.' }, { name: 'Records_Fuel_BKK', type: 'เบิกน้ำมัน กทม.' }, { name: 'Records_Tollway', type: 'เบิกทางด่วน' }, { name: 'Records_OT', type: 'สรุปทำงานล่วงเวลา' }, { name: 'Records_Req_OT', type: 'ขออนุมัติทำ OT' } ];
  
  sheetsToRead.forEach(s => {
    const sheet = ss.getSheetByName(s.name);
    if(sheet) {
      const data = sheet.getDataRange().getDisplayValues();
      const headers = data[0];
      const statusCol = headers.indexOf('สถานะ'); const remarkCol = headers.indexOf('หมายเหตุ');

      for(let i = 1; i < data.length; i++) {
        if(data[i][2] === username) {
          let rowData = data[i].slice();
          while(rowData.length > 0 && rowData[rowData.length - 1].trim() === "") { rowData.pop(); }
          
          let len = rowData.length;
          let link1 = rowData[len - 2] || ''; let link2 = rowData[len - 1] || '';
          let pdfLink = link2; let attachments = link1;
          if ((pdfLink === "" || !pdfLink.includes("http")) && link1.includes("http")) { pdfLink = link1; attachments = "-"; }
          
          let status = (statusCol > -1 && data[i][statusCol]) ? data[i][statusCol] : 'รออนุมัติ';
          let remark = (remarkCol > -1 && data[i][remarkCol]) ? data[i][remarkCol] : '';

          userDocs.push({ docId: data[i][0], date: data[i][1], type: s.type, attachments: attachments, link: pdfLink, status: status, remark: remark });
        }
      }
    }
  }); 
  return userDocs.reverse();
}

function getAllDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let allDocs = [];
  const sheetsToRead = [ { name: 'Records_Advance', type: 'เบิกเงินสำรอง' }, { name: 'Records_Fuel_Prov', type: 'ค่าน้ำมัน ตจว.' }, { name: 'Records_Expense_Prov', type: 'สรุปค่าใช้จ่าย ตจว.' }, { name: 'Records_Fuel_BKK', type: 'เบิกน้ำมัน กทม.' }, { name: 'Records_Tollway', type: 'เบิกทางด่วน' }, { name: 'Records_OT', type: 'สรุปทำงานล่วงเวลา' }, { name: 'Records_Req_OT', type: 'ขออนุมัติทำ OT' } ];
  
  sheetsToRead.forEach(s => {
    const sheet = ss.getSheetByName(s.name);
    if(sheet) {
      const data = sheet.getDataRange().getDisplayValues();
      const headers = data[0]; const statusCol = headers.indexOf('สถานะ');
      
      for(let i = 1; i < data.length; i++) {
          let rowData = data[i].slice();
          while(rowData.length > 0 && rowData[rowData.length - 1].trim() === "") { rowData.pop(); }

          let len = rowData.length;
          let link1 = rowData[len - 2] || ''; let link2 = rowData[len - 1] || '';
          let pdfLink = link2; let attachments = link1;
          if ((pdfLink === "" || !pdfLink.includes("http")) && link1.includes("http")) { pdfLink = link1; attachments = "-"; }
          
          let status = (statusCol > -1 && data[i][statusCol]) ? data[i][statusCol] : 'รออนุมัติ';
          allDocs.push({ docId: data[i][0], type: s.type, date: data[i][1], name: data[i][2], attachments: attachments, link: pdfLink, status: status });
      }
    }
  }); 
  return allDocs.reverse();
}

function deleteUserDocument(docId, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheetName = "";
    if (type === 'เบิกเงินสำรอง') sheetName = 'Records_Advance';
    else if (type === 'ค่าน้ำมัน ตจว.') sheetName = 'Records_Fuel_Prov';
    else if (type === 'สรุปค่าใช้จ่าย ตจว.') sheetName = 'Records_Expense_Prov';
    else if (type === 'เบิกน้ำมัน กทม.') sheetName = 'Records_Fuel_BKK';
    else if (type === 'เบิกทางด่วน') sheetName = 'Records_Tollway';
    else if (type === 'สรุปทำงานล่วงเวลา') sheetName = 'Records_OT';
    else if (type === 'ขออนุมัติทำ OT') sheetName = 'Records_Req_OT';
    else return { success: false, message: "ไม่พบประเภทเอกสาร" };

    const sheet = ss.getSheetByName(sheetName); if (!sheet) return { success: false, message: "ไม่พบแท็บข้อมูล" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0] === docId) { sheet.deleteRow(i + 1); return { success: true }; } }
    return { success: false, message: "ไม่พบรหัสเอกสาร" };
  } catch(e) { return { success: false, message: e.message }; }
}

// =====================================================================
// 🛡️ ระบบ Admin Panel
// =====================================================================
function getAdminUsers() { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); const data = sheet.getDataRange().getDisplayValues(); let users = []; for(let i = 1; i < data.length; i++) { users.push({ row: i + 1, username: data[i][0], password: data[i][1], name: data[i][2], dept: data[i][3], role: data[i][4] || 'user', position: data[i][5] || '', plate: data[i][6] || '', vType: data[i][7] || '', vBrand: data[i][8] || '' }); } return users; }
function saveUser(user) { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); if(user.row) { sheet.getRange(user.row, 1, 1, 9).setValues([[user.username, user.password, user.name, user.dept, user.role, user.position, user.plate, user.vType, user.vBrand]]); } else { sheet.appendRow([user.username, user.password, user.name, user.dept, user.role, user.position, user.plate, user.vType, user.vBrand]); } return getAdminUsers(); }
function deleteUser(row) { const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); sheet.deleteRow(row); return getAdminUsers(); }

// =====================================================================
// 🛠️ ตัวช่วยแปลงวันที่
// =====================================================================
function formatThaiShortDate(dateString) {
  if (!dateString) return ''; const parts = dateString.split('-'); if (parts.length !== 3) return dateString;
  const shortYear = (parseInt(parts[0]) + 543).toString().slice(-2); return parts[2] + '/' + parts[1] + '/' + shortYear;
}
function formatThaiMonthDate(dateString) {
  if (!dateString) return ''; const parts = dateString.split('-'); if (parts.length !== 3) return dateString;
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const mIndex = parseInt(parts[1]) - 1; const shortYear = (parseInt(parts[0]) + 543).toString().slice(-2);
  return parseInt(parts[2]) + '/' + months[mIndex] + '/' + shortYear;
}

// =====================================================================
// 📁 ระบบอัปโหลดไฟล์
// =====================================================================
function uploadSingleFile(base64Data, mimeType, fileName) {
  try {
    const folder = DriveApp.getFolderById(ATTACHMENT_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) { return "error"; }
}

function saveAttachments(attachmentsData, docId) {
  if (typeof attachmentsData === 'string') return attachmentsData;
  return "-";
}

// =====================================================================
// 📝 ระบบสร้าง PDF
// =====================================================================
function processAdvanceGeneral(data) { 
  try { 
    const config = CONFIG['form_advance_general']; const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName); 
    const docId = "ADV-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"); 
    const folder = DriveApp.getFolderById(config.folderId); const template = DriveApp.getFileById(config.templateId); 
    const newDoc = template.makeCopy("ใบเบิกเงิน_" + data.name + "_" + docId, folder); const doc = DocumentApp.openById(newDoc.getId()); const body = doc.getBody(); 
    body.replaceText('<<ClearDate>>', data.clearDate); body.replaceText('<<Name>>', data.name); body.replaceText('<<Dept>>', data.dept); body.replaceText('<<Position>>', data.position); body.replaceText('<<Job>>', data.job); body.replaceText('<<AdvDate>>', data.advDate); body.replaceText('<<AdvAmt>>', data.advAmt); 
    for (let i = 0; i < 12; i++) { let n = i + 1; body.replaceText('<<Item' + n + '>>', data.items[i] ? data.items[i].detail : ''); body.replaceText('<<Amt' + n + '>>', data.items[i] ? data.items[i].amount : ''); } 
    body.replaceText('<<Total>>', data.total); body.replaceText('<<Return>>', data.returnAmt); body.replaceText('<<Claim>>', data.claimAmt); body.replaceText('<<Clearer>>', data.clearer); body.replaceText('<<App1>>', data.app1); body.replaceText('<<Receiver>>', data.receiver); body.replaceText('<<App2>>', data.app2); 
    doc.saveAndClose(); const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); newDoc.setTrashed(true); const pdfUrl = pdfFile.getUrl(); 
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let rowData = [docId, data.clearDate, data.name, data.dept, data.position, data.job, data.advDate, data.advAmt]; 
    for (let i = 0; i < 12; i++) { rowData.push(data.items[i] ? data.items[i].detail : ''); rowData.push(data.items[i] ? data.items[i].amount : ''); } 
    rowData.push(data.total, data.returnAmt, data.claimAmt, data.clearer, data.app1, data.receiver, data.app2, attachmentUrls, pdfUrl); 
    sheet.appendRow(rowData); return { success: true, url: pdfUrl, docId: docId }; 
  } catch (error) { return { success: false, message: error.message }; } 
}

function processFuelProvincial(data) { 
  try { 
    const config = CONFIG['form_fuel_provincial']; const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName); 
    const docId = "FUEL-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"); 
    const folder = DriveApp.getFolderById(config.folderId); const template = DriveApp.getFileById(config.templateId); 
    const newDoc = template.makeCopy("ค่าน้ำมันตจว_" + data.respName + "_" + docId, folder); const doc = DocumentApp.openById(newDoc.getId()); const body = doc.getBody(); 
    body.replaceText('<<ReqDate>>', data.reqDate); body.replaceText('<<DepDate>>', data.depDate); body.replaceText('<<Time>>', data.time); body.replaceText('<<StartDate>>', data.startDate); body.replaceText('<<EndDate>>', data.endDate); body.replaceText('<<Duration>>', data.duration); body.replaceText('<<RespName>>', data.respName); body.replaceText('<<Position>>', data.position); 
    for (let i=0;i<5;i++) { let n=i+1; body.replaceText('<<Loc'+n+'>>', data.plans[i]?data.plans[i].loc:''); body.replaceText('<<Plan'+n+'>>', data.plans[i]?data.plans[i].plan:''); body.replaceText('<<Rem'+n+'>>', data.plans[i]?data.plans[i].rem:''); } 
    for (let i=0;i<6;i++) { let n=i+1; let ex=data.expenses[i]||{}; body.replaceText('<<Name'+n+'>>', ex.name||''); body.replaceText('<<Alw'+n+'>>', ex.alw||''); body.replaceText('<<Acc'+n+'>>', ex.acc||''); body.replaceText('<<Fuel'+n+'>>', ex.fuel||''); body.replaceText('<<Dist'+n+'>>', ex.dist||''); body.replaceText('<<Oth'+n+'>>', ex.oth||''); body.replaceText('<<Sum'+n+'>>', ex.sum||''); body.replaceText('<<RemEx'+n+'>>', ex.remEx||''); } 
    body.replaceText('<<GrandTotal>>', data.grandTotal); body.replaceText('<<ThaiText>>', data.thaiText); body.replaceText('<<Sign1>>', data.sign1); body.replaceText('<<Sign2>>', data.sign2); body.replaceText('<<Sign3>>', data.sign3); body.replaceText('<<Sign4>>', data.sign4); 
    doc.saveAndClose(); const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); newDoc.setTrashed(true); const pdfUrl = pdfFile.getUrl(); 
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    sheet.appendRow([docId, data.reqDate, data.respName, data.position, data.startDate, data.endDate, data.grandTotal, attachmentUrls, pdfUrl]); 
    return { success: true, url: pdfUrl, docId: docId }; 
  } catch (error) { return { success: false, message: error.message }; } 
}

function processAdvanceProvincial(data) {
  try {
    const config = CONFIG['form_advance_provincial']; const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const docId = "EXP-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    const folder = DriveApp.getFolderById(config.folderId); const template = DriveApp.getFileById(config.templateId);
    const newDoc = template.makeCopy("สรุปใช้จ่ายตจว_" + data.name + "_" + docId, folder); const doc = DocumentApp.openById(newDoc.getId()); const body = doc.getBody();
    body.replaceText('<<Date>>', data.docDate); body.replaceText('<<Name>>', data.name); body.replaceText('<<Plate>>', data.plate); body.replaceText('<<AdvRef>>', data.advRef); body.replaceText('<<Dept>>', data.dept); body.replaceText('<<Position>>', data.position); body.replaceText('<<ReturnDate>>', formatThaiShortDate(data.returnDate)); body.replaceText('<<StartDate>>', formatThaiShortDate(data.startDate)); body.replaceText('<<EndDate>>', formatThaiShortDate(data.endDate)); body.replaceText('<<TotalDays>>', data.totalDays); body.replaceText('<<AdvAmt>>', data.advAmt); body.replaceText('<<NumPeople>>', data.numPeople);
    for (let i=0;i<4;i++) { let n=i+1; let p=data.persons[i]||{}; body.replaceText('<<P'+n+'Name>>', p.name||''); body.replaceText('<<P'+n+'Alw>>', p.alw||''); body.replaceText('<<P'+n+'Acc>>', p.acc||''); body.replaceText('<<P'+n+'Dist>>', p.dist||''); body.replaceText('<<P'+n+'Total>>', p.total||''); }
    body.replaceText('<<WageName>>', data.wageName); body.replaceText('<<WageAmt>>', data.wageAmt); body.replaceText('<<ReceiptAmt>>', data.receiptAmt); body.replaceText('<<EntAmt>>', data.entAmt); body.replaceText('<<MatAmt>>', data.matAmt); body.replaceText('<<GrandTotal>>', data.grandTotal); body.replaceText('<<ReturnAmt>>', data.returnAmt); body.replaceText('<<ClaimAmt>>', data.claimAmt); body.replaceText('<<Sign1>>', data.sign1); body.replaceText('<<Sign2>>', data.sign2); body.replaceText('<<Sign3>>', data.sign3);
    doc.saveAndClose(); const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); newDoc.setTrashed(true); const pdfUrl = pdfFile.getUrl();
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    sheet.appendRow([docId, currentDate, data.name, data.dept, data.position, data.plate, data.advRef, formatThaiShortDate(data.startDate), formatThaiShortDate(data.endDate), data.advAmt, data.grandTotal, data.returnAmt, data.claimAmt, attachmentUrls, pdfUrl]);
    return { success: true, url: pdfUrl, docId: docId };
  } catch (error) { return { success: false, message: error.message }; }
}

function processTollway(data) { 
  try { 
    const config = CONFIG['form_tollway']; const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName); 
    const docId = "TOLL-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"); 
    const folder = DriveApp.getFolderById(config.folderId); const template = DriveApp.getFileById(config.templateId); 
    const newDoc = template.makeCopy("เบิกทางด่วน_" + data.name + "_" + docId, folder); const doc = DocumentApp.openById(newDoc.getId()); const body = doc.getBody(); 
    body.replaceText('<<Name>>', data.name); body.replaceText('<<Total>>', data.total); 
    for (let i=0;i<20;i++) { let n=i+1; let item=data.items[i]||{}; let rowNum=(item.price||item.from||item.to)?n:''; body.replaceText('<<No'+n+'>>', rowNum); body.replaceText('<<Date'+n+'>>', item.date||''); body.replaceText('<<From'+n+'>>', item.from||''); body.replaceText('<<To'+n+'>>', item.to||''); body.replaceText('<<Price'+n+'>>', item.price||''); body.replaceText('<<SoSe'+n+'>>', item.sose||''); } 
    doc.saveAndClose(); const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); newDoc.setTrashed(true); const pdfUrl = pdfFile.getUrl(); 
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"); let rowData = [docId, currentDate, data.name, data.total]; 
    for (let i=0;i<20;i++) { let n=i+1; let item=data.items[i]||{}; let rowNum=(item.price||item.from||item.to)?n:''; rowData.push(rowNum, item.date||'', item.from||'', item.to||'', item.price||'', item.sose||''); } 
    rowData.push(attachmentUrls, pdfUrl); 
    sheet.appendRow(rowData); return { success: true, url: pdfUrl, docId: docId }; 
  } catch (error) { return { success: false, message: error.message }; } 
}

function processReqOT(data) {
  try {
    const config = CONFIG['form_req_ot']; const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const docId = "REQ-OT-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    const folder = DriveApp.getFolderById(config.folderId); const template = DriveApp.getFileById(config.templateId);
    const newDoc = template.makeCopy("ขออนุมัติทำOT_" + data.name + "_" + docId, folder); const doc = DocumentApp.openById(newDoc.getId()); const body = doc.getBody();
    body.replaceText('<<Dept>>', data.dept || ''); body.replaceText('<<Job>>', data.job || '');
    body.replaceText('<<C1>>', data.reason === '1' ? '☑' : '☐'); body.replaceText('<<C2>>', data.reason === '2' ? '☑' : '☐'); body.replaceText('<<C3>>', data.reason === '3' ? '☑' : '☐');
    for (let i=0;i<10;i++) { let n=i+1; let item=data.items[i]||{}; body.replaceText('<<D'+n+'>>', formatThaiShortDate(item.d)||''); body.replaceText('<<N'+n+'>>', item.n||''); body.replaceText('<<ST'+n+'>>', item.st||''); body.replaceText('<<ET'+n+'>>', item.et||''); body.replaceText('<<Det'+n+'>>', item.det||''); }
    body.replaceText('<<Sign1>>', data.sign1 || ''); body.replaceText('<<Sign2>>', data.sign2 || ''); body.replaceText('<<Sign3>>', data.sign3 || ''); body.replaceText('<<Sign4>>', data.sign4 || '');
    doc.saveAndClose(); const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); newDoc.setTrashed(true); const pdfUrl = pdfFile.getUrl();
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    let reasonText = data.reason === '1' ? "งานเร่งด่วน" : (data.reason === '2' ? "งานต่อเนื่อง" : "งานพิเศษ");
    sheet.appendRow([docId, currentDate, data.name, data.dept, data.job, reasonText, attachmentUrls, pdfUrl]);
    return { success: true, url: pdfUrl, docId: docId };
  } catch (error) { return { success: false, message: error.message }; }
}

// 🌟 น้ำมัน กทม. (ระบบหั่นข้อมูลขึ้นหน้าใหม่ และกรองแถวว่าง)
function processFuelBKK(data) { 
  try { 
    const config = CONFIG['form_fuel_bkk']; 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName); 
    const docId = "BKK-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"); 
    const folder = DriveApp.getFolderById(config.folderId); 
    const template = DriveApp.getFileById(config.templateId); 
    const newDoc = template.makeCopy("เบิกน้ำมันกทม_" + data.name + "_" + docId, folder); 
    const doc = DocumentApp.openById(newDoc.getId()); 
    const body = doc.getBody(); 
    
    body.replaceText('<<Name>>', data.name); 
    body.replaceText('<<Dept>>', data.dept); 
    body.replaceText('<<Position>>', data.position); 
    body.replaceText('<<VType>>', data.vType); 
    body.replaceText('<<Plate>>', data.plate); 
    body.replaceText('<<VBrand>>', data.vBrand); 

    let items = data.items || [];
    let validItems = items.filter(item => item.date && item.date.trim() !== ""); // 🛠️ กรองแถวว่าง
    
    let chunkSize = 7; 
    let templateTable = null;
    let tableIndex = -1;
    for (let i = 0; i < body.getNumChildren(); i++) {
      let child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.TABLE && child.asText().getText().indexOf('<<D1>>') !== -1) {
        templateTable = child.copy();
        tableIndex = i;
        break;
      }
    }

    let numChunks = Math.ceil(validItems.length / chunkSize);
    if (numChunks === 0) numChunks = 1; 
    
    let insertOffset = 1; 
    
    for (let c = 0; c < numChunks; c++) {
      let chunk = validItems.slice(c * chunkSize, (c + 1) * chunkSize);
      
      if (c > 0 && templateTable !== null && tableIndex !== -1) {
        body.insertPageBreak(tableIndex + insertOffset++);
        body.insertTable(tableIndex + insertOffset++, templateTable.copy());
      }
      
      for (let i = 0; i < chunkSize; i++) { 
        let n = i + 1; 
        let item = chunk[i] || {}; 
        let dDate = item.date ? formatThaiMonthDate(item.date) : '';
        
        body.replaceText('<<D'+n+'>>', dDate); 
        body.replaceText('<<Loc'+n+'>>', item.loc || ''); 
        body.replaceText('<<JobNo'+n+'>>', item.jobNo || ''); 
        body.replaceText('<<Det'+n+'>>', item.det || ''); 
        body.replaceText('<<JobDept'+n+'>>', item.jobDept || ''); 
        body.replaceText('<<Dist'+n+'>>', item.dist || ''); 
        body.replaceText('<<Amt'+n+'>>', item.amt || ''); 
        body.replaceText('<<Rem'+n+'>>', item.rem || ''); 
      } 
    }

    body.replaceText('<<TotalDist>>', data.totalDist); 
    body.replaceText('<<SummaryText>>', data.summaryText); 
    body.replaceText('<<Sign1>>', data.sign1); 
    body.replaceText('<<Sign2>>', data.sign2); 
    
    doc.saveAndClose(); 
    
    const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); 
    newDoc.setTrashed(true); 
    const pdfUrl = pdfFile.getUrl(); 
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"); 
    sheet.appendRow([docId, currentDate, data.name, data.dept, data.position, data.plate, data.totalDist, data.totalAmt, attachmentUrls, pdfUrl]); 
    
    return { success: true, url: pdfUrl, docId: docId }; 
  } catch (error) { return { success: false, message: error.message }; } 
}

// 🌟 สรุป OT (ระบบหั่นข้อมูลขึ้นหน้าใหม่ และกรองแถวว่าง)
function processOT(data) { 
  try { 
    const config = CONFIG['form_ot']; 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName); 
    const docId = "OT-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss"); 
    const folder = DriveApp.getFolderById(config.folderId); 
    const template = DriveApp.getFileById(config.templateId); 
    const newDoc = template.makeCopy("ขอOT_" + data.name + "_" + docId, folder); 
    const doc = DocumentApp.openById(newDoc.getId()); 
    const body = doc.getBody(); 
    
    body.replaceText('<<Name>>', data.name); 
    body.replaceText('<<Position>>', data.position); 
    body.replaceText('<<Dept>>', data.dept); 
    body.replaceText('<<StartDate>>', data.startDate ? formatThaiShortDate(data.startDate) : ''); 
    body.replaceText('<<EndDate>>', data.endDate ? formatThaiShortDate(data.endDate) : ''); 

    let items = data.items || [];
    let validItems = items.filter(item => item.d && item.d.trim() !== ""); // 🛠️ กรองแถวว่าง
    
    let chunkSize = 10; 
    
    let templateTable = null;
    let tableIndex = -1;
    for (let i = 0; i < body.getNumChildren(); i++) {
      let child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.TABLE && child.asText().getText().indexOf('<<D1>>') !== -1) {
        templateTable = child.copy();
        tableIndex = i;
        break;
      }
    }

    let numChunks = Math.ceil(validItems.length / chunkSize);
    if (numChunks === 0) numChunks = 1; 
    
    let insertOffset = 1; 
    
    for (let c = 0; c < numChunks; c++) {
      let chunk = validItems.slice(c * chunkSize, (c + 1) * chunkSize);
      
      if (c > 0 && templateTable !== null && tableIndex !== -1) {
        body.insertPageBreak(tableIndex + insertOffset++);
        body.insertTable(tableIndex + insertOffset++, templateTable.copy());
      }
      
      for (let i = 0; i < chunkSize; i++) { 
        let n = i + 1; 
        let item = chunk[i] || {}; 
        let dDate = item.d ? formatThaiShortDate(item.d) : ''; 
        
        body.replaceText('<<D'+n+'>>', dDate); 
        body.replaceText('<<BS'+n+'>>', item.bs||''); 
        body.replaceText('<<BE'+n+'>>', item.be||''); 
        body.replaceText('<<BT'+n+'>>', item.bt||''); 
        body.replaceText('<<AS'+n+'>>', item.as||''); 
        body.replaceText('<<AE'+n+'>>', item.ae||''); 
        body.replaceText('<<AT'+n+'>>', item.at||''); 
        body.replaceText('<<WS'+n+'>>', item.ws||''); 
        body.replaceText('<<WE'+n+'>>', item.we||''); 
        body.replaceText('<<WT'+n+'>>', item.wt||''); 
        body.replaceText('<<HS'+n+'>>', item.hs||''); 
        body.replaceText('<<HE'+n+'>>', item.he||''); 
        body.replaceText('<<HT'+n+'>>', item.ht||''); 
        body.replaceText('<<Det'+n+'>>', item.det||''); 
      } 
    }

    body.replaceText('<<Sign1>>', data.sign1 || ''); 
    body.replaceText('<<Sign2>>', data.sign2 || ''); 
    body.replaceText('<<Sign3>>', data.sign3 || ''); 
    
    doc.saveAndClose(); 
    
    const pdfFile = folder.createFile(newDoc.getAs(MimeType.PDF)); 
    newDoc.setTrashed(true); 
    const pdfUrl = pdfFile.getUrl(); 
    
    let attachmentUrls = saveAttachments(data.attachments, docId);
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"); 
    sheet.appendRow([docId, currentDate, data.name, data.dept, data.position, data.startDate ? formatThaiShortDate(data.startDate) : '', data.endDate ? formatThaiShortDate(data.endDate) : '', attachmentUrls, pdfUrl]); 
    
    return { success: true, url: pdfUrl, docId: docId }; 
  } catch (error) { return { success: false, message: error.message }; } 
}

// =====================================================================
// 💾 ระบบบันทึกแบบร่าง (Save Drafts)
// =====================================================================
function saveDraft(username, formType, formDataJSON, existingDraftId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Drafts');
    if (!sheet) throw new Error("ไม่พบแท็บ Drafts");
    let currentDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    if (existingDraftId) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === existingDraftId && data[i][2] === username) {
          sheet.getRange(i + 1, 1, 1, 5).setValues([[existingDraftId, currentDate, username, formType, formDataJSON]]); 
          return { success: true, message: 'อัปเดตแบบร่างเรียบร้อยแล้ว', draftId: existingDraftId };
        }
      }
    }
    let newDraftId = "DRF-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    sheet.appendRow([newDraftId, currentDate, username, formType, formDataJSON]);
    return { success: true, message: 'บันทึกแบบร่างเรียบร้อยแล้ว', draftId: newDraftId };
  } catch (error) { return { success: false, message: error.message }; }
}

function getUserDrafts(username) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Drafts');
  if (!sheet) return [];
  let drafts = [];
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === username) {
      drafts.push({ draftId: data[i][0], date: data[i][1], formType: data[i][3], formData: data[i][4] });
    }
  }
  return drafts.reverse(); 
}

function deleteDraftRecord(draftId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Drafts');
    if (!sheet) return { success: false, message: "ไม่พบแท็บ Drafts" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === draftId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบรหัสแบบร่าง" };
  } catch(e) { return { success: false, message: e.message }; }
}

// =====================================================================
// 🚀 ฟังก์ชันหลักรับข้อมูลจากหน้าเว็บ (บันทึกข้อมูลดิบ & แจ้งเตือน LINE)
// =====================================================================
function submitDocument(formType, data) {
  try {
    let result; let docTypeName = "";

    if (formType === 'form_advance_general') { result = processAdvanceGeneral(data); docTypeName = "สรุปเบิกเงินสำรอง"; }
    else if (formType === 'form_fuel_provincial') { result = processFuelProvincial(data); docTypeName = "ค่าน้ำมัน ตจว."; }
    else if (formType === 'form_advance_provincial') { result = processAdvanceProvincial(data); docTypeName = "สรุปค่าใช้จ่าย ตจว."; }
    else if (formType === 'form_fuel_bkk') { result = processFuelBKK(data); docTypeName = "เบิกน้ำมัน กทม."; }
    else if (formType === 'form_tollway') { result = processTollway(data); docTypeName = "เบิกค่าทางด่วน"; }
    else if (formType === 'form_ot') { result = processOT(data); docTypeName = "สรุปทำงานล่วงเวลา (OT)"; }
    else if (formType === 'form_req_ot') { result = processReqOT(data); docTypeName = "ใบขออนุมัติทำ OT"; }
    else return { success: false, message: "ไม่พบประเภทฟอร์มที่ระบุ" }; 

    if (result && result.success) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let rawSheet = ss.getSheetByName('System_RawData');
        if (!rawSheet) {
          rawSheet = ss.insertSheet('System_RawData');
          rawSheet.appendRow(['Doc ID', 'Type', 'Raw JSON Data', 'Date']);
        }
        
        let rawData = Object.assign({}, data);
        if(rawData.attachments) delete rawData.attachments; 
        
        let jsonString = JSON.stringify(rawData);
        rawSheet.appendRow([result.docId, docTypeName, jsonString, new Date()]);

        let msg = `📝 มีเอกสารส่งเข้ามาใหม่ครับ!\n\n📌 ประเภท: ${docTypeName}\n👤 ผู้ขอเบิก: ${data.name || 'ไม่ระบุ'}\n📊 สถานะ: ⏳ รออนุมัติ\n\nรบกวนแอดมินตรวจสอบด้วยนะครับ 💻`;

        let adminUIDs = getAdminUIDs();
        if (adminUIDs.length > 0) {
          adminUIDs.forEach(uid => sendLinePushMessage(uid, msg));
        } else {
          if (typeof sendLineBot === "function") sendLineBot(msg + "\n(ยังไม่มีแอดมินผูกบัญชีส่วนตัว)");
        }
      } catch(errRaw) { console.log("บันทึก Raw Data สะดุด: " + errRaw); }
    }
    return result; 
  } catch (error) { return { success: false, message: "ระบบหลังบ้านขัดข้อง: " + error.message }; }
}

function getRawDataForEdit(docId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('System_RawData');
    if (!rawSheet) return { success: false, message: "ไม่พบฐานข้อมูลสำรอง กรุณาติดต่อแอดมิน" };
    
    const data = rawSheet.getDataRange().getValues();
    for (let i = data.length - 1; i > 0; i--) {
      if (data[i][0] === docId) return { success: true, rawJson: data[i][2] };
    }
    return { success: false, message: "ไม่พบข้อมูลเดิมของเอกสารนี้" };
  } catch (e) { return { success: false, message: e.message }; }
}

// ดึงรายชื่อ "สถานที่ปฏิบัติงาน" ที่ผู้ใช้เคยกรอกไว้ก่อนหน้า (สำหรับ autocomplete)
function getRecentLocations(username) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('System_RawData');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    let locations = [];
    for (let i = data.length - 1; i >= 1 && locations.length < 30; i--) {
      const type = data[i][1];
      if (type !== 'ค่าน้ำมัน ตจว.' && type !== 'เบิกน้ำมัน กทม.') continue;
      let raw;
      try { raw = JSON.parse(data[i][2]); } catch (e) { continue; }
      const ownerName = raw.name || raw.respName || '';
      if (ownerName && ownerName !== username) continue;
      if (Array.isArray(raw.plans)) raw.plans.forEach(p => { if (p && p.loc && p.loc.trim()) locations.push(p.loc.trim()); });
      if (Array.isArray(raw.items)) raw.items.forEach(it => { if (it && it.loc && it.loc.trim()) locations.push(it.loc.trim()); });
    }
    return [...new Set(locations)].slice(0, 20);
  } catch (e) { return []; }
}

function getUserUIDByName(searchText) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return "";
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let userCol = headers.findIndex(h => h.toString().toLowerCase().trim() === 'username');
    let nameCol = headers.findIndex(h => h.toString().toLowerCase().trim() === 'fullname');
    let uidCol = headers.findIndex(h => h.toString().toLowerCase().trim().includes('uid'));
    
    if (uidCol === -1) return "";
    let searchTxt = searchText.toString().toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      let rowUser = (userCol > -1) ? data[i][userCol].toString().toLowerCase().trim() : "";
      let rowName = (nameCol > -1) ? data[i][nameCol].toString().toLowerCase().trim() : "";
      if (rowUser === searchTxt || rowName === searchTxt) return data[i][uidCol] || "";
    }
  } catch (e) { console.log("Error finding UID: " + e.message); }
  return "";
}

function sendLinePushMessage(uid, text) {
  if (!uid || uid.trim() === "") return; 
  const url = 'https://api.line.me/v2/bot/message/push';
  const options = {
    'method': 'post',
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN },
    'payload': JSON.stringify({ 'to': uid, 'messages': [{'type': 'text', 'text': text}] })
  };
  try { UrlFetchApp.fetch(url, options); } catch(e) { console.log("Push Message Error: " + e.message); }
}

function updateDocumentStatus(docId, type, status, remark, adminName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formsMap = { 'เบิกเงินสำรอง': 'Records_Advance', 'ค่าน้ำมัน ตจว.': 'Records_Fuel_Prov', 'สรุปค่าใช้จ่าย ตจว.': 'Records_Expense_Prov', 'เบิกน้ำมัน กทม.': 'Records_Fuel_BKK', 'เบิกทางด่วน': 'Records_Tollway', 'สรุปทำงานล่วงเวลา': 'Records_OT', 'ขออนุมัติทำ OT': 'Records_Req_OT' };
    
    let sheetName = formsMap[type];
    if (!sheetName) throw new Error("ไม่พบประเภทเอกสาร: " + type);
    
    let sheet = ss.getSheetByName(sheetName);
    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    
    let statusCol = headers.indexOf('สถานะ'); let remarkCol = headers.indexOf('หมายเหตุ'); let approverCol = headers.indexOf('ผู้อนุมัติ(ระบบ)');
    let rowIndex = -1; let creatorName = "ไม่ระบุ";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) { rowIndex = i + 1; creatorName = data[i][2]; break; }
    }
    
    if (rowIndex === -1) throw new Error("ไม่พบรหัสเอกสารนี้ในระบบ");
    
    if(statusCol > -1) sheet.getRange(rowIndex, statusCol + 1).setValue(status);
    if(remarkCol > -1) sheet.getRange(rowIndex, remarkCol + 1).setValue(remark);
    if(approverCol > -1) sheet.getRange(rowIndex, approverCol + 1).setValue(adminName);
    
    let emoji = (status === 'อนุมัติแล้ว') ? '✅' : '❌';
    let msg = `${emoji} อัปเดตสถานะเอกสารของคุณครับ!\n\n📄 รหัส: ${docId}\n📌 ประเภท: ${type}\n📊 สถานะ: ${status}\n`;
    if (remark) msg += `📝 หมายเหตุ: ${remark}\n`;
    msg += `👤 ตรวจสอบโดย: ${adminName}`;

    let userUID = getUserUIDByName(creatorName);
    if (userUID && userUID !== "") {
      sendLinePushMessage(userUID, msg);
    } else {
      let publicMsg = `\n${emoji} อัปเดตสถานะเอกสาร!\nรหัส: ${docId}\nผู้ขอเบิก: ${creatorName}\nสถานะ: ${status}`;
      if(remark) publicMsg += `\nหมายเหตุ: ${remark}`;
      publicMsg += `\n\n💡 (คุณ ${creatorName} ยังไม่ได้ลงทะเบียนรับแจ้งเตือนส่วนตัวผ่าน LINE OA ของบริษัทครับ)`;
      if (typeof sendLineBot === "function") sendLineBot(publicMsg);
    }
    return { success: true };
  } catch (error) { return { success: false, message: error.message }; }
}

// อนุมัติ/ปฏิเสธหลายรายการพร้อมกัน (Bulk Approve/Reject)
function bulkUpdateDocumentStatus(items, status, remark, adminName) {
  let results = [];
  (items || []).forEach(item => {
    let res = updateDocumentStatus(item.docId, item.type, status, remark, adminName);
    results.push({ docId: item.docId, success: res.success, message: res.message || '' });
  });
  let successCount = results.filter(r => r.success).length;
  return { success: true, successCount: successCount, total: results.length, results: results };
}

function getAdminUIDs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let roleCol = headers.findIndex(h => {
      let text = h.toString().toLowerCase().trim();
      return text === 'role' || text === 'สิทธิ์' || text === 'บทบาท' || text.includes('สิทธิ');
    });
    let uidCol = headers.findIndex(h => h.toString().toLowerCase().trim().includes('uid'));
    
    if (roleCol === -1 || uidCol === -1) return [];

    let adminUIDs = [];
    for (let i = 1; i < data.length; i++) {
      let role = data[i][roleCol].toString().toLowerCase().trim();
      let uid = data[i][uidCol] ? data[i][uidCol].toString().trim() : "";
      if (role === 'admin' && uid !== "") adminUIDs.push(uid);
    }
    return adminUIDs;
  } catch (e) { return []; }
}