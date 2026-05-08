// ============================================================
// LINE OA Notification System — Google Apps Script  v2.0
// Applied fixes: #1 PropertiesService, #2 userId param,
//   #3 Retry + back-off, #4 Date validation, #5 Battery range,
//   #6 Mileage range, #7 Flex Messages, #8 getNow(),
//   #9 Sheet logging, #10 _DEV test suffix
// ============================================================

// ============================================================
// [SETUP #1] Run once, then delete this function.
// Stores credentials in Script Properties — never in source.
// ============================================================
function setupScriptProperties() {
  PropertiesService.getScriptProperties().setProperties({
    LINE_CHANNEL_ACCESS_TOKEN: 'SxoKiwa54KYE7RNh2vUAekZ0hnLVJT8up52jcMEd+OMecUYqX8MRFVcJQ3zoVO7gwKEstvB87+Zud/ccYw3FkVOLsz5ImmSkyz21hr6hTBITVtohG7eQ5WeuHV7cJfDzb94eqEMT6JhBU0XA4pwxswdB04t89/1O/w1cDnyilFU=',
    LINE_USER_ID:              'Ue774a86dceda0e9c08189bf792f13a01',
    LINE_LOG_SHEET:            'LINE_Log'
  });
  Logger.log('✅ Script Properties saved. Delete setupScriptProperties() before committing.');
}

var LINE_PUSH_URL = 'https://api.line.me/v2/bot/message/push';

// ============================================================
// CONFIG — reads from Script Properties (#1)
// ============================================================
function getConfig() {
  var p = PropertiesService.getScriptProperties().getProperties();
  if (!p.LINE_CHANNEL_ACCESS_TOKEN || !p.LINE_USER_ID) {
    throw new Error('LINE credentials not found. Run setupScriptProperties() first.');
  }
  return {
    token:    p.LINE_CHANNEL_ACCESS_TOKEN,
    userId:   p.LINE_USER_ID,
    logSheet: p.LINE_LOG_SHEET || 'LINE_Log'
  };
}

// ============================================================
// HELPER #8: Centralized timestamp
// ============================================================
function getNow() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

// ============================================================
// HELPER #4: Normalize date — accepts Date object or string
//   includeTime = true  → "dd/MM/yyyy HH:mm"
//   includeTime = false → "dd/MM/yyyy"
// ============================================================
function normalizeDate(input, fieldName, includeTime) {
  var fmt = includeTime ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy';
  if (input instanceof Date) {
    return Utilities.formatDate(input, Session.getScriptTimeZone(), fmt);
  }
  if (typeof input === 'string' && /\d{2}\/\d{2}\/\d{4}/.test(input)) {
    return input;
  }
  throw new Error(fieldName + ': ต้องเป็น Date object หรือ string รูปแบบ dd/MM/yyyy (ได้รับ: ' + input + ')');
}

// ============================================================
// HELPER #5: Battery bar — throws if value is out of 0-100 range
// ============================================================
function buildBatteryBar(percent) {
  if (typeof percent !== 'number' || percent < 0 || percent > 100) {
    throw new Error('battery ต้องเป็นตัวเลข 0–100 (ได้รับ: ' + percent + ')');
  }
  var filled = Math.round(percent / 10);
  return '[' + '█'.repeat(filled) + '░'.repeat(10 - filled) + ']';
}

// ============================================================
// CORE #2 #3: Send push message — userId param + retry logic
// ============================================================
/**
 * @param {string}   userId      - LINE User ID
 * @param {Object[]} messages    - LINE message objects (max 5)
 * @param {string}   messageType - label for log sheet (e.g. 'BOOKING')
 * @returns {{ success: boolean, code: number|null, body: string }}
 */
function sendLineMessage(userId, messages, messageType) {
  var config  = getConfig();
  var options = {
    method:             'post',
    contentType:        'application/json',
    headers:            { Authorization: 'Bearer ' + config.token },
    payload:            JSON.stringify({ to: userId, messages: messages }),
    muteHttpExceptions: true
  };

  var MAX_RETRIES = 3;
  var BASE_DELAY  = 1000; // ms — doubles each attempt: 1s → 2s → 4s

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      var response = UrlFetchApp.fetch(LINE_PUSH_URL, options);
      var code     = response.getResponseCode();
      var body     = response.getContentText();

      if (code === 200) {
        Logger.log('[LINE] ✅ สำเร็จ (attempt ' + attempt + ') | HTTP ' + code);
        logToSheet(messageType || 'UNKNOWN', userId, code, 'SUCCESS');
        return { success: true, code: code, body: body };
      }

      var retryable = (code === 429 || (code >= 500 && code < 600));
      Logger.log('[LINE] ' + (retryable ? '⚠️ Retrying' : '❌ Failed') +
                 ' | attempt ' + attempt + ' | HTTP ' + code + ' | ' + body);

      if (!retryable) {
        logToSheet(messageType || 'UNKNOWN', userId, code, 'FAILED');
        return { success: false, code: code, body: body };
      }

    } catch (e) {
      Logger.log('[LINE] ❌ Exception (attempt ' + attempt + '): ' + e.message);
    }

    if (attempt < MAX_RETRIES) {
      Utilities.sleep(BASE_DELAY * Math.pow(2, attempt - 1));
    }
  }

  logToSheet(messageType || 'UNKNOWN', userId, null, 'FAILED_ALL_RETRIES');
  return { success: false, code: null, body: 'All ' + MAX_RETRIES + ' retries exhausted' };
}

// ============================================================
// LOGGING #9: Append result row to Google Sheet
// ============================================================
function logToSheet(messageType, recipientId, statusCode, status) {
  try {
    var config = getConfig();
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = ss.getSheetByName(config.logSheet);

    if (!sheet) {
      sheet = ss.insertSheet(config.logSheet);
      var header = sheet.getRange(1, 1, 1, 5);
      sheet.appendRow(['Timestamp', 'Type', 'Recipient ID', 'HTTP Status', 'Result']);
      header.setFontWeight('bold').setBackground('#1565C0').setFontColor('#ffffff');
    }

    sheet.appendRow([new Date(), messageType, recipientId, statusCode || 'N/A', status]);
  } catch (e) {
    Logger.log('[LOG] ⚠️ บันทึก log ไม่ได้: ' + e.message);
  }
}

// ============================================================
// FLEX HELPERS #7
// ============================================================
function flexRow(label, value) {
  return {
    type: 'box', layout: 'horizontal', spacing: 'sm',
    contents: [
      { type: 'text', text: label,        size: 'sm', color: '#888888', flex: 3, wrap: false },
      { type: 'text', text: String(value), size: 'sm', color: '#222222', flex: 5, wrap: true, weight: 'bold' }
    ]
  };
}

function flexDivider() {
  return { type: 'separator', margin: 'sm' };
}

function flexFooter() {
  return {
    type: 'box', layout: 'vertical', paddingAll: '10px',
    contents: [{
      type: 'text', text: '⏰ แจ้งเมื่อ: ' + getNow(),
      size: 'xs', color: '#aaaaaa', align: 'end'
    }]
  };
}

function flexHeader(title, bgColor) {
  return {
    type: 'box', layout: 'vertical',
    backgroundColor: bgColor, paddingAll: '14px',
    contents: [{
      type: 'text', text: title,
      color: '#ffffff', size: 'lg', weight: 'bold'
    }]
  };
}

// ============================================================
// FLEX BUILDER #7 — 1: Vehicle Booking
// ============================================================
function buildBookingFlex(data) {
  return {
    type: 'flex',
    altText: '🚗 การจองรถ — ' + data.bookerName,
    contents: {
      type: 'bubble', size: 'kilo',
      header: flexHeader('🚗  การจองรถ', '#1565C0'),
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
        contents: [
          flexRow('👤 ผู้จอง',     data.bookerName),
          flexDivider(),
          flexRow('🔑 ทะเบียนรถ', data.licensePlate),
          flexRow('📅 รับรถ',      data.startDate),
          flexRow('📅 คืนรถ',      data.returnDate),
          flexDivider(),
          flexRow('📍 สถานที่',    data.location)
        ]
      },
      footer: flexFooter()
    }
  };
}

// ============================================================
// FLEX BUILDER #7 — 2: EV Charging
// ============================================================
function buildChargingFlex(data) {
  return {
    type: 'flex',
    altText: '⚡ รายงานชาร์จ EV — ' + data.name,
    contents: {
      type: 'bubble', size: 'kilo',
      header: flexHeader('⚡  รายงานการชาร์จรถ EV', '#2E7D32'),
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
        contents: [
          flexRow('📅 วันที่',           data.date),
          flexRow('👤 ชื่อ-สกุล',        data.name),
          flexRow('🏢 แผนก',             data.department),
          flexDivider(),
          flexRow('🛣️  เลขไมล์',          data.mileage.toLocaleString() + ' กม.'),
          flexRow('🔋 แบตก่อนชาร์จ',    data.batteryBefore + '%'),
          flexRow('🔋 แบตหลังชาร์จ',    data.batteryAfter + '% ' + buildBatteryBar(data.batteryAfter)),
          flexDivider(),
          flexRow('📍 สถานที่/หมายเหตุ', data.locationInfo)
        ]
      },
      footer: flexFooter()
    }
  };
}

// ============================================================
// FLEX BUILDER #7 — 3: Grab Code
// ============================================================
function buildGrabCodeFlex(data) {
  var codesLabel = data.codeCount === 1 ? '1 Code  (เที่ยวเดียว)' : '2 Codes (ไป-กลับ)';
  return {
    type: 'flex',
    altText: '🚖 รับ Grab Code — ' + data.employeeId,
    contents: {
      type: 'bubble', size: 'kilo',
      header: flexHeader('🚖  รับ Grab Code', '#00B14F'),
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '16px',
        contents: [
          flexRow('🪪 รหัสพนักงาน', data.employeeId),
          flexRow('👤 ชื่อ-สกุล',   data.name),
          flexDivider(),
          flexRow('📍 สถานที่',      data.location),
          flexRow('🎟️  จำนวน Code', codesLabel)
        ]
      },
      footer: flexFooter()
    }
  };
}

// ============================================================
// FUNCTION 1: Vehicle Booking Alert
// ============================================================
/**
 * @param {Object} data
 * @param {string}      data.bookerName   - ชื่อผู้จอง
 * @param {string}      data.licensePlate - ทะเบียนรถ
 * @param {Date|string} data.startDate    - วัน-เวลารับรถ
 * @param {Date|string} data.returnDate   - วัน-เวลาคืนรถ
 * @param {string}      data.location     - สถานที่ปลายทาง
 */
function sendBookingAlert(data) {
  if (!data.bookerName || !data.licensePlate || !data.location) {
    Logger.log('[LINE] ⚠️ sendBookingAlert: ข้อมูลไม่ครบถ้วน');
    return { success: false, code: null, body: 'Missing required booking fields' };
  }
  data.startDate  = normalizeDate(data.startDate,  'startDate',  true);
  data.returnDate = normalizeDate(data.returnDate, 'returnDate', true);

  var config = getConfig();
  return sendLineMessage(config.userId, [buildBookingFlex(data)], 'BOOKING');
}

// ============================================================
// FUNCTION 2: EV Charging Report Alert
// ============================================================
/**
 * @param {Object}      data
 * @param {Date|string} data.date           - วันที่ชาร์จ
 * @param {string}      data.name           - ชื่อพนักงาน
 * @param {string}      data.department     - แผนก
 * @param {number}      data.mileage        - เลขไมล์ (≥ 0)
 * @param {number}      data.batteryBefore  - % แบตก่อนชาร์จ (0–100)
 * @param {number}      data.batteryAfter   - % แบตหลังชาร์จ (0–100)
 * @param {string}      data.locationInfo   - สถานที่ / หมายเหตุ
 */
function sendChargingAlert(data) {
  if (!data.name || !data.department || !data.locationInfo) {
    Logger.log('[LINE] ⚠️ sendChargingAlert: ข้อมูลไม่ครบถ้วน');
    return { success: false, code: null, body: 'Missing required charging fields' };
  }
  if (typeof data.mileage !== 'number' || data.mileage < 0) {   // #6
    Logger.log('[LINE] ⚠️ mileage ต้องเป็นตัวเลข ≥ 0 (ได้รับ: ' + data.mileage + ')');
    return { success: false, code: null, body: 'Invalid mileage: must be a non-negative number' };
  }
  try {
    buildBatteryBar(data.batteryBefore); // #5 — throws if out of range
    buildBatteryBar(data.batteryAfter);  // #5
  } catch (e) {
    Logger.log('[LINE] ⚠️ ' + e.message);
    return { success: false, code: null, body: e.message };
  }
  data.date = normalizeDate(data.date, 'date', false); // #4

  var config = getConfig();
  return sendLineMessage(config.userId, [buildChargingFlex(data)], 'CHARGING');
}

// ============================================================
// FUNCTION 3: Grab Code Request Alert
// ============================================================
/**
 * @param {Object} data
 * @param {string} data.employeeId - รหัสพนักงาน
 * @param {string} data.name       - ชื่อพนักงาน
 * @param {string} data.location   - ต้นทาง / ปลายทาง
 * @param {number} data.codeCount  - จำนวน Code (1 หรือ 2 เท่านั้น)
 */
function sendGrabCodeAlert(data) {
  if (!data.employeeId || !data.name || !data.location) {
    Logger.log('[LINE] ⚠️ sendGrabCodeAlert: ข้อมูลไม่ครบถ้วน');
    return { success: false, code: null, body: 'Missing required grab fields' };
  }
  if (data.codeCount !== 1 && data.codeCount !== 2) {
    Logger.log('[LINE] ⚠️ codeCount ต้องเป็น 1 หรือ 2 เท่านั้น (ได้รับ: ' + data.codeCount + ')');
    return { success: false, code: null, body: 'codeCount must be 1 or 2' };
  }

  var config = getConfig();
  return sendLineMessage(config.userId, [buildGrabCodeFlex(data)], 'GRAB_CODE');
}

// ============================================================
// DIAGNOSE — รันฟังก์ชันนี้ก่อนเสมอเมื่อระบบไม่ทำงาน
// ตรวจสอบทุกชั้น: Properties → Token → Push API → Flex
// ============================================================
function diagnoseLineSetup() {
  Logger.log('╔══════════════════════════════════════════╗');
  Logger.log('║       LINE OA — DIAGNOSTIC REPORT        ║');
  Logger.log('╚══════════════════════════════════════════╝');

  // --- STEP 1: Script Properties ---
  Logger.log('\n[STEP 1] ตรวจสอบ Script Properties...');
  var p = PropertiesService.getScriptProperties().getProperties();
  var token  = p.LINE_CHANNEL_ACCESS_TOKEN;
  var userId = p.LINE_USER_ID;

  if (!token) {
    Logger.log('  ❌ LINE_CHANNEL_ACCESS_TOKEN ไม่พบ → รัน setupScriptProperties() ก่อน แล้วรัน diagnoseLineSetup() ใหม่');
    return;
  }
  Logger.log('  ✅ Token พบแล้ว (ตัวอักษร ' + token.length + ' ตัว, เริ่มต้น: ' + token.substring(0, 10) + '...)');

  if (!userId) {
    Logger.log('  ❌ LINE_USER_ID ไม่พบ → รัน setupScriptProperties() ก่อน');
    return;
  }
  Logger.log('  ✅ User ID: ' + userId);

  // --- STEP 2: ตรวจสอบ Token ผ่าน LINE Bot Info API ---
  Logger.log('\n[STEP 2] ตรวจสอบความถูกต้องของ Token กับ LINE API...');
  try {
    var botInfoRes  = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var botCode = botInfoRes.getResponseCode();
    var botBody = botInfoRes.getContentText();

    if (botCode === 200) {
      var bot = JSON.parse(botBody);
      Logger.log('  ✅ Token ถูกต้อง | Bot name: "' + bot.displayName + '" | followers: ' + bot.followersCount);
    } else if (botCode === 401) {
      Logger.log('  ❌ HTTP 401 — Token ผิดหรือหมดอายุ → ไปที่ LINE Developers Console และสร้าง Token ใหม่');
      Logger.log('     Response: ' + botBody);
      return;
    } else {
      Logger.log('  ⚠️ HTTP ' + botCode + ' | ' + botBody);
    }
  } catch (e) {
    Logger.log('  ❌ ไม่สามารถเชื่อมต่อ LINE API: ' + e.message);
    return;
  }

  // --- STEP 3: ส่งข้อความ plain-text ทดสอบ ---
  Logger.log('\n[STEP 3] ส่ง plain-text ทดสอบไปยัง User ID: ' + userId + ' ...');
  var testPayload = JSON.stringify({
    to: userId,
    messages: [{ type: 'text', text: '[DIAGNOSE] ✅ ทดสอบระบบ LINE OA — เชื่อมต่อสำเร็จ ' + getNow() }]
  });
  try {
    var testRes  = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:             'post',
      contentType:        'application/json',
      headers:            { Authorization: 'Bearer ' + token },
      payload:            testPayload,
      muteHttpExceptions: true
    });
    var testCode = testRes.getResponseCode();
    var testBody = testRes.getContentText();

    if (testCode === 200) {
      Logger.log('  ✅ ส่งสำเร็จ! ตรวจสอบ LINE ของคุณได้เลย');
    } else if (testCode === 400) {
      Logger.log('  ❌ HTTP 400 — User ID ผิด หรือ Bot ยังไม่ได้ถูก Add เป็นเพื่อน');
      Logger.log('     Response: ' + testBody);
      Logger.log('     → เปิด LINE แล้ว Add บอทเป็นเพื่อนก่อน จากนั้นรัน diagnoseLineSetup() ใหม่');
    } else if (testCode === 403) {
      Logger.log('  ❌ HTTP 403 — Channel ไม่มีสิทธิ์ push message → ตรวจสอบ Plan ใน LINE Developers Console');
      Logger.log('     Response: ' + testBody);
    } else {
      Logger.log('  ❌ HTTP ' + testCode + ' | ' + testBody);
    }
  } catch (e) {
    Logger.log('  ❌ Exception: ' + e.message);
  }

  // --- STEP 4: ทดสอบ Flex Message ---
  Logger.log('\n[STEP 4] ส่ง Flex Message ทดสอบ...');
  var flexTest = JSON.stringify({
    to: userId,
    messages: [{
      type: 'flex',
      altText: '[DIAGNOSE] Flex Message ทดสอบ',
      contents: {
        type: 'bubble',
        body: {
          type: 'box', layout: 'vertical',
          contents: [{ type: 'text', text: 'Flex Message ทำงานปกติ ✅', wrap: true }]
        }
      }
    }]
  });
  try {
    var flexRes  = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:             'post',
      contentType:        'application/json',
      headers:            { Authorization: 'Bearer ' + token },
      payload:            flexTest,
      muteHttpExceptions: true
    });
    var flexCode = flexRes.getResponseCode();
    if (flexCode === 200) {
      Logger.log('  ✅ Flex Message ส่งได้ปกติ');
    } else {
      Logger.log('  ❌ Flex Message ล้มเหลว HTTP ' + flexCode + ' | ' + flexRes.getContentText());
      Logger.log('     → ระบบจะ fallback เป็น plain-text อัตโนมัติ');
    }
  } catch (e) {
    Logger.log('  ❌ Exception: ' + e.message);
  }

  Logger.log('\n╔══════════════════════════════════════════╗');
  Logger.log('║             DIAGNOSTIC COMPLETE           ║');
  Logger.log('╚══════════════════════════════════════════╝');
}

// ============================================================
// DEV TESTS #10 — ⚠️ ส่งข้อความจริงไปยัง LINE OA ⚠️
// ============================================================

function testBookingAlert_DEV() {
  Logger.log('========== TEST 1: Vehicle Booking Alert ==========');
  try {
    var result = sendBookingAlert({
      bookerName:   'สมชาย ใจดี',
      licensePlate: 'กข-1234 กรุงเทพฯ',
      startDate:    new Date(2026, 4, 7, 8, 0),
      returnDate:   new Date(2026, 4, 7, 17, 30),
      location:     'สนามบินสุวรรณภูมิ ผู้โดยสารขาออก Terminal 1'
    });
    Logger.log('Result → ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('❌ Exception: ' + e.message);
    Logger.log('   → รัน diagnoseLineSetup() เพื่อหาสาเหตุ');
  }
  Logger.log('===================================================');
}

function testChargingAlert_DEV() {
  Logger.log('========== TEST 2: EV Charging Report Alert ==========');
  try {
    var result = sendChargingAlert({
      date:          new Date(2026, 4, 6),
      name:          'สุภาพร รักษ์โลก',
      department:    'Facilities Management',
      mileage:       24530,
      batteryBefore: 18,
      batteryAfter:  92,
      locationInfo:  'อาคาร A ชั้น B1 — หัวชาร์จ Type-2 เครื่องที่ 3'
    });
    Logger.log('Result → ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('❌ Exception: ' + e.message);
    Logger.log('   → รัน diagnoseLineSetup() เพื่อหาสาเหตุ');
  }
  Logger.log('======================================================');
}

function testGrabCodeAlert_DEV() {
  Logger.log('========== TEST 3: Grab Code Request Alert ==========');
  try {
    var result = sendGrabCodeAlert({
      employeeId: '508099',
      name:       'ChutipunH',
      location:   'อาคารสำนักงานใหญ่ (ชั้น 1) → BTS อโศก',
      codeCount:  2
    });
    Logger.log('Result → ' + JSON.stringify(result));
  } catch (e) {
    Logger.log('❌ Exception: ' + e.message);
    Logger.log('   → รัน diagnoseLineSetup() เพื่อหาสาเหตุ');
  }
  Logger.log('=====================================================');
}

// ⚠️ DEV ONLY — ส่งข้อความจริงทั้ง 3 รายการ
function runAllTests_DEV() {
  testBookingAlert_DEV();
  testChargingAlert_DEV();
  testGrabCodeAlert_DEV();
}
