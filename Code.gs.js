// 🟢 ใส่ Calendar ID ของคุณตรงนี้ (อย่าลืมเอามาใส่นะครับ!)
const CALENDAR_ID = 'dce80de3d80222989aaa3ba51b594063c47165f127d56464e46ad60f372713b5@group.calendar.google.com';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('⚡ ระบบจัดการ EV & Grab') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 1. ปรับปรุง formatToBE ให้แสดงผลแบบ 2 หลักเสมอ
function formatToBE(dateObj) {
  if (!dateObj) return "";
  let d = new Date(dateObj);
  let day = String(d.getDate()).padStart(2, '0');
  let month = String(d.getMonth() + 1).padStart(2, '0');
  let yearBE = d.getFullYear() + 543;
  return `${day}/${month}/${yearBE}`;
}

// 2. ปรับ Router ให้ใช้ Switch Case
function submitData(payload) {
  switch (payload.action) {
    case 'จองรถ EV':
      return handleBooking(payload);
    case 'รายงานใช้รถ':
      return handleReport(payload);
    default:
      return { status: "error", message: "ระบบไม่รู้จักคำสั่งนี้" };
  }
}

// 3. ระบบจองรถ EV (แก้ไขบั๊ก eventId)
function handleBooking(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EV_Bookings');

  // --- จัดการรูปแบบวันที่สำหรับ Google Calendar ---
  // การจองแบบทั้งวัน (All-Day) End Date ของ Calendar API ต้องบวกเพิ่ม 1 วันเสมอ
  const start = new Date(payload.startDate);
  start.setHours(0, 0, 0, 0);

  const end = new Date(payload.endDate);
  end.setDate(end.getDate() + 1); // บวกเพิ่ม 1 วันเพื่อให้ครอบคลุมวันจบ
  end.setHours(0, 0, 0, 0);

  try {
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) {
      return { status: "error", message: "ระบบขัดข้อง: ไม่พบปฏิทิน กรุณาตรวจสอบ Calendar ID" };
    }

    // ===============================================
    // 1. เช็คคิวใน Calendar ID1 (ดึงคิวทั้งหมดในช่วงวันที่เลือก)
    // ===============================================
    const events = cal.getEvents(start, end);
    let isConflict = false;

    for (let i = 0; i < events.length; i++) {
      const existingTitle = events[i].getTitle();
      
      // ถ้ารายชื่อ Event ในปฏิทิน มีทะเบียนรถคันนี้อยู่ แปลว่าไม่ว่าง
      if (existingTitle.includes(payload.carPlate)) {
        isConflict = true;
        break;
      }
    }

    // ===============================================
    // 3. ถ้าไม่ว่าง ให้กลับมาที่หน้า web แจ้งเตือนจองไม่สำเร็จ
    // ===============================================
    if (isConflict) {
      return { 
        status: "error", 
        message: `ไม่สามารถจองได้! รถทะเบียน ${payload.carPlate} ถูกจองในปฏิทินแล้วในช่วงวันที่ท่านเลือก` 
      };
    }

    // ===============================================
    // 2. ถ้าว่างให้ booking ลงใน Calendar ID1
    // ===============================================
    const eventTitle = `${payload.carPlate} -${payload.name} - ${payload.dept} `;
    const event = cal.createAllDayEvent(eventTitle, start, end); // สร้างลงปฏิทิน
    const eventId = event.getId();

    // (เพิ่มเติม) บันทึกประวัติลง Google Sheet เพื่อเก็บเป็นฐานข้อมูล
    const bookingId = `BK-${Utilities.formatDate(new Date(), "GMT+7", "yyMMddHHmmss")}`;
    const originalEndDate = new Date(payload.endDate); // ใช้วันที่เดิมลง Sheet จะได้ไม่งง
    
    sheet.appendRow([
      bookingId, 
      new Date(), 
      payload.name, 
      payload.carPlate, 
      start, 
      originalEndDate, 
      'Active', 
      eventId, 
      payload.dept, 
      payload.destination
    ]);

    // ส่งสถานะ success กลับไปที่ Web เพื่อแสดงแจ้งเตือนว่าสำเร็จ
    return { 
      status: "success", 
      message: `จองรถ ${payload.carPlate} สำเร็จเรียบร้อยแล้ว!` 
    };

  } catch(e) {
    return { status: "error", message: `ระบบขัดข้อง: ${e.message}` };
  }
}

// 📝 ระบบรายงานการใช้รถ EV (เพิ่มจากโค้ดเดิมของคุณ)
function handleReport(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('EV_Reports');
    
    sheet.appendRow([
      new Date(),
      payload.chargeDate ? new Date(payload.chargeDate) : '',
      payload.name,
      payload.dept,
      payload.mileage,
      payload.battBefore,
      payload.battAfter,
      payload.stationNetwork,
      payload.stationDetail,
      payload.extraInfo
    ]);

    return { status: "success", message: `บันทึกข้อมูลการ ${payload.action} เรียบร้อยแล้ว!` };
  } catch (e) {
    return { status: "error", message: "เกิดข้อผิดพลาดในการบันทึกรายงาน กรุณาลองใหม่" };
  }
}

// 4. ค้นหาพนักงานจาก Sheet
function checkEmployeeId(empId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const data = sheet.getDataRange().getValues().slice(1); 

  const employee = data.find(row => row[0].toString() === empId.toString());
  
  return employee 
    ? { status: "success", name: employee[1] } 
    : { status: "error" };
}

// 5. ระบบ Grab: แจกโค้ดและตัดสต๊อก (เพิ่ม block finally)
function processGrabRequest(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const poolSheet = ss.getSheetByName('Grab_Pool');
    const usedSheet = ss.getSheetByName('Grab_Used');

    const codes = poolSheet.getDataRange().getValues();
    const requiredCodes = parseInt(payload.tripType); 

    if (codes.length - 1 < requiredCodes) {
      return { status: "error", message: "ขออภัย โค้ด Grab ในระบบเหลือไม่เพียงพอสำหรับการเบิกนี้" };
    }

    let givenCodes = [];
    for(let i = 0; i < requiredCodes; i++) {
       givenCodes.push(codes[i + 1][0]); 
    }

    const codeString = givenCodes.join(", ");

    usedSheet.appendRow([new Date(), payload.name, payload.location, codeString, payload.empId]);
    poolSheet.deleteRows(2, requiredCodes);

    return {
      status: "success",
      message: `คุณได้รับโค้ด Grab จำนวน ${requiredCodes} โค้ด ดังนี้:\n\n${codeString}`
    };

  } catch (e) {
    return { status: "error", message: "ระบบขัดข้องชั่วคราว กรุณาลองใหม่" };
  } finally {
    lock.releaseLock(); 
  }
}
