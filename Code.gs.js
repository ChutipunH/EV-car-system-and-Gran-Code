// ฟังก์ชัน doGet จะถูกเรียกเมื่อมีคนเปิด URL ของ Web App
function doGet(e) {
  // เปิดหน้า Index.html เป็นหน้าแรกสุด
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('⚡ ระบบจัดการ EV & Grab') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันสำคัญ! ทำหน้าที่ดึงไฟล์ HTML ย่อยๆ มาแทรกในหน้า Index
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ฟังก์ชันแปลงวันที่ ค.ศ. เป็น พ.ศ.
function formatToBE(dateObj) {
  if (!dateObj) return "";
  let d = new Date(dateObj);
  let day = d.getDate();
  let month = d.getMonth() + 1;
  let yearBE = d.getFullYear() + 543;
  return day + "/" + month + "/" + yearBE;
}

// ฟังก์ชันรับข้อมูลจากหน้าเว็บและบันทึกลง Google Sheets
function submitData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();

    if (data.action === 'จองรถ EV') {
      const sheet = ss.getSheetByName('EV_Bookings');
      const bookingId = 'BK' + timestamp.getTime();
      sheet.appendRow([
        bookingId,
        timestamp,
        data.name,
        data.carRegis,
        data.startDate ? new Date(data.startDate) : '',
        data.endDate   ? new Date(data.endDate)   : '',
        'จอง',
        ''
      ]);

    } else if (data.action === 'รายงานใช้รถ') {
      const sheet = ss.getSheetByName('EV_Reports');
      sheet.appendRow([
        timestamp,
        data.chargeDate ? new Date(data.chargeDate) : '',
        data.name,
        data.dept,
        data.mileage,
        data.battBefore,
        data.battAfter,
        data.stationNetwork,
        data.stationDetail,
        data.extraInfo
      ]);
    }

    return {
      status: "success",
      message: "บันทึกข้อมูลการ " + data.action + " เรียบร้อยแล้ว!"
    };
  } catch (e) {
    Logger.log(e);
    return { status: "error", message: "เกิดข้อผิดพลาด กรุณาลองใหม่" };
  }
}

// --- ฟังก์ชันค้นหาพนักงานจาก Sheet: Employees ---
function checkEmployeeId(empId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) { // เริ่มที่ 1 เพื่อข้ามหัวตาราง
    if (data[i][0].toString() === empId.toString()) {
      return { status: "success", name: data[i][1] }; // ส่งชื่อกลับไป
    }
  }
  return { status: "error" }; // หาไม่เจอ
}

// --- ฟังก์ชันแจกโค้ด Grab และตัดสต๊อก ---
function processGrabRequest(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // ล็อคป้องกันคนกดพร้อมกัน
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const poolSheet = ss.getSheetByName('Grab_Pool');
    const usedSheet = ss.getSheetByName('Grab_Used');

    const codes = poolSheet.getDataRange().getValues();
    const requiredCodes = parseInt(payload.tripType); // จำนวนที่ขอ (1 หรือ 2)

    // เช็คว่ามีโค้ดเหลือพอแจกไหม (ลบ 1 คือลบหัวตารางออก)
    if (codes.length - 1 < requiredCodes) {
      return { status: "error", message: "ขออภัย โค้ด Grab ในระบบเหลือไม่เพียงพอสำหรับการเบิกนี้" };
    }

    let givenCodes = [];
    // ดึงโค้ดออกมาตามจำนวนที่ต้องการ
    for(let i = 0; i < requiredCodes; i++) {
       givenCodes.push(codes[i + 1][0]); // ดึงจากบรรทัดที่ 2 เป็นต้นไป
    }

    // รวมโค้ดเป็นข้อความ (กรณีได้ 2 โค้ด จะมีลูกน้ำคั่น)
    const codeString = givenCodes.join(", ");

    // บันทึกลง Sheet ประวัติการใช้งาน [เวลา, ชื่อ, สถานที่, โค้ด, รหัสพนักงาน]
    usedSheet.appendRow([new Date(), payload.name, payload.location, codeString, payload.empId]);

    // ลบโค้ดที่แจกไปแล้วออกจาก Pool (ลบบรรทัดที่ 2 ออกตามจำนวนที่แจก)
    poolSheet.deleteRows(2, requiredCodes);

    lock.releaseLock();

    return {
      status: "success",
      message: "คุณได้รับโค้ด Grab จำนวน " + requiredCodes + " โค้ด ดังนี้:\n\n" + codeString
    };

  } catch (e) {
    return { status: "error", message: "ระบบขัดข้องชั่วคราว กรุณาลองใหม่" };
  }
}