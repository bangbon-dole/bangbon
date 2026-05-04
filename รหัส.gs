// การตั้งค่าหลัก
const SPREADSHEET_ID = '1GxOLj13jFWBMgn_nqHNDMxhpwjb2NRww8Rm4r6_Xqvc'; 
const SHEET_NAME = 'ชีต1'; 
const LOGO_FILE_ID = '1AQ79fhA_-JNXfYIbWXYQXmtKa-VNAiWq';

function doGet(e) {
  var page = e.parameter.page;
  var fileName = (page === 'report') ? 'report' : 'Index';
  var title = (page === 'report') ? 'ระบบรายงานการลงเวลา' : 'ระบบลงเวลาปฏิบัติงาน สกร.ระดับเขตบางบอน';

  var html = HtmlService.createTemplateFromFile(fileName);
  
  // ดึงโลโก้จาก Google Drive เป็น Base64
  try {
    var imageBlob = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
    var base64Image = Utilities.base64Encode(imageBlob.getBytes());
    html.logoData = "data:" + imageBlob.getContentType() + ";base64," + base64Image;
  } catch (err) {
    html.logoData = ""; 
  }

  return html.evaluate()
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// บันทึกข้อมูลลง Google Sheets
function saveData(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const sheet = ss.getSheetByName(SHEET_NAME); 
    if (!sheet) return { success: false, message: "ไม่พบชื่อชีต: " + SHEET_NAME };

    const now = new Date();
    const timestamp = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy HH:mm:ss");
    const todayStr = Utilities.formatDate(now, "GMT+7", "dd/MM/yyyy");
    
    // ตรวจสอบการลงชื่อซ้ำในวันเดียวกัน
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      const isDuplicate = existingData.some(row => {
        const rowDate = (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy") : row[0].toString().split(' ')[0];
        return rowDate === todayStr && row[1] === data.name;
      });

      if (isDuplicate) {
        return { success: false, message: "❌ คุณ " + data.name + " ได้บันทึกข้อมูลของวันนี้ไปเรียบร้อยแล้ว" };
      }
    }

    const monthNames = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
    const currentMonth = monthNames[now.getMonth()] + " " + (now.getFullYear() + 543);
    
    sheet.appendRow([timestamp, data.name, data.position, data.status, data.remark, currentMonth]);
    return { success: true, message: "✅ บันทึกข้อมูลเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

// รายชื่อพนักงาน
function getStaffNames() {
  const names = [
    "นางสาวบุสรา นิลเต๊ะ", 
    "นางผกามาศ เรืองจิรัษเฐียร", 
    "นางสาวอรอุมา สุวัฒนกุลชัย",
    "นางสาวรัชนี ปาละโค", 
    "ว่าที่ร้อยตรีหญิงเขมจิรา สวัสดี", 
    "นางสาวรุ่งนภา สุขคุณา",
    "นายภูวณัฏฐ์ ไชยปัญญาวิชญ์", 
    "นางสาวธิติมา กวาวมณี", 
    "นายวิทยา เจือจันทร์",
    "นายอรรถพล กรุงแก้ว", 
    "นางสาวสุภาวดี คงเกลี้ยง", 
    "นายณดล จาริยพานิช",
    "นางสาวสิมิลัน อภิศักดิ์มนตรี", 
    "นางสาวขวัญสุดา จำปา", 
    "นางสมพร แสงงาม"
  ];
  return names; // ส่งค่ากลับโดยไม่ใช้ .sort() เพื่อรักษาลำดับเดิม
}

// กรองข้อมูลสำหรับรายงาน
function getFilteredData(searchName, searchMonth, searchYear) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); 
  
  const filtered = rows.filter(row => {
  const dateVal = row[0]; 
  const nameVal = row[1];
  let m = "";
  let y = "";

  if (dateVal instanceof Date && !isNaN(dateVal)) {
    m = (dateVal.getMonth() + 1).toString(); // แปลงเป็น String เพื่อให้เทียบกับค่าจากหน้าเว็บได้แม่นยำ
    y = dateVal.getFullYear().toString();
  } else if (typeof dateVal === 'string' && dateVal.includes('/')) {
    // กรณีวันที่ใน Sheet ถูกบันทึกเป็นข้อความ (String)
    m = parseInt(dateVal.split('/')[1], 10).toString();
    y = dateVal.split('/')[2].split(' ')[0];
  }

  // ใช้ == แทน === เพื่อความยืดหยุ่นในการเทียบ String และ Number
  const matchName = (searchName === "" || nameVal === searchName);
  const matchMonth = (searchMonth === "" || m == searchMonth); 
  const matchYear = (searchYear === "" || y == searchYear || (parseInt(y)+543) == searchYear);
// ใส่ไว้ก่อนบรรทัด return matchName...
console.log("แถวนี้เดือนคือ: " + m + " | เดือนที่หาคือ: " + searchMonth);
  return matchName && matchMonth && matchYear;
});
  
  return filtered.map(row => ({
    dateTime: (row[0] instanceof Date) ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy HH:mm:ss") : row[0],
    name: row[1], position: row[2], status: row[3], note: row[4], month: row[5]
  }));
}
