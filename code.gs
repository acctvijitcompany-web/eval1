/** =========================
 * ตั้งค่าเบื้องต้น
 * ======================= */
const SPREADSHEET_NAME = 'แบบประเมินการทำงานของพนักงาน (Responses)'; // ชื่อไฟล์ใน Google Drive
const SHEET_NAME = 'Responses'; // ชื่อชีตด้านในไฟล์

/**
 * คืนค่า sheet สำหรับบันทึกข้อมูล
 * - ถ้ายังไม่มีไฟล์ จะสร้างใหม่ให้
 * - เก็บ ID ไฟล์ไว้ใน Script Properties ครั้งต่อไปจะเปิดจาก ID เดิม
 */
function getResponseSheet() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('RESP_SHEET_ID');
  let ss;

  if (sheetId) {
    // ถ้ามี ID แล้ว ลองเปิดไฟล์
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      // ถ้าเปิดไม่ได้ (ลบไฟล์ไปแล้ว ฯลฯ) ให้สร้างใหม่
      ss = SpreadsheetApp.create(SPREADSHEET_NAME);
      props.setProperty('RESP_SHEET_ID', ss.getId());
    }
  } else {
    // ยังไม่เคยมีไฟล์มาก่อน → สร้างใหม่
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    props.setProperty('RESP_SHEET_ID', ss.getId());
  }

  // หา/สร้างชีตสำหรับเก็บข้อมูล
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }

  // ถ้ายังไม่มีหัวตาราง ให้สร้างหัวตารางรอบแรก
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'ความตรงต่อเวลา',
      'คุณภาพของงานโดยรวม',
      'มารยาทและการสื่อสาร',
      'ความรวดเร็วและความพร้อม',
      'ความพึงพอใจโดยรวม',
      'ข้อเสนอแนะเพิ่มเติม'
    ]);
  }

  return sheet;
}

/**
 * หน้าเว็บฟอร์ม
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('แบบประเมินการทำงานของพนักงาน');
}

/**
 * รับข้อมูลจากฟอร์ม HTML และบันทึกลงชีต
 */
function submitForm(data) {
  try {
    const sheet = getResponseSheet();

    sheet.appendRow([
      new Date(),
      data.q1,
      data.q2,
      data.q3,
      data.q4,
      data.q5,
      data.comment
    ]);

    // ส่งข้อความกลับให้ front-end เอาไปแสดงใน SweetAlert ได้
    return 'OK';
  } catch (e) {
    console.error('Error in submitForm:', e);
    throw e; // ส่ง error กลับไปให้ withFailureHandler จัดการ
  }
}
