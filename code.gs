/***********************
 * CONFIG
 ***********************/
const EVAL_SPREADSHEET_ID = '1sTkczKwDbeEP_0COWU13YpJRlVJZ5dVo83QIkN91XYw'; // ID สเปรดชีต Responses
const EVAL_SHEET_NAME     = 'Responses';

/***********************
 * ENTRY (หน้าแบบประเมิน)
 ***********************/
function doGet(e) {
  // ป้องกันกรณีรันจาก Editor แล้วไม่มี e ส่งมา
  e = e || {};
  e.parameter = e.parameter || {};

  var t = HtmlService.createTemplateFromFile('evaluation');

  t.evalId  = e.parameter.evalId  || '';
  t.eventId = e.parameter.eventId || '';
  t.docId   = e.parameter.docId   || '';

  return t
    .evaluate()
    .setTitle('แบบประเมินการทำงานของพนักงาน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***********************
 * INTERNAL – เปิดสเปรดชีต + ชีต Responses
 ***********************/
function getEvalSheet_() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(EVAL_SPREADSHEET_ID);
  } catch (err) {
    Logger.log('openById error: ' + err);
    throw new Error(
      'ไม่สามารถเปิดไฟล์สเปรดชีตแบบประเมินได้\n' +
      'กรุณาตรวจสอบว่า:\n' +
      '1) ใช้ Spreadsheet ID ถูกต้อง\n' +
      '2) บัญชีที่รัน Apps Script มีสิทธิ์เข้าถึงไฟล์นั้น'
    );
  }

  let sheet = ss.getSheetByName(EVAL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EVAL_SHEET_NAME);
    sheet.appendRow([
      'Timestamp',
      'ความตรงต่อเวลา',
      'คุณภาพของงานโดยรวม',
      'มารยาทและการสื่อสาร',
      'ความรวดเร็วและความพร้อม',
      'ความพึงพอใจโดยรวม',
      'ข้อเสนอแนะเพิ่มเติม',
      'EvalId',
      'EventId',
      'DocId'
    ]);
  }
  return sheet;
}

/***********************
 * SAVE EVALUATION
 ***********************/
function saveEvaluation(data) {
  const sheet = getEvalSheet_();

  sheet.appendRow([
    new Date(),
    data.q1,
    data.q2,
    data.q3,
    data.q4,
    data.q5,
    data.comment || '',
    data.evalId  || '',
    data.eventId || '',
    data.docId   || ''
  ]);

  if (data.docId) {
    appendEvaluationToDoc_(data);
  }

  return 'OK';
}

function appendEvaluationToDoc_(data) {
  try {
    const doc  = DocumentApp.openById(data.docId);
    const body = doc.getBody();

    const headingText = 'ผลการประเมินการปฏิบัติงาน';

    let found = false;
    const paras = body.getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      if (paras[i].getText().trim() === headingText) {
        found = true;
        break;
      }
    }

    if (!found) {
      body.appendParagraph('');
      const h = body.appendParagraph(headingText);
      h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    }

    const lines = [
      'ความตรงต่อเวลา: '        + (data.q1 || '-'),
      'คุณภาพของงานโดยรวม: '    + (data.q2 || '-'),
      'มารยาทและการสื่อสาร: '   + (data.q3 || '-'),
      'ความรวดเร็วและความพร้อม: ' + (data.q4 || '-'),
      'ความพึงพอใจโดยรวม: '     + (data.q5 || '-'),
      'ข้อเสนอแนะเพิ่มเติม: '    + (data.comment || '-')
    ];

    lines.forEach(function (txt) {
      body.appendParagraph(txt);
    });

    doc.saveAndClose();
  } catch (err) {
    Logger.log('appendEvaluationToDoc_ error: ' + err);
  }
}
