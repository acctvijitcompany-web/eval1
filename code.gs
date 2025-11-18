/**********************
 * CONFIG
 **********************/
const CALENDAR_ID  = 'acct.vijitcompany@gmail.com';        // ปฏิทินงาน
const FOLDER_ID    = '1Hx9gx75C-XtmFW4-v8zpzwwbQFTQb1jj';  // โฟลเดอร์เก็บ PDF+Doc
const TZ           = 'Asia/Bangkok';
const COMPANY_NAME = 'บริษัท วิจิตร เอ แอนด์ ที จำกัด';

// ใส่ URL Web App ของหน้าแบบประเมิน (deployment ล่าสุด) ตรงนี้
// เช่น https://script.google.com/macros/s/xxxxx/exec
const EVAL_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxh5XcVAgoVLqxxtq4vxXxs2UyK7714jaL0VqexNQMFbqqTcjPsGmNKl9AgzFkIjnJs3g/exec';

/**********************
 * ENTRY – DASHBOARD
 **********************/
function doGet(e) {
  var t = HtmlService.createTemplateFromFile('Index'); // ชื่อไฟล์ HTML -> Index.html
  return t
    .evaluate()
    .setTitle('ระบบบันทึกการปฏิบัติงานประจำวัน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**********************
 * DASHBOARD DATA
 * baseDateIso: yyyy-MM-dd (วันที่เลือกในหน้าเว็บ) – แสดงล่วงหน้า 3 วัน
 **********************/
function getDashboardData(baseDateIso) {
  var start = new Date();
  if (baseDateIso) {
    var p = baseDateIso.split('-');
    start = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }
  start.setHours(0, 0, 0, 0);

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var todayKey = Utilities.formatDate(today, TZ, 'yyyy-MM-dd');

  var end = new Date(start);
  end.setDate(end.getDate() + 3); // ล่วงหน้า 3 วัน

  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  var events = cal.getEvents(start, end);

  var statusMap = getStatusMap_();
  var now = new Date();

  var list = events
    .map(function (ev, idx) {
      var s = ev.getStartTime();
      var e = ev.getEndTime();
      var dateKey = Utilities.formatDate(s, TZ, 'yyyy-MM-dd');

      var key = ev.getId();
      var st = statusMap[key] || {};

      var workStatus = st.status || (st.pdfUrl ? 'done' : ''); // ถ้ามี PDF แตะว่า done
      var statusText;
      if (workStatus === 'done') {
        statusText = 'ทำงานเสร็จแล้ว (บันทึกแล้ว)';
      } else if (workStatus === 'working') {
        statusText = 'กำลังทำงาน';
      } else {
        statusText = 'ยังไม่บันทึก';
      }

      var timeRange =
        Utilities.formatDate(s, TZ, 'HH:mm') +
        ' - ' +
        Utilities.formatDate(e, TZ, 'HH:mm');

      return {
        id: key,
        order: idx + 1,
        title: ev.getTitle(),
        description: ev.getDescription() || '',
        location: ev.getLocation() || '',
        startIso: s.toISOString(),
        endIso: e.toISOString(),
        dateLabel: formatThaiDateShort_(s),
        timeRange: timeRange,

        // สถานะ & ไฟล์
        statusText: statusText,
        workStatus: workStatus,          // '', 'working', 'done'
        pdfUrl: st.pdfUrl || '',
        qrUrl: st.qrUrl || '',
        evalStatus: st.evalStatus || '', // 'pending', 'done' (ถ้ามี)
        canEdit: dateKey === todayKey    // บันทึกได้เฉพาะ "วันนี้"
      };
    })
    .sort(function (a, b) {
      return new Date(a.startIso) - new Date(b.startIso);
    });

  var summary = { total: list.length, pending: 0, logged: 0 };
  list.forEach(function (r) {
    if (r.workStatus === 'done') summary.logged++;
    else summary.pending++;
  });

  return {
    nowIso: now.toISOString(),
    nowLabel: formatThaiDateTime_(now),
    events: list,
    summary: summary
  };
}

/**********************
 * SAVE WORK RESULT – บันทึก + PDF + QR
 **********************/
function saveWorkResult(form) {
  var eventId = form.eventId;
  var note    = form.note || '';
  var lat     = form.lat  || '';
  var lng     = form.lng  || '';

  var cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  var event = cal.getEventById(eventId);
  if (!event) throw new Error('ไม่พบงานใน Calendar');

  // รูปผลงาน (สูงสุด 6 รูป)
  var blobs = [];
  if (form.workImages) {
    var f = form.workImages;
    if (Object.prototype.toString.call(f) === '[object Array]') {
      for (var i = 0; i < f.length && i < 6; i++) {
        blobs.push(f[i]);
      }
    } else {
      blobs.push(f);
    }
  }

  var detail = { note: note, lat: lat, lng: lng };
  var pdfObj = createWorkPdf_(event, detail, blobs);
  var pdfFile = pdfObj.pdfFile;
  var docId   = pdfObj.docId;

  var st = event.getStartTime();
  var dateText = formatThaiDateShort_(st);

  // สร้าง URL แบบประเมินให้ผูกกับเอกสารนี้
  var evalId = new Date().getTime().toString() + '_' + eventId;
  var evalUrl =
    EVAL_WEBAPP_URL +
    '?evalId=' + encodeURIComponent(evalId) +
    '&eventId=' + encodeURIComponent(eventId) +
    '&docId=' + encodeURIComponent(docId);

  // ใช้ quickchart.io สร้างรูป QR
  var qrUrl =
    'https://quickchart.io/qr?text=' +
    encodeURIComponent(evalUrl) +
    '&size=300';

  var sheet = getOrCreateSheet_('WorkLog');
  sheet.appendRow([
    new Date(),        // 0 Timestamp
    eventId,           // 1 EventId
    dateText,          // 2 DateText
    event.getTitle(),  // 3 Title
    note,              // 4 Note
    lat,               // 5 Lat
    lng,               // 6 Lng
    pdfFile.getUrl(),  // 7 PdfUrl
    qrUrl,             // 8 QrUrl
    'pending',         // 9 EvalStatus
    docId,             //10 DocId
    'done'             //11 Status (ทำงานเสร็จแล้ว)
  ]);

  return {
    pdfUrl: pdfFile.getUrl(),
    qrUrl: qrUrl
  };
}

/**********************
 * UPDATE WORK STATUS – pending / working / done
 * ใช้สำหรับปุ่ม "เริ่มงาน"
 **********************/
function setWorkStatus(eventId, status) {
  if (!eventId) throw new Error('eventId ว่าง');

  var sh = getOrCreateSheet_('WorkLog');
  var values = sh.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < values.length; i++) {
    if (values[i][1] === eventId) { // column B = EventId
      rowIndex = i + 1; // แถวจริง (เริ่มที่ 1)
    }
  }

  var cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  var event = cal.getEventById(eventId);

  if (rowIndex === -1) {
    // ยังไม่มีบันทึกเลย → สร้างแถวใหม่เก็บสถานะอย่างเดียว
    var st = event.getStartTime();
    var dateText = formatThaiDateShort_(st);
    sh.appendRow([
      new Date(),        // 0 Timestamp
      eventId,           // 1 EventId
      dateText,          // 2 DateText
      event ? event.getTitle() : '', // 3 Title
      '',  // 4 Note
      '',  // 5 Lat
      '',  // 6 Lng
      '',  // 7 PdfUrl
      '',  // 8 QrUrl
      '',  // 9 EvalStatus
      '',  //10 DocId
      status //11 Status
    ]);
  } else {
    sh.getRange(rowIndex, 12).setValue(status); // col 12 = Status
  }

  return { status: status };
}

/**********************
 * CREATE PDF – รายงาน + รูปภาพ
 **********************/
function createWorkPdf_(event, detail, imageBlobs) {
  var folder = DriveApp.getFolderById(FOLDER_ID);

  var st = event.getStartTime();
  var et = event.getEndTime();
  var title    = event.getTitle();
  var location = event.getLocation() || '-';

  var dateText = formatThaiDateShort_(st);
  var timeRange =
    Utilities.formatDate(st, TZ, 'HH:mm') +
    ' - ' +
    Utilities.formatDate(et, TZ, 'HH:mm');

  var docName =
    'รายงานการปฏิบัติงาน_' +
    Utilities.formatDate(st, TZ, 'yyyyMMdd_HHmm');
  var doc  = DocumentApp.create(docName);
  var docId = doc.getId();
  var body = doc.getBody();
  body.clear();

  // หัวบริษัท
  var companyPara = body.appendParagraph(COMPANY_NAME);
  companyPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  companyPara.editAsText().setBold(true).setFontSize(20);

  body.appendParagraph('');
  var titlePara = body.appendParagraph('ใบรายงานการปฏิบัติงาน');
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titlePara.editAsText().setBold(true).setFontSize(16);
  body.appendParagraph('');

  // ตารางข้อมูลหลัก
  var infoTable = body.appendTable([
    ['ชื่องาน', title],
    ['วันที่', dateText],
    ['เวลา', timeRange],
    ['สถานที่', location]
  ]);
  infoTable.setBorderWidth(0);
  for (var r = 0; r < infoTable.getNumRows(); r++) {
    infoTable.getCell(r, 0).editAsText().setBold(true);
  }
  body.appendParagraph('');

  // รายละเอียด + พิกัด
  var dTitle = body.appendParagraph('รายละเอียดการปฏิบัติงาน');
  dTitle.setBold(true);
  if (detail.lat && detail.lng) {
    body.appendParagraph('พิกัดสถานที่: ' + detail.lat + ', ' + detail.lng);
    body.appendParagraph(''); // เว้น 1 บรรทัด
  }
  body.appendParagraph(detail.note || '-');
  body.appendParagraph('');

  // รูปภาพผลงาน
  var validImages = [];
  if (imageBlobs && imageBlobs.length) {
    imageBlobs.forEach(function (b, idx) {
      try {
        if (!b || typeof b.getBytes !== 'function') return;

        var ct = b.getContentType && b.getContentType();
        if (!ct) return;

        if (ct.indexOf('image/') === 0) {
          if (ct !== 'image/png' &&
              ct !== 'image/jpeg' &&
              ct !== 'image/gif') {
            b = Utilities.newBlob(
              b.getBytes(),
              'image/png',
              (b.getName() || 'image') + '.png'
            );
          }
          validImages.push(b);
        } else if (ct === 'application/octet-stream' ||
                   ct === 'binary/octet-stream') {
          b = Utilities.newBlob(
            b.getBytes(),
            'image/png',
            (b.getName() || 'image') + '.png'
          );
          validImages.push(b);
        }
      } catch (err) {
        Logger.log('Skip invalid image #' + (idx + 1) + ' : ' + err);
      }
    });
  }

  if (validImages.length) {
    body.appendPageBreak();

    var imgTitle = body.appendParagraph('รูปภาพผลงาน');
    imgTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    imgTitle.setBold(true);
    body.appendParagraph('');

    var sizePt = 6 * 72; // 6 นิ้ว

    validImages.slice(0, 6).forEach(function (blob, i) {
      try {
        var cap = body.appendParagraph('รูปที่ ' + (i + 1));
        cap.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        cap.setBold(true);

        var p = body.appendParagraph('');
        p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var img = p.appendInlineImage(blob);
        img.setWidth(sizePt).setHeight(sizePt);

        body.appendParagraph('');
      } catch (err) {
        Logger.log('Error inserting image #' + (i + 1) + ' : ' + err);
      }
    });
  }

  doc.saveAndClose();

  var pdfBlob = doc.getAs(MimeType.PDF);
  var pdfName =
    'รายงานการปฏิบัติงาน_' +
    Utilities.formatDate(st, TZ, 'yyyyMMdd_HHmm') +
    '.pdf';

  var pdfFile = folder.createFile(pdfBlob.setName(pdfName));
  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  // ย้ายไฟล์เอกสารต้นฉบับเข้าโฟลเดอร์เดียวกัน
  var docFile = DriveApp.getFileById(docId);
  folder.addFile(docFile);
  try {
    DriveApp.getRootFolder().removeFile(docFile);
  } catch (e) {
    Logger.log('Cannot remove file from root: ' + e);
  }

  return { pdfFile: pdfFile, docId: docId };
}

/**********************
 * SHEET HELPERS
 **********************/
function getLogSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LOG_SPREADSHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // ถ้าเปิดไม่ได้จะไปสร้างใหม่ด้านล่าง
    }
  }

  var active = SpreadsheetApp.getActive();
  if (active) {
    props.setProperty('LOG_SPREADSHEET_ID', active.getId());
    return active;
  }

  var ss = SpreadsheetApp.create('WorkLog');
  props.setProperty('LOG_SPREADSHEET_ID', ss.getId());
  return ss;
}

function getOrCreateSheet_(name) {
  var ss = getLogSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (name === 'WorkLog' && sh.getLastRow() === 0) {
    sh.appendRow([
      'Timestamp',  // A
      'EventId',    // B
      'DateText',   // C
      'Title',      // D
      'Note',       // E
      'Lat',        // F
      'Lng',        // G
      'PdfUrl',     // H
      'QrUrl',      // I
      'EvalStatus', // J
      'DocId',      // K
      'Status'      // L
    ]);
  }
  return sh;
}

function getStatusMap_() {
  var sh = getOrCreateSheet_('WorkLog');
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var eventId = row[1];
    if (!eventId) continue;
    map[eventId] = {
      pdfUrl: row[7] || '',
      qrUrl: row[8] || '',
      evalStatus: row[9] || '',
      docId: row[10] || '',
      status: row[11] || ''
    };
  }
  return map;
}

/**********************
 * DATE HELPERS
 **********************/
function formatThaiDateShort_(date) {
  var thMonths = [
    'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
  ];
  var d = new Date(date);
  var day = d.getDate();
  var month = thMonths[d.getMonth()];
  var year = d.getFullYear() + 543;
  return day + ' ' + month + ' ' + year;
}

function formatThaiDateTime_(date) {
  var dStr = formatThaiDateShort_(date);
  var tStr = Utilities.formatDate(date, TZ, 'HH:mm');
  return dStr + ' เวลา ' + tStr + ' น.';
}
