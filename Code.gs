const CONFIG = {
  SPREADSHEET_ID: '1w0JZfHJAAQdDy8OfvOxTz4pkPh9I-BdB0d-Mq1rt8lQ',
  SHEET_NAME: 'baocao',
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  FIXED_TOTAL: 1,
  FIXED_CLASS_SIZE: 350,
  DEFAULT_CLASS_NAME: 'TSA',
  DEFAULT_CONTENT: 'Luyện đề',
  DEFAULT_QLL: 'Đăng Minh',
  QLL_OPTIONS: ['Đăng Minh', 'Đăng Khoa'],
  HEADERS: [
    'Date',
    'Lớp học',
    'Môn học',
    'Thầy/cô',
    'Tổng',
    'Sĩ số lớp',
    'Ca học',
    'QLL',
    'Link record buổi học (youtube)',
    'Buổi số (các em ghi đủ số buổi nhé )'
  ],
  SESSION_COLUMN_INDEX: 10
};

function doGet(e) {
  var action = getParam_(e, 'action');
  var prefix = sanitizePrefix_(getParam_(e, 'prefix'));

  if (action === 'init') {
    return output_(getInitialData_(), prefix);
  }

  if (action === 'saveLesson') {
    var payload = {
      sessionNumber: getParam_(e, 'sessionNumber'),
      className: getParam_(e, 'className') || CONFIG.DEFAULT_CLASS_NAME,
      subject: getParam_(e, 'subject'),
      teacher: getParam_(e, 'teacher'),
      date: getParam_(e, 'date'),
      qll: getParam_(e, 'qll') || CONFIG.DEFAULT_QLL,
      classTime: getParam_(e, 'classTime'),
      youtubeLink: getParam_(e, 'youtubeLink'),
      lessonContent: getParam_(e, 'lessonContent') || CONFIG.DEFAULT_CONTENT
    };
    return output_(saveLesson_(payload), prefix);
  }

  return output_({
    success: true,
    message: 'Apps Script API is running.',
    usage: {
      init: '?action=init',
      saveLesson: '?action=saveLesson&sessionNumber=B01&subject=To%C3%A1n%20%C4%90%E1%BA%A1i...'
    }
  }, prefix);
}

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var body = JSON.parse(raw);
    var action = body.action || '';

    if (action === 'saveLesson') {
      return output_(saveLesson_(body.payload || {}), '');
    }

    return output_({ success: false, message: 'Action không hợp lệ.' }, '');
  } catch (error) {
    return output_({
      success: false,
      message: error && error.message ? error.message : 'Có lỗi khi xử lý request.'
    }, '');
  }
}

function getInitialData_() {
  ensureSheetSetup_();

  return {
    success: true,
    nextSessionNumber: getNextSessionNumber_(),
    yesterdayInput: getYesterdayInVietnam_(),
    defaultClassName: CONFIG.DEFAULT_CLASS_NAME,
    defaultContent: CONFIG.DEFAULT_CONTENT,
    defaultQll: CONFIG.DEFAULT_QLL,
    qllOptions: CONFIG.QLL_OPTIONS,
    infoMessage: 'Đã kết nối sheet baocao.'
  };
}

function saveLesson_(payload) {
  try {
    ensureSheetSetup_();
    validatePayload_(payload);

    var sheet = getSheet_();
    var normalizedSession = normalizeSessionNumber_(payload.sessionNumber);
    var qll = String(payload.qll || CONFIG.DEFAULT_QLL).trim() || CONFIG.DEFAULT_QLL;

    var row = [
      formatDateForSheet_(payload.date),
      String(payload.className || CONFIG.DEFAULT_CLASS_NAME).trim() || CONFIG.DEFAULT_CLASS_NAME,
      String(payload.subject || '').trim(),
      String(payload.teacher || '').trim(),
      CONFIG.FIXED_TOTAL,
      CONFIG.FIXED_CLASS_SIZE,
      String(payload.classTime || '').trim(),
      qll,
      String(payload.youtubeLink || '').trim(),
      normalizedSession
    ];

    var nextRow = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

    return {
      success: true,
      message: 'Đã export dữ liệu sang sheet baocao thành công.',
      nextSessionNumber: getNextSessionNumber_()
    };
  } catch (error) {
    return {
      success: false,
      message: error && error.message ? error.message : 'Có lỗi xảy ra khi export.'
    };
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  return sheet;
}

function ensureSheetSetup_() {
  var sheet = getSheet_();

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  } else {
    var existingHeaders = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).getValues()[0];
    var headerIsBlank = existingHeaders.every(function (v) {
      return String(v || '').trim() === '';
    });

    if (headerIsBlank) {
      sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    }
  }

  sheet.setFrozenRows(1);
}

function getNextSessionNumber_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 'B01';
  }

  var values = sheet
    .getRange(2, CONFIG.SESSION_COLUMN_INDEX, lastRow - 1, 1)
    .getValues()
    .flat();

  var maxNumber = 0;

  values.forEach(function (value) {
    var extracted = extractSessionNumber_(value);
    if (extracted > maxNumber) {
      maxNumber = extracted;
    }
  });

  return formatSessionNumber_(maxNumber + 1);
}

function extractSessionNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  var text = String(value).trim().toUpperCase();
  var match = text.match(/(\d+)/);
  return match ? Number(match[1]) : 0;
}

function formatSessionNumber_(num) {
  return 'B' + String(Number(num) || 1).padStart(2, '0');
}

function normalizeSessionNumber_(value) {
  var text = String(value || '').trim().toUpperCase();

  if (!text) {
    throw new Error('Thiếu số buổi.');
  }

  if (/^B\d+$/.test(text)) {
    return formatSessionNumber_(text.replace('B', ''));
  }

  if (/^\d+$/.test(text)) {
    return formatSessionNumber_(text);
  }

  return text;
}

function getYesterdayInVietnam_() {
  var todayInVietnam = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  var parts = todayInVietnam.split('-').map(Number);
  var vietnamDateAsUtc = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  vietnamDateAsUtc.setUTCDate(vietnamDateAsUtc.getUTCDate() - 1);
  return Utilities.formatDate(vietnamDateAsUtc, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function formatDateForSheet_(dateValue) {
  var text = String(dateValue || '').trim();
  if (!text) return '';

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    return text;
  }

  var parts = text.split('-');
  if (parts.length === 3) {
    return parts[2] + '/' + parts[1] + '/' + parts[0];
  }

  return text;
}

function validatePayload_(payload) {
  if (!payload) throw new Error('Không nhận được dữ liệu từ form.');
  if (!payload.sessionNumber) throw new Error('Thiếu số buổi.');
  if (!payload.subject) throw new Error('Thiếu môn học.');
  if (!payload.teacher) throw new Error('Thiếu thầy/cô.');
  if (!payload.date) throw new Error('Thiếu ngày.');
  if (!payload.qll) throw new Error('Thiếu QLL.');
  if (!payload.classTime) throw new Error('Thiếu ca học.');
  if (!payload.youtubeLink) throw new Error('Thiếu link Youtube.');
}

function getParam_(e, key) {
  return e && e.parameter && Object.prototype.hasOwnProperty.call(e.parameter, key)
    ? e.parameter[key]
    : '';
}

function sanitizePrefix_(prefix) {
  return String(prefix || '').replace(/[^\w.$]/g, '');
}

function output_(obj, prefix) {
  var text = JSON.stringify(obj);

  if (prefix) {
    return ContentService
      .createTextOutput(prefix + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JSON);
}
