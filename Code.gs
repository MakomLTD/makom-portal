// =====================================================
//  MAKOM – פורטל פרויקטים | Google Apps Script
//  Code.gs  v4.0
//
//  ★ שלב אחד בלבד לפני שהכל עובד:
//    1. פתחו Apps Script Editor → הדביקו קוד זה
//    2. הפעילו את הפונקציה fullSetup() (▶ Run)
//    3. אשרו הרשאות
//    4. פרסו: Deploy → New Deployment → Web App → Anyone
//    5. העתיקו את ה-URL לכלי מחולל הקישורים
// =====================================================

const SHEET_NAME = 'projects';

// קורא SPREADSHEET_ID מ-Script Properties (נשמר על-ידי fullSetup)
function getSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
}

// ─── נקודת כניסה ─────────────────────────────────────
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const SPREADSHEET_ID = getSpreadsheetId_();

  // [1] JSON API — לרענון חי מצד הלקוח
  if (params.action === 'data') {
    try {
      return jsonResponse_({
        ok: true,
        projects: getProjects_(params),
        updatedAt: new Date().toISOString()
      });
    } catch (error) {
      return jsonResponse_({ ok: false, error: String(error.message || error), projects: [] });
    }
  }

  // [2] בדיקת תקינות פריסה
  if (params.test === '1') {
    return HtmlService
      .createHtmlOutput(
        '<div dir="rtl" style="font-family:Arial,sans-serif;padding:32px;color:#11193F">' +
        '<h1>✅ Apps Script עובד</h1>' +
        '<p>הפריסה תקינה. עכשיו פתחו את הכתובת ללא ?test=1 כדי לראות את הפורטל.</p></div>'
      )
      .setTitle('בדיקת Apps Script')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // [3] תבנית HTML מלאה
  try {
    const template = HtmlService.createTemplateFromFile('index');
    template.projectsJson = JSON.stringify(getProjects_(params));
    template.scriptUrl    = ScriptApp.getService().getUrl();
    template.clientId     = params.client || '';
    template.tokenId      = params.token  || '';

    return template
      .evaluate()
      .setTitle('פורטל בקרת פרויקטים')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService
      .createHtmlOutput(buildErrorHtml_(error))
      .setTitle('שגיאה בטעינת הפורטל')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ─── שכבת נתונים ──────────────────────────────────────
function getProjects_(params) {
  const SPREADSHEET_ID = getSpreadsheetId_();
  if (!SPREADSHEET_ID) {
    return getSampleProjects_();
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('לא נמצא גיליון בשם "' + SHEET_NAME + '"');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values.shift().map(normalizeHeader_);
  const client  = String(params.client || '').trim().toLowerCase();
  const token   = String(params.token  || '').trim();

  return values
    .map(rowToObject_(headers))
    .filter(r => String(r.active || '').toUpperCase() !== 'FALSE')
    .filter(r => !client || String(r.client_id || '').trim().toLowerCase() === client)
    .filter(r => !token  || String(r.token     || '').trim()               === token)
    .map(r => ({
      project_name:     r.project_name,
      status:           r.status,
      description:      r.description,
      site_url:         r.site_url,
      youtube_url:      r.youtube_url,
      image_url:        r.image_url,
      location:         r.location,
      lat:              r.lat,
      lng:              r.lng,
      planned_end_date: r.planned_end_date,
      actual_end_date:  r.actual_end_date,
      tags:             r.tags,
      progress:         r.progress,
      risk:             r.risk,
      next_step:        r.next_step
    }));
}

function rowToObject_(headers) {
  return function(row) {
    return headers.reduce((obj, h, i) => { obj[h] = row[i] || ''; return obj; }, {});
  };
}

function normalizeHeader_(header) {
  const clean = String(header || '').trim();
  const map = {
    'לקוח':'client_id', 'מזהה לקוח':'client_id',
    'טוקן':'token',
    'פעיל':'active',
    'שם פרויקט':'project_name', 'פרויקט':'project_name',
    'סטטוס':'status',
    'תיאור':'description',
    'קישור אתר':'site_url', 'קישור':'site_url',
    'קישור יוטיוב':'youtube_url', 'סרטונים':'youtube_url', 'רשימת סרטונים':'youtube_url',
    'תמונה':'image_url', 'קישור תמונה':'image_url',
    'מיקום':'location',
    'קו רוחב':'lat',
    'קו אורך':'lng',
    'תאריך סיום מתוכנן':'planned_end_date', 'סיום מתוכנן':'planned_end_date', 'תאריך סיום':'planned_end_date',
    'תאריך סיום בפועל':'actual_end_date',   'סיום בפועל':'actual_end_date',
    'תגיות':'tags',
    'התקדמות':'progress', 'אחוז התקדמות':'progress',
    'סיכון':'risk',       'רמת סיכון':'risk',
    'פעולה הבאה':'next_step', 'צעד הבא':'next_step',
    'client_id':'client_id', 'token':'token', 'active':'active',
    'project_name':'project_name', 'status':'status', 'description':'description',
    'site_url':'site_url', 'youtube_url':'youtube_url', 'image_url':'image_url',
    'location':'location', 'lat':'lat', 'lng':'lng',
    'planned_end_date':'planned_end_date', 'actual_end_date':'actual_end_date',
    'tags':'tags', 'progress':'progress', 'risk':'risk', 'next_step':'next_step'
  };
  return map[clean] || clean;
}

function getSampleProjects_() {
  return [];
}

function setupSheet() {
  const SPREADSHEET_ID = getSpreadsheetId_();
  if (!SPREADSHEET_ID) {
    throw new Error('הפעל קודם fullSetup() — לא נמצא SPREADSHEET_ID בהגדרות.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  _populateSheet_(ss);
}

function fullSetup() {
  const props = PropertiesService.getScriptProperties();

  const existing = DriveApp.getFoldersByName('MAKOM פורטל');
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder('MAKOM פורטל');
  Logger.log('📁 תיקייה: ' + folder.getUrl());

  const tempSs = SpreadsheetApp.create('MAKOM פורטל – פרויקטים');
  const ssFile = DriveApp.getFileById(tempSs.getId());
  ssFile.moveTo(folder);
  ssFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const sheetId = tempSs.getId();
  props.setProperty('SPREADSHEET_ID', sheetId);
  Logger.log('📊 Sheet ID: ' + sheetId);
  Logger.log('📊 Sheet URL: ' + tempSs.getUrl());

  _populateSheet_(tempSs);

  Logger.log('');
  Logger.log('✅ הכל מוכן!');
  Logger.log('👉 עכשיו פרסם: Deploy → New Deployment → Web App → Anyone');
  Logger.log('👉 העתק את ה-Web App URL לכלי מחולל הקישורים.');
}

function _populateSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const sheets = ss.getSheets();
  if (sheets.length > 1) {
    const blankSheet = sheets.find(function(s) {
      return s.getName() === 'Sheet1' || s.getName() === 'גיליון1';
    });
    if (blankSheet) try { ss.deleteSheet(blankSheet); } catch(e) {}
  }
  sheet.clear();

  const headers = [
    'client_id','token','active','project_name','status','description',
    'site_url','youtube_url','image_url','location','lat','lng',
    'planned_end_date','actual_end_date','tags','progress','risk','next_step'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#11193F').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log('✅ כותרות Sheet נוצרו — הוסיפו שורות לקוחות ופרויקטים ידנית.');
}

function debugProjects() {
  const rows = getProjects_({ client: '', token: '' });
  Logger.log(JSON.stringify(rows, null, 2));
  return rows;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildErrorHtml_(error) {
  const msg = escapeHtml_(error && (error.stack || error.message) ? (error.stack || error.message) : String(error));
  return '<div dir="rtl" style="font-family:Arial,sans-serif;padding:32px;line-height:1.7;color:#11193F">' +
    '<h1>שגיאה בטעינת הפורטל</h1>' +
    '<p>בדקו ששם קובץ ה-HTML ב-Apps Script הוא <strong>index</strong> ושבוצעה פריסה כגרסה חדשה.</p>' +
    '<pre style="direction:ltr;text-align:left;white-space:pre-wrap;background:#FDEDEA;color:#7A2B21;border:1px solid #F4B4AA;border-radius:12px;padding:14px">' + msg + '</pre>' +
    '</div>';
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
