/**
 * PTW Sample Test 1 — Results logger (Google Apps Script Web App)
 * -----------------------------------------------------------------
 * Writes one row per test attempt to the Google Sheet
 * "PTW Sample Test 1 — Results Log" (owner: imrannow@gmail.com).
 *
 *  - "register" event  -> new row: candidate details, Status = Started
 *  - "result"   event  -> same row updated: score, %, PASS/FAIL, per-question answers
 *
 * DEPLOY (one time):
 *  1. Open the sheet > Extensions > Apps Script. Replace everything with this file. Save.
 *  2. Deploy > New deployment > type "Web app".
 *       Execute as: Me (imrannow@gmail.com)
 *       Who has access: Anyone
 *  3. Authorise when asked, copy the Web app URL (ends in /exec)
 *     and paste it into SHEET_ENDPOINT at the top of the HTML file's <script>.
 *  4. After editing this script later: Deploy > Manage deployments > Edit > Version: New version.
 */

const SPREADSHEET_ID = '1XwpaiKUg-ui0CyE2huWoKD_nRspPuDfeVIgsP-JrslQ';
const SHEET_NAME = '';            // '' = first tab of the spreadsheet
const SHARED_KEY = 'NBTC-PTW-2026'; // must match LOG_KEY in the HTML file
const TIMEZONE = 'Asia/Kuwait';

const HEADERS = ['Attempt ID','Registered At','First Name','Last Name','Phone','Email','Status',
  'Started At','Submitted At','Time Taken','Submission Type','Score','Percentage','Result',
  'Correct','Incorrect','Unanswered','Flagged'];
for (let i = 1; i <= 50; i++) HEADERS.push('Q' + i);
HEADERS.push('Browser');

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = SHEET_NAME ? (ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME)) : ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  if (sh.getFrozenRows() === 0) {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0b56a7').setFontColor('#ffffff');
  }
  return sh;
}

function fmt_(ms) {
  return ms ? Utilities.formatDate(new Date(Number(ms)), TIMEZONE, 'yyyy-MM-dd HH:mm:ss') : '';
}

function clean_(v, max) {
  // Strip leading formula characters so candidate input can never run as a formula.
  let s = String(v == null ? '' : v).slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function findRow_(sh, attemptId) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === attemptId) return i + 2;
  }
  return -1;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== SHARED_KEY) return json_({ ok: false, error: 'bad key' });
    if (!d.attemptId) return json_({ ok: false, error: 'no attempt id' });

    const sh = getSheet_();
    const id = clean_(d.attemptId, 64);
    let row = findRow_(sh, id);

    if (d.type === 'register') {
      if (row === -1) {
        const r = new Array(HEADERS.length).fill('');
        r[0] = id;
        r[1] = fmt_(d.registeredAt);
        r[2] = clean_(d.firstName, 60);
        r[3] = clean_(d.lastName, 60);
        r[4] = clean_(d.phone, 30);
        r[5] = clean_(d.email, 120);
        r[6] = 'Started';
        r[7] = fmt_(d.startedAt);
        r[HEADERS.length - 1] = clean_(d.browser, 250);
        sh.appendRow(r);
      }
      return json_({ ok: true });
    }

    if (d.type === 'result') {
      if (row === -1) { // register event never arrived — create the row now
        sh.appendRow([id, fmt_(d.registeredAt), clean_(d.firstName, 60), clean_(d.lastName, 60),
          clean_(d.phone, 30), clean_(d.email, 120)]);
        row = sh.getLastRow();
      }
      const res = [
        'Submitted', fmt_(d.startedAt), fmt_(d.submittedAt), clean_(d.timeTaken, 20),
        clean_(d.submissionType, 30), Number(d.score), Number(d.percentage) / 100, clean_(d.result, 10),
        Number(d.correct), Number(d.incorrect), Number(d.unanswered), Number(d.flagged)
      ];
      sh.getRange(row, 7, 1, res.length).setValues([res]);
      sh.getRange(row, 13).setNumberFormat('0%');
      const answers = (d.answers || []).slice(0, 50).map(a => clean_(a, 10));
      while (answers.length < 50) answers.push('');
      sh.getRange(row, 19, 1, 50).setValues([answers]);
      sh.getRange(row, 14).setFontWeight('bold')
        .setFontColor(d.result === 'PASS' ? '#0f6b3c' : '#b3261e');
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown type' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, service: 'PTW Sample Test 1 logger' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
