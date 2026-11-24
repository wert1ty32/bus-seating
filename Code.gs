/**
 * Бекенд розсадки в автобусі (Google Apps Script, прив'язаний до Google Таблиці).
 *
 * Встановлення:
 *  1. Створи Google Таблицю → Розширення → Apps Script.
 *  2. Встав цей код замість вмісту Code.gs і збережи.
 *  3. Deploy → New deployment → тип "Web app":
 *       Execute as: Me    |    Who has access: Anyone
 *  4. Скопіюй URL, що закінчується на /exec, і встав у index.html (SCRIPT_URL).
 *  Після кожної зміни коду: Deploy → Manage deployments → Edit → Version: New version.
 */

const SHEET_NAME = 'Розсадка';
const TOTAL_SEATS = 67;
const BLOCKED_SEATS = []; // напр. [1, 2] — місця, які не можна обрати (керівник, гід)

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Час', 'Місце', 'Прізвище', "Ім'я", 'Ключ']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
  }
  return sh;
}

function readRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[1] !== '' && r[1] !== null)
    .map(r => ({
      seat: Number(r[1]),
      last: String(r[2]),
      first: String(r[3]),
      key: String(r[4] || makeKey_(r[3], r[2])),
    }));
}

function makeKey_(first, last) {
  return (String(last) + '|' + String(first)).toLowerCase().replace(/\s+/g, ' ').trim();
}

function cleanName_(s) {
  s = String(s || '').trim().replace(/\s+/g, ' ');
  if (!/^[A-Za-z][A-Za-z' -]{0,39}$/.test(s)) return null;
  return s.split(/([ -])/).map(p => p.length > 1 || /[A-Za-z]/.test(p)
    ? p.charAt(0).toUpperCase() + p.slice(1).toLowerCase() : p).join('');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function publicTaken_(rows) {
  return rows.map(r => ({ seat: r.seat, name: r.last + ' ' + r.first }));
}

function doGet() {
  const rows = readRows_(getSheet_());
  return json_({ ok: true, total: TOTAL_SEATS, blocked: BLOCKED_SEATS, taken: publicTaken_(rows) });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const first = cleanName_(data.first);
  const last = cleanName_(data.last);
  const seat = Number(data.seat);
  if (!first || !last) return json_({ ok: false, error: 'bad_name' });
  const key = makeKey_(first, last);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ ok: false, error: 'busy' });
  try {
    const sh = getSheet_();
    const rows = readRows_(sh);
    const mine = rows.find(r => r.key === key);

    // Перевірка без бронювання: чи вже є місце в цієї людини
    if (data.action === 'check') {
      return json_({ ok: true, seat: mine ? mine.seat : null, taken: publicTaken_(rows) });
    }

    if (mine) return json_({ ok: false, error: 'already', seat: mine.seat, taken: publicTaken_(rows) });
    if (!Number.isInteger(seat) || seat < 1 || seat > TOTAL_SEATS || BLOCKED_SEATS.includes(seat)) {
      return json_({ ok: false, error: 'bad_seat' });
    }
    if (rows.some(r => r.seat === seat)) {
      return json_({ ok: false, error: 'taken', taken: publicTaken_(rows) });
    }

    sh.appendRow([new Date(), seat, last, first, key]);
    SpreadsheetApp.flush();
    rows.push({ seat, first, last, key });
    return json_({ ok: true, seat, first, last, taken: publicTaken_(rows) });
  } finally {
    lock.releaseLock();
  }
}
