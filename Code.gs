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

function doGet(e) {
  if (e && e.parameter && e.parameter.kind === 'hotel') return hotelGet_();
  const rows = readRows_(getSheet_());
  return json_({ ok: true, total: TOTAL_SEATS, blocked: BLOCKED_SEATS, taken: publicTaken_(rows) });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  if (data.kind === 'hotel') return hotelPost_(data);

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

/* ===================== ГОТЕЛЬ ===================== */

const HOTEL_SHEET = 'Поселення';
// Номер кімнати → кількість ліжок. Має збігатися з ROOMS у hotel.html.
const HOTEL_ROOMS = {
  '1': 4, '2': 4, '3': 4, '4': 4, '5': 4, '6': 4, '7': 4, '8': 4, '9': 4, '10': 4,
  '11': 4, '12': 4, '13': 4, '14': 4, '15': 4, '16': 4, '17': 4, '18': 4, '19': 4, '20': 4,
};

function hotelSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOTEL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HOTEL_SHEET);
    sh.appendRow(['Час', 'Кімната', 'Ліжко', 'Прізвище', "Ім'я", 'Ключ']);
    sh.setFrozenRows(1);
    sh.getRange('A1:F1').setFontWeight('bold');
  }
  return sh;
}

function hotelRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues()
    .filter(r => r[1] !== '' && r[1] !== null)
    .map(r => ({
      room: String(r[1]),
      bed: Number(r[2]),
      last: String(r[3]),
      first: String(r[4]),
      key: String(r[5] || makeKey_(r[4], r[3])),
    }));
}

function hotelPublic_(rows) {
  return rows.map(r => ({ room: r.room, bed: r.bed, name: r.last + ' ' + r.first }));
}

function hotelGet_() {
  return json_({ ok: true, rooms: HOTEL_ROOMS, taken: hotelPublic_(hotelRows_(hotelSheet_())) });
}

function hotelPost_(data) {
  const first = cleanName_(data.first);
  const last = cleanName_(data.last);
  if (!first || !last) return json_({ ok: false, error: 'bad_name' });
  const key = makeKey_(first, last);
  const room = String(data.room || '');
  const bed = Number(data.bed);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ ok: false, error: 'busy' });
  try {
    const sh = hotelSheet_();
    const rows = hotelRows_(sh);
    const mine = rows.find(r => r.key === key);

    if (data.action === 'check') {
      return json_({ ok: true, room: mine ? mine.room : null, bed: mine ? mine.bed : null, taken: hotelPublic_(rows) });
    }
    if (mine) return json_({ ok: false, error: 'already', room: mine.room, bed: mine.bed, taken: hotelPublic_(rows) });
    const cap = HOTEL_ROOMS[room];
    if (!cap || !Number.isInteger(bed) || bed < 1 || bed > cap) return json_({ ok: false, error: 'bad_seat' });
    if (rows.some(r => r.room === room && r.bed === bed)) {
      return json_({ ok: false, error: 'taken', taken: hotelPublic_(rows) });
    }

    sh.appendRow([new Date(), room, bed, last, first, key]);
    SpreadsheetApp.flush();
    rows.push({ room, bed, first, last, key });
    return json_({ ok: true, room, bed, first, last, taken: hotelPublic_(rows) });
  } finally {
    lock.releaseLock();
  }
}
