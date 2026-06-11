/**
 * ====== НАСТРОЙКИ ======
 * Структура листов (уже существуют в таблице):
 *   PROMOCODES: code | used (TRUE/FALSE)
 *   CONFIG:     key  | value   (строки: start, end, dailyWins)
 *   USERS:      userId | date  (одна строка = одна попытка)
 *   WIN_LOG:    date | code | userId
 */
const SPREADSHEET_ID = 'ВСТАВЬТЕ_СЮДА_ID_ВАШЕЙ_ТАБЛИЦЫ';

const SHEET_PROMO   = 'PROMOCODES';
const SHEET_CONFIG  = 'CONFIG';
const SHEET_USERS   = 'USERS';
const SHEET_WINLOG  = 'WIN_LOG';

const MAX_ATTEMPTS_PER_DAY = 3;
const WIN_CHANCE           = 0.18; // шанс выигрыша на одну попытку

/**
 * Веб-приложение: принимает запрос от игры.
 * Тело запроса (Content-Type: text/plain, чтобы избежать CORS-preflight),
 * JSON вида: { "ip": "1.2.3.4" }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const userId = (body.ip || '').toString().trim();

    if (!userId) {
      return jsonResponse({ ok: false, error: 'no_ip', message: 'IP не передан' });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const tz = Session.getScriptTimeZone();
      const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

      const config = readConfig(ss, tz);

      if (config.start && todayStr < config.start) {
        return jsonResponse({ ok: false, error: 'not_started', message: 'Акция ещё не началась' });
      }
      if (config.end && todayStr > config.end) {
        return jsonResponse({ ok: false, error: 'ended', message: 'Акция завершена' });
      }

      const usersSheet = ss.getSheetByName(SHEET_USERS);
      const winLogSheet = ss.getSheetByName(SHEET_WINLOG);

      const attemptsToday = countByDate(usersSheet, 1, 0, userId, todayStr, tz);
      const userWonToday  = countByDate(winLogSheet, 0, 2, userId, todayStr, tz) > 0;

      if (userWonToday) {
        return jsonResponse({
          ok: false,
          error: 'already_won',
          message: 'Вы уже выиграли сегодня! Приходите завтра.',
          attemptsLeft: 0
        });
      }

      if (attemptsToday >= MAX_ATTEMPTS_PER_DAY) {
        return jsonResponse({
          ok: false,
          error: 'limit_reached',
          message: 'Лимит попыток на сегодня исчерпан (3 в день)',
          attemptsToday: attemptsToday,
          attemptsLeft: 0
        });
      }

      const attemptNumber = attemptsToday + 1;
      usersSheet.appendRow([userId, new Date()]);

      const winsToday = countByDate(winLogSheet, 0, null, null, todayStr, tz);

      let win = false;
      let promoCode = '';

      // Общий дневной лимит призов на всех пользователей (config.dailyWins, обычно = 1)
      if (winsToday < config.dailyWins && Math.random() < WIN_CHANCE) {
        promoCode = pickPromoCode(ss);
        if (promoCode) {
          win = true;
          winLogSheet.appendRow([new Date(), promoCode, userId]);
        }
      }

      return jsonResponse({
        ok: true,
        win: win,
        promoCode: win ? promoCode : null,
        attemptNumber: attemptNumber,
        attemptsLeft: MAX_ATTEMPTS_PER_DAY - attemptNumber
      });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: 'server_error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ ok: true, message: 'Game stats API работает' });
}

/**
 * Читает CONFIG (key/value) и приводит даты к строкам yyyy-MM-dd
 */
function readConfig(ss, tz) {
  const data = ss.getSheetByName(SHEET_CONFIG).getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (!key) continue;
    if (value instanceof Date) {
      cfg[key] = Utilities.formatDate(value, tz, 'yyyy-MM-dd');
    } else {
      cfg[key] = value;
    }
  }
  cfg.dailyWins = Number(cfg.dailyWins || 0);
  return cfg;
}

/**
 * Считает строки за сегодняшнюю дату (колонка dateCol), опционально
 * фильтруя по совпадению значения userId в колонке userCol.
 */
function countByDate(sheet, dateCol, userCol, userId, todayStr, tz) {
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][dateCol];
    if (!(rowDate instanceof Date)) continue;
    if (Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd') !== todayStr) continue;
    if (userCol !== null && data[i][userCol] !== userId) continue;
    count++;
  }
  return count;
}

/**
 * Берёт случайный неиспользованный промокод из PROMOCODES и помечает used = TRUE.
 * Возвращает '' если коды закончились.
 */
function pickPromoCode(ss) {
  const sheet = ss.getSheetByName(SHEET_PROMO);
  const data = sheet.getDataRange().getValues();

  const available = [];
  for (let i = 1; i < data.length; i++) {
    const code = data[i][0];
    const used = data[i][1];
    const isUsed = used === true || String(used).toUpperCase() === 'TRUE';
    if (code && !isUsed) {
      available.push({ row: i + 1, code: String(code) });
    }
  }
  if (available.length === 0) return '';

  const chosen = available[Math.floor(Math.random() * available.length)];
  sheet.getRange(chosen.row, 2).setValue(true);
  return chosen.code;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
