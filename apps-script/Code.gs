/**
 * Прийом карт перевірки стану ведення учнівських зошитів.
 * Скрипт прив'язаний до Google Таблиці:
 *   аркуш «Карти» — один рядок на кожну надіслану карту (вчитель + предмет);
 *   аркуш «Класи» — один рядок на кожен клас із карти;
 *   аркуш «Зведення» — підсумки за предметами (формули, оновлюються самі).
 * Повторне надсилання тієї самої карти (той самий ID) оновлює її рядки, а не дублює.
 * Власнику таблиці надходить лист-сповіщення.
 *
 * Розгортання: Розгорнути → Нове розгортання → Тип «Вебзастосунок»;
 * «Виконувати від імені»: Я; «Хто має доступ»: Усі. Скопіюйте адресу, що закінчується на /exec,
 * і вставте її в index.html (CONFIG.endpoint).
 */
const SHEET_ID = '1ubqUdJEmxGPhftctTZzQ1aAvcp7NEFxcTZd8wGfCmt4';
const SHEET_CARDS = 'Карти';
const SHEET_CLASSES = 'Класи';
const SHEET_SUMMARY = 'Зведення';
const NOTIFY = true;
const MAX_BYTES = 200000;

const REQ_CODES = ['В1', 'В2', 'В3', 'В4', 'В5', 'В6', 'В7', 'В8', 'В9', 'В10', 'В11', 'В12', 'В13'];
const CARD_HEADERS = ['Надіслано', 'ID', 'ПІБ', 'Предмет', 'Email', 'Період', 'Класи', 'Учнів', 'Мають робочий зошит', '% робочих',
  'Мають зошит для контрольних', '% контрольних', 'Не здали зошит / фото', 'Дотримання вимог, %', 'Рівень']
  .concat(REQ_CODES)
  .concat(['Типові недоліки', 'Вжиті заходи', 'Потрібна допомога', 'Коментар', 'Надсилань']);
const CLASS_HEADERS = ['Надіслано', 'ID карти', 'ПІБ', 'Предмет', 'Клас', 'Учнів', 'Мають робочий зошит', 'Мають зошит для контрольних',
  'Перевірок за період', 'Як перевірялися', 'Не здали зошит / фото', '% робочих'];

function doGet() {
  return json_({ ok: true, service: 'notebooks-card' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > MAX_BYTES) return json_({ ok: false, error: 'empty_or_too_large' });
    const p = JSON.parse(raw);
    if (!p || !p.id || !p.data || !String(p.data.fio || '').trim()) return json_({ ok: false, error: 'bad_payload' });
    const id = String(p.id).slice(0, 64);
    const d = p.data, s = p.summary || {};
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shC = sheet_(ss, SHEET_CARDS, CARD_HEADERS);
    const shK = sheet_(ss, SHEET_CLASSES, CLASS_HEADERS);
    const now = new Date();
    const fio = String(d.fio).trim();
    const subject = String(d.subject || '');

    // Карта: оновити або додати
    let rowIdx = -1, count = 0;
    const last = shC.getLastRow();
    if (last > 1) {
      const ids = shC.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === id) { rowIdx = i + 2; break; } }
      if (rowIdx > 0) count = Number(shC.getRange(rowIdx, CARD_HEADERS.indexOf('Надсилань') + 1).getValue()) || 0;
    }
    const req = d.req || {};
    const row = [now, id, fio, subject, d.email || '', d.period || '', s.classes || '', s.students || 0, s.withBook || 0,
      s.bookPct == null ? '' : s.bookPct / 100, s.withControl || 0, s.controlPct == null ? '' : s.controlPct / 100,
      s.missing || 0, s.reqPct == null ? '' : s.reqPct / 100, s.level || '']
      .concat(REQ_CODES.map(function (c) { const v = req[c]; return v === 'na' ? 'н/з' : (typeof v === 'number' ? v : ''); }))
      .concat([(s.defects || []).join('; '), d.measures || '', d.needs || '', d.comment || '', count + 1]);
    if (rowIdx > 0) shC.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else { shC.appendRow(row); rowIdx = shC.getLastRow(); }
    shC.getRange(rowIdx, CARD_HEADERS.indexOf('% робочих') + 1).setNumberFormat('0%');
    shC.getRange(rowIdx, CARD_HEADERS.indexOf('% контрольних') + 1).setNumberFormat('0%');
    shC.getRange(rowIdx, CARD_HEADERS.indexOf('Дотримання вимог, %') + 1).setNumberFormat('0%');

    // Класи: видалити старі рядки цієї карти й додати нові
    const lastK = shK.getLastRow();
    if (lastK > 1) {
      const kIds = shK.getRange(2, 2, lastK - 1, 1).getValues();
      for (let i = kIds.length - 1; i >= 0; i--) { if (String(kIds[i][0]) === id) shK.deleteRow(i + 2); }
    }
    (p.classes || []).forEach(function (c) {
      const st = parseInt(c.students, 10), wb = parseInt(c.withBook, 10);
      shK.appendRow([now, id, fio, subject, c.cls || '', num_(c.students), num_(c.withBook), num_(c.withControl),
        num_(c.checks), c.modeLabel || '', num_(c.missing), (st > 0 && !isNaN(wb)) ? wb / st : '']);
      shK.getRange(shK.getLastRow(), CLASS_HEADERS.length).setNumberFormat('0%');
    });

    summarySheet_(ss);

    if (NOTIFY) {
      const me = Session.getEffectiveUser().getEmail();
      if (me) {
        MailApp.sendEmail({
          to: me,
          subject: (count > 0 ? 'Оновлено карту перевірки зошитів: ' : 'Нова карта перевірки зошитів: ') + fio,
          htmlBody: '<p><b>' + esc_(fio) + '</b>, ' + esc_(subject) + '</p>' +
            '<p>Класи: ' + esc_(s.classes || '—') + '<br>Мають робочі зошити: ' + (s.bookPct == null ? '—' : s.bookPct + ' %') +
            '<br>Дотримання єдиних вимог: ' + (s.reqPct == null ? '—' : s.reqPct + ' % (' + esc_(s.level) + ')') + '</p>' +
            '<p><a href="' + ss.getUrl() + '">Таблиця відповідей</a></p>'
        });
      }
    }
    return json_({ ok: true, updated: count > 0 });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) { }
  }
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3eaf5').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Аркуш «Зведення»: підсумки за предметами формулами з аркушів «Карти» і «Класи». */
function summarySheet_(ss) {
  let sh = ss.getSheetByName(SHEET_SUMMARY);
  if (sh) return;
  sh = ss.insertSheet(SHEET_SUMMARY);
  const subjects = ['Українська мова', 'Українська література', 'Українська мова та література', 'Зарубіжна література', 'Математика (алгебра, геометрія)', 'Англійська мова'];
  const head = ['Предмет', 'Карт надіслано', 'Учнів', 'Мають робочий зошит', '% робочих', 'Мають зошит для контрольних', '% контрольних', 'Не здали зошит / фото', 'Середнє дотримання вимог'];
  sh.appendRow(head);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e3eaf5').setWrap(true);
  subjects.concat(['Усього']).forEach(function (subj, i) {
    const r = i + 2;
    const crit = subj === 'Усього' ? '"<>"' : 'A' + r;
    const k = "'" + SHEET_CLASSES + "'!", c = "'" + SHEET_CARDS + "'!";
    sh.getRange(r, 1, 1, 9).setFormulas([[
      '',
      '=COUNTIF(' + c + 'D:D,' + crit + ')' + (subj === 'Усього' ? '-1' : ''),
      '=SUMIF(' + k + 'D:D,' + crit + ',' + k + 'F:F)',
      '=SUMIF(' + k + 'D:D,' + crit + ',' + k + 'G:G)',
      '=IFERROR(D' + r + '/C' + r + ',"")',
      '=SUMIF(' + k + 'D:D,' + crit + ',' + k + 'H:H)',
      '=IFERROR(F' + r + '/C' + r + ',"")',
      '=SUMIF(' + k + 'D:D,' + crit + ',' + k + 'K:K)',
      '=IFERROR(AVERAGEIF(' + c + 'D:D,' + crit + ',' + c + 'N:N),"")'
    ]]);
    sh.getRange(r, 1).setValue(subj);
  });
  sh.getRange(2, 5, subjects.length + 1, 1).setNumberFormat('0%');
  sh.getRange(2, 7, subjects.length + 1, 1).setNumberFormat('0%');
  sh.getRange(2, 9, subjects.length + 1, 1).setNumberFormat('0%');
  sh.getRange(subjects.length + 2, 1, 1, 9).setFontWeight('bold');
  sh.setColumnWidth(1, 260);
  sh.setFrozenRows(1);
}

function num_(v) { const n = parseInt(v, 10); return isNaN(n) ? '' : n; }

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Запустіть один раз вручну перед розгортанням: створить аркуші й попросить дозволи. */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  sheet_(ss, SHEET_CARDS, CARD_HEADERS);
  sheet_(ss, SHEET_CLASSES, CLASS_HEADERS);
  summarySheet_(ss);
  const me = Session.getEffectiveUser().getEmail();
  Logger.log('Готово. Сповіщення надходитимуть на: ' + me);
  if (false) MailApp.sendEmail(me, 'test', 'test');
}
