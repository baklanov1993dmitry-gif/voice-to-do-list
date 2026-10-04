/*************************************************************
 * СИСТЕМА ДНЕВНЫХ ОТЧЁТОВ С ИИ-АНАЛИЗОМ
 * Google Sheets + Apps Script + Claude API
 *
 * Листы: "Ввод", "Журнал", "Анализ"
 * Журнал: A=Дата B=Начало C=Конец D=Часы E=Описание F=Категория G=Корр.действия
 *************************************************************/

// Модели Claude (актуальны на май 2026):
var MODEL_FAST = 'claude-haiku-4-5-20251001'; // дёшево/быстро — для категоризации
var MODEL_SMART = 'claude-sonnet-4-6';        // умнее — для недельного анализа

// Часовой пояс зашит явно, а не берётся из настроек проекта Apps Script —
// та настройка ненадёжна (может не совпадать с реальным местоположением
// и давать сдвиг на несколько часов во всех расчётах времени).
var TZ = 'Asia/Novosibirsk';

/* ============================================================
 * 1) МГНОВЕННЫЙ РАЗБОР ФРАЗЫ ПРИ ВВОДЕ (без API, локально)
 * ============================================================ */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Ввод') return;
  if (e.range.getA1Notation() !== 'A1') return;

  var text = (e.value || '').toString().trim();
  if (!text || text === 'Пиши сюда…') return;

  var p = parseEntry(text);
  var log = e.source.getSheetByName('Журнал');
  // Дата, Начало, Конец, Часы, Описание, Категория(пусто), Корр.действия(пусто)
  log.appendRow([p.date, p.start, p.end, p.hours, p.description, '', '']);
  e.range.clearContent();
}

function parseEntry(text) {
  var now = new Date();
  var tz = TZ;
  var date = Utilities.formatDate(now, tz, 'dd.MM.yyyy');
  var start = '', end = '', hours = '', description = text;

  // Интервал: "с 14:00 до 16:00", "14:00-16:00", "14-16"
  var interval = text.match(/(?:с\s*)?(\d{1,2})(?::(\d{2}))?\s*(?:до|-|–|—)\s*(\d{1,2})(?::(\d{2}))?/i);
  if (interval) {
    var h1 = parseInt(interval[1], 10), m1 = interval[2] ? parseInt(interval[2], 10) : 0;
    var h2 = parseInt(interval[3], 10), m2 = interval[4] ? parseInt(interval[4], 10) : 0;
    start = pad(h1) + ':' + pad(m1);
    end = pad(h2) + ':' + pad(m2);
    var diff = (h2 * 60 + m2) - (h1 * 60 + m1);
    if (diff < 0) diff += 24 * 60;
    hours = Math.round((diff / 60) * 100) / 100;
    description = text.replace(interval[0], '').replace(/^[\s,.:;-]+/, '').trim();
    return { date: date, start: start, end: end, hours: hours, description: description };
  }

  // Длительность: "1,5 часа", "2 ч", "30 минут", "полтора часа", "полчаса"
  var totalMinutes = 0, found = false;
  var hMatch = text.match(/(\d+(?:[.,]\d+)?)\s*(?:час(?:а|ов)?|ч\b)/i);
  if (hMatch) { totalMinutes += parseFloat(hMatch[1].replace(',', '.')) * 60; found = true; }
  var mMatch = text.match(/(\d+)\s*(?:минут[аы]?|мин\b|м\b)/i);
  if (mMatch) { totalMinutes += parseInt(mMatch[1], 10); found = true; }
  if (/полтора\s*час/i.test(text)) { totalMinutes += 90; found = true; }
  if (/полчаса/i.test(text)) { totalMinutes += 30; found = true; }

  if (found) {
    hours = Math.round((totalMinutes / 60) * 100) / 100;
    description = text;
    if (hMatch) description = description.replace(hMatch[0], '');
    if (mMatch) description = description.replace(mMatch[0], '');
    description = description.replace(/полтора\s*час(?:а)?/i, '')
                             .replace(/полчаса/i, '')
                             .replace(/^[\s,.:;-]+/, '').trim();
  }
  return { date: date, start: start, end: end, hours: hours, description: description };
}

function pad(n) { return (n < 10 ? '0' : '') + n; }


/* ============================================================
 * 2) МЕНЮ В ТАБЛИЦЕ
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Отчёты')
    .addItem('Категоризировать новые задачи', 'категоризироватьЗадачи')
    .addItem('Сделать анализ недели', 'анализНедели')
    .addItem('Сводка за сегодня', 'дневнаяСводка')
    .addSeparator()
    .addItem('Оформить таблицу (запустить один раз)', 'оформитьТаблицу')
    .addItem('Очистить весь журнал (осторожно!)', 'очиститьЖурналСПодтверждением')
    .addToUi();
}


/* ============================================================
 * 3) ХРАНЕНИЕ API-КЛЮЧА
 * ============================================================ */

// ГЛАВНЫЙ способ сохранить ключ — безопасно запускать прямо из редактора Apps Script,
// не обращается к getUi(), поэтому не падает независимо от того, откуда запущена.
// 1) Впиши свой ключ вместо текста внутри кавычек ниже.
// 2) Сохрани файл (дискета).
// 3) Сверху выбери в списке функций "сохранитьКлючНапрямую" и нажми "Выполнить".
// 4) Проверь через "проверитьКлюч" (см. ниже), что всё получилось.
// 5) Сотри свой ключ обратно на плейсхолдер и снова сохрани файл — не держи ключ открытым в коде.
function сохранитьКлючНапрямую() {
  var МОЙ_КЛЮЧ = 'sk-ant-ВСТАВЬ_СЮДА_СВОЙ_КЛЮЧ';
  if (МОЙ_КЛЮЧ.indexOf('ВСТАВЬ_СЮДА') !== -1) {
    throw new Error('Сначала впиши свой настоящий ключ вместо плейсхолдера в переменной МОЙ_КЛЮЧ и сохрани файл.');
  }
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', МОЙ_КЛЮЧ);
  console.log('Ключ сохранён в защищённое хранилище.');
}

// Проверка: показывает, сохранён ли ключ (не сам ключ, а только факт и последние 4 символа).
function проверитьКлюч() {
  var k = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!k) {
    console.log('Ключ НЕ найден. Запусти сохранитьКлючНапрямую.');
  } else {
    console.log('Ключ найден, длина ' + k.length + ' символов, оканчивается на: ...' + k.slice(-4));
  }
}

// Альтернативный способ через диалог — работает только если вызван из меню таблицы
// "📋 Отчёты" (там, где Google даёт доступ к getUi()), а не из редактора.
function сохранитьКлюч() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Введите Anthropic API ключ (sk-ant-...)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('ANTHROPIC_KEY', resp.getResponseText().trim());
    ui.alert('Ключ сохранён.');
  }
}

function getKey_() {
  var k = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_KEY');
  if (!k) throw new Error('API-ключ не найден. Запусти функцию сохранитьКлючНапрямую.');
  return k;
}


/* ============================================================
 * 4) ВЫЗОВ CLAUDE API
 * ============================================================ */
function callClaude_(model, systemPrompt, userText, maxTokens) {
  var payload = {
    model: model,
    max_tokens: maxTokens || 1024,
    system: systemPrompt,
    messages: [{ role: 'user', content: userText }]
  };
  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': getKey_(),
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (data.error) throw new Error('API: ' + data.error.message);
  return (data.content || []).map(function (b) { return b.text || ''; }).join('\n').trim();
}


/* ============================================================
 * 5) КАТЕГОРИЗАЦИЯ ЗАДАЧ (столбец F)
 * ============================================================ */
function категоризироватьЗадачи() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Журнал пуст.'); return; }

  var rng = log.getRange(2, 5, last - 1, 2).getValues(); // E (описание) + F (категория)
  var pending = [];
  for (var i = 0; i < rng.length; i++) {
    var descr = (rng[i][0] || '').toString().trim();
    var cat = (rng[i][1] || '').toString().trim();
    if (descr && !cat) pending.push({ row: i + 2, descr: descr });
  }
  if (!pending.length) { SpreadsheetApp.getUi().alert('Новых задач без категории нет.'); return; }

  var list = pending.map(function (p, idx) { return (idx + 1) + '. ' + p.descr; }).join('\n');
  var system = 'Ты помощник по тайм-трекингу. Тебе дают пронумерованный список задач. ' +
    'Придумай для них компактный набор осмысленных категорий (на русском, 1-3 слова) ' +
    'и присвой каждой задаче одну категорию. Старайся переиспользовать категории, ' +
    'не плоди лишние. Ответь ТОЛЬКО валидным JSON-массивом вида ' +
    '[{"n":1,"category":"..."}], без markdown и пояснений.';

  var raw = callClaude_(MODEL_FAST, system, list, 1500);
  raw = raw.replace(/```json|```/g, '').trim();

  var arr;
  try { arr = JSON.parse(raw); }
  catch (err) { SpreadsheetApp.getUi().alert('Не удалось разобрать ответ ИИ:\n' + raw); return; }

  arr.forEach(function (item) {
    var p = pending[item.n - 1];
    if (p && item.category) log.getRange(p.row, 6).setValue(item.category);
  });
  SpreadsheetApp.getUi().alert('Категории проставлены для ' + arr.length + ' задач.');
}


/* ============================================================
 * 6) АНАЛИЗ НЕДЕЛИ (лист "Анализ")
 * ============================================================ */
function анализНедели() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Журнал пуст.'); return; }

  // Берём записи за последние 7 дней
  var values = log.getRange(2, 1, last - 1, 7).getValues();
  var weekAgo = new Date(); weekAgo.setDate(weekAgo.getDate() - 7);

  var rows = [];
  var tz1 = TZ;
  values.forEach(function (r) {
    var d = parseRuDate_(r[0]);
    if (d && d >= weekAgo) {
      var dateStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz1, 'dd.MM.yyyy') : r[0];
      rows.push('Дата ' + dateStr + ' | Часы ' + r[3] + ' | ' + r[4] +
                (r[5] ? ' | Категория: ' + r[5] : '') +
                (r[6] ? ' | Корр.действие: ' + r[6] : ''));
    }
  });
  if (!rows.length) { SpreadsheetApp.getUi().alert('За последнюю неделю записей нет.'); return; }

  var system = 'Ты аналитик личной продуктивности. Тебе дают журнал задач за неделю. ' +
    'Сгруппируй задачи по типам, посчитай суммарные часы по каждому типу, ' +
    'выдели наблюдения (на что ушло больше всего времени, перекосы, что повторяется) ' +
    'и дай 3-5 коротких практичных рекомендаций. Пиши по-русски, структурно, кратко.';

  var analysis = callClaude_(MODEL_SMART, system, rows.join('\n'), 2000);

  var sheet = ss.getSheetByName('Анализ');
  var stamp = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setNumberFormat('@').setValue('Анализ недели — ' + stamp + '\n\n' + analysis);
  SpreadsheetApp.getUi().alert('Анализ готов — смотри лист «Анализ».');
}

// Короткая сводка за сегодняшний день (в отличие от недельного анализа — только текущая дата).
function дневнаяСводка() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) { SpreadsheetApp.getUi().alert('Журнал пуст.'); return; }

  var tz = TZ;
  var today = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
  var values = log.getRange(2, 1, last - 1, 7).getValues();

  var rows = [];
  values.forEach(function (r) {
    var dateStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'dd.MM.yyyy') : r[0];
    if (dateStr === today) {
      rows.push('Часы ' + r[3] + ' | ' + r[4] + (r[5] ? ' | Категория: ' + r[5] : ''));
    }
  });
  if (!rows.length) { SpreadsheetApp.getUi().alert('За сегодня записей нет.'); return; }

  var system = 'Ты аналитик личной продуктивности. Тебе дают список задач за ОДИН день. ' +
    'Кратко (3-5 предложений) сгруппируй по смыслу, посчитай суммарные часы, отметь, ' +
    'на что ушло больше всего времени. Пиши по-русски, без длинных списков.';

  var analysis = callClaude_(MODEL_FAST, system, rows.join('\n'), 800);

  var sheet = ss.getSheetByName('Анализ');
  var stamp = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setNumberFormat('@').setValue('Дневная сводка — ' + stamp + '\n\n' + analysis);
  SpreadsheetApp.getUi().alert('Сводка готова — смотри лист «Анализ».');
}

// Для авто-триггера по воскресеньям (без диалогов alert)
function еженедельныйАнализ() {
  try { категоризироватьЗадачиТихо_(); } catch (e) {}
  try { анализНеделиТихо_(); } catch (e) {}
}

function parseRuDate_(s) {
  if (s instanceof Date) return s;
  var m = (s || '').toString().match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (!m) return null;
  return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}

/* Тихие версии для авто-триггера (без UI alert) */
function категоризироватьЗадачиТихо_() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) return;
  var rng = log.getRange(2, 5, last - 1, 2).getValues();
  var pending = [];
  for (var i = 0; i < rng.length; i++) {
    var descr = (rng[i][0] || '').toString().trim();
    var cat = (rng[i][1] || '').toString().trim();
    if (descr && !cat) pending.push({ row: i + 2, descr: descr });
  }
  if (!pending.length) return;
  var listText = pending.map(function (p, idx) { return (idx + 1) + '. ' + p.descr; }).join('\n');
  var system = 'Ты помощник по тайм-трекингу. Дают пронумерованный список задач. ' +
    'Придумай компактный набор категорий (русский, 1-3 слова) и присвой каждой одну. ' +
    'Переиспользуй категории. Ответь ТОЛЬКО JSON-массивом [{"n":1,"category":"..."}].';
  var raw = callClaude_(MODEL_FAST, system, listText, 1500).replace(/```json|```/g, '').trim();
  var arr = JSON.parse(raw);
  arr.forEach(function (item) {
    var p = pending[item.n - 1];
    if (p && item.category) log.getRange(p.row, 6).setValue(item.category);
  });
}

function анализНеделиТихо_() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) return;
  var values = log.getRange(2, 1, last - 1, 7).getValues();
  var weekAgo = new Date(); weekAgo.setDate(weekAgo.getDate() - 7);
  var rows = [];
  var tz2 = TZ;
  values.forEach(function (r) {
    var d = parseRuDate_(r[0]);
    if (d && d >= weekAgo) {
      var dateStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz2, 'dd.MM.yyyy') : r[0];
      rows.push('Дата ' + dateStr + ' | Часы ' + r[3] + ' | ' + r[4] +
                (r[5] ? ' | Категория: ' + r[5] : '') +
                (r[6] ? ' | Корр.действие: ' + r[6] : ''));
    }
  });
  if (!rows.length) return;
  var system = 'Ты аналитик личной продуктивности. Дают журнал задач за неделю. ' +
    'Сгруппируй по типам, посчитай часы по каждому типу, выдели наблюдения ' +
    'и дай 3-5 практичных рекомендаций. По-русски, кратко.';
  var analysis = callClaude_(MODEL_SMART, system, rows.join('\n'), 2000);
  var sheet = ss.getSheetByName('Анализ');
  var stamp = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setNumberFormat('@').setValue('Анализ недели — ' + stamp + '\n\n' + analysis);
}


/* ============================================================
 * 7) ВЕБ-СТРАНИЦА ВВОДА (открыл — сразу диктуешь, без выбора ячейки)
 * ============================================================ */
// doGet обслуживает два разных случая:
// 1) обычное открытие ссылки — отдаёт старую HTML-страницу (запасной вход);
// 2) ?action=stats — отдаёт JSON со статистикой для внешней PWA-страницы на GitHub Pages.
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === 'stats') {
    return ContentService.createTextOutput(JSON.stringify(получитьСтатистику_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'list') {
    return ContentService.createTextOutput(JSON.stringify(получитьВсеЗаписи_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === 'timecheck') {
    // Проверяет ИМЕННО развёрнутый веб-адрес (тот, по которому стучится
    // приложение) — в отличие от кнопки "Выполнить" в редакторе, которая
    // всегда берёт код напрямую и может не совпадать с тем, что реально
    // отдаёт сайту развёрнутая версия.
    return ContentService.createTextOutput(JSON.stringify({
      tz: TZ,
      nowInTz: Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm:ss'),
      rawServerTime: new Date().toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('Страница')
    .setTitle('Запись задачи')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// doPost — вход для внешней PWA-страницы: принимает JSON {text: "..."} в теле запроса,
// сохраняет запись, возвращает JSON с результатом. Тело — обязательно text/plain на
// стороне отправителя (не application/json), иначе браузер шлёт предварительный
// CORS-запрос (preflight), который Apps Script не поддерживает.
// doPost — вход для внешней PWA-страницы. Тело запроса — JSON с полем action:
// "save" (по умолчанию, для обратной совместимости) — {text}
// "delete" — {ids: [...]} — удаляет одну или несколько записей по ID
// "update" — {id, description, start, end} — правит существующую запись
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'save';

    if (action === 'save') {
      var result = записатьЗадачу_(body.text || '');
      return ContentService.createTextOutput(JSON.stringify({ ok: true, message: result.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'delete') {
      var delResult = удалитьЗаписи_(body.ids || []);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, deleted: delResult.deleted }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'update') {
      var updResult = обновитьЗапись_(body);
      return ContentService.createTextOutput(JSON.stringify(updResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'неизвестное действие' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: (err && err.message) ? err.message : String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Вызывается со старой HTML-страницы через google.script.run — обёртка над общей логикой.
function добавитьЗапись(text) {
  return записатьЗадачу_(text).message;
}

// Общая логика: разбирает фразу (время + суть) одним запросом к ИИ, пишет в Журнал.
// Если ИИ недоступен — резервный путь на простых правилах, чтобы запись не терялась.
// Каждой записи присваивается уникальный ID (столбец I) — по нему потом ищем
// запись для удаления или редактирования, независимо от того, как сдвинутся
// номера строк после удаления других записей.
function записатьЗадачу_(text) {
  text = (text || '').toString().trim();
  if (!text) return { message: 'Пусто' };

  var tz = TZ;
  var today = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
  var nowStr = Utilities.formatDate(new Date(), tz, 'HH:mm');
  var prevEnd = предыдущийКонец_(today);

  var p;
  var debugInfo = '';
  try {
    p = разобратьЗаписьЧерезAI_(text, today, nowStr, prevEnd);
    debugInfo = 'ИИ OK: ' + (p.debugRaw || '');

    // Страховка от редких галлюцинаций времени на коротких/безликих фразах
    // ("Толчок", "Проверка" — без единого намёка на время). Если в САМОМ
    // тексте нет явного интервала ("с 14 до 16") — не доверяем догадке ИИ
    // насчёт часов на циферблате, а сами детерминированно ставим конец =
    // текущий момент. Длительность (если она была явно сказана словами,
    // например "последний час") по-прежнему берём у ИИ — с этим он
    // справляется надёжно, здесь ошибок не наблюдалось.
    var hasExplicitInterval = /\d{1,2}(:\d{2})?\s*(?:до|-|–|—)\s*\d{1,2}(:\d{2})?/i.test(text);
    var overrideApplied = !hasExplicitInterval;
    if (!hasExplicitInterval) {
      p.end = nowStr;
      var aiHours = (typeof p.hours === 'number' && p.hours > 0 && p.hours <= 12) ? p.hours : null;
      if (aiHours) {
        p.start = вычестьЧасы_(nowStr, aiHours);
        p.hours = aiHours;
      } else if (prevEnd) {
        p.start = prevEnd;
        p.hours = вычислитьЧасы_(prevEnd, nowStr);
      } else {
        p.start = '';
        p.hours = '';
      }
    }
  } catch (e) {
    // Резервный путь на regex-правилах — на случай сбоя API, запись не теряем.
    debugInfo = 'ИИ ОШИБКА: ' + (e && e.message ? e.message : e);
    p = parseEntry(text);
    p.date = p.date || today;
    if (p.hours && !p.start && !p.end) {
      p.end = nowStr;
      p.start = вычестьЧасы_(nowStr, p.hours);
    }
    if (!p.start && !p.hours) {
      p.end = nowStr;
      if (prevEnd) { p.start = prevEnd; p.hours = вычислитьЧасы_(prevEnd, nowStr); }
    }
  }

  var id = Utilities.getUuid();
  var log = SpreadsheetApp.getActive().getSheetByName('Журнал');
  var newRow = log.getLastRow() + 1;
  // ВАЖНО: формат ставим ПЕРЕД записью значений, а не после. Если писать через
  // обычный appendRow(), Google Таблица иногда успевает "на лету" распознать
  // "14:07" как настоящее время и превратить в служебное значение, привязанное
  // к условной дате 1899 года — а при обратном чтении с учётом часового пояса
  // это давало нестандартный сдвиг (историческое, не круглое смещение для тех
  // широт). Принудительный текстовый формат ДО записи не даёт Таблице вообще
  // начать это распознавание.
  log.getRange(newRow, 1, 1, 3).setNumberFormat('@'); // A=Дата, B=Начало, C=Конец — только эти три
  log.getRange(newRow, 1, 1, 9).setValues([[p.date, p.start, p.end, p.hours, p.description, '', '', debugInfo, id]]);
  var nowStrFinal = Utilities.formatDate(new Date(), TZ, 'HH:mm:ss'); // ещё раз, прямо перед записью — для сверки с nowStr выше
  return {
    message: 'Записано: ' + (p.hours ? p.hours + ' ч — ' : '') + p.description,
    hours: p.hours, description: p.description, id: id
  };
}

// Удаляет одну или несколько записей по их ID (столбец I). Удаляет снизу вверх,
// чтобы номера строк остальных удаляемых записей не сдвигались по ходу дела.
function удалитьЗаписи_(ids) {
  if (!ids || !ids.length) return { deleted: 0 };
  var log = SpreadsheetApp.getActive().getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) return { deleted: 0 };

  var idCol = log.getRange(2, 9, last - 1, 1).getValues();
  var rowsToDelete = [];
  for (var i = 0; i < idCol.length; i++) {
    var val = (idCol[i][0] || '').toString();
    if (ids.indexOf(val) !== -1) rowsToDelete.push(i + 2);
  }
  rowsToDelete.sort(function (a, b) { return b - a; }); // снизу вверх
  rowsToDelete.forEach(function (r) { log.deleteRow(r); });
  return { deleted: rowsToDelete.length };
}

// Правит существующую запись по ID: описание и/или время. Часы пересчитываются
// автоматически из начала/конца, если оба заданы.
function обновитьЗапись_(payload) {
  var id = payload.id;
  if (!id) return { ok: false, error: 'нет id' };

  var log = SpreadsheetApp.getActive().getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) return { ok: false, error: 'журнал пуст' };

  var idCol = log.getRange(2, 9, last - 1, 1).getValues();
  var rowNum = -1;
  for (var i = 0; i < idCol.length; i++) {
    if ((idCol[i][0] || '').toString() === id) { rowNum = i + 2; break; }
  }
  if (rowNum === -1) return { ok: false, error: 'запись не найдена' };

  var start = (payload.start || '').toString().trim();
  var end = (payload.end || '').toString().trim();
  var description = (payload.description || '').toString().trim();

  if (start) log.getRange(rowNum, 2).setValue(start);
  if (end) log.getRange(rowNum, 3).setValue(end);
  if (start && end) log.getRange(rowNum, 4).setValue(вычислитьЧасы_(start, end));
  if (description) log.getRange(rowNum, 5).setValue(description);

  return { ok: true };
}

// Статистика для главного экрана PWA: часы и число записей за сегодня + 3 последние
// записи (не только сегодняшние — если человек не писал со вчера, это будет видно).
function получитьСтатистику_() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  var tz = TZ;
  var today = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
  if (last < 2) return { todayHours: 0, todayCount: 0, last3: [] };

  var values = log.getRange(2, 1, last - 1, 9).getValues();
  var todayHours = 0, todayCount = 0;
  var rows = [];
  values.forEach(function (r) {
    var dateStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'dd.MM.yyyy') : r[0];
    if (dateStr === today) {
      todayCount++;
      if (typeof r[3] === 'number') todayHours += r[3];
    }
    if (r[4]) { // есть описание — валидная запись
      rows.push(строкаВЗапись_(r, tz));
    }
  });

  var last3 = rows.slice(-5).reverse(); // последние 5 (было 3) — меньше пустого места на экране, самая новая первая

  return {
    todayHours: Math.round(todayHours * 100) / 100,
    todayCount: todayCount,
    last3: last3
  };
}

// Полный список записей для экрана "Все записи" — самые новые сверху.
// Ограничение 500 строк — с запасом на годы вперёд при разумной нагрузке;
// если понадобится больше, можно легко поднять лимит здесь.
function получитьВсеЗаписи_() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName('Журнал');
  var last = log.getLastRow();
  var tz = TZ;
  if (last < 2) return { entries: [] };

  var values = log.getRange(2, 1, last - 1, 9).getValues();
  var rows = [];
  values.forEach(function (r) {
    if (r[4]) rows.push(строкаВЗапись_(r, tz));
  });

  rows.reverse(); // новые сверху
  if (rows.length > 500) rows = rows.slice(0, 500);

  return { entries: rows };
}

// Общий разбор строки листа "Журнал" в аккуратный объект для фронтенда.
// r[8] — ID записи (может отсутствовать у самых старых строк, созданных до
// введения ID, тогда фронтенд просто не даст их редактировать/удалять).
function строкаВЗапись_(r, tz) {
  var dateStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'dd.MM.yyyy') : r[0];
  var startStr = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'HH:mm') : (r[1] || '');
  var endStr = (r[2] instanceof Date) ? Utilities.formatDate(r[2], tz, 'HH:mm') : (r[2] || '');
  return {
    id: r[8] || '',
    date: dateStr,
    start: startStr,
    end: endStr,
    hours: (typeof r[3] === 'number') ? r[3] : '',
    description: r[4] || '',
    category: r[5] || ''
  };
}

// Один запрос к ИИ: чистит текст от мусора/ошибок распознавания И одновременно
// определяет время задачи — интервал, длительность (в т.ч. словами: "последний час",
// "полдня") или, если время не сказано, продолжает от конца предыдущей записи.
// Модель — MODEL_SMART: эта функция разбирает КАЖДУЮ запись, точность важнее экономии.
function разобратьЗаписьЧерезAI_(text, dateStr, nowStr, prevEnd) {
  var system =
    'Ты помощник по учёту рабочего времени. Пользователь надиктовал голосом (иногда с ошибками ' +
    'распознавания речи и словами-паразитами: "ну", "эээ", "короче", "в общем", "как бы", "значит") ' +
    'описание задачи, которой он занимался. Тебе даны: сегодняшняя дата, текущее время (момент записи) ' +
    'и время окончания предыдущей задачи за сегодня (если было).\n\n' +
    'Разбери текст и верни JSON с полями:\n' +
    '- description: КОРОТКОЕ название задачи (2-5 слов) в виде отглагольного существительного — ' +
    'например "Отладка голосовых отчётов". Это НЕ пересказ фразы, а заголовок по сути дела.\n' +
    '  ЗАПРЕЩЕНО оставлять в description: любые обрывки о времени и длительности (слова "час", "часа", ' +
    '"минут", "последний", "последние", "следующие", "с... до...") и глаголы самого факта работы ' +
    '("занимался", "делал", "потратил", "провёл", "работал над" — оставляй только ЧЕМ занимался, не сам факт занятости), ' +
    'вводные слова и обращения ("привет", "как дела", "короче говоря", "в общем", "ну вот").\n' +
    '  ВАЖНО: если человек упомянул НЕСКОЛЬКО разных дел одной фразой — сохрани ВСЕ их в description, ' +
    'через "и" или запятую. Никогда не отбрасывай тему только потому, что она вторая или третья по счёту — ' +
    'убирать можно только настоящий мусор (паразиты, повторы), но не содержание.\n' +
    '  ВАЖНО: НЕ обобщай сказанное в абстрактную категорию или ярлык. Пересказывай конкретно, ЧТО именно ' +
    'сделал человек, своими словами, но не теряя сути. НИКОГДА не пиши общие ярлыки вроде "личные дела", ' +
    '"бытовые вопросы", "разное", "рабочие задачи" — это стирает смысл, и потом человек не поймёт, что тут ' +
    'было зашифровано. Пиши конкретное действие/предмет, а не категорию, к которой оно относится.\n' +
    '  Пример 1: вход "последний час занимался отладкой голосовых отчётов" → description "Отладка голосовых отчётов".\n' +
    '  Пример 2: вход "последние полтора часа занимался отладкой отчётов" → description "Отладка отчётов".\n' +
    '  Пример 3: вход "ну короче два часа делал отчёт по клиенту X в общем" → description "Отчёт по клиенту X".\n' +
    '  Пример 4 (несколько дел — сохранить оба): вход "решал проблему отчётов и заодно развешивал бельё" ' +
    '→ description "Решение проблемы отчётов и развешивание белья" (НЕ "Решение проблемы отчётов").\n' +
    '  Пример 5 (не обобщать в абстрактную категорию): вход "помылся и покушал" → description "Приём пищи и душ" ' +
    '(НЕ "Личные дела и гигиена" — это слишком абстрактно, по такой записи потом не восстановишь, что реально было).\n' +
    '- start: время начала "ЧЧ:ММ" или ""\n' +
    '- end: время окончания "ЧЧ:ММ" или ""\n' +
    '- hours: число часов (дробное, округли до сотых) или ""\n\n' +
    'Правила определения времени (по приоритету):\n' +
    '1. Явный интервал ("с 14 до 16") — используй как start/end, вычисли hours.\n' +
    '2. Длительность, цифрами или словами ("2 часа", "полчаса", "полтора часа", "последний час", ' +
    '"весь день" = 8 часов, "полдня" = 4 часа) — end = данное текущее время, ' +
    'start = end минус длительность, hours = длительность.\n' +
    '3. Время не упомянуто вовсе — end = данное текущее время. Если дано время конца предыдущей ' +
    'задачи — start = это время, hours = разница start и end. Если предыдущего времени нет — start = "", hours = "".\n\n' +
    'Прежде чем ответить, проверь себя: нет ли в description хоть одного слова из запрещённого списка выше. ' +
    'Ответь ТОЛЬКО валидным JSON, без пояснений и markdown-обёртки.';

  var userMsg =
    'Сегодняшняя дата: ' + dateStr + '\n' +
    'Текущее время (момент записи): ' + nowStr + '\n' +
    'Время окончания предыдущей задачи сегодня: ' + (prevEnd || 'нет данных') + '\n' +
    'Текст пользователя: ' + text;

  var raw = callClaude_(MODEL_FAST, system, userMsg, 300);
  var rawOriginal = raw; // сохраняем для диагностики до всякой обработки
  raw = raw.replace(/```json|```/g, '').trim();
  var obj = JSON.parse(raw);

  var hoursVal = (obj.hours === '' || obj.hours === null || obj.hours === undefined)
    ? '' : Number(obj.hours);
  if (typeof hoursVal === 'number' && isNaN(hoursVal)) hoursVal = ''; // никогда не пишем NaN в таблицу

  return {
    date: dateStr,
    start: (obj.start || '').toString().trim(),
    end: (obj.end || '').toString().trim(),
    hours: hoursVal,
    description: (obj.description || text).toString().trim(),
    debugRaw: rawOriginal
  };
}

// Ищет время конца последней записи за указанную дату (столбец C). null, если нет.
// Важно: если ячейка когда-либо распознавалась Таблицей как время, getValues()
// вернёт настоящий объект Date, а не текст — его нужно форматировать, а не toString().
function предыдущийКонец_(dateStr) {
  var log = SpreadsheetApp.getActive().getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last < 2) return null;
  var tz = TZ;
  var row = log.getRange(last, 1, 1, 3).getValues()[0]; // A=Дата, B=Начало, C=Конец
  var rowDate = row[0];
  var rowDateStr = (rowDate instanceof Date)
    ? Utilities.formatDate(rowDate, tz, 'dd.MM.yyyy')
    : (rowDate || '').toString();
  if (rowDateStr !== dateStr) return null; // последняя запись не за сегодня

  var rawEnd = row[2];
  var rowEnd = (rawEnd instanceof Date)
    ? Utilities.formatDate(rawEnd, tz, 'HH:mm')
    : (rawEnd || '').toString().trim();
  if (!rowEnd) return null;

  // Защита от заражённого времени: если "конец" предыдущей записи оказывается
  // ПОЗЖЕ текущего момента — это бессмыслица (время не идёт вспять), значит в
  // той записи мусорное время. Не продолжаем от него цепочку, иначе порча
  // расползётся на все следующие записи. Это и была причина многодневного бага.
  var nowStr = Utilities.formatDate(new Date(), tz, 'HH:mm');
  if (вМинуты_(rowEnd) > вМинуты_(nowStr)) return null;

  return rowEnd;
}

// Переводит "ЧЧ:ММ" в число минут от начала суток (для сравнения времён).
function вМинуты_(hhmm) {
  var p = (hhmm || '').split(':');
  if (p.length < 2) return -1;
  return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
}

// Отнимает hours (число, может быть дробным) от времени ЧЧ:ММ. Возвращает ЧЧ:ММ.
function вычестьЧасы_(timeStr, hours) {
  var t = timeStr.split(':');
  var totalMin = parseInt(t[0], 10) * 60 + parseInt(t[1], 10);
  totalMin -= Math.round(hours * 60);
  totalMin = ((totalMin % (24 * 60)) + 24 * 60) % (24 * 60); // на случай ухода в минус (через полночь)
  var h = Math.floor(totalMin / 60), m = totalMin % 60;
  return pad(h) + ':' + pad(m);
}

// Разница между двумя временами ЧЧ:ММ в часах (с округлением до сотых).
function вычислитьЧасы_(start, end) {
  var s = start.split(':'), e = end.split(':');
  var sMin = parseInt(s[0], 10) * 60 + parseInt(s[1], 10);
  var eMin = parseInt(e[0], 10) * 60 + parseInt(e[1], 10);
  var diff = eMin - sMin;
  if (diff < 0) diff += 24 * 60; // через полночь
  return Math.round((diff / 60) * 100) / 100;
}

/* ============================================================
 * 8) ОФОРМЛЕНИЕ ТАБЛИЦЫ — запускается один раз (или повторно после
 * добавления новых строк), делает лист читаемым: шапка, ширины
 * столбцов, чередующаяся заливка, границы. Заодно принудительно
 * делает столбцы даты/времени текстовыми, чтобы Таблица больше
 * никогда не подменяла их служебными объектами Date.
 * ============================================================ */
function оформитьТаблицу() {
  var ss = SpreadsheetApp.getActive();
  оформитьЛистЖурнал_(ss.getSheetByName('Журнал'));
  оформитьЛистАнализ_(ss.getSheetByName('Анализ'));
  SpreadsheetApp.getUi().alert('Оформление применено.');
}

function оформитьЛистЖурнал_(sheet) {
  var lastCol = 9; // A..I (H — диагностика ИИ, I — ID записи, скрыт)
  var lastRow = Math.max(sheet.getLastRow(), 2);

  // Подписи для столбцов H и I, если их ещё нет (не перезаписываем, если уже подписаны).
  if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('Диагностика ИИ');
  if (!sheet.getRange(1, 9).getValue()) sheet.getRange(1, 9).setValue('ID');

  // Старым записям (созданным до появления системы ID) присваиваем ID сейчас —
  // без этого несколько записей с одинаковым "пустым" ID могли задеть друг
  // друга при удалении/редактировании по ID.
  назначитьIdСтарымЗаписям_(sheet);

  // Шапка
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Ширины столбцов
  sheet.setColumnWidth(1, 100);  // Дата
  sheet.setColumnWidth(2, 70);   // Начало
  sheet.setColumnWidth(3, 70);   // Конец
  sheet.setColumnWidth(4, 60);   // Часы
  sheet.setColumnWidth(5, 320);  // Описание
  sheet.setColumnWidth(6, 140);  // Категория
  sheet.setColumnWidth(7, 220);  // Корр. действия
  sheet.setColumnWidth(8, 260);  // Диагностика (H)
  sheet.hideColumns(9);          // ID (I) — служебный, скрываем от глаз

  // Принудительно текстовый формат для дат/времени — исключает баг с автопревращением
  // строк "23:14" / "06.08.2026" в служебные объекты Date при обратном чтении.
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1000), 4).setNumberFormat('@');

  // Чередующаяся заливка строк для читаемости
  try {
    var banding = sheet.getRange(1, 1, lastRow, lastCol).getBandings();
    banding.forEach(function (b) { b.remove(); });
    sheet.getRange(1, 1, lastRow, lastCol)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  } catch (e) { /* бандинг необязателен, не прерываем оформление */ }

  // Тонкие границы
  sheet.getRange(1, 1, lastRow, lastCol)
    .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

  // Перенос текста в описании и корр. действиях, чтобы длинные фразы не обрезались
  sheet.getRange(2, 5, Math.max(lastRow - 1, 1), 1).setWrap(true);
  sheet.getRange(2, 7, Math.max(lastRow - 1, 1), 1).setWrap(true);
}

function оформитьЛистАнализ_(sheet) {
  sheet.setColumnWidth(1, 900);
  var lastRow = Math.max(sheet.getLastRow(), 50);
  sheet.getRange(1, 1, lastRow, 1)
    .setWrap(true)
    .setVerticalAlignment('top')
    .setFontSize(11);
}

/* ============================================================
 * 9) РЕЗЕРВНАЯ КОПИЯ НА ПОЧТУ — независимая страховка на случай
 * потери доступа к этому Google-аккаунту. Отправляет таблицу как
 * Excel-файл на указанную почту. Запусти вручную для проверки,
 * потом поставь на еженедельный триггер (см. инструкцию).
 * ============================================================ */
function отправитьРезервнуюКопию() {
  var ПОЧТА_ДЛЯ_БЭКАПОВ = 'ВПИШИ_СЮДА_СВОЮ_ПОЧТУ@example.com'; // лучше НЕ этот же Google-аккаунт

  var ss = SpreadsheetApp.getActive();
  var file = DriveApp.getFileById(ss.getId());
  var stamp = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy');
  var blob = file.getAs(MimeType.MICROSOFT_EXCEL).setName('Дневник_задач_' + stamp + '.xlsx');

  MailApp.sendEmail({
    to: ПОЧТА_ДЛЯ_БЭКАПОВ,
    subject: 'Резервная копия дневника задач — ' + stamp,
    body: 'Автоматическая резервная копия таблицы. Файл в приложении, открывается в Excel.',
    attachments: [blob]
  });
}

/* ============================================================
 * 10) АВТОПРИВЯЗКА ID К СТАРЫМ ЗАПИСЯМ
 * Записи, созданные до появления системы ID (столбец I), имеют там
 * пустую ячейку. Если таких записей несколько, они все "неотличимы"
 * друг от друга при поиске по ID — удаление/редактирование одной
 * могло случайно задеть все такие записи разом. Эта функция один раз
 * проходит по столбцу I и заполняет только пустые ячейки, ничего не
 * трогая у записей, где ID уже есть. Безопасно запускать повторно.
 * ============================================================ */
function назначитьIdСтарымЗаписям_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var idRange = sheet.getRange(2, 9, last - 1, 1);
  var ids = idRange.getValues();
  var changed = false;
  for (var i = 0; i < ids.length; i++) {
    if (!ids[i][0]) { ids[i][0] = Utilities.getUuid(); changed = true; }
  }
  if (changed) idRange.setValues(ids);
}

/* ============================================================
 * 11) ДИАГНОСТИКА ВРЕМЕНИ — запусти вручную из редактора (кнопка
 * "Выполнить"), результат смотри в панели журнала выполнения снизу.
 * Показывает, что именно ТЕКУЩИЙ сохранённый код считает "сейчас" —
 * это проверка независимо от того, была ли переразвёрнута веб-версия.
 * ============================================================ */
function проверитьВремя() {
  Logger.log('TZ в коде = ' + TZ);
  Logger.log('Сервер считает "сейчас" в этом часовом поясе: ' + Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm:ss'));
  Logger.log('Сырое время сервера (без учёта TZ): ' + new Date().toString());
}

/* ============================================================
 * 12) НАДЁЖНАЯ ПОЛНАЯ ОЧИСТКА ЖУРНАЛА — удаляет структурно САМИ строки
 * (включая скрытый столбец ID), а не просто содержимое видимых столбцов.
 * Это исключает ситуацию, когда после ручного удаления в скрытом столбце
 * I остаются старые значения, и getLastRow() продолжает "видеть" мнимые
 * старые строки, из-за чего новая запись попадает не во 2-ю строку,
 * а куда-то дальше.
 * ============================================================ */
function очиститьЖурналСПодтверждением() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Точно очистить весь журнал?',
    'Это удалит ВСЕ строки с данными на листе "Журнал" безвозвратно (заголовок останется). Отменить нельзя.',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = SpreadsheetApp.getActive().getSheetByName('Журнал');
  var last = log.getLastRow();
  if (last > 1) log.deleteRows(2, last - 1);
  ui.alert('Журнал очищен. Осталась только строка заголовков.');
}
