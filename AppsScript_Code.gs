/**
 * ============================================================
 *  ОПЛАТЫ ЗИКМЕС — мост между Google Таблицей и Tampermonkey-скриптом
 * ============================================================
 *
 * УСТАНОВКА:
 * 1. Откройте вашу Google Таблицу ("Оплаты Зикмес").
 * 2. Меню "Расширения" -> "Apps Script".
 * 3. Удалите весь код-заглушку (function myFunction(){...}) и вставьте сюда ВЕСЬ этот файл.
 * 4. Нажмите "Сохранить" (значок дискеты).
 * 5. Нажмите "Развернуть" (Deploy) -> "Новое развертывание" (New deployment).
 * 6. Тип развертывания: "Веб-приложение" (Web app).
 * 7. Настройки:
 *      - "Описание": любое (например "v1")
 *      - "Выполнять как": "Я (ваш email)"
 *      - "У кого есть доступ": "Все" (Anyone) — это обязательно, иначе Tampermonkey не сможет постучаться
 * 8. Нажмите "Развернуть". Google попросит дать разрешения — разрешите (это ваш собственный скрипт).
 * 9. Скопируйте появившийся "URL веб-приложения" (выглядит как https://script.google.com/macros/s/XXXXX/exec)
 * 10. Этот URL нужно будет вставить в Tampermonkey-скрипт в переменную APPS_SCRIPT_URL.
 *
 * ВАЖНО: если в будущем вы измените код этого файла, нужно сделать НОВОЕ развертывание
 * (Развернуть -> Управление развертываниями -> карандаш редактировать -> поменять версию на "Новая" -> Развернуть),
 * иначе изменения не подхватятся по старому URL.
 */

// Названия столбцов (нумерация с 1, как в таблице: A=1, B=2, C=3, D=4, E=5, F=6)
var COL_DATE = 1;     // A — Дата платежа
var COL_FIO = 2;      // B — Ф.И.О.
var COL_ORDER_ID = 3; // C — АЙДИ заказа
var COL_AMOUNT = 4;   // D — Сумма
var COL_METHOD = 5;   // E — Метод оплаты
var COL_STATUS = 6;   // F — Статус (Обработан)

var METHOD_FILTER = 'Расчётный счёт'; // обрабатываем только такой метод оплаты
var STATUS_DONE = 'Обработан';
var MIN_AMOUNT = 5; // платежи строго меньше этой суммы не переносим

/**
 * Точка входа для GET-запросов (получение списка необработанных строк)
 * Пример: GET {URL}?action=getPending
 */
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'getPending') {
      return jsonResponse(getPendingRows_());
    }
    return jsonResponse({error: 'Unknown action: ' + action});
  } catch (err) {
    return jsonResponse({error: String(err)});
  }
}

/**
 * Точка входа для POST-запросов (отметить строку(и) как обработанные)
 * Тело запроса (JSON): { "action": "markProcessed", "rows": [5, 7, 12] }
 * где rows — номера строк в таблице (1-based, как в самой Google Sheets)
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'markProcessed') {
      return jsonResponse(markProcessed_(body.rows || []));
    }
    return jsonResponse({error: 'Unknown action: ' + action});
  } catch (err) {
    return jsonResponse({error: String(err)});
  }
}

/**
 * Возвращает все необработанные строки АКТИВНОГО (текущего) листа,
 * у которых метод оплаты = "Расчётный счёт" и статус (F) пустой.
 */
function getPendingRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return {sheetName: sheetName, rows: []};
  }

  var range = sheet.getRange(1, 1, lastRow, COL_STATUS);
  var values = range.getValues();

  var result = [];
  for (var i = 0; i < values.length; i++) {
    var rowNum = i + 1; // 1-based номер строки в таблице
    var row = values[i];

    var dateVal = row[COL_DATE - 1];
    var fio = row[COL_FIO - 1];
    var orderId = row[COL_ORDER_ID - 1];
    var amount = row[COL_AMOUNT - 1];
    var method = row[COL_METHOD - 1];
    var status = row[COL_STATUS - 1];

    // Пропускаем строки без ID заказа (заголовки, пустые строки, разделители)
    if (!orderId || String(orderId).trim() === '') continue;

    // Пропускаем уже обработанные
    var statusStr = String(status).trim();
    if (statusStr === STATUS_DONE) continue;

    // Берём только нужный метод оплаты
    var methodStr = String(method).trim();
    if (methodStr !== METHOD_FILTER) continue;

    // Сумма — может быть с запятой как десятичным разделителем (169,33)
    var amountStr = String(amount).trim().replace(',', '.');
    var amountNum = parseFloat(amountStr);
    if (isNaN(amountNum)) continue;

    // Платежи меньше минимальной суммы не переносим
    if (amountNum < MIN_AMOUNT) continue;

    // ID заказа — оставляем только цифры
    var orderIdClean = String(orderId).trim().replace(/[^0-9]/g, '');
    if (orderIdClean === '') continue;

    // Дата — приводим к строке dd.mm.yyyy
    var dateStr = formatDate_(dateVal);

    result.push({
      row: rowNum,
      date: dateStr,
      fio: String(fio || ''),
      orderId: orderIdClean,
      amount: amountNum,
      method: methodStr
    });
  }

  return {sheetName: sheetName, rows: result};
}

/**
 * Приводит значение даты (может быть Date-объект или строка) к формату dd.mm.yyyy
 */
function formatDate_(val) {
  if (val instanceof Date) {
    var d = val.getDate();
    var m = val.getMonth() + 1;
    var y = val.getFullYear();
    return pad2_(d) + '.' + pad2_(m) + '.' + y;
  }
  var s = String(val || '').trim();
  return s; // уже строка вида 01.06.2026
}

function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}

/**
 * Ставит "Обработан" в столбец F для указанных номеров строк на активном листе.
 */
function markProcessed_(rowNumbers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var done = [];
  var failed = [];

  for (var i = 0; i < rowNumbers.length; i++) {
    var r = rowNumbers[i];
    try {
      sheet.getRange(r, COL_STATUS).setValue(STATUS_DONE);
      done.push(r);
    } catch (err) {
      failed.push({row: r, error: String(err)});
    }
  }

  return {done: done, failed: failed};
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
