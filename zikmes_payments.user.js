// ==UserScript==
// @name         Зикмес — Автопривязка платежей из Google Таблицы
// @namespace    zikmes-payments
// @version      1.0
// @description  Кнопка на странице ok-crm.com: читает Google Таблицу, находит платежи "Ожидает оплаты" или "Просрочен" по ID заказа и заполняет их (сумма, метод, дата, статус "Оплачен"), затем отмечает строку в таблице как "Обработан". Платежи меньше 5 (минимальной суммы) не переносятся.
// @match        https://a.ok-crm.com/*
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ==========================================================
  // НАСТРОЙКИ — ОБЯЗАТЕЛЬНО ЗАМЕНИТЕ URL НИЖЕ НА СВОЙ!
  // ==========================================================
  // Вставьте сюда URL вашего развернутого Google Apps Script Web App
  // (получаете его после Развернуть -> Новое развертывание -> Веб-приложение)
  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/ВАШ_ID_РАЗВЕРТЫВАНИЯ/exec';

  const DELAY_MS = 500; // задержка между обработкой заказов
  const PAYMENT_METHOD_VALUE = '7'; // 7 = Расчетный счет (см. <select id="payment-payment_method">)
  const PAYMENT_STATUS_PAID = '3';  // 3 = Оплачен
  const PAYMENT_STATUS_PENDING = '1'; // 1 = Ожидает оплаты
  const PAYMENT_STATUS_OVERDUE = '6'; // 6 = Просрочен
  const ELIGIBLE_STATUSES = [PAYMENT_STATUS_PENDING, PAYMENT_STATUS_OVERDUE]; // какие статусы платежей считаем подходящими для привязки

  // ==========================================================
  // UI: плавающая кнопка + панель статуса
  // ==========================================================
  function injectUI() {
    const btn = document.createElement('button');
    btn.textContent = '💳 Привязать платежи из таблицы';
    btn.id = 'zikmes-payments-btn';
    btn.style.cssText = `
      position: fixed; bottom: 20px; right: 20px; z-index: 999999;
      background: #28a745; color: white; border: none; border-radius: 6px;
      padding: 12px 18px; font-size: 14px; font-weight: bold; cursor: pointer;
      box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    `;
    document.body.appendChild(btn);

    const panel = document.createElement('div');
    panel.id = 'zikmes-payments-panel';
    panel.style.cssText = `
      position: fixed; bottom: 70px; right: 20px; z-index: 999999;
      background: white; border: 1px solid #ccc; border-radius: 6px;
      padding: 12px; width: 420px; max-height: 60vh; overflow-y: auto;
      box-shadow: 0 2px 12px rgba(0,0,0,0.3); font-size: 13px; font-family: sans-serif;
      display: none; color: #222;
    `;
    document.body.appendChild(panel);

    btn.addEventListener('click', () => {
      panel.style.display = 'block';
      panel.innerHTML = '';
      runBatch(panel, btn);
    });
  }

  function log(panel, text, color) {
    const line = document.createElement('div');
    line.textContent = text;
    if (color) line.style.color = color;
    panel.appendChild(line);
    panel.scrollTop = panel.scrollHeight;
  }

  // ==========================================================
  // Главный процесс
  // ==========================================================
  async function runBatch(panel, btn) {
    btn.disabled = true;
    btn.textContent = '⏳ Обработка...';

    log(panel, 'Запрашиваю список платежей из таблицы...');
    let pending;
    try {
      const resp = await fetch(`${APPS_SCRIPT_URL}?action=getPending`);
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      pending = data.rows;
      log(panel, `Лист "${data.sheetName}": найдено ${pending.length} необработанных строк.`);
    } catch (err) {
      log(panel, 'ОШИБКА при получении данных из таблицы: ' + err.message, 'red');
      btn.disabled = false;
      btn.textContent = '💳 Привязать платежи из таблицы';
      return;
    }

    if (pending.length === 0) {
      log(panel, 'Нечего обрабатывать — все строки уже обработаны или не подходят по фильтру.', 'gray');
      btn.disabled = false;
      btn.textContent = '💳 Привязать платежи из таблицы';
      return;
    }

    let okCount = 0, errCount = 0, notFoundCount = 0;
    const successRows = [];

    for (const item of pending) {
      try {
        log(panel, `Заказ ${item.orderId} (${item.fio}, ${item.amount} руб.)...`);
        const result = await processOrder(item);
        if (result === 'ok') {
          log(panel, `  ✅ Платёж заполнен и сохранён.`, 'green');
          okCount++;
          successRows.push(item.row);
        } else if (result === 'not_found') {
          log(panel, `  ⚠️ Заказ найден, но нет платежа со статусом "Ожидает оплаты"/"Просрочен".`, 'orange');
          notFoundCount++;
        }
      } catch (err) {
        log(panel, `  ❌ Ошибка: ${err.message}`, 'red');
        errCount++;
      }
      await sleep(DELAY_MS);
    }

    // Отмечаем успешные строки как обработанные
    if (successRows.length > 0) {
      try {
        await fetch(APPS_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // избегаем preflight CORS
          body: JSON.stringify({ action: 'markProcessed', rows: successRows })
        });
        log(panel, `Таблица обновлена: ${successRows.length} строк помечены "Обработан".`, 'blue');
      } catch (err) {
        log(panel, 'ОШИБКА при обновлении таблицы: ' + err.message, 'red');
      }
    }

    log(panel, `--- ИТОГО: обработано ${okCount}, ошибок ${errCount}, не найдено платежей ${notFoundCount} ---`, 'black');
    btn.disabled = false;
    btn.textContent = '💳 Привязать платежи из таблицы';
  }

  // ==========================================================
  // Обработка одного заказа: найти платёж "Ожидает оплаты" и заполнить его
  // ==========================================================
  async function processOrder(item) {
    // 1. Открываем страницу заказа, чтобы найти ID платежа со статусом "Ожидает оплаты"
    const orderHtml = await fetchText(`/order/view/${item.orderId}`);
    const orderDoc = new DOMParser().parseFromString(orderHtml, 'text/html');

    const paymentId = findPendingPaymentId(orderDoc);
    if (!paymentId) {
      return 'not_found';
    }

    // 2. Загружаем форму редактирования этого платежа (получаем свежий CSRF-токен)
    const formHtml = await fetchText(`/payment/update/${paymentId}?fancybox=true`);
    const formDoc = new DOMParser().parseFromString(formHtml, 'text/html');

    const csrfInput = formDoc.querySelector('#payment-form input[name="_csrf"]');
    if (!csrfInput) {
      throw new Error('Не найден CSRF-токен в форме платежа');
    }
    const csrf = csrfInput.value;

    const currencySelect = formDoc.querySelector('#payment-currency_id');
    const currencyId = currencySelect ? currencySelect.value : '1';

    // 3. Конвертируем дату dd.mm.yyyy в unix timestamp (как делает сама форма)
    const dateTimestamp = await convertDateToTimestamp(item.date);

    // 4. Отправляем сохранение формы
    const body = new URLSearchParams();
    body.set('_csrf', csrf);
    body.set('Payment[currency_id]', currencyId);
    body.set('Payment[amount]', item.amount.toFixed(2));
    body.set('Payment[payment_method]', PAYMENT_METHOD_VALUE);
    body.set('date-payment-date-disp', item.date);
    body.set('Payment[date]', dateTimestamp);
    body.set('Payment[status]', PAYMENT_STATUS_PAID);

    // сохраняем дату истечения как было (если есть скрытое поле)
    const expireHidden = formDoc.querySelector('#payment-expire_date');
    const expireDisp = formDoc.querySelector('#payment-expire_date-disp');
    if (expireHidden && expireHidden.value) {
      body.set('expire_date-payment-expire_date-disp', expireDisp ? expireDisp.value : '');
      body.set('Payment[expire_date]', expireHidden.value);
    }

    const orderIdField = formDoc.querySelector('#payment-order_id');
    body.set('Payment[order_id]', orderIdField ? orderIdField.value : item.orderId);

    const saveResp = await fetch(`/payment/update/${paymentId}?fancybox=true`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      credentials: 'same-origin',
      body: body.toString()
    });

    if (!saveResp.ok) {
      throw new Error(`Сервер вернул статус ${saveResp.status} при сохранении платежа`);
    }

    return 'ok';
  }

  // ==========================================================
  // Находит ID платежа со статусом "Ожидает оплаты" (1) или "Просрочен" (6) на странице заказа.
  // Каждый select#payment-N-status лежит внутри своей собственной <form>,
  // где рядом есть скрытое поле editableKey с реальным ID платежа в базе.
  // ==========================================================
  function findPendingPaymentId(doc) {
    const statusSelects = [...doc.querySelectorAll('select[id^="payment-"][id$="-status"]')];

    for (const sel of statusSelects) {
      const selectedOption = sel.querySelector('option[selected]');
      const isEligible = selectedOption && ELIGIBLE_STATUSES.includes(selectedOption.value);
      if (!isEligible) continue;

      // Ищем editableKey в пределах той же формы, что и сам select
      const form = sel.closest('form');
      if (!form) continue;

      const editableKeyInput = form.querySelector('input[name="editableKey"]');
      const editableAttrInput = form.querySelector('input[name="editableAttribute"]');

      if (editableKeyInput && editableAttrInput && editableAttrInput.value === 'status') {
        return editableKeyInput.value;
      }
    }
    return null;
  }

  // ==========================================================
  // Конвертация даты dd.mm.yyyy -> unix timestamp
  // (считаем как полночь UTC этой даты — совпадает с тем, что обычно
  // сохраняет datecontrol при saveFormat="U")
  // ==========================================================
  async function convertDateToTimestamp(dateStr) {
    // dateStr формата "01.06.2026"
    const parts = dateStr.split('.');
    if (parts.length !== 3) {
      return String(Math.floor(Date.now() / 1000));
    }
    const [dd, mm, yyyy] = parts;
    const dateObj = new Date(Date.UTC(parseInt(yyyy), parseInt(mm) - 1, parseInt(dd)));
    return String(Math.floor(dateObj.getTime() / 1000));
  }

  function fetchText(url) {
    return fetch(url, { credentials: 'same-origin' }).then(r => {
      if (!r.ok) throw new Error(`Не удалось загрузить ${url}: статус ${r.status}`);
      return r.text();
    });
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // ==========================================================
  // Запуск
  // ==========================================================
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectUI);
  } else {
    injectUI();
  }
})();
