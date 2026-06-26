// ==UserScript==
// @name         CRM status to Telegram
// @namespace    crm-status-monitor
// @version      1.0
// @match        https://a.ok-crm.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      api.telegram.org
// @connect      a.ok-crm.com
// ==/UserScript==

(function () {
  'use strict';

  var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzug3rDGl3VP6319GHU3KLF-P566g42OzS1F2Gyo-OyixlzDpV2zsyfCrXvBr4yfvENCw/exec';
  var APPS_SCRIPT_TOKEN = 'shtenli1';

  var TELEGRAM_BOT_TOKEN = '8864410133:AAFMh0pWwS5POriOVMGYEYAbcG2FjTinIjg';
  var TELEGRAM_CHAT_ID = '308245897';

  var CRM_LIST_URL_TEMPLATE =
    'https://a.ok-crm.com/order/index?OrderSearch%5Bid%5D={id}' +
    '&OrderSearch%5BproductLinks%5D=' +
    '&OrderSearch%5Binfo_payment_method%5D=' +
    '&OrderSearch%5Bbank%5D=' +
    '&OrderSearch%5Bcreated_at%5D=' +
    '&OrderSearch%5BrouteDate%5D=' +
    '&OrderSearch%5Bdelivery_method%5D=' +
    '&OrderSearch%5Bstock_id%5D=' +
    '&OrderSearch%5Bstatus%5D=' +
    '&OrderSearch%5Bdealer_id%5D=' +
    '&OrderSearch%5BtagValuesArr%5D=' +
    '&OrderSearch%5Bclient_phones%5D=' +
    '&OrderSearch%5Bdelivery_address%5D=' +
    '&OrderSearch%5Bcomment%5D=';

  var TARGET_STATUS_KEYWORDS = ['сдан'];
  var NOTIFY_TIMEZONE_OFFSET_HOURS = 3;
  var NOTIFY_START_HOUR = 9;
  var NOTIFY_END_HOUR = 17;
  var NOTIFY_WORKDAYS = [1, 2, 3, 4, 5];
  var DELAY_BETWEEN_ORDERS_MS = 800;

  var STORAGE_KEY = 'crm_monitor_last_statuses';
  var PENDING_KEY = 'crm_monitor_pending_notifications';
  var LOG_PREFIX = '[CRM-Monitor]';

  var KNOWN_STATUSES = [
    'Сдан',
    'Новый',
    'Заморожен',
    'Готов ехать',
    'В пути',
    'Отказ',
    'Не заполнена клиентом',
    'Загружен'
  ];

  function log() {
    var args = Array.prototype.slice.call(arguments);
    args.unshift(LOG_PREFIX);
    console.log.apply(console, args);
  }

  function getStoredStatuses() {
    var raw = GM_getValue(STORAGE_KEY, '{}');
    try { return JSON.parse(raw); } catch (e) { return {}; }
  }

  function saveStoredStatuses(obj) {
    GM_setValue(STORAGE_KEY, JSON.stringify(obj));
  }

  function getPendingNotifications() {
    var raw = GM_getValue(PENDING_KEY, '[]');
    try { return JSON.parse(raw); } catch (e) { return []; }
  }

  function savePendingNotifications(arr) {
    GM_setValue(PENDING_KEY, JSON.stringify(arr));
  }

  function isWithinNotifyWindow() {
    var now = new Date();
    var utcMs = now.getTime() + now.getTimezoneOffset() * 60000;
    var minskMs = utcMs + NOTIFY_TIMEZONE_OFFSET_HOURS * 3600000;
    var minskDate = new Date(minskMs);
    var dayOfWeek = minskDate.getUTCDay();
    var hour = minskDate.getUTCHours();
    if (NOTIFY_WORKDAYS.indexOf(dayOfWeek) === -1) return false;
    if (hour < NOTIFY_START_HOUR || hour >= NOTIFY_END_HOUR) return false;
    return true;
  }

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function fetchOrderIds() {
    return new Promise(function (resolve, reject) {
      var url = APPS_SCRIPT_URL + '?token=' + encodeURIComponent(APPS_SCRIPT_TOKEN);
      GM_xmlhttpRequest({
        method: 'GET',
        url: url,
        onload: function (response) {
          try {
            var data = JSON.parse(response.responseText);
            if (data.error) { reject(new Error('Apps Script error: ' + data.error)); return; }
            resolve(data.ids || []);
          } catch (e) { reject(e); }
        },
        onerror: function (err) { reject(err); }
      });
    });
  }

  var unrecognizedValuesSeen = {};

  function extractStatusFromListHtml(html, orderId) {
    try {
      var parser = new DOMParser();
      var doc = parser.parseFromString(html, 'text/html');
      var row = doc.querySelector('tr[data-key="' + orderId + '"]');
      if (!row) {
        log('Order ' + orderId + ': row not found. HTML length: ' + html.length);
        return null;
      }
      var cells = row.querySelectorAll('td');
      var allCellsText = [];
      for (var i = 0; i < cells.length; i++) {
        var text = (cells[i].textContent || '').trim();
        allCellsText.push(text);
        for (var j = 0; j < KNOWN_STATUSES.length; j++) {
          if (KNOWN_STATUSES[j].toLowerCase() === text.toLowerCase()) {
            return KNOWN_STATUSES[j];
          }
        }
      }
      log('Order ' + orderId + ': status not recognized. Cell texts: [' + allCellsText.join(' | ') + ']');
      return null;
    } catch (e) {
      log('Parse error:', e);
      return null;
    }
  }

  function fetchOrderStatus(orderId) {
    return new Promise(function (resolve) {
      var url = CRM_LIST_URL_TEMPLATE.replace('{id}', encodeURIComponent(orderId));
      GM_xmlhttpRequest({
        method: 'GET',
        url: url,
        onload: function (response) {
          if (response.status !== 200) {
            resolve(null);
            return;
          }
          var status = extractStatusFromListHtml(response.responseText, orderId);
          resolve(status);
        },
        onerror: function () {
          resolve(null);
        }
      });
    });
  }

  function isTargetStatus(statusText) {
    if (!statusText) return false;
    var lower = statusText.toLowerCase();
    for (var i = 0; i < TARGET_STATUS_KEYWORDS.length; i++) {
      if (lower.indexOf(TARGET_STATUS_KEYWORDS[i]) !== -1) return true;
    }
    return false;
  }

  function sendTelegramMessage(text) {
    return new Promise(function (resolve, reject) {
      var url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
      GM_xmlhttpRequest({
        method: 'POST',
        url: url,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: text }),
        onload: function (response) {
          if (response.status === 200) { resolve(); } else { reject(new Error('Telegram HTTP ' + response.status)); }
        },
        onerror: function (err) { reject(err); }
      });
    });
  }

  function flushPendingNotifications() {
    return new Promise(function (resolveOuter) {
      var pending = getPendingNotifications();
      if (pending.length === 0) { resolveOuter(); return; }

      var stillPending = [];
      var i = 0;

      function next() {
        if (i >= pending.length) {
          savePendingNotifications(stillPending);
          resolveOuter();
          return;
        }
        var item = pending[i];
        i++;
        sendTelegramMessage(item.text).then(function () {
          log('Pending notification sent: ' + item.text);
          sleep(DELAY_BETWEEN_ORDERS_MS).then(next);
        }).catch(function (e) {
          log('Failed to send pending notification, keeping in queue:', e);
          stillPending.push(item);
          sleep(DELAY_BETWEEN_ORDERS_MS).then(next);
        });
      }
      next();
    });
  }

  function runCheckCycle() {
    log('Starting check cycle...');

    return flushPendingNotifications().then(function () {
      return fetchOrderIds();
    }).then(function (orderIds) {
      log('Got ' + orderIds.length + ' IDs to check');
      var storedStatuses = getStoredStatuses();
      var newlyDoneOrderIds = [];
      var i = 0;

      function processNext() {
        if (i >= orderIds.length) {
          log('Scan complete. Found ' + newlyDoneOrderIds.length + ' newly done orders');
          return sendFinalNotification();
        }
        var orderId = orderIds[i];
        i++;
        return fetchOrderStatus(orderId).then(function (currentStatus) {
          if (currentStatus !== null) {
            var previousStatus = storedStatuses[orderId];
            var isNewlyDone = isTargetStatus(currentStatus) && !isTargetStatus(previousStatus);
            storedStatuses[orderId] = currentStatus;
            saveStoredStatuses(storedStatuses);

            if (isNewlyDone) {
              newlyDoneOrderIds.push(orderId);
            }
          }
          return sleep(DELAY_BETWEEN_ORDERS_MS).then(processNext);
        });
      }

      function sendFinalNotification() {
        if (newlyDoneOrderIds.length === 0) {
          log('Check cycle finished. Nothing to notify.');
          return Promise.resolve();
        }

        var lines = newlyDoneOrderIds.map(function (orderId) {
          return 'Заказ ' + orderId + ' - сдан';
        });
        var messageText = lines.join('\n');

        return sendTelegramMessage(messageText).then(function () {
          log('Notification sent for ' + newlyDoneOrderIds.length + ' orders: ' + newlyDoneOrderIds.join(', '));
          log('Check cycle finished.');
        }).catch(function (e) {
          log('Failed to send final notification, queuing as one item:', e);
          var pending = getPendingNotifications();
          pending.push({ orderId: 'batch', text: messageText, detectedAt: Date.now() });
          savePendingNotifications(pending);
          log('Check cycle finished (queued due to send error).');
        });
      }

      return processNext();
    }).catch(function (e) {
      log('Check cycle error:', e);
    });
  }

  function createButton() {
    var btn = document.createElement('button');
    btn.textContent = 'CRM Monitor: Run check';
    btn.style.position = 'fixed';
    btn.style.top = '70px';
    btn.style.right = '20px';
    btn.style.zIndex = '999999';
    btn.style.padding = '10px 16px';
    btn.style.backgroundColor = '#2e7d32';
    btn.style.color = '#ffffff';
    btn.style.border = 'none';
    btn.style.borderRadius = '6px';
    btn.style.fontSize = '14px';
    btn.style.fontWeight = 'bold';
    btn.style.cursor = 'pointer';
    btn.style.boxShadow = '0 2px 6px rgba(0,0,0,0.3)';
    btn.id = 'crm-monitor-run-btn';

    btn.addEventListener('click', function () {
      btn.disabled = true;
      btn.textContent = 'Running...';
      btn.style.backgroundColor = '#888888';
      runCheckCycle().then(function () {
        btn.textContent = 'Done! Click to run again';
        btn.style.backgroundColor = '#2e7d32';
        btn.disabled = false;
      }).catch(function (e) {
        log('Run failed:', e);
        btn.textContent = 'Error! Check console';
        btn.style.backgroundColor = '#c62828';
        btn.disabled = false;
      });
    });

    document.body.appendChild(btn);
    log('Run button added to page (top-right corner).');
  }

  function init() {
    try {
      log('Script loaded on page: ' + location.href);
      createButton();
      log('INIT DONE OK.');
    } catch (e) {
      log('INIT FAILED WITH ERROR:', e, e && e.message, e && e.stack);
    }
  }

  try {
    init();
  } catch (e) {
    console.log('[CRM-Monitor] TOP LEVEL INIT CALL FAILED:', e);
  }
})();
