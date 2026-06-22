#!/usr/bin/env python3
"""
Полный пайплайн: CRM -> МТС VATS -> Google Таблица

КОМАНДЫ:

1. Выгрузить заказы из CRM и подготовить файл номеров для МТС:
   python scrape_crm.py crm

2. Объединить отчёт МТС с заказами из CRM (после того как обзвон завершился):
   python scrape_crm.py merge ОТЧЕТ_МТС.csv

Установка:
    pip install selenium beautifulsoup4 webdriver-manager openpyxl pandas
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import csv
import time
import sys
import re
from datetime import datetime

import requests

CRM_OUTPUT_FILE   = "crm_orders.xlsx"      # выгрузка из CRM (с привязкой к заказу)
MTS_NUMBERS_FILE  = "mts_numbers.csv"      # файл со списком номеров для загрузки в МТС
MERGED_OUTPUT_FILE = "crm_with_results.xlsx"  # итоговый файл с результатами обзвона

# URL веб-приложения Google Apps Script (созданного через "Развернуть -> Веб-приложение")
GOOGLE_SHEET_WEBHOOK = "https://script.google.com/macros/s/AKfycbyLoeIWSWVnlVME4xbbYSyFm_mjVWM7gpjUvGHDbp9mG4lBWny4R9X9W878j6gC03fucw/exec"


def send_to_google_sheet(action: str, rows: list[dict]):
    """Отправляет данные в Google Таблицу через Apps Script веб-приложение."""
    try:
        resp = requests.post(
            GOOGLE_SHEET_WEBHOOK,
            json={"action": action, "rows": rows},
            timeout=60,
        )
        result = resp.json()
        if result.get("success"):
            print(f"  ☁️  Google Таблица обновлена: {result.get('message')}")
        else:
            print(f"  ⚠️  Ошибка записи в Google Таблицу: {result.get('message')}")
    except Exception as e:
        print(f"  ⚠️  Не удалось связаться с Google Таблицей: {e}")

ROUTES = [
    {"name": "б",     "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=17"},
    {"name": "а",     "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=30"},
    {"name": "в",     "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=58"},
    {"name": "минск", "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=7768"},
    {"name": "6м1",   "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=13876"},
    {"name": "11м",   "url": "https://a.ok-crm.com/route/cabinet-calls?route_id=15064"},
]


# ============================================================
#  ЧАСТЬ 1 — Выгрузка из CRM
# ============================================================

def create_driver():
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except ImportError:
        service = Service()

    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver


def wait_for_manual_login(driver):
    driver.get(ROUTES[0]["url"])
    print()
    print("=" * 55)
    print("  Залогинься через Google в открывшемся браузере.")
    print("  Дождись пока загрузится таблица с заказами.")
    print("  Нажми Enter здесь когда увидишь таблицу.")
    print("=" * 55)
    input("  -> Нажми Enter: ")
    print()


def scroll_and_collect(driver) -> list[dict]:
    print("  Жду загрузки строк...")
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "tr.grid-order-widget"))
        )
    except Exception:
        print("  Строки не появились за 20 сек, пробую парсить...")

    collected = {}
    scroll_attempts = 0
    max_attempts = 30
    last_count = 0

    while scroll_attempts < max_attempts:
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        order_rows = soup.find_all("tr", class_="grid-order-widget")

        for tr in order_rows:
            tds = tr.find_all("td")
            if len(tds) < 6:
                continue
            row_id = clean(tds[1].get_text())
            if not row_id:
                continue
            if row_id not in collected:
                collected[row_id] = {
                    "ID":              row_id,
                    "Товары":          clean(tds[2].get_text()),
                    "Цены":            clean(tds[3].get_text()),
                    "Телефон_сырой":   clean(tds[4].get_text()),
                    "Адрес доставки":  clean(tds[5].get_text()),
                }

        current_count = len(collected)
        print(f"  Собрано: {current_count}", end="\r")

        if current_count == last_count:
            scroll_attempts += 1
        else:
            scroll_attempts = 0
            last_count = current_count

        try:
            driver.execute_script("""
                var container = document.querySelector('.kv-grid-container')
                             || document.querySelector('.table-responsive')
                             || document.querySelector('.grid-view')
                             || document.body;
                container.scrollTop += 500;
                window.scrollBy(0, 300);
            """)
        except Exception:
            pass

        time.sleep(0.5)

    print()
    return list(collected.values())


def clean(text: str) -> str:
    return " ".join(text.split()).strip()


def extract_phone(raw: str) -> str:
    """
    Извлекает первый номер телефона из строки и нормализует в формат 375XXXXXXXXX.
    Примеры входа: '+375 44 709-98-51, +375 2235 55-437' -> '375447099851'
                   'Иванов Иван+375 29 698-18-39'          -> '375296981839'
    """
    # Ищем все последовательности похожие на номер с +375 или 375 или 80
    matches = re.findall(r'(?:\+?375|80)[\d\s\-\(\)]{7,}', raw)
    if not matches:
        return ""

    first = matches[0]
    digits = re.sub(r'\D', '', first)  # убираем всё кроме цифр

    # Нормализуем к формату 375XXXXXXXXX (12 цифр)
    if digits.startswith("00375"):
        digits = digits[2:]
    if digits.startswith("80"):
        digits = "375" + digits[2:]
    if digits.startswith("375") and len(digits) == 12:
        return digits
    if digits.startswith("375") and len(digits) > 12:
        return digits[:12]

    return digits  # вернём как есть если не подошло под правило, для проверки вручную


def write_sheet(ws, rows: list[dict]):
    headers = ["ID", "Товары", "Цены", "Телефон клиента", "Адрес доставки"]
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(rows, 2):
        phone = extract_phone(row["Телефон_сырой"])
        ws.cell(row=row_idx, column=1, value=row["ID"])
        ws.cell(row=row_idx, column=2, value=row["Товары"])
        ws.cell(row=row_idx, column=3, value=row["Цены"])
        ws.cell(row=row_idx, column=4, value=phone)
        ws.cell(row=row_idx, column=5, value=row["Адрес доставки"])

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 45


def cmd_crm():
    """Команда: выгрузить из CRM и подготовить файл номеров для МТС."""
    driver = create_driver()
    all_rows = []  # все строки со всех маршрутов, для файла МТС

    try:
        wait_for_manual_login(driver)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        total = 0
        for route in ROUTES:
            print(f"\n[Маршрут: {route['name']}] {route['url']}")
            driver.get(route["url"])
            time.sleep(2)

            rows = scroll_and_collect(driver)
            print(f"  Итого: {len(rows)} записей")
            total += len(rows)

            ws = wb.create_sheet(title=route["name"])
            write_sheet(ws, rows)

            for row in rows:
                row["Маршрут"] = route["name"]
            all_rows.extend(rows)

        ws_info = wb.create_sheet(title="Инфо")
        ws_info["A1"] = "Последнее обновление:"
        ws_info["B1"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        ws_info["A2"] = "Всего записей:"
        ws_info["B2"] = total

        wb.save(CRM_OUTPUT_FILE)
        print(f"\n✅ CRM-выгрузка сохранена -> {CRM_OUTPUT_FILE} ({total} записей)")

        # Отправляем все строки в Google Таблицу (лист "Заказы")
        print("\n☁️  Отправляю данные в Google Таблицу...")
        gsheet_rows = []
        for row in all_rows:
            phone = extract_phone(row["Телефон_сырой"])
            gsheet_rows.append({
                "ID":              row["ID"],
                "Товары":          row["Товары"],
                "Цены":            row["Цены"],
                "Телефон клиента": phone,
                "Адрес доставки":  row["Адрес доставки"],
                "Маршрут":         row["Маршрут"],
            })
        send_to_google_sheet("upload_crm", gsheet_rows)

        # Готовим файл номеров для МТС: только номера, без заголовка, уникальные, формат 375XXXXXXXXX
        phones = set()
        for row in all_rows:
            phone = extract_phone(row["Телефон_сырой"])
            if re.fullmatch(r"375\d{9}", phone):
                phones.add(phone)
            else:
                print(f"  ⚠️  Пропущен некорректный номер (заказ {row['ID']}): '{row['Телефон_сырой']}'")

        # Сохраняем .xlsx-версию с числовым форматом (0 знаков после запятой),
        # чтобы при открытии в Excel номера не превращались в 3,75E+11
        mts_xlsx_file = MTS_NUMBERS_FILE.replace(".csv", ".xlsx")
        wb_mts = openpyxl.Workbook()
        ws_mts = wb_mts.active
        ws_mts.title = "Номера"
        for i, phone in enumerate(sorted(phones), 1):
            cell = ws_mts.cell(row=i, column=1, value=int(phone))
            cell.number_format = "0"  # числовой формат, без дробной части
        ws_mts.column_dimensions["A"].width = 18
        wb_mts.save(mts_xlsx_file)

        print(f"✅ Файл для МТС сохранён -> {mts_xlsx_file} ({len(phones)} уникальных номеров)")
        print("\nТеперь загрузи этот файл в задание автоинформирования в МТС VATS.")

    finally:
        driver.quit()
        print("Браузер закрыт.")


# ============================================================
#  ЧАСТЬ 2 — Объединение отчёта МТС с заказами CRM
# ============================================================

def load_crm_data() -> dict:
    """Загружает все заказы из CRM-выгрузки в словарь {телефон: [список заказов]}."""
    if not openpyxl.load_workbook:
        pass

    wb = openpyxl.load_workbook(CRM_OUTPUT_FILE, data_only=True)
    phone_to_orders = {}

    for sheet_name in wb.sheetnames:
        if sheet_name == "Инфо":
            continue
        ws = wb[sheet_name]
        headers = [cell.value for cell in ws[1]]

        try:
            id_col = headers.index("ID")
            goods_col = headers.index("Товары")
            price_col = headers.index("Цены")
            phone_col = headers.index("Телефон клиента")
            addr_col = headers.index("Адрес доставки")
        except ValueError:
            continue

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[phone_col]:
                continue
            phone = str(row[phone_col]).strip()
            order = {
                "ID":              row[id_col],
                "Товары":          row[goods_col],
                "Цены":            row[price_col],
                "Телефон клиента": phone,
                "Адрес доставки":  row[addr_col],
                "Маршрут":         sheet_name,
            }
            phone_to_orders.setdefault(phone, []).append(order)

    return phone_to_orders


ANSWER_LABELS = {
    "1": "Клиент готов принять",
    "2": "Клиент не готов принять",
    "3": "Связь с оператором",
}


def load_mts_report(filepath: str) -> dict:
    """
    Загружает отчёт МТС (CSV или Excel) в формате:
    Номер | Часовой пояс | Статус | Количество попыток | Результат последнего вызова |
    Длительность прослушивания | Аудиофайл | Ответ

    Возвращает {телефон: {"статус": ..., "ответ": ..., "ответ_текст": ...}}
    """
    phone_to_result = {}

    if filepath.lower().endswith(".csv"):
        with open(filepath, "r", encoding="utf-8-sig") as f:
            sample = f.read(2048)
            f.seek(0)
            delimiter = ";" if sample.count(";") > sample.count(",") else ","
            reader = csv.DictReader(f, delimiter=delimiter)
            for row in reader:
                _store_mts_row(row, phone_to_result)
    else:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            row_dict = dict(zip(headers, row))
            _store_mts_row(row_dict, phone_to_result)

    return phone_to_result


def _store_mts_row(row: dict, phone_to_result: dict):
    """Извлекает данные из одной строки отчёта МТС и сохраняет в словарь по номеру."""
    phone_raw = str(row.get("Номер", "")).strip()
    phone = normalize_phone(phone_raw)
    if not phone:
        return

    status = str(row.get("Статус", "")).strip()
    answer_raw = row.get("Ответ", "")
    answer_str = str(answer_raw).strip() if answer_raw is not None else ""

    answer_label = ANSWER_LABELS.get(answer_str, "")
    if status.lower() != "успешно":
        result_text = f"Звонок не успешен ({status})" if status else "Нет данных"
    elif answer_label:
        result_text = answer_label
    elif answer_str:
        result_text = f"Ответ: {answer_str}"
    else:
        result_text = "Прослушано, без ответа"

    phone_to_result[phone] = result_text


def normalize_phone(raw: str) -> str:
    digits = re.sub(r'\D', '', raw)
    if digits.startswith("80"):
        digits = "375" + digits[2:]
    if digits.startswith("00375"):
        digits = digits[2:]
    if not digits.startswith("375"):
        return ""
    return digits[:12]


def cmd_merge(report_path: str):
    """Команда: объединить отчёт МТС с заказами CRM."""
    print(f"📥 Загружаю заказы из {CRM_OUTPUT_FILE}...")
    phone_to_orders = load_crm_data()
    total_orders = sum(len(v) for v in phone_to_orders.values())
    print(f"   Загружено заказов: {total_orders} (уникальных номеров: {len(phone_to_orders)})")

    print(f"📥 Загружаю отчёт МТС из {report_path}...")
    phone_to_result = load_mts_report(report_path)
    print(f"   Загружено результатов: {len(phone_to_result)}")

    # Собираем итоговую таблицу
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Результаты"

    headers = ["ID", "Товары", "Цены", "Телефон клиента", "Адрес доставки", "Маршрут", "Результат обзвона"]
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    row_idx = 2
    matched = 0
    gsheet_results = []
    for phone, orders in phone_to_orders.items():
        result = phone_to_result.get(phone, "Нет данных / не дозвонились")
        if phone in phone_to_result:
            matched += 1
        for order in orders:
            ws.cell(row=row_idx, column=1, value=order["ID"])
            ws.cell(row=row_idx, column=2, value=order["Товары"])
            ws.cell(row=row_idx, column=3, value=order["Цены"])
            ws.cell(row=row_idx, column=4, value=order["Телефон клиента"])
            ws.cell(row=row_idx, column=5, value=order["Адрес доставки"])
            ws.cell(row=row_idx, column=6, value=order["Маршрут"])
            ws.cell(row=row_idx, column=7, value=result)
            row_idx += 1

            gsheet_results.append({
                "ID":               order["ID"],
                "Товары":           order["Товары"],
                "Цены":             order["Цены"],
                "Телефон клиента":  order["Телефон клиента"],
                "Адрес доставки":   order["Адрес доставки"],
                "Маршрут":          order["Маршрут"],
                "Результат обзвона": result,
            })

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 45
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 30

    wb.save(MERGED_OUTPUT_FILE)
    print(f"\n✅ Готово! -> {MERGED_OUTPUT_FILE}")
    print(f"   Сопоставлено по номеру: {matched} из {len(phone_to_orders)}")

    print("\n☁️  Отправляю результаты в Google Таблицу...")
    send_to_google_sheet("upload_results", gsheet_results)
    print("   Открой свою Google Таблицу — лист 'Результаты обзвона' обновлён.")


# ============================================================
#  ТОЧКА ВХОДА
# ============================================================

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    command = sys.argv[1]

    if command == "crm":
        cmd_crm()
    elif command == "merge":
        if len(sys.argv) < 3:
            print("Укажи путь к файлу отчёта МТС:")
            print("  python scrape_crm.py merge ОТЧЕТ.csv")
            sys.exit(1)
        cmd_merge(sys.argv[2])
    else:
        print(f"Неизвестная команда: {command}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
