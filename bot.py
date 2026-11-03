"""Telegram-бот товарного отчета (long polling).

Команды:
    /updated   - дата обновления отчета из pipeline_state.csv (строка tovarnyi_otchet_tgbot)
    /report    - берет дату из pipeline_state.csv, находит на Drive папку с этой датой (ДД.ММ.ГГГГ)
                 и отправляет из нее по файлу на каждый доступный филиал ({филиал}_{as_of}.xlsx),
                 а при доступе all - общий отчет report.xlsx по всем филиалам
    /my_store - список доступных филиалов

Доступ берется из Google Таблицы Users (колонки user_id, access): user_id - числовой id в Telegram,
access - название филиала (значение в поле "магазин" отчета) или all.

Запуск:
    python bot.py            # слушать команды
    python bot.py --updated  # вывести дату обновления и выйти (без Telegram)
"""
import argparse
import csv
import io
import logging
import os
import time
from collections import defaultdict
from pathlib import Path

import requests
from dotenv import load_dotenv

from app import DATE_FORMAT, REPORT_FILENAME, REPORTS_FOLDER_ID, filial_filename, get_state_date
from drive_client import download_bytes, export_sheet_csv, find_child, get_drive_service, list_children

BASE_DIR = Path(__file__).parent

load_dotenv(BASE_DIR / ".env")

CREDS_PATH = BASE_DIR / os.environ["DRIVE_CREDS_PATH"]
TOKEN_PATH = BASE_DIR / os.environ["DRIVE_TOKEN_PATH"]
BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
USERS_SHEET_ID = os.environ["USERS_SHEET_ID"]
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"
POLL_TIMEOUT = 50  # сек, long polling

ACCESS_ALL = "all"

logger = logging.getLogger("bot")


def get_updated_date(service=None) -> str | None:
    """Дата обновления отчета из pipeline_state.csv (ДД.ММ.ГГГГ); None, если отчет еще не формировался."""
    state_date = get_state_date(service or get_drive_service(CREDS_PATH, TOKEN_PATH))
    return f"{state_date:{DATE_FORMAT}}" if state_date is not None else None


def send_message(chat_id: int, text: str) -> None:
    response = requests.post(f"{API_URL}/sendMessage", json={"chat_id": chat_id, "text": text}, timeout=30)
    response.raise_for_status()


def send_document(chat_id: int, filename: str, data: bytes) -> None:
    response = requests.post(
        f"{API_URL}/sendDocument",
        data={"chat_id": chat_id},
        files={"document": (filename, data, XLSX_MIMETYPE)},
        timeout=120,
    )
    response.raise_for_status()


def get_access(user_id: int) -> list[str]:
    """Филиалы пользователя из таблицы Users (в порядке строк); ['all'] - все филиалы; [] - нет доступа."""
    service = get_drive_service(CREDS_PATH, TOKEN_PATH)
    rows = csv.DictReader(io.StringIO(export_sheet_csv(service, USERS_SHEET_ID)))
    access = defaultdict(list)
    for row in rows:
        access[row["user_id"].strip()].append(row["access"].strip())
    filials = access.get(str(user_id), [])
    if any(f.lower() == ACCESS_ALL for f in filials):
        return [ACCESS_ALL]
    return filials


def handle_updated(chat_id: int, user_id: int) -> None:
    as_of = get_updated_date()
    send_message(chat_id, f"Данные обновлены по {as_of}" if as_of else "Отчет еще не формировался")


def handle_my_store(chat_id: int, user_id: int) -> None:
    filials = get_access(user_id)
    if not filials:
        send_message(chat_id, f"Нет доступа. Ваш user_id: {user_id}")
    elif filials == [ACCESS_ALL]:
        send_message(chat_id, "Доступен отчет по всем филиалам")
    else:
        send_message(chat_id, ", ".join(filials))


def handle_report(chat_id: int, user_id: int) -> None:
    filials = get_access(user_id)
    if not filials:
        send_message(chat_id, f"Нет доступа. Ваш user_id: {user_id}")
        return

    service = get_drive_service(CREDS_PATH, TOKEN_PATH)
    as_of = get_updated_date(service)
    folder_id = find_child(service, REPORTS_FOLDER_ID, as_of, folder=True) if as_of else None
    if folder_id is None:
        send_message(chat_id, f"Отчет за {as_of} не найден" if as_of else "Отчет еще не формировался")
        return
    files = {child["name"]: child["id"] for child in list_children(service, folder_id)}

    if filials == [ACCESS_ALL]:
        if REPORT_FILENAME not in files:
            send_message(chat_id, f"Отчет за {as_of} не найден")
            return
        send_document(chat_id, f"Все филиалы_{as_of}.xlsx", download_bytes(service, files[REPORT_FILENAME]).getvalue())
        return

    no_data = []
    for filial in filials:
        file_id = files.get(filial_filename(filial))
        if file_id is None:
            no_data.append(filial)
        else:
            send_document(chat_id, f"{filial}_{as_of}.xlsx", download_bytes(service, file_id).getvalue())
    if no_data:
        send_message(chat_id, f"Нет данных в отчете по филиалам: {', '.join(no_data)}")


HANDLERS = {"/updated": handle_updated, "/report": handle_report, "/my_store": handle_my_store}


def handle_update(update: dict) -> None:
    message = update.get("message") or {}
    text = message.get("text", "")
    if not text.startswith("/"):
        return
    command = text.split()[0].split("@")[0].lower()  # "/updated@bot_name" -> "/updated"
    handler = HANDLERS.get(command)
    if handler:
        handler(message["chat"]["id"], message["from"]["id"])


def run() -> None:
    offset = None
    logger.info("Бот запущен")
    while True:
        try:
            response = requests.get(
                f"{API_URL}/getUpdates",
                params={"timeout": POLL_TIMEOUT, "offset": offset},
                timeout=POLL_TIMEOUT + 10,
            )
            response.raise_for_status()
            for update in response.json()["result"]:
                offset = update["update_id"] + 1
                try:
                    handle_update(update)
                except Exception:
                    logger.exception("Ошибка обработки update %s", update["update_id"])
        except requests.RequestException:
            logger.exception("Ошибка getUpdates, повтор через 5 сек")
            time.sleep(5)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    logging.getLogger("httpx").setLevel(logging.WARNING)
    parser = argparse.ArgumentParser()
    parser.add_argument("--updated", action="store_true", help="вывести дату обновления и выйти")
    args = parser.parse_args()
    if args.updated:
        print(get_updated_date() or "Отчет еще не формировался")
    else:
        run()
