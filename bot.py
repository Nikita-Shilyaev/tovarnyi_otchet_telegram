"""Telegram-бот товарного отчета (long polling).

Команды:
    /updated - название листа report.xlsx на Google Drive (там записана максимальная дата_факта)
    /report  - отправляет в чат файл report.xlsx с Google Drive

Запуск:
    python bot.py            # слушать команды
    python bot.py --updated  # вывести дату обновления и выйти (без Telegram)
"""
import argparse
import logging
import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

from drive_client import download_bytes, get_drive_service

BASE_DIR = Path(__file__).parent

load_dotenv(BASE_DIR / ".env")

CREDS_PATH = BASE_DIR / os.environ["DRIVE_CREDS_PATH"]
TOKEN_PATH = BASE_DIR / os.environ["DRIVE_TOKEN_PATH"]
BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
REPORT_XLSX_ID = os.environ["REPORT_XLSX_ID"]
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"
POLL_TIMEOUT = 50  # сек, long polling

logger = logging.getLogger("bot")


def get_updated_date() -> str:
    """Дата обновления данных - название листа отчета на Drive (ДД.ММ.ГГГГ)."""
    service = get_drive_service(CREDS_PATH, TOKEN_PATH)
    workbook = load_workbook(download_bytes(service, REPORT_XLSX_ID), read_only=True)
    return workbook.sheetnames[0]


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


def handle_updated(chat_id: int) -> None:
    send_message(chat_id, f"Данные обновлены по {get_updated_date()}")


def handle_report(chat_id: int) -> None:
    service = get_drive_service(CREDS_PATH, TOKEN_PATH)
    send_document(chat_id, "report.xlsx", download_bytes(service, REPORT_XLSX_ID).getvalue())


HANDLERS = {"/updated": handle_updated, "/report": handle_report}


def handle_update(update: dict) -> None:
    message = update.get("message") or {}
    text = message.get("text", "")
    if not text.startswith("/"):
        return
    command = text.split()[0].split("@")[0].lower()  # "/updated@bot_name" -> "/updated"
    handler = HANDLERS.get(command)
    if handler:
        handler(message["chat"]["id"])


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
        print(get_updated_date())
    else:
        run()
