"""Витрина отчета: план / факт / остаток / средняя цена по месяц - магазин - бит литера - категория.

Читает план (csv), дневной план, Sales_sku_target и Ostatki_target с Google Drive.
Результат кладет на Drive в папку с датой обновления данных (ДД.ММ.ГГГГ - последняя дата продаж) внутри
REPORTS_FOLDER_ID: report.xlsx (все филиалы) и по файлу на каждую комбинацию филиалов пользователей из таблицы
Users ({филиал1}_{филиал2}_....xlsx, филиалы по алфавиту). После загрузки всех файлов
записывает эту дату в pipeline_state.csv (source=tovarnyi_otchet_tgbot) - по ней бот находит папку.
Листы report.xlsx: "План факт по филиалам" (первый): план на год и месяц
(дневной план), факт YTD/MTD (все продажи филиала), выполнение плана и факт прошлого года за те же даты;
"План факт по категориям" - витрина; "Артикулы" - остаток и скорость продаж с учетом наличия по магазину, артикулу,
цвету и размеру (все артикулы и магазины, одинаковый во всех файлах); "Обозначения" - расшифровка YTD/MTD/LY/YoY. В файлах комбинаций те же листы,
на листах по филиалам и по категориям - только филиалы комбинации.
Месяцы: все месяцы плана с января 2026 по месяц последней даты продаж.
- факт: сумма_со_скидкой за месяц (для текущего месяца - MTD по as_of);
- остаток: срез на 1-е число месяца; если среза на 1-е нет (текущий месяц) - последний срез внутри месяца.

Запуск:
    python app.py                  # MTD на последнюю дату в продажах
    python app.py --as-of 2026-10-03
    python app.py --schedule       # висит в фоне и обновляет отчет сразу и далее каждые 24 часа
"""
import argparse
import csv
import io
import logging
import os
import re
import textwrap
import uuid
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from apscheduler.schedulers.blocking import BlockingScheduler
from dotenv import load_dotenv
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from drive_client import (download_bytes, export_sheet_csv, find_or_create_subfolder, get_drive_service,
                          update_file_bytes, upload_bytes)

BASE_DIR = Path(__file__).parent
logger = logging.getLogger("app")

load_dotenv(BASE_DIR / ".env")

CREDS_PATH = BASE_DIR / os.environ["DRIVE_CREDS_PATH"]
TOKEN_PATH = BASE_DIR / os.environ["DRIVE_TOKEN_PATH"]
SALES_SKU_ID = os.environ["SALES_SKU_TARGET_ID"]
OSTATKI_ID = os.environ["OSTATKI_TARGET_ID"]

PLAN_ID = os.environ["PLAN_PRODAZH_ID"]
DAILY_PLAN_ID = os.environ["DAILY_PLAN_ID"]  # daily_plan_updated_06082026.csv: date, магазин, план, месячный_план
REPORTS_FOLDER_ID = os.environ["REPORTS_FOLDER_ID"]  # папка, в которой создаются папки с датами отчета
PIPELINE_STATE_ID = os.environ["PIPELINE_STATE_ID"]
USERS_SHEET_ID = os.environ["USERS_SHEET_ID"]  # Google Таблица доступов: user_id, access (филиал или all)
ACCESS_ALL = "all"
XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CSV_MIMETYPE = "text/csv"

REPORT_FILENAME = "report.xlsx"
DATE_FORMAT = "%d.%m.%Y"  # имя папки отчета
# строка отчета в pipeline_state.csv (source, layer, updated_at); updated_at в ISO, как у остальных строк
STATE_SOURCE = "tovarnyi_otchet_tgbot"
STATE_LAYER = "mart"

# форматы Excel: колонка -> (формат, вид отображения для расчета ширины)
CURRENCY = r"[$₸-43F]"  # тенге (kk-KZ)
FINANCIAL = rf'_-* #,##0\ {CURRENCY}_-;\-* #,##0\ {CURRENCY}_-;_-* "-"\ {CURRENCY}_-;_-@_-'  # финансовый с валютой, без десятичных знаков
COLUMN_FORMATS = {
    "доля бит литеры": ("0%", ("percent", 0)),
    "доля категории внутри бит литеры": ("0%", ("percent", 0)),
    "выполнение плана": ("0.00%", ("percent", 2)),
    "план категории": (FINANCIAL, ("financial", 0)),
    "факт_mtd": (FINANCIAL, ("financial", 0)),
    "плана осталось выполнить": (FINANCIAL, ("financial", 0)),
    "средняя цена продажи с учетом скидки": (FINANCIAL, ("financial", 0)),
    "скорость продаж с учетом наличия": ("0.00", ("fixed", 2)),
}
# заливка флага "возможность выполнить план" (стандартные цвета палитры Excel)
FLAG_COLUMN = "возможность выполнить план"
FLAG_COLORS = {
    "да, достаточный запас": "92D050",  # светло-зеленый
    "да, рекомендуется пополнить запасы": "FFFF00",  # желтый
    "нет, остатков не хватает": "FF0000",  # красный
}
HEADER_LINES = 3  # высота строки заголовков = столько обычных строк
ROW_HEIGHT = 15  # высота обычной строки в Excel, pt

# Средняя ежемесячная скидка от розничной цены (память проекта, 2026-10-04)
DISCOUNT = {"ЮИ": 0.10, "ЗИ": 0.10, "СГ": 0.0}

# колонки плана в csv в порядке листа "План факт по категориям" ("план месяц" на листе не выводится)
PLAN_COLUMNS = ["месяц", "магазин", "бит литера", "категория", "доля бит литеры",
                "доля категории внутри бит литеры", "план"]
# переименование колонок на листе "План факт по категориям" (расчеты идут по исходным названиям)
REPORT_RENAME = {"план": "план категории", "плана осталось": "плана осталось выполнить"}
# в файлах филиалов эти столбцы листа "План факт по категориям" скрыты (видны только в report.xlsx)
FILIAL_HIDDEN_COLUMNS = ["доля бит литеры", "доля категории внутри бит литеры"]

# остаток / шт продать осталось: от этого значения запас считается достаточным
STOCK_COVER_OK = 3

KEY = ["магазин", "бит_литера", "категория"]


def calc_fact(sales: pd.DataFrame, month_start: pd.Timestamp, as_of: pd.Timestamp) -> pd.Series:
    month_end = month_start + pd.offsets.MonthEnd(0)
    period = sales[(sales["date"] >= month_start) & (sales["date"] <= min(month_end, as_of))]
    return period.groupby(KEY)["сумма_со_скидкой"].sum().rename("факт_mtd")


def pick_snapshot(snapshots: list, month_start: pd.Timestamp) -> pd.Timestamp:
    """Срез на 1-е число месяца; если его нет - последний срез внутри месяца."""
    month_end = month_start + pd.offsets.MonthEnd(0)
    in_month = [d for d in snapshots if month_start <= d <= month_end]
    if not in_month:
        raise SystemExit(f"Нет среза остатков за {month_start:%m-%Y}")
    return month_start if month_start in in_month else max(in_month)


def calc_stock(snapshot: pd.Timestamp, ostatki: pd.DataFrame) -> pd.DataFrame:
    """Остаток шт и средневзвешенная (по остатку) розничная цена на последнем срезе."""
    last = ostatki[ostatki["date"] == snapshot].copy()
    last["сумма_розница"] = last["розничная_цена"] * last["остаток"]
    g = last.groupby(KEY)[["остаток", "сумма_розница"]].sum()
    g["розничная_цена_ср"] = g["сумма_розница"] / g["остаток"].where(g["остаток"] > 0)
    return g.rename(columns={"остаток": "остаток_шт"})[["остаток_шт", "розничная_цена_ср"]]


def read_drive_csv(service, file_id: str, **kwargs) -> pd.DataFrame:
    """Скачивает csv с Google Drive по id и читает его в DataFrame."""
    return pd.read_csv(download_bytes(service, file_id), **kwargs)


def display_length(value, display=None) -> int:
    """Длина значения так, как оно отображается в ячейке."""
    if value is None or value == "" or (isinstance(value, float) and np.isnan(value)):
        return 0
    kind, digits = display or (None, 0)
    if kind == "percent":
        return len(f"{value * 100:.{digits}f}%")
    if kind == "fixed":
        return len(f"{value:.{digits}f}")
    if kind == "financial":
        return len(f"{value:,.{digits}f}") + 4  # " ₸" и отступы формата _- и -_
    if isinstance(value, float):
        return len(f"{value:.10g}")
    return len(str(value))


def header_width(header: str, min_width: int) -> int:
    """Минимальная ширина, при которой заголовок переносится не более чем на HEADER_LINES строк."""
    width = max(min_width, max(len(word) for word in header.split()))
    while len(textwrap.wrap(header, width)) > HEADER_LINES:
        width += 1
    return width


def format_sheet(ws, columns: list, table_name: str = "report", hidden: tuple = ()) -> None:
    """Таблица Excel: стиль таблицы, заголовки, форматы чисел, ширина столбцов по значениям; hidden - скрытые столбцы."""
    last_col, last_row = get_column_letter(ws.max_column), ws.max_row
    table = Table(displayName=table_name, ref=f"A1:{last_col}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.row_dimensions[1].height = ROW_HEIGHT * HEADER_LINES

    for idx, name in enumerate(columns, start=1):
        letter = get_column_letter(idx)
        ws[f"{letter}1"].alignment = Alignment(wrap_text=True, vertical="center")
        number_format, display = COLUMN_FORMATS.get(name, (None, None))
        width = 0
        for (cell,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            if number_format:
                cell.number_format = number_format
            width = max(width, display_length(cell.value, display))
        # запас под фильтр в заголовке; заголовок должен влезать в HEADER_LINES строк
        ws.column_dimensions[letter].width = header_width(name, max(width, 6) + 3)
        ws.column_dimensions[letter].hidden = name in hidden
        if name == FLAG_COLUMN and last_row > 1:
            for text, color in FLAG_COLORS.items():
                fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                ws.conditional_formatting.add(
                    f"{letter}2:{letter}{last_row}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill)
                )


REPORT_SHEET = "План факт по категориям"
PLAN_FACT_SHEET = "План факт по филиалам"
PLAN_FACT_LABEL = "показатель"
# строки листа план-факт: суммы - финансовый формат, доли - проценты без знаков после запятой
PLAN_FACT_AMOUNTS = ["план год", "план месяц", "факт YTD", "факт MTD", "факт YTD LY", "факт MTD LY"]
PLAN_FACT_RATIOS = ["выполнение плана год", "выполнение плана месяц", "откл факт YTD YoY", "откл факт MTD YoY"]
PLAN_FACT_ROWS = ["план год", "план месяц", "факт YTD", "факт MTD", "выполнение плана год", "выполнение плана месяц",
                  "факт YTD LY", "факт MTD LY", "откл факт YTD YoY", "откл факт MTD YoY"]
GLOSSARY_SHEET = "Обозначения"
GLOSSARY = pd.DataFrame(
    [
        ("YTD", "период с начала текущего года по дату отчета"),
        ("MTD", "период с первого числа текущего месяца по дату отчета"),
        ("LY", "показатель за такой же период в прошлом году"),
        ("YoY", "сравнение показателей за текущий год к прошлому году за одинаковый период"),
        ("месяцев с остатком", "за последние 12 месяцев до даты отчета: в скольких месяцах артикул был в наличии в магазине "
                               "(остаток больше 0 на конец месяца)"),
        ("скорость продаж с учетом наличия", "среднее кол-во продаж артикула в шт за месяц, когда он был в наличии: "
                                             "продажи шт за 12 мес / месяцев с остатком; пусто, если месяцев с остатком 0"),
    ],
    columns=["обозначение", "описание"],
)
GLOSSARY_WRAP_WIDTH = 50  # ширина столбца "описание" (символы), текст переносится по словам
# условное форматирование листа план-факт (цвета стандартных правил Excel)
PLAN_FACT_BAR_ROWS = ["выполнение плана год", "выполнение плана месяц"]  # голубая гистограмма, сплошная заливка
PLAN_FACT_BAR_COLOR = "FF008AEF"  # "Голубая гистограмма"
PLAN_FACT_SIGN_ROWS = ["откл факт YTD YoY", "откл факт MTD YoY"]  # <= 0 - красный, > 0 - зеленый
SIGN_NEGATIVE = (PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), Font(color="9C0006", bold=True))
SIGN_POSITIVE = (PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), Font(color="006100"))


def sum_by_filial(df: pd.DataFrame, column: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
    period = df[(df["date"] >= start) & (df["date"] <= end)]
    return period.groupby("магазин")[column].sum()


def calc_plan_fact(sales: pd.DataFrame, daily_plan: pd.DataFrame, filials: list, as_of: pd.Timestamp) -> pd.DataFrame:
    """План / факт по филиалам: строки - показатели, столбцы - филиалы.

    План год / месяц - сумма дневного плана за весь год / месяц as_of.
    Факт YTD - с 1 января по as_of, MTD - с 1-го числа месяца по as_of; LY - те же даты прошлого года.
    Факт - все продажи филиала (сумма_со_скидкой).
    выполнение плана год / месяц = факт YTD / план год, факт MTD / план месяц; откл факт YTD / MTD YoY = факт / факт LY - 1 (доли).
    """
    year_start, month_start = as_of.replace(month=1, day=1), as_of.replace(day=1)
    ly = pd.DateOffset(years=1)
    rows = {
        "план год": sum_by_filial(daily_plan, "план", year_start, year_start + pd.offsets.YearEnd(0)),
        "план месяц": sum_by_filial(daily_plan, "план", month_start, month_start + pd.offsets.MonthEnd(0)),
        "факт YTD": sum_by_filial(sales, "сумма_со_скидкой", year_start, as_of),
        "факт MTD": sum_by_filial(sales, "сумма_со_скидкой", month_start, as_of),
        "факт YTD LY": sum_by_filial(sales, "сумма_со_скидкой", year_start - ly, as_of - ly),
        "факт MTD LY": sum_by_filial(sales, "сумма_со_скидкой", month_start - ly, as_of - ly),
    }
    t = pd.DataFrame(rows).reindex(filials)
    for name in ["факт YTD", "факт MTD", "факт YTD LY", "факт MTD LY"]:
        t[name] = t[name].fillna(0)
    ratio = lambda a, b: t[a] / t[b].where(t[b] > 0)  # нет плана / нет продаж LY - пусто
    t["выполнение плана год"] = ratio("факт YTD", "план год")
    t["выполнение плана месяц"] = ratio("факт MTD", "план месяц")
    t["откл факт YTD YoY"] = ratio("факт YTD", "факт YTD LY") - 1
    t["откл факт MTD YoY"] = ratio("факт MTD", "факт MTD LY") - 1
    return t[PLAN_FACT_ROWS].T.rename_axis(PLAN_FACT_LABEL).reset_index()


def format_plan_fact_sheet(ws, columns: list) -> None:
    """Лист план-факт: таблица Excel, форматы по строкам (суммы / доли), ширина столбцов по значениям,
    условное форматирование строк выполнения плана (гистограмма) и откл факт YTD / MTD YoY (знак)."""
    last_col, last_row = get_column_letter(ws.max_column), ws.max_row
    table = Table(displayName="plan_fact", ref=f"A1:{last_col}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    ws.row_dimensions[1].height = ROW_HEIGHT * HEADER_LINES
    widths = [0] * len(columns)
    for row in ws.iter_rows(min_row=2):
        label = row[0].value
        number_format, display = (FINANCIAL, ("financial", 0)) if label in PLAN_FACT_AMOUNTS else ("0%", ("percent", 0))
        for idx, cell in enumerate(row):
            if idx:
                cell.number_format = number_format
            widths[idx] = max(widths[idx], display_length(cell.value, display if idx else None))
    for idx, name in enumerate(columns):
        letter = get_column_letter(idx + 1)
        ws[f"{letter}1"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[letter].width = header_width(name, max(widths[idx], 6) + 3)

    if ws.max_column < 2:
        return
    for (cell,) in ws.iter_rows(min_row=2, max_col=1):
        cells = f"B{cell.row}:{last_col}{cell.row}"
        if cell.value in PLAN_FACT_BAR_ROWS:
            # шкала 0..100% плана (а не min..max по строке): в файле филиала одна ячейка, полоса была бы всегда полной
            ws.conditional_formatting.add(cells, DataBarRule(start_type="num", start_value=0, end_type="num",
                                                             end_value=1, color=PLAN_FACT_BAR_COLOR))
        elif cell.value in PLAN_FACT_SIGN_ROWS:
            first = f"B{cell.row}"  # ISNUMBER: пустые ячейки (нет продаж LY) не красим
            for formula, (fill, font) in [(f"AND(ISNUMBER({first}),{first}<=0)", SIGN_NEGATIVE),
                                          (f"AND(ISNUMBER({first}),{first}>0)", SIGN_POSITIVE)]:
                ws.conditional_formatting.add(cells, FormulaRule(formula=[formula], fill=fill, font=font))


X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"
DATA_BAR_RE = re.compile(
    r'(<conditionalFormatting sqref="([^"]+)"><cfRule type="dataBar" priority="\d+"><dataBar>(.*?)</dataBar>)(</cfRule>)'
)


def solid_data_bars(xlsx: bytes) -> bytes:
    """Делает гистограммы сплошной заливкой.

    openpyxl пишет только гистограмму Excel 2007 (всегда градиент); сплошная заливка задается расширением x14
    (gradient="0"), которое дописываем в xml листов: ссылка x14:id в cfRule и описание правила в extLst листа.
    """
    src = zipfile.ZipFile(io.BytesIO(xlsx))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and b'type="dataBar"' in data:
                xml, x14_rules = data.decode("utf-8"), []

                def add_ext(match):
                    rule_id = f"{{{uuid.uuid4()}}}".upper()
                    cfvo = "".join(f'<x14:cfvo type="{t}"><xm:f>{v}</xm:f></x14:cfvo>'
                                   for t, v in re.findall(r'<cfvo type="(\w+)" val="([^"]*)"', match.group(3)))
                    x14_rules.append(
                        f'<x14:conditionalFormatting xmlns:xm="{XM_NS}"><x14:cfRule type="dataBar" id="{rule_id}">'
                        f'<x14:dataBar minLength="0" maxLength="100" gradient="0">{cfvo}'
                        f'<x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>'
                        f'</x14:cfRule><xm:sqref>{match.group(2)}</xm:sqref></x14:conditionalFormatting>'
                    )
                    return (f'{match.group(1)}<extLst><ext uri="{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}" '
                            f'xmlns:x14="{X14_NS}"><x14:id>{rule_id}</x14:id></ext></extLst>{match.group(4)}')

                xml = DATA_BAR_RE.sub(add_ext, xml)
                ext = (f'<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" xmlns:x14="{X14_NS}">'
                       f'<x14:conditionalFormattings>{"".join(x14_rules)}</x14:conditionalFormattings></ext></extLst>')
                data = xml.replace("</worksheet>", ext + "</worksheet>").encode("utf-8")
            dst.writestr(item, data)
    return out.getvalue()


ARTICLES_SHEET = "Артикулы"
# изделие: разные бит литера / поставщик / категория / подкатегория = разные изделия
ITEM_KEY = ["артикул", "бит литера", "поставщик", "категория", "подкатегория"]
# прочие признаки - самое частое значение по изделию (одинаковые во всех строках изделия)
ITEM_ATTRS = ["коллекция", "вставка"]
# строка листа: магазин + изделие + цвет + размер
VARIANT = ["цвет", "размер"]
NO_SIZE = -1.0  # пустой размер в ключе группировки (NaN в ключах groupby теряется)
ARTICLES_COLUMNS = ["магазин", "артикул", "поставщик", "бит литера", "категория", "подкатегория", "коллекция", "вставка",
                    "цвет", "размер", "продажи шт за 12 мес", "месяцев с остатком",
                    "скорость продаж с учетом наличия", "текущий остаток"]
# цвет в остатках пишется по-разному ("Белое", "БЕЛЫЙ", "<>", "жолтое") - приводим к словарю продаж
COLOR_STEMS = [("бел", "Белый"), ("жел", "Желтый"), ("жол", "Желтый"), ("крас", "Красный"), ("роз", "Розовый"),
               ("комб", "Комбинированный"), ("серебр", "Серебро"), ("золот", "Золото")]


def normalize_color(value) -> str:
    """Цвет к единому словарю; пусто, мусор ("<>", "---", "не указан", размеры) и "нет данных" - пустая строка."""
    text = str(value).strip().lower() if pd.notna(value) else ""
    return next((color for stem, color in COLOR_STEMS if text.startswith(stem)), "")


def barcode_attrs(sales: pd.DataFrame, ostatki: pd.DataFrame) -> pd.DataFrame:
    """Признаки изделия по штрихкоду: последняя продажа, для непроданных - последняя запись в остатках (по всем срезам).

    Один источник признаков для продаж и остатков, чтобы продажи и остаток одного изделия попали в одну строку
    (в старых срезах остатков признаки проданных штрихкодов бывают устаревшими, например поставщик).
    Подкатегории в остатках нет - берется самая частая по артикулу в продажах.
    """
    cols = ["штрихкод", "артикул", "бит_литера", "поставщик", "категория", "коллекция", "вставка", "цвет", "размер"]
    from_stock = ostatki.rename(columns={"камень": "вставка"}).sort_values("date").drop_duplicates("штрихкод", keep="last")[cols]
    from_sales = sales.sort_values("date").drop_duplicates("штрихкод", keep="last")[cols]
    # артикул в остатках и продажах записан по-разному ("544020301429" / "544020301429P", "TR679" / "R-TR679"):
    # по общим штрихкодам переводим артикулы остатков в формат продаж (самое частое соответствие)
    pairs = from_stock[["штрихкод", "артикул"]].merge(from_sales[["штрихкод", "артикул"]], on="штрихкод", suffixes=("", "_sales"))
    pairs = pairs[pairs["артикул"] != pairs["артикул_sales"]]
    art_map = pairs.groupby("артикул")["артикул_sales"].agg(lambda col: col.mode().iat[0])
    from_stock["артикул"] = from_stock["артикул"].map(art_map).fillna(from_stock["артикул"])
    from_stock = from_stock[~from_stock["штрихкод"].isin(from_sales["штрихкод"])]
    attrs = pd.concat([from_sales, from_stock], ignore_index=True).rename(columns={"бит_литера": "бит литера"})
    attrs["цвет"] = attrs["цвет"].map(normalize_color)
    attrs["размер"] = pd.to_numeric(attrs["размер"], errors="coerce")
    subcategory = sales.dropna(subset=["подкатегория"]).groupby("артикул")["подкатегория"].agg(lambda col: col.mode().iat[0])
    attrs["подкатегория"] = attrs["артикул"].map(subcategory)
    attrs[ITEM_KEY[1:]] = attrs[ITEM_KEY[1:]].fillna("")  # пустые признаки - пустая строка (NaN в ключах группировки теряется)
    return attrs.set_index("штрихкод")


def calc_articles(sales: pd.DataFrame, ostatki: pd.DataFrame, snapshot: pd.Timestamp, as_of: pd.Timestamp) -> pd.DataFrame:
    """Лист "Артикулы": строка - магазин, изделие (артикул, бит литера, поставщик, категория, подкатегория),
    цвет и размер; скорость продаж с учетом наличия и текущий остаток.

    продажи шт за 12 мес - количество (нетто с возвратами) за 12 месяцев до даты отчета
    (as_of - 1 год ... as_of - 1 день).
    месяцев с остатком - за тот же период число месяцев с остатком > 0 на конец месяца (срезы остатков на конец
    месяца хранятся под 1-м числом этого месяца: 2026-07-01 = остаток на 31.07.2026; берутся только полностью
    прошедшие месяцы периода, недельные срезы текущего месяца не считаются).
    скорость продаж с учетом наличия = продажи шт за 12 мес / месяцев с остатком (пусто при 0 месяцев).
    текущий остаток - срез остатков на дату отчета (тот же, что "дата_остатка" текущего месяца в витрине).
    Строки: магазин - изделие - цвет - размер с остатком > 0 или с продажами за период; все бит литеры и магазины.
    Порядок: магазин, артикул, продажи шт за 12 мес по убыванию.
    """
    attrs = barcode_attrs(sales, ostatki)
    attrs = attrs[~attrs.index.duplicated(keep="first")]  # платки и монеты могут делить штрихкод
    attrs["размер"] = attrs["размер"].fillna(NO_SIZE)
    period = sales[(sales["date"] >= as_of - pd.DateOffset(years=1)) & (sales["date"] < as_of)]
    start = as_of - pd.DateOffset(years=1)
    month_ends = ostatki[(ostatki["date"].dt.day == 1) & (ostatki["date"] >= start)
                         & (ostatki["date"] + pd.offsets.MonthEnd(0) < as_of) & (ostatki["остаток"] > 0)]
    stock_rows = ostatki[ostatki["date"] == snapshot]
    keyed = lambda df: df[["магазин", "штрихкод"]].join(attrs[ITEM_KEY + ITEM_ATTRS + VARIANT], on="штрихкод")
    sold_rows = keyed(period).assign(qty=period["количество"].values)
    stock_rows = keyed(stock_rows).assign(qty=stock_rows["остаток"].values)
    month_rows = keyed(month_ends).assign(month=month_ends["date"].values)
    group = ["магазин", *ITEM_KEY, *VARIANT]
    t = pd.concat(
        [sold_rows.groupby(group)["qty"].sum().rename("продажи шт за 12 мес"),
         stock_rows.groupby(group)["qty"].sum().rename("текущий остаток")],
        axis=1,
    ).fillna(0)
    stock_months = month_rows.groupby(group)["month"].nunique().rename("месяцев с остатком")
    t = t.join(stock_months).fillna({"месяцев с остатком": 0})
    t = t[(t["текущий остаток"] > 0) | (t["продажи шт за 12 мес"] != 0)].reset_index()
    mode = lambda col: col.mode().iat[0] if col.notna().any() else None
    item_attrs = pd.concat([stock_rows, sold_rows]).groupby(ITEM_KEY)[ITEM_ATTRS].agg(mode)
    t = t.merge(item_attrs, left_on=ITEM_KEY, right_index=True, how="left")
    t["размер"] = t["размер"].mask(t["размер"] == NO_SIZE)  # без размера - пустая ячейка
    t["месяцев с остатком"] = t["месяцев с остатком"].astype(int)
    months = t["месяцев с остатком"]
    t["скорость продаж с учетом наличия"] = t["продажи шт за 12 мес"] / months.where(months > 0)
    for col in ["продажи шт за 12 мес", "текущий остаток"]:
        if (t[col] % 1 == 0).all():  # штуки целые - пишем без ".0"
            t[col] = t[col].astype(int)
    t = t.sort_values(["магазин", "артикул", "продажи шт за 12 мес"], ascending=[True, True, False], kind="stable")
    return t[ARTICLES_COLUMNS]


def format_glossary_sheet(ws) -> None:
    """Лист "Обозначения": "обозначение" по ширине текста, "описание" фиксированной ширины с переносом;
    высота строки по числу строк переноса (оценка с небольшим запасом на ширину букв)."""
    ws.column_dimensions["A"].width = max(len("обозначение"), GLOSSARY["обозначение"].str.len().max()) + 3
    ws.column_dimensions["B"].width = GLOSSARY_WRAP_WIDTH
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        lines = len(textwrap.wrap(str(row[1].value), GLOSSARY_WRAP_WIDTH - 3))
        ws.row_dimensions[row[0].row].height = ROW_HEIGHT * max(lines, 1)


def build_report_xlsx(df: pd.DataFrame, plan_fact: pd.DataFrame = None, articles: pd.DataFrame = None,
                      hidden_columns: tuple = ()) -> bytes:
    """xlsx в памяти: лист план-факт по филиалам (если передан, первым), лист витрины по категориям
    (hidden_columns на нем скрыты), лист "Артикулы" (если передан) и (вместе с план-фактом) лист "Обозначения"."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if plan_fact is not None:
            plan_fact.to_excel(writer, sheet_name=PLAN_FACT_SHEET, index=False)
            format_plan_fact_sheet(writer.sheets[PLAN_FACT_SHEET], list(plan_fact.columns))
        df.to_excel(writer, sheet_name=REPORT_SHEET, index=False)
        format_sheet(writer.sheets[REPORT_SHEET], list(df.columns), hidden=hidden_columns)
        if articles is not None:
            articles.to_excel(writer, sheet_name=ARTICLES_SHEET, index=False)
            format_sheet(writer.sheets[ARTICLES_SHEET], list(articles.columns), table_name="articles")
        if plan_fact is not None:
            GLOSSARY.to_excel(writer, sheet_name=GLOSSARY_SHEET, index=False)
            format_glossary_sheet(writer.sheets[GLOSSARY_SHEET])
    return solid_data_bars(buffer.getvalue())


def read_access(service) -> dict[str, list[str]]:
    """Доступы из таблицы Users: user_id -> филиалы по алфавиту (без повторов) или ['all'].

    Порядок строк в таблице не важен: одинаковый набор филиалов у разных пользователей дает одну комбинацию.

    Строки с пустым user_id (пользователь еще не зарегистрирован в боте) собираются под ключом "",
    чтобы файл его комбинации филиалов был готов к моменту регистрации.
    """
    access = defaultdict(set)
    for row in csv.DictReader(io.StringIO(export_sheet_csv(service, USERS_SHEET_ID))):
        user_id, filial = (row.get("user_id") or "").strip(), (row.get("access") or "").strip()
        if filial:
            access[user_id].add(filial)
    return {user_id: [ACCESS_ALL] if any(f.lower() == ACCESS_ALL for f in filials) else sorted(filials)
            for user_id, filials in access.items()}


def access_combos(access: dict[str, list[str]]) -> list[tuple]:
    """Уникальные комбинации филиалов пользователей (без доступа all), без повторов."""
    return list(dict.fromkeys(tuple(filials) for filials in access.values() if filials != [ACCESS_ALL]))


def combo_name(filials) -> str:
    """Имя отчета комбинации филиалов: филиалы через "_" по алфавиту (общее для витрины и бота)."""
    return "_".join("".join("_" if ch in '\\/:*?"<>|' else ch for ch in filial).strip() for filial in filials)


def combo_filename(filials) -> str:
    """Имя файла комбинации филиалов в папке отчета."""
    return combo_name(filials) + ".xlsx"


def read_pipeline_state(service) -> pd.DataFrame:
    return read_drive_csv(service, PIPELINE_STATE_ID, dtype=str, keep_default_na=False)


def get_state_date(service) -> pd.Timestamp | None:
    """Дата обновления отчета из pipeline_state.csv; None, если строки отчета нет."""
    state = read_pipeline_state(service)
    row = state[state["source"] == STATE_SOURCE]
    return pd.Timestamp(row["updated_at"].iloc[0]) if not row.empty else None


def set_state_date(service, as_of: pd.Timestamp) -> None:
    """Записывает дату отчета в pipeline_state.csv (обновляет строку отчета или добавляет ее в конец)."""
    state = read_pipeline_state(service)
    value = f"{as_of:%Y-%m-%d}"
    mask = state["source"] == STATE_SOURCE
    if mask.any():
        state.loc[mask, "updated_at"] = value
    else:
        state.loc[len(state)] = {"source": STATE_SOURCE, "layer": STATE_LAYER, "updated_at": value}
    data = state.to_csv(index=False, lineterminator="\r\n").encode("utf-8")
    update_file_bytes(service, PIPELINE_STATE_ID, data, CSV_MIMETYPE)


def update_report(as_of_arg: str = None) -> None:
    service = get_drive_service(CREDS_PATH, TOKEN_PATH)
    plan = read_drive_csv(service, PLAN_ID, dtype={"месяц": str}).rename(columns={"бит литера": "бит_литера"})

    sales = read_drive_csv(service, SALES_SKU_ID, usecols=["date", *KEY, "штрихкод", "артикул", "поставщик", "подкатегория",
                                                            "коллекция", "вставка", "цвет", "размер",
                                                            "количество", "сумма_со_скидкой"],
                           parse_dates=["date"], dtype={"артикул": str, "штрихкод": str})
    as_of = pd.Timestamp(as_of_arg) if as_of_arg else sales["date"].max()

    ostatki = read_drive_csv(service, OSTATKI_ID, usecols=["date", *KEY, "штрихкод", "артикул", "поставщик", "коллекция",
                                                             "камень", "цвет", "размер", "розничная_цена", "остаток"],
                             parse_dates=["date"],
                             # текстовые столбцы явно строками: в длинных пустых участках pandas иначе
                             # выводит float и предупреждает о смешанных типах (DtypeWarning)
                             dtype={col: str for col in [*KEY, "штрихкод", "артикул", "поставщик", "коллекция",
                                                         "камень", "цвет"]})
    snapshots = sorted(ostatki["date"].unique())

    parts = []
    # месяцы в отчете от последнего к первому; "месяц" - строка ММ-ГГГГ, сортируем по дате, а не по строке
    months = sorted(
        ((pd.Timestamp(f"{month[3:]}-{month[:2]}-01"), plan_m) for month, plan_m in plan.groupby("месяц")),
        key=lambda item: item[0],
        reverse=True,
    )
    for month_start, plan_m in months:
        if month_start > as_of:
            continue
        snapshot = pick_snapshot(snapshots, month_start)
        part = (
            plan_m.merge(calc_fact(sales, month_start, as_of), on=KEY, how="left")
            .merge(calc_stock(snapshot, ostatki), on=KEY, how="left")
        )
        part["дата_факта"] = min(month_start + pd.offsets.MonthEnd(0), as_of).date()
        part["дата_остатка"] = snapshot.date()
        parts.append(part)

    # витрина строится по плану: филиалы и категории без плана в отчет не входят
    out = pd.concat(parts, ignore_index=True)
    out["факт_mtd"] = out["факт_mtd"].fillna(0)
    out["остаток_шт"] = out["остаток_шт"].fillna(0)
    out["выполнение плана"] = out["факт_mtd"] / out["план"].where(out["план"] > 0)
    left = out["план"] - out["факт_mtd"]
    out["плана осталось"] = left.where(left >= 0)  # при перевыполнении плана пусто
    # без остатка (розничная_цена_ср пустая) средняя цена пустая
    price = out["розничная_цена_ср"] * (1 - out["бит_литера"].map(DISCOUNT))
    out["средняя цена продажи с учетом скидки"] = (price / 10).round() * 10  # округление до десятков
    price_col = out["средняя цена продажи с учетом скидки"]
    out["шт продать осталось"] = np.ceil(out["плана осталось"] / price_col.where(price_col > 0))  # вверх до целых
    need, stock = out["шт продать осталось"], out["остаток_шт"]
    enough = stock >= need * STOCK_COVER_OK  # остаток / шт продать >= 3 (при need=0 тоже верно)
    no_stock = (stock == 0) & (out["плана осталось"] > 0)  # цены нет (остаток 0), но план не выполнен
    out["возможность выполнить план"] = np.select(
        [no_stock, need.isna(), stock < need, ~enough],
        ["нет, остатков не хватает", None, "нет, остатков не хватает", "да, рекомендуется пополнить запасы"],
        default="да, достаточный запас",
    )
    out.loc[need.isna() & ~no_stock, "возможность выполнить план"] = np.nan

    # колонки плана как в csv (без "план месяц"), справа расчетные; затем переименование для листа
    out = out.rename(columns={"бит_литера": "бит литера"})
    out = out[[*PLAN_COLUMNS, "факт_mtd", "выполнение плана", "плана осталось", "остаток_шт", "средняя цена продажи с учетом скидки", "шт продать осталось", "возможность выполнить план", "дата_факта", "дата_остатка"]]
    out = out.rename(columns=REPORT_RENAME)
    as_of_name = f"{as_of:{DATE_FORMAT}}"
    folder_id = find_or_create_subfolder(service, REPORTS_FOLDER_ID, as_of_name)
    # филиалы - столбцы в порядке витрины (только филиалы с планом по категориям)
    daily_plan = read_drive_csv(service, DAILY_PLAN_ID, usecols=["date", "магазин", "план"], parse_dates=["date"])
    plan_fact = calc_plan_fact(sales, daily_plan, list(out["магазин"].unique()), as_of)
    articles = calc_articles(sales, ostatki, pick_snapshot(snapshots, as_of.replace(day=1)), as_of)
    upload_bytes(service, folder_id, REPORT_FILENAME, build_report_xlsx(out, plan_fact, articles), XLSX_MIMETYPE)
    # по файлу на каждую комбинацию филиалов из таблицы Users: в файле только филиалы комбинации
    # (лист "Артикулы" общий); филиалы без строк в витрине (нет плана) пропускаются
    present = set(out["магазин"])
    combos_done = 0
    for combo in access_combos(read_access(service)):
        filials = [filial for filial in combo if filial in present]
        if not filials:
            print(f"Нет данных в отчете по филиалам: {', '.join(combo)} - файл не сформирован")
            continue
        combo_rows = out[out["магазин"].isin(filials)]
        combo_plan_fact = plan_fact[[PLAN_FACT_LABEL, *filials]]  # столбцы по алфавиту
        upload_bytes(service, folder_id, combo_filename(combo),
                     build_report_xlsx(combo_rows, combo_plan_fact, articles, FILIAL_HIDDEN_COLUMNS), XLSX_MIMETYPE)
        combos_done += 1

    # дата в pipeline_state - только после загрузки всех файлов, чтобы бот не увидел неполную папку;
    # пересчет за прошлую дату (--as-of) не откатывает дату, которую видит бот
    state_date = get_state_date(service)
    if state_date is None or as_of >= state_date:
        set_state_date(service, as_of)
    else:
        print(f"pipeline_state не изменен: там более поздняя дата {state_date:{DATE_FORMAT}}")

    in_period = sales[(sales["date"] >= pd.Timestamp("2026-01-01")) & (sales["date"] <= as_of)]
    print(f"Папка {as_of_name} (Drive): report.xlsx {len(out)} строк, филиалов {out['магазин'].nunique()}, "
          f"месяцев {out['месяц'].nunique()}, факт по {as_of.date()}; файлов по комбинациям филиалов {combos_done}")
    print(f"Продажи вне плана (не вошли в витрину): {in_period['сумма_со_скидкой'].sum() - out['факт_mtd'].sum():,.0f}")


def run_scheduler() -> None:
    """Обновляет отчет каждые 24 часа. Ошибка в одном запуске не останавливает расписание."""
    scheduler = BlockingScheduler()
    scheduler.add_job(update_report, "interval", hours=24, next_run_time=datetime.now())  # первый запуск сразу
    logger.info("Планировщик запущен: обновление отчета сразу и далее каждые 24 часа")
    scheduler.start()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--as-of", help="дата, по которую считать MTD (по умолчанию - последняя дата продаж)")
    parser.add_argument("--schedule", action="store_true", help="обновлять отчет каждые 24 часа")
    args = parser.parse_args()
    if args.schedule:
        logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
        run_scheduler()
    else:
        update_report(args.as_of)


if __name__ == "__main__":
    main()
