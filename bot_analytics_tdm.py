"""
bot_analytics_tdm.py — бот аналитики для TDM.
Использует официальную библиотеку messenger_bot_api (messenger_bot_api-2.0.6).

Установка:
    pip install messenger_bot_api-2_0_6-py3-none-any.whl
    pip install selenium pandas openpyxl requests python-pptx pillow

Переменные окружения (.env):
    TDM_TOKEN           — токен бота
    TDM_API_URL         — https://api.tdm.mos.ru
    TDM_SSE_URL         — https://pusher.tdm.mos.ru
    TDM_FILE_URL        — https://fileupload.tdm.mos.ru
    TDM_WORKSPACE_ID    — ID пространства (обычно -1)
    login_MM, password_MM, login_NG, password_NG
"""

import logging
import os
import re
import shutil
import threading
import time
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Agg")  # headless backend - fixes tkinter thread crash

import pandas as pd
pd.set_option('future.no_silent_downcasting', True)   # подавляем FutureWarning fillna

import requests as req_lib
import pythoncom          # COM-инициализация для win32com в потоках
from dotenv import load_dotenv

# ── Официальная библиотека TDM ────────────────────────────────────────────────
from messenger_bot_api import (
    Application,
    MessageHandler,
    CommandHandler,
    ClickButtonEventHandler,
    MessageRequest,
    InlineMessageButton,
)
from messenger_bot_api.event import MessageBotEvent

# ── Модули бота ───────────────────────────────────────────────────────────────
from oati import process_file_OATI
from week_svod import parcing_data_MM_sync, process_file_MM_week
from mji_svod import parcing_MWI_sync, MWI_choosing_files, MWI_process_file, create_pivot_and_pdf
from mmonitor import parcing_data_MM_sync as mm_parse_sync, choosing_time_MM, process_file_MM
from ng_otvety import (
    choosing_time_NG, process_ng_prosroki_file, parcing_data_sync,
    personalizating_table_osn, personalizating_table_prosrok,
    personalizating_table_eight_day, personalizating_table_seven_day,
    personalizating_table_six_day, personalizating_table_five_day,
    personalizating_table_weekend, add_run_delete_and_save_files,
)
from lk_prefect import parcing_data_lk_prefekta_sync, process_lk_prefekta_file
# make_tsafap_oati и make_ng импортируются лениво внутри обработчиков,
# чтобы отсутствие selenium/pillow не валило запуск бота.

# ─────────────────────────────────────────────────────────────────────────────
load_dotenv()

TOKEN        = os.getenv("TDM_TOKEN", "")
API_URL      = os.getenv("TDM_API_URL", "https://api.tdm.mos.ru")
SSE_URL      = os.getenv("TDM_SSE_URL", "https://pusher.tdm.mos.ru")
FILE_URL     = os.getenv("TDM_FILE_URL", "https://fileupload.tdm.mos.ru")
WORKSPACE_ID = int(os.getenv("TDM_WORKSPACE_ID", "-1"))

home_dir  = os.path.expanduser("~")
directory = os.path.join(home_dir, "Downloads")

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("tdm_analytics_bot")

# ── Праздничные / нерабочие дни ───────────────────────────────────────────────
excluded_dates = [
    # Майские праздники (1-3 мая и 9-11 мая)
    "01.05.2026", "02.05.2026", "03.05.2026",
    "09.05.2026", "10.05.2026", "11.05.2026",
    # Май (добавлены только неохваченные)
    "16.05.2026", "17.05.2026", "23.05.2026", "24.05.2026", "30.05.2026", "31.05.2026",
    # Июнь (добавлены только неохваченные)
    "06.06.2026", "07.06.2026", "20.06.2026", "21.06.2026", "27.06.2026", "28.06.2026", "12.06.2026", "13.06.2026", "14.06.2026",
    # Июль
    "04.07.2026", "05.07.2026", "11.07.2026", "12.07.2026", "18.07.2026", "19.07.2026", "25.07.2026", "26.07.2026",
    # Август
    "01.08.2026", "02.08.2026", "08.08.2026", "09.08.2026", "15.08.2026", "16.08.2026", "22.08.2026", "23.08.2026",
    "29.08.2026", "30.08.2026",
    # Сентябрь
    "05.09.2026", "06.09.2026", "12.09.2026", "13.09.2026", "19.09.2026", "20.09.2026", "26.09.2026", "27.09.2026",
    # Октябрь (добавлены только неохваченные)
    "10.10.2026", "11.10.2026", "17.10.2026", "18.10.2026", "24.10.2026", "25.10.2026",
    # Ноябрь (добавлены только неохваченные)
    "14.11.2026", "15.11.2026", "21.11.2026", "22.11.2026", "28.11.2026", "29.11.2026",
    # Декабрь
    "05.12.2026", "06.12.2026", "12.12.2026", "13.12.2026", "19.12.2026", "20.12.2026", "26.12.2026", "27.12.2026",
    # Канун Нового 2027 года
    "31.12.2026"
]


# ─────────────────────────────────────────────────────────────────────────────
# Состояние пользователей
# ─────────────────────────────────────────────────────────────────────────────
_user_states: dict[int, dict] = {}

def state(group_id: int) -> dict:
    if group_id not in _user_states:
        _user_states[group_id] = {}
    return _user_states[group_id]


# ─────────────────────────────────────────────────────────────────────────────
# ГЛОБАЛЬНАЯ блокировка бота
#
# Пока какой-то сценарий выполняется, бот «занят» и принадлежит одной группе
# (_busy_owner). Все остальные группы/пользователи получают BUSY_MESSAGE и ждут.
# Владелец может продолжать свой многошаговый сценарий (ввод дат / файл) или
# отменить его командой /start.
#
# _busy_cancellable:
#   True  — бот просто ждёт ввода от пользователя (даты/файл) → /start отменяет.
#   False — идёт обработка в потоке → отменить нельзя (поток не прервать).
# ─────────────────────────────────────────────────────────────────────────────
_global_lock = threading.Lock()
_busy_owner: int | None = None
_busy_cancellable: bool = False

BUSY_MESSAGE = "⏳ Бот сейчас выполняет другой запрос. Пожалуйста, дождитесь его завершения."


def acquire_lock(group_id: int) -> bool:
    """Атомарно захватывает бота для группы. False — если уже занят кем-то."""
    global _busy_owner, _busy_cancellable
    with _global_lock:
        if _busy_owner is not None:
            return False
        _busy_owner = group_id
        _busy_cancellable = False
        return True


def release_lock():
    """Полностью освобождает бота."""
    global _busy_owner, _busy_cancellable
    with _global_lock:
        _busy_owner = None
        _busy_cancellable = False


def set_cancellable(value: bool):
    global _busy_cancellable
    with _global_lock:
        _busy_cancellable = value


def lock_owner() -> int | None:
    with _global_lock:
        return _busy_owner


def is_cancellable() -> bool:
    with _global_lock:
        return _busy_cancellable


def _reset_flow_state(st: dict):
    """Сбрасывает все флаги ожидания многошаговых сценариев."""
    st["waiting_for_dates"]       = False
    st["waiting_for_file"]        = False
    st["waiting_for_oati_file"]   = False
    st["waiting_for_oati_photo_file"]   = False
    st["waiting_for_tsafap_photo_file"] = False
    st["waiting_for_ng_file"]     = False
    st["processing_step"]         = None
    st["dates"]                   = None


# ─────────────────────────────────────────────────────────────────────────────
# Кнопки главного меню
# ─────────────────────────────────────────────────────────────────────────────
MENU_BUTTONS: list[InlineMessageButton] = [
    InlineMessageButton(id=1, label="🏢 ЛК префекта (НГ)",        callback_message="lk_prefekt"),
    InlineMessageButton(id=2, label="📊 Монитор в Работе (ММ)",    callback_message="mm_monitor"),
    InlineMessageButton(id=3, label="📈 Ответы в работе (НГ)",     callback_message="ng_answers"),
    InlineMessageButton(id=4, label="📄 Свод МЖИ (НГ)",            callback_message="mji_svod"),
    InlineMessageButton(id=5, label="📎 Еженедельный свод",         callback_message="week_svod"),
    InlineMessageButton(id=6, label="🅾️ Слайд ОАТИ",              callback_message="oati"),
    InlineMessageButton(id=7, label="📸 Фото ОАТИ",               callback_message="oati_photo"),
    InlineMessageButton(id=8, label="📸 Фото ЦАФАП",              callback_message="tsafap_photo"),
    InlineMessageButton(id=9, label="📷 Фото НГ",                  callback_message="ng_photo"),
    InlineMessageButton(id=10, label="❓ Объяснение команд",        callback_message="explain"),
]

# BBCode-разметка поддерживается библиотекой через TextFormatter
EXPLANATION_TEXT = (
    "📋 [b]Объяснение команд:[/b]\n\n"
    "🏢 [b]ЛК префекта (НГ)[/b] — отчёт по заявкам ЛК Префекта (все районы)\n\n"
    "📊 [b]Монитор в Работе (ММ)[/b] — сводный отчёт Монитора ММ\n\n"
    "📈 [b]Ответы в работе (НГ)[/b] — отчёт с просрочками по дням\n\n"
    "📄 [b]Свод МЖИ (НГ)[/b] — актуальные данные по заявкам МЖИ\n\n"
    "📎 [b]Еженедельный свод[/b] — свод с Монитора Мэра для презентаций\n\n"
    "🅾️ [b]Слайд ОАТИ[/b] — создание презентационного слайда ОАТИ\n\n"
    "📸 [b]Фото ОАТИ[/b] — презентация с фото по выгрузке ОАТИ\n\n"
    "📸 [b]Фото ЦАФАП[/b] — презентация с фото по выгрузке ЦАФАП\n\n"
    "📷 [b]Фото НГ[/b] — презентация с фото по выгрузке НГ"
)


# ─────────────────────────────────────────────────────────────────────────────
# Вспомогательные обёртки над API библиотеки
# ─────────────────────────────────────────────────────────────────────────────

def reply_menu(event: MessageBotEvent, text: str = "Выберите команду:"):
    event.reply_text_message(MessageRequest(text, buttons=MENU_BUTTONS))


def reply_loading(event: MessageBotEvent, text: str) -> int | None:
    resp = event.reply_text(text)
    return (resp or {}).get("messageId")


def update_loading(event: MessageBotEvent, msg_id: int | None, text: str):
    if not msg_id:
        return
    try:
        event._request.update_text_message(
            event.workspace_id, event.group_id, msg_id, text
        )
    except Exception as e:
        logger.warning(f"update_loading: {e}")


def delete_msg(event: MessageBotEvent, msg_id: int | None):
    if msg_id:
        try:
            event.delete_messages_from_current_group([msg_id])
        except Exception as e:
            logger.warning(f"delete_msg: {e}")


def send_file_safe(event: MessageBotEvent, path: str, caption: str = "",
                   retries: int = 3, delay: int = 8) -> bool:
    """
    Отправляет файл с повторными попытками.
    TDM иногда отвечает 504 (таймаут шлюза) на крупных файлах — повтор обычно помогает.
    """
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            event.reply_file(path, caption)
            if attempt > 1:
                logger.info(f"Файл отправлен со {attempt}-й попытки: {os.path.basename(path)}")
            return True
        except Exception as e:
            last_err = e
            logger.warning(
                f"send_file_safe попытка {attempt}/{retries} "
                f"[{os.path.basename(path)}]: {e}"
            )
            if attempt < retries:
                time.sleep(delay)

    logger.error(f"send_file_safe не удалось [{os.path.basename(path)}]: {last_err}")
    try:
        event.reply_text(
            f"❌ Не удалось отправить «{os.path.basename(path)}» — сервер вернул ошибку "
            f"(возможно, файл слишком большой). Попробуйте сформировать ещё раз."
        )
    except Exception:
        pass
    return False


def _cleanup_run(pptx_path: str | None, uploaded_path: str | None):
    """
    Удаляет всё временное после генерации презентации:
      * сам .pptx и весь временный каталог запуска (родитель .pptx),
      * загруженный пользователем Excel-файл.
    """
    try:
        if pptx_path and os.path.exists(pptx_path):
            run_dir = os.path.dirname(pptx_path)
            shutil.rmtree(run_dir, ignore_errors=True)
            logger.info(f"Удалён временный каталог: {run_dir}")
    except Exception as e:
        logger.warning(f"cleanup run dir: {e}")

    try:
        if uploaded_path and os.path.exists(uploaded_path):
            os.remove(uploaded_path)
            logger.info(f"Удалён загруженный файл: {uploaded_path}")
    except Exception as e:
        logger.warning(f"cleanup uploaded: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# Утилиты
# ─────────────────────────────────────────────────────────────────────────────

def choosing_time_frame_MM() -> tuple[str, str]:
    today         = datetime.now()
    weekday       = today.weekday()
    start_of_week = today - timedelta(days=weekday)
    offset        = weekday if weekday <= 5 else 5
    start_day     = (today - timedelta(days=1)
                     if weekday == 0
                     else start_of_week + timedelta(days=(weekday - offset)))
    return (
        start_day.strftime("%d%m%Y") + "2100",
        today.strftime("%d%m%Y")     + "2100",
    )


def latest_xlsx() -> str | None:
    files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]
    if not files:
        return None
    files.sort(key=lambda x: os.path.getmtime(os.path.join(directory, x)))
    return os.path.join(directory, files[-1])


# ─────────────────────────────────────────────────────────────────────────────
# Запуск задач в потоках (с автоснятием глобальной блокировки)
# ─────────────────────────────────────────────────────────────────────────────

def run_final_job(event: MessageBotEvent, fn, *args):
    """Терминальная задача: по завершении (успех/ошибка) блокировка снимается."""
    def wrapped():
        pythoncom.CoInitialize()
        try:
            fn(event, *args)
        except Exception as e:
            logger.error(f"{getattr(fn, '__name__', fn)} упал: {e}", exc_info=True)
            try:
                reply_menu(event, f"❌ Внутренняя ошибка: {str(e)[:100]}\n\nВыберите команду:")
            except Exception:
                pass
        finally:
            pythoncom.CoUninitialize()
            release_lock()
    threading.Thread(target=wrapped, daemon=True).start()


def run_step_job(event: MessageBotEvent, fn, *args):
    """Промежуточный шаг: при краше снимает блокировку, при успехе — НЕТ."""
    def wrapped():
        pythoncom.CoInitialize()
        try:
            fn(event, *args)
        except Exception as e:
            logger.error(f"{getattr(fn, '__name__', fn)} упал: {e}", exc_info=True)
            try:
                reply_menu(event, f"❌ Внутренняя ошибка: {str(e)[:100]}\n\nВыберите команду:")
            except Exception:
                pass
            release_lock()
        finally:
            pythoncom.CoUninitialize()
    threading.Thread(target=wrapped, daemon=True).start()


def _extract_file_info(event: MessageBotEvent) -> dict:
    """Извлекает информацию о файле из сырого payload события."""
    payload = event.get_payload_data() or {}

    messages = payload.get("messages") or []
    if messages:
        msg = messages[0]
        logger.debug(f"Payload messages[0] keys: {list(msg.keys())}")

        file_info = msg.get("file") or {}
        if file_info.get("resourceRef") is not None:
            logger.info(f"Файл найден в messages[0].file: {file_info.get('fileName')}")
            return file_info

        for item in (msg.get("multimedia") or []):
            fi = item.get("file") or {}
            if fi.get("resourceRef") is not None:
                logger.info(f"Файл найден в messages[0].multimedia[].file: {fi.get('fileName')}")
                return fi

        media_file = (msg.get("media") or {}).get("file") or {}
        if media_file.get("resourceRef") is not None:
            logger.info(f"Файл найден в messages[0].media.file: {media_file.get('fileName')}")
            return media_file

    root_file = payload.get("file") or {}
    if root_file.get("resourceRef") is not None:
        logger.info(f"Файл найден в payload.file: {root_file.get('fileName')}")
        return root_file

    logger.warning(f"Файл не найден в payload. Ключи: {list(payload.keys())}")
    if messages:
        logger.warning(f"messages[0] полностью: {messages[0]}")
    return {}


def _has_file(event: MessageBotEvent) -> bool:
    info = _extract_file_info(event)
    return bool(info and info.get("resourceRef") is not None)


def _download_incoming_file(event: MessageBotEvent) -> tuple[str | None, str]:
    """Скачивает прикреплённый зашифрованный файл через FILE_URL."""
    file_info    = _extract_file_info(event)
    resource_ref = file_info.get("resourceRef")
    filename     = file_info.get("fileName", "incoming.xlsx")

    if not resource_ref:
        logger.error("resourceRef не найден, скачивание невозможно")
        return None, ""

    if isinstance(resource_ref, str):
        logger.warning(f"resourceRef пришёл как строка: {resource_ref}")
        resource_ref = {"fileId": resource_ref}

    logger.info(f"Скачиваю файл '{filename}', resourceRef: {resource_ref}")

    dl_url  = f"{FILE_URL}/api/v1/download/secret/decryptable"
    headers = {"Authorization": TOKEN}

    try:
        r = req_lib.post(
            dl_url,
            json={"resourceRef": resource_ref},
            headers=headers,
            timeout=120,
        )
        r.raise_for_status()
    except Exception as e:
        logger.error(f"Ошибка скачивания файла: {e}")
        return None, filename

    content = r.content

    if content[:2] != b'PK':
        logger.error(
            f"Файл не является валидным xlsx! "
            f"Первые 16 байт: {content[:16].hex()} (ожидалось 504B...)"
        )
        logger.error(f"HTTP статус: {r.status_code}, Content-Type: {r.headers.get('Content-Type')}")
        try:
            logger.error(f"Тело ответа (текст): {r.text[:500]}")
        except Exception:
            pass
        return None, filename

    temp_dir  = os.path.join(directory, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    temp_path = os.path.join(temp_dir, filename)

    with open(temp_path, "wb") as f:
        f.write(content)

    logger.info(f"✅ Файл успешно сохранён: {temp_path} ({len(content)} байт)")
    return temp_path, filename


# ─────────────────────────────────────────────────────────────────────────────
# Обработчики команд (запускаются в отдельных потоках)
# ─────────────────────────────────────────────────────────────────────────────

def _do_lk_prefekt(event: MessageBotEvent):
    msg_id = reply_loading(event, "🏢 Выгружаю ЛК Префекта...\n📥 Подключаюсь к порталу...")

    success = parcing_data_lk_prefekta_sync(
        on_error=lambda t: update_loading(event, msg_id, f"❌ {t}")
    )
    if not success:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Ошибка выгрузки ЛК Префекта.\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "🏢 ЛК Префекта...\n⚙️ Обрабатываю данные...")

    filepath = latest_xlsx()
    if not filepath:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Excel-файл не найден.\n\nВыберите команду:")
        return

    try:
        processed = process_lk_prefekta_file(directory, "Все районы", filepath)
    except Exception as e:
        logger.error(f"process_lk_prefekta_file: {e}")
        processed = None

    if not processed:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Нет данных или ошибка обработки.\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "🏢 ЛК Префекта...\n📤 Отправляю файл...")
    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, processed, f"🏢 ЛК Префекта (все районы) на {ts}")
    reply_menu(event, "✅ Отчёт ЛК Префекта отправлен!\n\nВыберите следующую команду:")


def _do_mm_monitor(event: MessageBotEvent):
    msg_id = reply_loading(event, "📊 Выгружаю Монитор в Работе...\n📥 Подключаюсь к порталу...")

    start_date, end_date = choosing_time_frame_MM()
    success = mm_parse_sync(
        start_date, end_date,
        on_error=lambda t: update_loading(event, msg_id, f"❌ {t}")
    )
    if not success:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Ошибка выгрузки ММ.\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "📊 Монитор в Работе...\n⚙️ Обрабатываю данные...")

    filepath = latest_xlsx()
    if not filepath:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Excel-файл не найден.\n\nВыберите команду:")
        return

    try:
        timenow = choosing_time_MM()
        excel_path, pdf_path = process_file_MM(filepath, timenow)
    except Exception as e:
        logger.error(f"process_file_MM: {e}")
        delete_msg(event, msg_id)
        reply_menu(event, f"❌ Ошибка обработки: {str(e)[:100]}\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "📊 Монитор в Работе...\n📤 Отправляю файлы...")
    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    if pdf_path and os.path.exists(pdf_path):
        send_file_safe(event, pdf_path,   f"📊 Монитор в Работе (PDF) на {ts}")
    send_file_safe(event, excel_path, f"📋 Монитор в Работе (Excel) на {ts}")
    reply_menu(event, "✅ Монитор в Работе отправлен!\n\nВыберите следующую команду:")


def _do_ng_answers(event: MessageBotEvent):
    msg_id = reply_loading(event, "📈 Выгружаю Ответы в работе...\n📥 Подключаюсь к порталу...")

    success = parcing_data_sync(
        on_error=lambda t: update_loading(event, msg_id, f"❌ {t}")
    )
    if not success:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Ошибка выгрузки НГ.\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "📈 Ответы в работе...\n⚙️ Обрабатываю данные...")

    filepath = latest_xlsx()
    if not filepath:
        delete_msg(event, msg_id)
        reply_menu(event, "❌ Excel-файл не найден.\n\nВыберите команду:")
        return

    try:
        timenow = choosing_time_NG()
        process_ng_prosroki_file(timenow, filepath, excluded_dates)

        update_loading(event, msg_id, "📈 Ответы в работе...\n🎨 Форматирую таблицы...")
        personalizating_table_osn(timenow)
        personalizating_table_prosrok(timenow)
        personalizating_table_eight_day(timenow)
        personalizating_table_seven_day(timenow)
        personalizating_table_six_day(timenow)
        personalizating_table_five_day(timenow)
        personalizating_table_weekend(timenow)

        update_loading(event, msg_id, "📈 Ответы в работе...\n📄 Создаю PDF...")
        pdf_path, first_sheet_path, full_path = add_run_delete_and_save_files(timenow)

    except Exception as e:
        logger.error(f"ng_answers: {e}")
        delete_msg(event, msg_id)
        reply_menu(event, f"❌ Ошибка: {str(e)[:150]}\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "📈 Ответы в работе...\n📤 Отправляю файлы...")
    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, pdf_path,         f"📊 Ответы в работе (PDF) на {ts}")
    send_file_safe(event, first_sheet_path, f"📋 СВОД (Excel) на {ts}")
    if os.path.exists(full_path):
        send_file_safe(event, full_path, f"📁 Детальные данные на {ts}")
    reply_menu(event, "✅ Ответы в работе отправлены!\n\nВыберите следующую команду:")


def _do_mji_svod(event: MessageBotEvent):
    msg_id = reply_loading(event, "📄 Выгружаю Свод МЖИ...\n📥 Подключаюсь к порталу...")

    count = parcing_MWI_sync(
        on_error=lambda t: update_loading(event, msg_id, f"❌ {t}")
    )

    update_loading(event, msg_id, "📄 Свод МЖИ...\n⚙️ Обрабатываю данные...")

    try:
        df         = MWI_process_file(MWI_choosing_files(directory, count))
        timenow    = datetime.now().strftime("%H-%M")
        excel_file = os.path.join(
            directory,
            f"СВОД МЖИ {datetime.now().strftime('%d.%m.%y')} на {timenow}.xlsx"
        )
        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="МЖИ", index=False)

        update_loading(event, msg_id, "📄 Свод МЖИ...\n📊 Создаю сводную таблицу...")
        pdf_path, ok, _ = create_pivot_and_pdf(excel_file, directory)

    except Exception as e:
        logger.error(f"mji_svod: {e}")
        delete_msg(event, msg_id)
        reply_menu(event, f"❌ Ошибка: {str(e)[:150]}\n\nВыберите команду:")
        return

    update_loading(event, msg_id, "📄 Свод МЖИ...\n📤 Отправляю файлы...")
    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, excel_file, f"📊 Свод МЖИ (Excel) на {ts}")
    if pdf_path and os.path.exists(pdf_path):
        send_file_safe(event, pdf_path, f"📄 Свод МЖИ (PDF) на {ts}")
    reply_menu(event, "✅ Свод МЖИ отправлен!\n\nВыберите следующую команду:")


def _do_week_svod_parse(event: MessageBotEvent, date1: str, date2: str, msg_id: int | None):
    """Шаг 2 — парсинг портала. При УСПЕХЕ блокировку НЕ снимаем (ждём файл)."""
    success = parcing_data_MM_sync(
        date1 + "2100", date2 + "2059",
        on_error=lambda t: update_loading(event, msg_id, f"❌ {t}")
    )
    if not success:
        delete_msg(event, msg_id)
        _reset_flow_state(state(event.group_id))
        release_lock()
        reply_menu(event, "❌ Ошибка выгрузки данных.\n\nВыберите команду:")
        return

    update_loading(
        event, msg_id,
        f"✅ Данные за {date1}–{date2} выгружены!\n\n"
        "📤 Отправьте городскую выгрузку (Excel-файл).\n\n"
        "Для отмены отправьте /start"
    )
    st = state(event.group_id)
    st["waiting_for_file"]   = True
    st["processing_step"]    = "first_file"
    st["instruction_msg_id"] = msg_id
    set_cancellable(True)


def _do_week_svod_process(event: MessageBotEvent, user_file_path: str):
    """Шаг 3 — обрабатываем пользовательский файл + файл с портала."""
    st           = state(event.group_id)
    date1, date2 = st.get("dates") or ("", "")

    msg_id = reply_loading(event, "⚙️ Обрабатываю еженедельный свод...")

    portal_files = sorted(
        [
            os.path.join(directory, f)
            for f in os.listdir(directory)
            if f.endswith(".xlsx") and os.path.isfile(os.path.join(directory, f))
        ],
        key=os.path.getmtime,
        reverse=True,
    )

    if not portal_files:
        delete_msg(event, msg_id)
        _reset_flow_state(st)
        reply_menu(event, "❌ Файл с портала не найден в Downloads.\n\nВыберите команду:")
        return

    downloaded = portal_files[0]
    logger.info(f"Файл с портала: {downloaded}")
    logger.info(f"Пользовательский файл: {user_file_path}")

    try:
        output_path = process_file_MM_week(user_file_path, downloaded)
    except Exception as e:
        logger.error(f"process_file_MM_week: {e}", exc_info=True)
        delete_msg(event, msg_id)
        _reset_flow_state(st)
        reply_menu(event, f"❌ Ошибка обработки: {str(e)[:100]}\n\nВыберите команду:")
        return

    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, output_path, f"📎 Еженедельный свод за {date1}–{date2} (выгрузка: {ts})")
    reply_menu(event, "✅ Еженедельный свод отправлен!\n\nВыберите следующую команду:")

    _reset_flow_state(st)
    try:
        os.remove(user_file_path)
    except Exception:
        pass


def _do_oati(event: MessageBotEvent, file_path: str):
    """Слайд ОАТИ."""
    msg_id = reply_loading(event, "🅾️ Обрабатываю файл ОАТИ...")

    try:
        ppt_path, stats_text = process_file_OATI(file_path)
    except Exception as e:
        logger.error(f"process_file_OATI: {e}")
        delete_msg(event, msg_id)
        _reset_flow_state(state(event.group_id))
        reply_menu(event, f"❌ Ошибка ОАТИ: {str(e)[:150]}\n\nВыберите команду:")
        return

    delete_msg(event, msg_id)

    send_file_safe(event, ppt_path, "🅾️ Слайд ОАТИ")
    event.reply_text_message(
        MessageRequest(f"[b]Статистика ОАТИ:[/b]\n{stats_text}")
    )
    reply_menu(event, "✅ Слайд ОАТИ готов!\n\nВыберите следующую команду:")

    _reset_flow_state(state(event.group_id))
    try:
        os.remove(file_path)
    except Exception:
        pass


def _do_photo_presentation(event: MessageBotEvent, file_path: str, system: str):
    """Общий обработчик фото-презентаций ОАТИ/ЦАФАП. По завершении чистит всё временное."""
    from make_tsafap_oati import process_tsafap_oati_file

    msg_id = reply_loading(
        event,
        f"📸 Создаю презентацию {system}...\n"
        "📥 Скачиваю фото с портала, это может занять несколько минут..."
    )

    pptx_path = None
    try:
        pptx_path = process_tsafap_oati_file(file_path, system=system)
    except Exception as e:
        logger.error(f"process_tsafap_oati_file [{system}]: {e}", exc_info=True)
        delete_msg(event, msg_id)
        _reset_flow_state(state(event.group_id))
        reply_menu(event, f"❌ Ошибка Фото {system}: {str(e)[:150]}\n\nВыберите команду:")
        _cleanup_run(pptx_path, file_path)
        return

    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, pptx_path, f"📸 Фото {system} на {ts}")
    reply_menu(event, f"✅ Презентация {system} отправлена!\n\nВыберите следующую команду:")

    _reset_flow_state(state(event.group_id))
    _cleanup_run(pptx_path, file_path)


def _do_oati_photo(event: MessageBotEvent, file_path: str):
    _do_photo_presentation(event, file_path, system="ОАТИ")


def _do_tsafap_photo(event: MessageBotEvent, file_path: str):
    _do_photo_presentation(event, file_path, system="ЦАФАП")


def _do_ng_photo(event: MessageBotEvent, file_path: str):
    """Фото НГ — презентация по выгрузке. По завершении чистит всё временное."""
    from make_ng import process_ng_file

    msg_id = reply_loading(
        event,
        "📷 Создаю презентацию НГ...\n"
        "📥 Скачиваю фото, это может занять несколько минут..."
    )

    pptx_path = None
    try:
        pptx_path = process_ng_file(file_path)
    except Exception as e:
        logger.error(f"process_ng_file: {e}", exc_info=True)
        delete_msg(event, msg_id)
        _reset_flow_state(state(event.group_id))
        reply_menu(event, f"❌ Ошибка Фото НГ: {str(e)[:150]}\n\nВыберите команду:")
        _cleanup_run(pptx_path, file_path)
        return

    delete_msg(event, msg_id)

    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    send_file_safe(event, pptx_path, f"📷 Фото НГ на {ts}")
    reply_menu(event, "✅ Презентация НГ отправлена!\n\nВыберите следующую команду:")

    _reset_flow_state(state(event.group_id))
    _cleanup_run(pptx_path, file_path)


# ─────────────────────────────────────────────────────────────────────────────
# Старт сценариев (вызывается, когда бот СВОБОДЕН)
# ─────────────────────────────────────────────────────────────────────────────

def _begin(event: MessageBotEvent) -> bool:
    """Пытается захватить бота. Если занят — отвечает «подождите»."""
    if acquire_lock(event.group_id):
        return True
    event.reply_text(BUSY_MESSAGE)
    return False


def _start_heavy(event: MessageBotEvent, fn, *args):
    """Однокнопочный тяжёлый сценарий (lk/mm/ng/mji)."""
    if not _begin(event):
        return
    set_cancellable(False)
    run_final_job(event, fn, *args)


def _start_week(event: MessageBotEvent):
    """Еженедельный свод — шаг 1: спрашиваем даты."""
    if not _begin(event):
        return
    st = state(event.group_id)
    _reset_flow_state(st)
    st["waiting_for_dates"] = True
    set_cancellable(True)
    event.reply_text(
        "📎 Еженедельный свод\n\n"
        "Введите две даты через пробел в формате дд.мм.гггг:\n"
        "Пример: 01.01.2025 07.01.2025\n\n"
        "Для отмены отправьте /start"
    )


def _start_file_scenario(event: MessageBotEvent, state_key: str, prompt: str):
    """Универсальный старт сценария, ожидающего Excel-файл
    (ОАТИ / Фото ОАТИ-ЦАФАП / Фото НГ)."""
    if not _begin(event):
        return
    st = state(event.group_id)
    _reset_flow_state(st)
    st[state_key] = True
    set_cancellable(True)
    event.reply_text(prompt)


# ─────────────────────────────────────────────────────────────────────────────
# Роутер входящих сообщений
# ─────────────────────────────────────────────────────────────────────────────
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")

# Какие флаги означают «ждём файл» и какой обработчик запускать
FILE_FLOWS = {
    "waiting_for_oati_file":         _do_oati,
    "waiting_for_oati_photo_file":   _do_oati_photo,
    "waiting_for_tsafap_photo_file": _do_tsafap_photo,
    "waiting_for_ng_file":           _do_ng_photo,
}


def on_message(event: MessageBotEvent):
    """Единственный обработчик для MessageHandler и ClickButtonEventHandler."""
    st              = state(event.group_id)
    text            = (event.message_text or "").strip()
    has_file        = _has_file(event)
    is_button_click = getattr(event, "has_selected_button", False)

    owner = lock_owner()

    # ══════════════════════════════════════════════════════════════════════
    # 1) Бот занят ДРУГОЙ группой — абсолютно все ждут
    # ══════════════════════════════════════════════════════════════════════
    if owner is not None and owner != event.group_id:
        if is_button_click or has_file:
            event.reply_text(BUSY_MESSAGE)
        else:
            logger.info(f"Бот занят (владелец {owner}), игнор от {event.group_id}: {text[:30]!r}")
        return

    # ══════════════════════════════════════════════════════════════════════
    # 2) Бот занят НАМИ — продолжаем наш многошаговый сценарий
    # ══════════════════════════════════════════════════════════════════════
    if owner == event.group_id:

        # ── Ожидаем файл ──
        if has_file:
            file_info = _extract_file_info(event)
            fname     = file_info.get("fileName", "").lower()

            active_file_flow = next((k for k in FILE_FLOWS if st.get(k)), None)
            week_file = st.get("waiting_for_file") and st.get("processing_step") == "first_file"

            if active_file_flow or week_file:
                if not (fname.endswith(".xlsx") or fname.endswith(".xls")):
                    event.reply_text("❌ Пожалуйста, отправьте файл Excel (.xlsx или .xls)")
                    return

                set_cancellable(False)
                file_path, _ = _download_incoming_file(event)

                if not file_path:
                    _reset_flow_state(st)
                    release_lock()
                    event.reply_text("❌ Не удалось скачать файл. Попробуйте команду заново.")
                    reply_menu(event)
                    return

                if active_file_flow:
                    handler = FILE_FLOWS[active_file_flow]
                    st[active_file_flow] = False
                    run_final_job(event, handler, file_path)
                else:
                    st["waiting_for_file"] = False
                    run_final_job(event, _do_week_svod_process, file_path)
                return

            event.reply_text("⏳ Идёт обработка предыдущего запроса, дождитесь завершения.")
            return

        # ── Ожидаем даты для еженедельного свода ──
        if st.get("waiting_for_dates"):
            parts = text.split()
            if len(parts) != 2 or not DATE_RE.match(parts[0]) or not DATE_RE.match(parts[1]):
                event.reply_text(
                    "❌ Введите две даты через пробел в формате дд.мм.гггг\n"
                    "Пример: 01.01.2025 07.01.2025\n\n"
                    "Для отмены отправьте /start"
                )
                return
            try:
                datetime.strptime(parts[0], "%d.%m.%Y")
                datetime.strptime(parts[1], "%d.%m.%Y")
            except ValueError:
                event.reply_text("❌ Некорректные даты. Проверьте ввод.\n\nДля отмены — /start")
                return

            st["waiting_for_dates"] = False
            st["dates"]             = (parts[0], parts[1])
            set_cancellable(False)
            msg_id = reply_loading(
                event, f"⏳ Выгружаю данные за {parts[0]}–{parts[1]}...\nПодождите 1–2 минуты."
            )
            run_step_job(event, _do_week_svod_parse, parts[0], parts[1], msg_id)
            return

        # ── Владелец прислал что-то постороннее во время обработки ──
        event.reply_text(
            "⏳ Сейчас выполняется ваш предыдущий запрос.\n"
            "Дождитесь ответа или отправьте /start для отмены."
        )
        return

    # ══════════════════════════════════════════════════════════════════════
    # 3) Бот СВОБОДЕН (owner is None)
    # ══════════════════════════════════════════════════════════════════════

    if has_file:
        event.reply_text("⚠️ Файл получен, но я не ожидаю файл сейчас.")
        reply_menu(event)
        return

    cmd_map = {
        "lk_prefekt": lambda: _start_heavy(event, _do_lk_prefekt),
        "mm_monitor": lambda: _start_heavy(event, _do_mm_monitor),
        "ng_answers": lambda: _start_heavy(event, _do_ng_answers),
        "mji_svod":   lambda: _start_heavy(event, _do_mji_svod),
        "week_svod":  lambda: _start_week(event),
        "oati":       lambda: _start_file_scenario(
            event, "waiting_for_oati_file",
            "🅾️ Слайд ОАТИ\n\nОтправьте Excel-файл с выгрузкой ОАТИ.\n\nДля отмены отправьте /start"
        ),
        "oati_photo": lambda: _start_file_scenario(
            event, "waiting_for_oati_photo_file",
            "📸 Фото ОАТИ\n\nОтправьте Excel-выгрузку ОАТИ.\n\nДля отмены отправьте /start"
        ),
        "tsafap_photo": lambda: _start_file_scenario(
            event, "waiting_for_tsafap_photo_file",
            "📸 Фото ЦАФАП\n\nОтправьте Excel-выгрузку ЦАФАП.\n\nДля отмены отправьте /start"
        ),
        "ng_photo": lambda: _start_file_scenario(
            event, "waiting_for_ng_file",
            "📷 Фото НГ\n\nОтправьте Excel-файл с выгрузкой НГ.\n\nДля отмены отправьте /start"
        ),
        "explain":    lambda: event.reply_text_message(
            MessageRequest(EXPLANATION_TEXT, buttons=MENU_BUTTONS)
        ),
    }

    if text in cmd_map:
        cmd_map[text]()
        return

    if is_button_click:
        logger.warning(f"Неизвестная команда от кнопки: {text}")
        reply_menu(event, "Выберите команду из меню:")
        return

    logger.info(f"Игнорирую обычное сообщение: {text[:50]!r}")


def on_start(event: MessageBotEvent):
    """Обработчик /start. Также служит командой ОТМЕНЫ ввода дат/файла."""
    st    = state(event.group_id)
    owner = lock_owner()

    if owner == event.group_id:
        if is_cancellable():
            delete_msg(event, st.get("instruction_msg_id"))
            _reset_flow_state(st)
            release_lock()
            reply_menu(event, "🚫 Текущий сценарий отменён.\n\nВыберите команду:")
        else:
            event.reply_text("⏳ Идёт обработка, дождитесь её завершения.")
        return

    if owner is not None:
        event.reply_text(BUSY_MESSAGE)
        return

    _reset_flow_state(st)
    reply_menu(event, "👋 Привет! Выберите команду:")


# ─────────────────────────────────────────────────────────────────────────────
# Точка входа
# ─────────────────────────────────────────────────────────────────────────────

def main():
    if not TOKEN:
        logger.error("TDM_TOKEN не задан в .env")
        return

    logger.info("Запуск TDM бота аналитики...")

    app = Application(
        token=TOKEN,
        request_kwargs={
            "api_base_url":         API_URL,
            "sse_base_url":         SSE_URL,
            "file_upload_base_url": FILE_URL,
            "workspace_id":         WORKSPACE_ID,
        },
    )

    app.add_handler(CommandHandler("start", on_start))
    app.add_handler(ClickButtonEventHandler(on_message))
    app.add_handler(MessageHandler(on_message))

    logger.info("🤖 TDM бот запущен. Ожидаю сообщения...")
    app.start()

    try:
        threading.Event().wait()
    except KeyboardInterrupt:
        logger.info("Остановка по Ctrl+C...")
        app.stop()


if __name__ == "__main__":
    main()
