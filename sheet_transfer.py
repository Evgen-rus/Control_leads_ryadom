# -*- coding: utf-8 -*-

"""
Скрипт для синхронизации данных между Google-таблицами.

Переносит только новые строки из источника в приёмник на основе уникальности 
по столбцу D (Телефон Лида). Копирует данные из столбцов A-G.

Требования:
  - python-dotenv для загрузки .env файла
  - google-api-python-client google-auth-httplib2 google-auth-oauthlib
  - Файл .env с настройками
  - Файл credentials.json для Google Sheets API
"""

import os
import logging
from typing import List, Set, Optional
from datetime import datetime, timedelta
import pytz
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# Загружаем переменные из .env файла
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    logging.warning("python-dotenv не установлен. Используем системные переменные окружения.")

# Константы
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
PHONE_COLUMN_INDEX = 3  # Столбец D (Телефон Лида) - индекс 3
DATE_COLUMN_INDEX = 0   # Столбец A (Дата Лида) - индекс 0
MAX_COLUMNS = 7  # Копируем столбцы A-G (индексы 0-6)

# Московский часовой пояс
MOSCOW_TZ = pytz.timezone('Europe/Moscow')

# Настройки из переменных окружения
SRC_ID = os.getenv("SRC_ID")
DST_ID = os.getenv("DST_ID")
SRC_SHEET = os.getenv("SRC_SHEET")
DST_SHEET = os.getenv("DST_SHEET")

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)


def create_sheets_service():
    """
    Создаёт сервис для работы с Google Sheets API с оптимизированными настройками.
    
    Returns:
        googleapiclient.discovery.Resource: Сервис Google Sheets
        
    Raises:
        RuntimeError: Если не найден файл credentials или переменная окружения
        Exception: При ошибках авторизации
    """
    try:
        creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE")
        if not creds_file:
            raise RuntimeError("Не задана переменная среды GOOGLE_CREDENTIALS_FILE")
        
        # Преобразуем относительный путь в абсолютный от корня проекта
        if not os.path.isabs(creds_file):
            project_root = os.path.dirname(os.path.abspath(__file__))
            creds_file = os.path.join(project_root, creds_file)
        
        if not os.path.exists(creds_file):
            raise RuntimeError(f"Файл credentials не найден: {creds_file}")
        
        logger.info(f"Загружаем credentials из: {creds_file}")
        creds = service_account.Credentials.from_service_account_file(
            creds_file, scopes=SCOPES
        )
        
        # Создаём сервис с отключённым кэшем discovery для лучшей производительности
        service = build("sheets", "v4", credentials=creds, cache_discovery=False)
        logger.info("Сервис Google Sheets успешно создан")
        return service
        
    except Exception as e:
        logger.error(f"Ошибка создания сервиса Google Sheets: {e}")
        raise


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=1, max=10),
    retry=retry_if_exception_type((HttpError, ConnectionError, TimeoutError)),
    before_sleep=lambda retry_state: logger.info(f"Повторная попытка через {retry_state.next_action.sleep} секунд...")
)
def get_sheet_data(service, spreadsheet_id: str, sheet_name: str) -> List[List[str]]:
    """
    Получает данные из указанного листа Google Таблицы.
    Автоматически повторяет попытки при временных ошибках.
    
    Args:
        service: Сервис Google Sheets
        spreadsheet_id: ID таблицы
        sheet_name: Название листа
        
    Returns:
        List[List[str]]: Список строк с данными
        
    Raises:
        HttpError: При критических ошибках API Google Sheets
        ConnectionError: При проблемах с сетью
        TimeoutError: При таймаутах
    """
    logger.info(f"Читаем данные из листа '{sheet_name}' (ID: {spreadsheet_id[:10]}...)")
    
    response = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()
    
    rows = response.get("values", [])
    logger.info(f"Получено {len(rows)} строк из листа '{sheet_name}'")
    return rows


def normalize_row(row: List[str], max_columns: int) -> List[str]:
    """
    Нормализует строку до указанного количества столбцов.
    Добавляет пустые ячейки если строка короче, обрезает если длиннее.
    
    Args:
        row: Исходная строка
        max_columns: Максимальное количество столбцов
        
    Returns:
        List[str]: Нормализованная строка
    """
    normalized = row[:max_columns]  # Обрезаем до max_columns
    while len(normalized) < max_columns:  # Дополняем пустыми ячейками
        normalized.append("")
    return normalized


def parse_date(date_str: str) -> Optional[datetime]:
    """
    Парсит дату из строки в формате YYYY-MM-DD.
    
    Args:
        date_str: Строка с датой
        
    Returns:
        datetime или None если дата не распознана
    """
    if not date_str or not date_str.strip():
        return None
    
    try:
        # Пробуем разные форматы даты
        formats = [
            "%Y-%m-%d",      # 2025-07-15
            "%d.%m.%Y",      # 15.07.2025
            "%d/%m/%Y",      # 15/07/2025
            "%Y.%m.%d",      # 2025.07.15
        ]
        
        date_str = date_str.strip()
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        logger.debug(f"Не удалось распознать дату: '{date_str}'")
        return None
        
    except Exception as e:
        logger.debug(f"Ошибка при парсинге даты '{date_str}': {e}")
        return None


def normalize_phone_number(phone: str) -> str:
    """
    Нормализует номер телефона, приводя к формату 79123456789.
    Убирает все нечисловые символы, обрабатывает различные форматы.
    
    Args:
        phone: Исходный номер телефона в любом формате
        
    Returns:
        str: Нормализованный номер в формате 79123456789 или пустая строка если невалидный
        
    Examples:
        '+7 (912) 345-67-89' -> '79123456789'
        '8-912-345-67-89' -> '79123456789'
        '+79123456789' -> '79123456789'
        '89123456789' -> '79123456789'
    """
    if not phone or not isinstance(phone, str):
        return ""
    
    # Убираем все нечисловые символы
    digits_only = ''.join(filter(str.isdigit, phone))
    
    # Проверяем минимальную длину (должно быть как минимум 10 цифр)
    if len(digits_only) < 10:
        logger.debug(f"Слишком короткий номер после очистки: '{phone}' -> '{digits_only}'")
        return ""
    
    # Обрабатываем различные форматы российских номеров
    if len(digits_only) == 11:
        # 79123456789 или 89123456789
        if digits_only.startswith('7'):
            return digits_only  # Уже в нужном формате
        elif digits_only.startswith('8'):
            return '7' + digits_only[1:]  # Заменяем 8 на 7
    elif len(digits_only) == 10:
        # 9123456789 (без кода страны)
        if digits_only.startswith('9'):
            return '7' + digits_only
    
    # Если номер слишком длинный или не соответствует российскому формату
    if len(digits_only) > 11:
        logger.debug(f"Слишком длинный номер: '{phone}' -> '{digits_only}'")
        return ""
    
    logger.debug(f"Не удалось нормализовать номер: '{phone}' -> '{digits_only}'")
    return ""


def find_recent_data_start_index(rows: List[List[str]]) -> int:
    """
    Находит индекс первой строки с датой >= (сегодня - ANALYSIS_DAYS_DEPTH дней).
    Оптимизирует проверку приёмника, анализируя только недавние данные.
    Использует московский часовой пояс для определения текущей даты.
    
    Args:
        rows: Список строк с данными (без заголовка)
        
    Returns:
        int: Индекс первой строки для анализа (0 если не найдено)
    """
    if not rows:
        return 0
    
    # Получаем глубину анализа из настроек (по умолчанию 1 день)
    analysis_days = int(os.getenv('ANALYSIS_DAYS_DEPTH', 1))
    
    # Вычисляем дату начала анализа по московскому времени
    moscow_now = datetime.now(MOSCOW_TZ)
    cutoff_date = moscow_now - timedelta(days=analysis_days)
    cutoff_date = cutoff_date.date()
    
    logger.info(f"Ищем данные начиная с {cutoff_date.strftime('%Y-%m-%d')} (глубина: {analysis_days} дней)")
    logger.info(f"Текущее московское время: {moscow_now.strftime('%Y-%m-%d %H:%M:%S %Z')}")
    
    # Проходим строки с конца (новые данные обычно в конце)
    recent_start_index = 0
    found_recent = False
    
    for i in range(len(rows) - 1, -1, -1):
        row = rows[i]
        
        # Проверяем наличие даты в столбце A
        if len(row) > DATE_COLUMN_INDEX and row[DATE_COLUMN_INDEX].strip():
            date_str = row[DATE_COLUMN_INDEX].strip()
            parsed_date = parse_date(date_str)
            
            if parsed_date:
                row_date = parsed_date.date()
                
                # Если дата >= cutoff_date, запоминаем индекс
                if row_date >= cutoff_date:
                    recent_start_index = i
                    found_recent = True
                else:
                    # Если дата < cutoff_date, прекращаем поиск
                    break
    
    if found_recent:
        total_rows = len(rows)
        recent_rows = total_rows - recent_start_index
        logger.info(f"Найдены недавние данные: строки {recent_start_index+1}-{total_rows} ({recent_rows} строк)")
        logger.info(f"Оптимизация: пропускаем {recent_start_index} старых строк ({recent_start_index/total_rows*100:.1f}%)")
    else:
        logger.info("Недавние данные не найдены, анализируем все строки")
    
    return recent_start_index


def extract_phone_numbers(rows: List[List[str]], start_index: int = 0) -> Set[str]:
    """
    Извлекает нормализованные номера телефонов из строк (столбец D, индекс 3).
    Может анализировать только часть строк для оптимизации.
    Номера приводятся к единому формату 79123456789 для корректного сравнения.
    
    Args:
        rows: Список строк с данными
        start_index: Индекс первой строки для анализа (по умолчанию 0)
        
    Returns:
        Set[str]: Множество нормализованных номеров телефонов
    """
    phones = set()
    analyzed_rows = rows[start_index:] if start_index > 0 else rows
    invalid_phones = 0
    
    for row in analyzed_rows:
        if len(row) > PHONE_COLUMN_INDEX and row[PHONE_COLUMN_INDEX].strip():
            original_phone = row[PHONE_COLUMN_INDEX].strip()
            normalized_phone = normalize_phone_number(original_phone)
            
            if normalized_phone:
                phones.add(normalized_phone)
                if original_phone != normalized_phone:
                    logger.debug(f"Нормализован номер: '{original_phone}' -> '{normalized_phone}'")
            else:
                invalid_phones += 1
                logger.debug(f"Пропущен некорректный номер: '{original_phone}'")
    
    if start_index > 0:
        logger.info(f"Найдено {len(phones)} уникальных нормализованных телефонов в {len(analyzed_rows)} недавних строках")
    else:
        logger.info(f"Найдено {len(phones)} уникальных нормализованных телефонов во всех {len(rows)} строках")
    
    if invalid_phones > 0:
        logger.info(f"Пропущено {invalid_phones} некорректных номеров телефонов")
    
    return phones


def filter_new_rows(src_rows: List[List[str]], existing_phones: Set[str]) -> List[List[str]]:
    """
    Фильтрует новые строки, которых ещё нет в приёмнике.
    Сравнение происходит по нормализованным номерам телефонов (столбец D).
    В приёмник записываются строки с нормализованными номерами в формате 79123456789.
    Строки без номера телефона или с некорректными номерами пропускаются.
    
    Args:
        src_rows: Строки из источника (без заголовка)
        existing_phones: Множество существующих нормализованных телефонов в приёмнике
        
    Returns:
        List[List[str]]: Список новых строк для добавления с нормализованными номерами
    """
    new_rows = []
    skipped_no_phone = 0
    skipped_duplicate = 0
    skipped_invalid = 0
    
    for row in src_rows:
        # Нормализуем строку до 7 столбцов (A-G)
        normalized_row = normalize_row(row, MAX_COLUMNS)
        
        # Проверяем наличие телефона в столбце D
        if len(normalized_row) > PHONE_COLUMN_INDEX:
            original_phone = normalized_row[PHONE_COLUMN_INDEX].strip()
            
            # Пропускаем строки без номера телефона
            if not original_phone:
                skipped_no_phone += 1
                logger.debug(f"Пропущена строка без телефона: {normalized_row[:3]}...")
                continue
            
            # Нормализуем номер телефона
            normalized_phone = normalize_phone_number(original_phone)
            
            # Пропускаем строки с некорректными номерами
            if not normalized_phone:
                skipped_invalid += 1
                logger.debug(f"Пропущена строка с некорректным номером '{original_phone}': {normalized_row[:3]}...")
                continue
            
            # Пропускаем строки с уже существующими телефонами
            if normalized_phone in existing_phones:
                skipped_duplicate += 1
                logger.debug(f"Пропущена дублирующая строка (нормализованный телефон: {normalized_phone}): {normalized_row[:3]}...")
                continue
            
            # Заменяем оригинальный номер на нормализованный в строке для записи
            normalized_row[PHONE_COLUMN_INDEX] = normalized_phone
            
            # Добавляем новую строку с нормализованным номером
            new_rows.append(normalized_row)
            
            if original_phone != normalized_phone:
                logger.debug(f"Новая строка с нормализованным номером: {normalized_row[:3]}... ('{original_phone}' -> '{normalized_phone}')")
            else:
                logger.debug(f"Новая строка: {normalized_row[:3]}... (телефон: {normalized_phone})")
        else:
            # Строка слишком короткая (нет столбца D)
            skipped_no_phone += 1
            logger.debug(f"Пропущена короткая строка (нет столбца D): {normalized_row}")
    
    logger.info(f"Обработано строк: {len(src_rows)}")
    logger.info(f"Новых строк для добавления: {len(new_rows)}")
    logger.info(f"Пропущено строк без телефона: {skipped_no_phone}")
    logger.info(f"Пропущено строк с некорректными номерами: {skipped_invalid}")
    logger.info(f"Пропущено дублирующих строк: {skipped_duplicate}")
    
    return new_rows


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=1, max=10),
    retry=retry_if_exception_type((HttpError, ConnectionError, TimeoutError)),
    before_sleep=lambda retry_state: logger.info(f"Повторная попытка через {retry_state.next_action.sleep} секунд...")
)
def append_rows_to_sheet(service, spreadsheet_id: str, sheet_name: str, rows: List[List[str]]) -> int:
    """
    Добавляет строки в конец указанного листа.
    Автоматически повторяет попытки при временных ошибках.
    
    Args:
        service: Сервис Google Sheets
        spreadsheet_id: ID таблицы
        sheet_name: Название листа
        rows: Строки для добавления
        
    Returns:
        int: Количество добавленных строк
        
    Raises:
        HttpError: При критических ошибках API Google Sheets
        ConnectionError: При проблемах с сетью
        TimeoutError: При таймаутах
    """
    if not rows:
        logger.info("Новых строк нет — ничего не добавляем")
        return 0
    
    logger.info(f"Добавляем {len(rows)} строк в лист '{sheet_name}'")
    
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=sheet_name,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows}
    ).execute()
    
    logger.info(f"Успешно добавлено {len(rows)} строк")
    return len(rows)


def sync_and_return_new_rows() -> List[List[str]]:
    """
    Выполняет синхронизацию данных между таблицами и возвращает новые строки.
    Функция для интеграции с Telegram-ботом.
    
    Returns:
        List[List[str]]: Список новых строк, добавленных в приёмник
        
    Raises:
        Exception: При критических ошибках синхронизации
    """
    try:
        logger.info("=== Запуск синхронизации для Telegram-бота ===")
        logger.info(f"Источник: {SRC_SHEET} (ID: {SRC_ID[:10]}...)")
        logger.info(f"Приёмник: {DST_SHEET} (ID: {DST_ID[:10]}...)")
        
        # Создаём сервис Google Sheets
        service = create_sheets_service()
        
        # Читаем данные из источника и приёмника
        src_rows = get_sheet_data(service, SRC_ID, SRC_SHEET)
        dst_rows = get_sheet_data(service, DST_ID, DST_SHEET)
        
        # Проверяем, что источник не пуст
        if not src_rows:
            logger.warning("Источник пуст — нечего синхронизировать")
            return []
        
        if len(src_rows) <= 1:
            logger.warning("В источнике только заголовок — нет данных для синхронизации")
            return []
        
        logger.info("=== Применение оптимизации по датам ===")
        
        # Проверяем настройку оптимизации источника
        optimize_source = os.getenv('OPTIMIZE_SOURCE', 'false').lower() == 'true'
        logger.info(f"Оптимизация источника: {'включена' if optimize_source else 'отключена'}")
        
        # Оптимизация источника (если включена)
        src_data_rows = src_rows[1:]  # Пропускаем заголовок
        logger.info(f"ИСТОЧНИК: общее количество строк данных: {len(src_data_rows)}")
        
        if optimize_source:
            src_recent_start_index = find_recent_data_start_index(src_data_rows)
            src_recent_rows = src_data_rows[src_recent_start_index:]
            
            logger.info(f"ИСТОЧНИК: будет обработано {len(src_recent_rows)} недавних строк")
            if src_recent_start_index > 0:
                logger.info(f"ИСТОЧНИК: пропущено {src_recent_start_index} старых строк")
        else:
            src_recent_rows = src_data_rows
            logger.info(f"ИСТОЧНИК: будут обработаны все {len(src_recent_rows)} строк")
        
        # Оптимизация приёмника (всегда включена)
        dst_data_rows = dst_rows[1:] if len(dst_rows) > 1 else []
        logger.info(f"ПРИЁМНИК: общее количество строк данных: {len(dst_data_rows)}")
        
        dst_recent_start_index = find_recent_data_start_index(dst_data_rows)
        existing_phones = extract_phone_numbers(dst_data_rows, dst_recent_start_index)
        
        logger.info("=== Фильтрация новых записей ===")
        
        # Фильтруем новые строки из обработанных данных источника
        new_rows = filter_new_rows(src_recent_rows, existing_phones)
        
        # Добавляем новые строки в приёмник
        if new_rows:
            append_rows_to_sheet(service, DST_ID, DST_SHEET, new_rows)
            logger.info(f"=== Синхронизация завершена. Добавлено {len(new_rows)} новых строк ===")
        else:
            logger.info("=== Синхронизация завершена. Новых строк не найдено ===")
        
        return new_rows
        
    except Exception as e:
        logger.error(f"Критическая ошибка в sync_and_return_new_rows(): {e}")
        raise


def main():
    """
    Основная функция скрипта.
    Выполняет синхронизацию данных между таблицами.
    """
    try:
        logger.info("=== Запуск синхронизации данных ===")
        logger.info(f"Источник: {SRC_SHEET} (ID: {SRC_ID[:10]}...)")
        logger.info(f"Приёмник: {DST_SHEET} (ID: {DST_ID[:10]}...)")
        
        # Создаём сервис Google Sheets
        service = create_sheets_service()
        
        # Читаем данные из источника и приёмника
        src_rows = get_sheet_data(service, SRC_ID, SRC_SHEET)
        dst_rows = get_sheet_data(service, DST_ID, DST_SHEET)
        
        # Проверяем, что источник не пуст
        if not src_rows:
            logger.warning("Источник пуст — нечего синхронизировать")
            return
        
        if len(src_rows) <= 1:
            logger.warning("В источнике только заголовок — нет данных для синхронизации")
            return
        
        logger.info("=== Применение оптимизации по датам ===")
        
        # Проверяем настройку оптимизации источника
        optimize_source = os.getenv('OPTIMIZE_SOURCE', 'false').lower() == 'true'
        logger.info(f"Оптимизация источника: {'включена' if optimize_source else 'отключена'}")
        
        # Оптимизация источника (если включена)
        src_data_rows = src_rows[1:]  # Пропускаем заголовок
        logger.info(f"ИСТОЧНИК: общее количество строк данных: {len(src_data_rows)}")
        
        if optimize_source:
            src_recent_start_index = find_recent_data_start_index(src_data_rows)
            src_recent_rows = src_data_rows[src_recent_start_index:]
            
            logger.info(f"ИСТОЧНИК: будет обработано {len(src_recent_rows)} недавних строк")
            if src_recent_start_index > 0:
                logger.info(f"ИСТОЧНИК: пропущено {src_recent_start_index} старых строк")
        else:
            src_recent_rows = src_data_rows
            logger.info(f"ИСТОЧНИК: будут обработаны все {len(src_recent_rows)} строк")
        
        # Оптимизация приёмника (всегда включена)
        dst_data_rows = dst_rows[1:] if len(dst_rows) > 1 else []
        logger.info(f"ПРИЁМНИК: общее количество строк данных: {len(dst_data_rows)}")
        
        dst_recent_start_index = find_recent_data_start_index(dst_data_rows)
        existing_phones = extract_phone_numbers(dst_data_rows, dst_recent_start_index)
        
        logger.info("=== Фильтрация новых записей ===")
        
        # Фильтруем новые строки из обработанных данных источника
        new_rows = filter_new_rows(src_recent_rows, existing_phones)
        
        # Добавляем новые строки в приёмник
        added_count = append_rows_to_sheet(service, DST_ID, DST_SHEET, new_rows)
        
        logger.info(f"=== Синхронизация завершена. Добавлено {added_count} новых строк ===")
        
    except Exception as e:
        logger.error(f"Критическая ошибка в main(): {e}")
        raise


if __name__ == "__main__":
    main()