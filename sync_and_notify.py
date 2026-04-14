#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Объединённый скрипт для синхронизации Google-таблиц и отправки уведомлений в Telegram.

Последовательность выполнения:
1. Синхронизация данных между таблицами (sheet_transfer.py)
2. Отправка уведомлений о новых лидах в Telegram через Telegram Bot API

Требования:
- Настроенные переменные окружения в .env файле
- Файл credentials.json для Google Sheets API
- Токен Telegram-бота и Chat ID
"""

import logging
import os
import time
import sys
from datetime import datetime
from typing import List
from pathlib import Path
import requests
from dotenv import load_dotenv
from sheet_transfer import sync_and_return_new_rows

# Создаём папку для логов (абсолютный путь рядом со скриптом)
logs_dir = (Path(__file__).resolve().parent / "logs")
logs_dir.mkdir(exist_ok=True)

# Настройка логирования с записью в файл
log_filename = logs_dir / f"sync_and_notify_{datetime.now().strftime('%Y%m%d')}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(str(log_filename), encoding='utf-8'),  # Запись в файл
        logging.StreamHandler()  # Вывод в консоль
    ],
    force=True  # Принудительно переопределяем конфигурацию, если она уже была настроена при импорте
)
logger = logging.getLogger(__name__)


def escape_html(text: str) -> str:
    """Экранирует HTML-символы для безопасной отправки в Telegram."""
    if not text:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def send_telegram_message(
    token: str,
    chat_id: int,
    text: str,
    logger_instance: logging.Logger
) -> bool:
    """Отправляет одно сообщение через Telegram Bot API."""
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML",
    }

    try:
        response = requests.post(url, json=payload, timeout=30)
        response.raise_for_status()

        response_data = response.json()
        if not response_data.get("ok"):
            logger_instance.error(
                "Telegram API вернул ошибку: %s",
                response_data.get("description", "Неизвестная ошибка")
            )
            return False

        return True
    except requests.RequestException as error:
        logger_instance.error("Ошибка запроса к Telegram Bot API: %s", error)
        return False
    except ValueError as error:
        logger_instance.error("Некорректный ответ Telegram Bot API: %s", error)
        return False


def notify_rows_data(new_rows: List[List[str]]) -> bool:
    """
    Отправляет уведомления в Telegram для переданных строк данных.
    
    Args:
        new_rows (List[List[str]]): Список новых строк для отправки
        
    Returns:
        bool: True если успешно, False при ошибке
    """

    # Загружаем переменные окружения
    load_dotenv(override=True)

    telegram_bot_token = os.getenv('TELEGRAM_BOT_TOKEN_ASSISTANT')
    telegram_chat_id = os.getenv('TELEGRAM_CHAT_ID')

    if not telegram_bot_token:
        logger.error("TELEGRAM_BOT_TOKEN_ASSISTANT не найден в переменных окружения")
        return False

    if not telegram_chat_id:
        logger.error("TELEGRAM_CHAT_ID не найден в переменных окружения")
        return False

    # Преобразуем Chat ID в число для правильной работы с Telegram API
    try:
        telegram_chat_id = int(telegram_chat_id)
    except ValueError:
        logger.error(f"TELEGRAM_CHAT_ID должен быть числом, получено: {telegram_chat_id}")
        return False

    if not new_rows:
        logger.info("Новых строк для Telegram нет")
        return True

    try:
        logger.info(f"Найдено {len(new_rows)} новых лидов для отправки в Telegram")
        logger.info(f"Используем Chat ID: {telegram_chat_id}")

        for i, row in enumerate(new_rows, 1):
            try:
                # Формируем сообщение с проверкой длины строки
                name = row[2] if len(row) > 2 else "Не указано"
                phone = row[3] if len(row) > 3 else "Не указано"
                comment = row[4] if len(row) > 4 else "Не указано"
                additional_comment = row[5] if len(row) > 5 else "Не указано"
                audio_link = row[6] if len(row) > 6 else "Не указано"
                date = row[0] if len(row) > 0 else "Не указано"

                # Экранируем HTML-символы для безопасности
                name_escaped = escape_html(name)
                phone_escaped = escape_html(phone)
                comment_escaped = escape_html(comment)
                additional_comment_escaped = escape_html(additional_comment)
                audio_link_escaped = escape_html(audio_link)
                date_escaped = escape_html(date)

                message = (
                    f"Новый лид: {name_escaped} ({phone_escaped})\n\n"
                    f"Имя: {name_escaped}\n\n"
                    f"Телефон: {phone_escaped}\n\n"
                    f"Комментарий: {comment_escaped}\n\n"
                    f"Доп. комментарий: {additional_comment_escaped}\n\n"
                    f"Ссылка на запись: {audio_link_escaped}\n\n"
                    f"Дата лида: {date_escaped}"
                )

                send_success = send_telegram_message(
                    token=telegram_bot_token,
                    chat_id=telegram_chat_id,
                    text=message,
                    logger_instance=logger
                )

                if send_success:
                    logger.info(
                        f"Отправлено уведомление {i}/{len(new_rows)} для лида: {name} ({phone})"
                    )
                else:
                    logger.error(f"Не удалось отправить уведомление {i}/{len(new_rows)}")

                # Небольшая задержка снижает риск ограничения со стороны Telegram API.
                time.sleep(1)

            except Exception as e:
                logger.error(f"Ошибка при отправке строки {i}: {e}")
                logger.error(f"Данные строки: {row}")

        logger.info(f"Завершена отправка уведомлений. Обработано {len(new_rows)} лидов.")
        return True

    except Exception as e:
        logger.error(f"Критическая ошибка при отправке в Telegram: {e}")
        return False


def main():
    """
    Основная функция для запуска синхронизации и отправки уведомлений.
    """
    try:
        start_time = datetime.now()
        logger.info("=== ЗАПУСК ЦИКЛА СИНХРОНИЗАЦИИ И УВЕДОМЛЕНИЙ ===")
        logger.info(f"Время запуска: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Логи записываются в: {log_filename}")
        
        # Счётчики для статистики
        sync_success = False
        telegram_success = False
        new_rows = []
        
        # Этап 1: Синхронизация Google-таблиц
        try:
            logger.info("🔄 ЭТАП 1: Синхронизация Google-таблиц")
            new_rows = sync_and_return_new_rows()
            sync_success = True
            logger.info(f"✅ Этап 1 завершён успешно. Найдено новых лидов: {len(new_rows)}")
        except Exception as e:
            logger.error(f"❌ Ошибка на этапе синхронизации: {e}")
            sync_success = False
            new_rows = []
        
        # Этап 2: Уведомления в Telegram (только если есть новые данные)
        if sync_success and new_rows:
            try:
                logger.info("📱 ЭТАП 2: Отправка уведомлений в Telegram")
                telegram_success = notify_rows_data(new_rows)
                if telegram_success:
                    logger.info("✅ Этап 2 завершён успешно")
                else:
                    logger.error("❌ Этап 2 завершён с ошибками")
            except Exception as e:
                logger.error(f"❌ Ошибка на этапе уведомлений в Telegram: {e}")
                telegram_success = False
        else:
            logger.info("📱 ЭТАП 2: Пропущен (нет новых данных)")
            telegram_success = True  # Считаем успешным, так как нет данных для обработки
        
        # Итоговая статистика
        end_time = datetime.now()
        duration = (end_time - start_time).total_seconds()
        
        logger.info("=" * 60)
        logger.info("📊 ИТОГОВАЯ СТАТИСТИКА")
        logger.info("=" * 60)
        logger.info(f"Время выполнения: {duration:.2f} секунд")
        logger.info(f"Синхронизация: {'✅ Успешно' if sync_success else '❌ Ошибка'}")
        logger.info(f"Найдено новых лидов: {len(new_rows)}")
        logger.info(f"Telegram уведомления: {'✅ Успешно' if telegram_success else '❌ Ошибка'}")
        
        logger.info("=" * 60)
        logger.info("🎉 ЦИКЛ СИНХРОНИЗАЦИИ И УВЕДОМЛЕНИЙ ЗАВЕРШЁН")
        
    except Exception as e:
        logger.error(f"Критическая ошибка в main(): {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()