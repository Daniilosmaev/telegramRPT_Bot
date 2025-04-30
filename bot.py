import logging
import os
import json
import asyncio
from collections import defaultdict
import re # Импортируем библиотеку для регулярных выражений

# --- Библиотеки Aiogram ---
from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import (ReplyKeyboardMarkup, KeyboardButton,
                           InlineKeyboardMarkup, InlineKeyboardButton,
                           ParseMode, ContentType)
from aiogram.utils.exceptions import MessageNotModified, CantParseEntities, MessageToDeleteNotFound, MessageCantBeDeleted

# --- Библиотека для работы с Word ---
from docx import Document
# --- КОНФИГУРАЦИЯ ---
# !!! НЕ ЗАБУДЬ ВСТАВИТЬ СВОИ ЗНАЧЕНИЯ !!!
API_TOKEN = "8199872713:AAHwZq0lkZGvysbkaik6lFdpedeFjWlqay0" # Ваш API токен
ADMIN_IDS = [1093014764] # Список ID администраторов
# !!! ------------------------------------ !!!
DATA_DIR = "data"
SCHEDULE_FILENAME = "schedule.docx"
DB_FILENAME = "bot_data.json"

# --- Настройка логирования ---
# Установим уровень DEBUG для подробной диагностики парсинга
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
log = logging.getLogger(__name__)

# --- Инициализация бота и диспетчера ---
# ЭТОТ БЛОК ДОЛЖЕН БЫТЬ ДО ВСЕХ @dp.message_handler и @dp.callback_query_handler !!!
storage = MemoryStorage()
bot = Bot(token=API_TOKEN, parse_mode=ParseMode.HTML) # Установим HTML как основной режим парсинга
dp = Dispatcher(bot, storage=storage)

# --- Пути к файлам данных ---
schedule_path = os.path.join(DATA_DIR, SCHEDULE_FILENAME)
db_path = os.path.join(DATA_DIR, DB_FILENAME)

# --- Функции для работы с данными (JSON файл) ---
def load_bot_data():
    if not os.path.exists(DATA_DIR):
        try: os.makedirs(DATA_DIR); log.info(f"Создана директория: {DATA_DIR}")
        except OSError as e: log.error(f"Ошибка создания директории {DATA_DIR}: {e}"); return {'admission_file_id': None, 'events_text': "Ошибка: Не удалось создать директорию для данных."}
    try:
        with open(db_path, 'r', encoding='utf-8') as f: data = json.load(f)
        data.setdefault('admission_file_id', None); data.setdefault('events_text', "Информация о мероприятиях еще не добавлена."); log.info(f"Данные успешно загружены из {db_path}"); return data
    except FileNotFoundError: log.warning(f"Файл {db_path} не найден. Создаю новый."); return {'admission_file_id': None, 'events_text': "Информация о мероприятиях еще не добавлена."}
    except json.JSONDecodeError: log.error(f"Ошибка декодирования JSON в файле {db_path}."); return {'admission_file_id': None, 'events_text': "Информация о мероприятиях еще не добавлена."}
    except Exception as e: log.error(f"Неизвестная ошибка при загрузке данных из {db_path}: {e}"); return {'admission_file_id': None, 'events_text': "Ошибка загрузки данных."}

def save_bot_data(data):
    if not os.path.exists(DATA_DIR):
        try: os.makedirs(DATA_DIR); log.info(f"Создана директория: {DATA_DIR}")
        except OSError as e: log.error(f"Ошибка создания директории {DATA_DIR} при сохранении: {e}"); return
    try:
        with open(db_path, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=4)
        log.info(f"Данные успешно сохранены в {db_path}")
    except Exception as e: log.error(f"Ошибка сохранения данных в {db_path}: {e}", exc_info=True)

bot_data = load_bot_data()

# --- Клавиатуры ---
kb_student = ReplyKeyboardMarkup(resize_keyboard=True)
kb_student.add(KeyboardButton("Набор на 2026 год"), KeyboardButton("Мероприятия"))
kb_student.add(KeyboardButton("Расписание"), KeyboardButton("❓ Задать вопрос"))

kb_admin = ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
kb_admin.add(KeyboardButton("Загрузить файл 'Набор на 2026'"), KeyboardButton("Загрузить текст 'Мероприятия'"),
             KeyboardButton("Загрузить файл 'Расписание' (.docx)"), KeyboardButton("Выйти из админ-панели"))

# --- Состояния FSM ---
class AdminStates(StatesGroup): waiting_for_admission_file = State(); waiting_for_events_text = State(); waiting_for_schedule_file = State()
class StudentStates(StatesGroup): waiting_for_question = State()

# --- ОБНОВЛЕННЫЕ Функции парсинга расписания ---

def get_group_names_from_schedule(docx_path: str) -> list[str] | None:
    """
    Читает имена групп из заголовков таблиц в DOCX файле.
    Ищет ячейки, содержащие "группа", и извлекает название группы из той же ячейки
    (ищет текст после "группа" в ячейке).
    Возвращает отсортированный список уникальных имен групп или None при ошибке чтения файла.
    Возвращает пустой список, если группы не найдены.
    """
    if not os.path.exists(docx_path):
        log.error(f"Файл расписания не найден: {docx_path}")
        return None
    try:
        doc = Document(docx_path)
        all_groups = set()
        log.info(f"Читаем имена групп из {docx_path}")

        for i, table in enumerate(doc.tables):
            log.info(f"Обработка таблицы #{i} для поиска групп.")
            if not table.rows:
                log.info(f"Таблица #{i} пустая, пропускаю.")
                continue
            try:
                header_cells = table.rows[0].cells
                log.info(f"Чтение заголовка таблицы #{i} с {len(header_cells)} ячейками.")

                for idx, cell in enumerate(header_cells):
                    cell_text = cell.text.strip()
                    cell_text_upper = cell_text.upper()
                    keyword = "ГРУППА"

                    # Измененная логика: ищем вхождение "ГРУППА" и берем текст после него
                    if keyword in cell_text_upper:
                        log.debug(f"Найдено слово '{keyword}' в ячейке #{idx} таблицы #{i}. Полный текст ячейки: '{cell_text}'")

                        keyword_index = cell_text_upper.find(keyword)
                        # Берем весь текст после слова "группа" и удаляем пробелы/переводы строк
                        group_name = cell_text[keyword_index + len(keyword):].strip()

                        group_name_upper = group_name.upper()

                        if group_name_upper: # Добавляем, только если имя не пустое
                            all_groups.add(group_name_upper)
                            log.debug(f"Извлечено и добавлено название группы: '{group_name_upper}'")
                        else:
                            log.debug(f"Текст после слова '{keyword}' в ячейке #{idx} таблицы #{i} пустой.")


                log.info(f"Найденные уникальные группы после обработки таблицы #{i}: {all_groups}")
            except Exception as e: # Ловим более общие ошибки чтения ячеек
                log.error(f"Ошибка чтения заголовков в таблице #{i} файла {docx_path}: {e}", exc_info=True)
                continue # Продолжаем обработку других таблиц


        if not all_groups:
             log.warning(f"Не найдено ни одного имени группы в заголовках таблиц файла {docx_path}")
             return [] # Возвращаем пустой список

        sorted_groups = sorted(list(all_groups))
        log.info(f"Итоговый список найденных и отсортированных групп: {sorted_groups}")
        return sorted_groups


    except Exception as e:
        log.error(f"Не удалось открыть/прочитать DOCX файл {docx_path} для получения имен групп: {e}", exc_info=True)
        return None # Ошибка чтения файла


async def parse_schedule(docx_path: str, target_group: str) -> str | None:
    """
    ОБНОВЛЕНО: Парсит DOCX файл для ЗАДАННОЙ группы, учитывая структуру с | группа | ИМЯ |.
    Возвращает отформатированную строку для Telegram или None/сообщение об ошибке.
    Формат вывода:
    День:
      Занятие1 (текст из ячейки)
      Занятие2
      ...
    """
    target_group_upper = target_group.strip().upper() # Сразу приводим к верхнему регистру
    log.info(f"Начинаю парсинг расписания для группы: '{target_group_upper}' в файле {docx_path}") # Изменил на INFO
    try:
        if not os.path.exists(docx_path): raise FileNotFoundError(f"Файл не найден: {docx_path}")
        doc = Document(docx_path)
    except Exception as e:
        log.error(f"Не удалось открыть/прочитать DOCX файл {docx_path}: {e}", exc_info=True)
        return f"❌ Ошибка чтения файла расписания.\nСообщите администратору.\n({e})"

    schedule_data = defaultdict(list)
    target_group_column_index = -1
    target_group_found_in_header = False # Флаг для отслеживания, нашли ли группу в заголовке

    for i, table in enumerate(doc.tables):
        log.debug(f"Обработка таблицы #{i} для парсинга расписания.")
        if not table.rows: continue

        try: header_cells = table.rows[0].cells
        except Exception as e: log.warning(f"Ошибка чтения заголовка табл.{i}: {e}"); continue

        # Ищем столбец с нужной группой в заголовке (индекс ячейки с данными)
        # Используем ту же логику поиска, что и в get_group_names_from_schedule
        current_table_group_index = -1
        for idx, cell in enumerate(header_cells):
            cell_text = cell.text.strip()
            cell_text_upper = cell_text.upper()
            keyword = "ГРУППА"

            if keyword in cell_text_upper:
                 keyword_index = cell_text_upper.find(keyword)
                 group_name_in_header = cell_text[keyword_index + len(keyword):].strip().upper()

                 # --- ИЗМЕНЕНИЕ ЗДЕСЬ: Сравниваем названия групп без пробелов ---
                 group_name_in_header_clean = group_name_in_header.replace(" ", "")
                 target_group_upper_clean = target_group_upper.replace(" ", "")

                 if group_name_in_header_clean == target_group_upper_clean:
                     # Найден столбец данных для этой группы
                     current_table_group_index = idx
                     target_group_found_in_header = True
                     log.debug(f"Найдена группа '{target_group_upper}' в таблице #{i}, столбец с данными: {current_table_group_index}")
                     break # Нашли нужную группу в этой таблице
                 # --- КОНЕЦ ИЗМЕНЕНИЯ ---


        if current_table_group_index == -1:
            log.debug(f"Группа '{target_group_upper}' не найдена в заголовках таблицы #{i}. Пропускаю эту таблицу.")
            continue # Группа не найдена в этой таблице, ищем в следующей

        # Определяем столбец дня недели (предполагаем, что он первый, индекс 0)
        day_column_index = 0
        # Проверяем, что столбец дня существует и содержит текст "дата\nдень" или только "день"
        # УБРАЛ ПРОВЕРКУ len(cells) <= day_column_index ЗДЕСЬ, т.к. она будет ниже после получения cells
        if "ДЕНЬ" not in header_cells[day_column_index].text.strip().upper():
             # Если в первой ячейке заголовка нет слова "ДЕНЬ", предполагаем, что это все равно ячейка дня
             log.warning(f"В таблице #{i} в ячейке дня с индексом {day_column_index} не найдено 'день'. Все равно использую ее как столбец дня.")
             pass # Продолжаем, используя day_column_index = 0

        current_day = None
        days_of_week = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"] # Для поиска дня

        # --- УЛУЧШЕННАЯ Логика для обновления current_day ---
        for row_idx, row in enumerate(table.rows[1:], start=1): # Обход строк данных (начиная со второй строки таблицы, индекс 1)
            try:
                cells = row.cells # Успешное получение ячеек
                # ПРОВЕРКА ТЕПЕРЬ ЗДЕСЬ, ПОСЛЕ УСПЕШНОГО ПОЛУЧЕНИЯ cells
                if len(cells) <= max(day_column_index, current_table_group_index):
                     log.debug(f"Строка {row_idx} в таблице #{i} слишком короткая ({len(cells)} ячеек) для парсинга дня или группы. Пропускаю.")
                     continue

                day_cell_text_raw = cells[day_column_index].text # Сырой текст ячейки дня
                day_cell_text_stripped = day_cell_text_raw.strip() # Обрезанный текст ячейки дня
                log.debug(f"Ячейка дня в строке {row_idx}, табл.{i}. Сырой текст: '{day_cell_text_raw}', Обрезанный текст: '{day_cell_text_stripped}'")

                if day_cell_text_stripped: # Если ячейка дня не пустая
                    found_day_in_cell = False
                    # Проверяем каждую строку в ячейке на предмет названия дня
                    for line in day_cell_text_raw.split('\n'): # Используем сырой текст для split
                        stripped_line = line.strip()
                        # Ищем название дня недели (без учета регистра)
                        if stripped_line.upper() in days_of_week:
                            current_day = stripped_line.capitalize() # Нашли день, капитализируем и устанавливаем
                            log.debug(f"Обновлен текущий день недели: {current_day} (строка {row_idx} в таблице #{i})")
                            found_day_in_cell = True
                            break # Нашли день в этой ячейке, дальше в этой ячейке искать не нужно

                    # Оставим день из предыдущей строки, если текущая ячейка дня не пустая, но не содержит названия дня недели.
                    # Это нужно для строк с занятиями под тем же днем.
                    if not found_day_in_cell and current_day is not None:
                         log.debug(f"Ячейка дня {row_idx} в табл.{i} содержит текст ('{day_cell_text_stripped}'), но не содержит названия дня недели. Сохраняю предыдущий день '{current_day}'.")
                    elif not found_day_in_cell and current_day is None:
                         log.debug(f"Ячейка дня {row_idx} в табл.{i} содержит текст ('{day_cell_text_stripped}'), но не содержит названия дня недели и предыдущий день не установлен. День не определен.")

                else:
                     log.debug(f"Ячейка дня {row_idx} в табл.{i} пустая. Сохраняю предыдущий день '{current_day}'.")

            # --- Конец УЛУЧШЕННОЙ Логики для обновления current_day ---


                if current_day is None:
                     log.debug(f"День еще не определен для строки {row_idx} в табл.{i}. Пропускаю парсинг занятий в этой строке.")
                     continue # Если день еще не определен, пропускаем эту строку

                # Извлекаем данные для нашей группы
                lesson_cell_text_raw = cells[current_table_group_index].text # Текст как есть
                lesson_cell_text_stripped = lesson_cell_text_raw.strip() # Текст после strip()

                # *** ЭТИ СТРОКИ НУЖНЫ ДЛЯ ДИАГНОСТИКИ (ОСТАВЛЯЕМ ИХ) ***
                log.debug(f"Строва {row_idx}, день '{current_day}', группа '{target_group_upper}', столбец {current_table_group_index}.")
                log.debug(f"  => Сырой текст ячейки урока: '{lesson_cell_text_raw}'")
                log.debug(f"  => Текст ячейки урока после strip(): '{lesson_cell_text_stripped}'")
                # *** КОНЕЦ ДИАГНОСТИЧЕСКИХ СТРОК ***


                # *** ИЗМЕНЕНИЕ ЗДЕСЬ: Обработка текста ячейки для вывода ***
                if lesson_cell_text_stripped: # Добавляем только если текст после strip() не пустой
                     # Use HTML escape FIRST
                     escaped_lesson = await escape_html(lesson_cell_text_stripped)

                     # Заменяем внутренние символы новой строки (\n) на пробелы для объединения текста
                     # Удаляем пробелы с концов каждой строки перед объединением
                     lines = escaped_lesson.split('\n')
                     formatted_lesson_content = " ".join(line.strip() for line in lines if line.strip())

                     # Добавляем отступ и отформатированный контент в список занятий для дня
                     schedule_data[current_day].append(f"  - {formatted_lesson_content}")

                     log.debug(f"  => Добавлено занятие для {current_day}: '{lesson_cell_text_stripped}' -> formatted as '{formatted_lesson_content}'")
                else:
                     log.debug(f"  => Текст ячейки урока после strip() пустой, занятие не найдено.")

                # --- КОНЕЦ ИЗМЕНЕНИЯ ---

            except Exception as e:
                # Теперь этот except ловит ошибки при обработке ячеек или данных строки
                log.warning(f"Ошибка обработки строки {row_idx} в табл.{i}: {e}")
                # Continue к следующей строке
                continue


        # Если нашли группу в этой таблице, дальше можно не искать
        # Это оптимизация. Если расписание для группы может быть разбросано по разным таблицам, закомментируйте этот break.
        if target_group_found_in_header:
             log.debug(f"Группа найдена и обработана в таблице #{i}, прекращаю поиск в других таблицах.")
             break


    # --- Формирование результата ---
    # Проверяем target_group_found_in_header, т.к. current_table_group_index сбрасывается для каждой таблицы
    if not target_group_found_in_header:
        log.warning(f"Группа '{target_group_upper}' не найдена в заголовках ни одной таблицы файла {docx_path}")
        # Сравниваем очищенные имена для вывода сообщения об ошибке
        # Найдем ближайшее похожее имя, если возможно? Или просто используем оригинальное?
        # Оставим оригинальное, так как оно было на кнопке.
        return f"😕 Расписание для группы <b>{await escape_html(target_group)}</b> не найдено в файле." # Возвращаем сообщение, если группа не найдена вообще


    if not schedule_data:
         log.info(f"Для группы '{target_group_upper}' не найдено записей в расписании.")
         # Возвращаем более точное сообщение, если группа найдена в заголовке, но данных нет
         return f"ℹ️ Для группы <b>{await escape_html(target_group)}</b> нет записей в расписании на эту неделю."


    output_lines = [f"<b>Расписание для группы: {target_group}</b>\n"] # Используем HTML для жирного шрифта
    # Сортируем дни в правильном порядке
    days_order = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"]
    # Приводим ключи словаря к верхнему регистру для сравнения с days_order
    schedule_data_upper_keys = {day.upper(): lessons for day, lessons in schedule_data.items()}
    # Сортируем по порядку дней, а затем по оригинальному ключу на случай дней вне списка (хотя по формату файла такого быть не должно)
    sorted_days = sorted(schedule_data_upper_keys.keys(), key=lambda day_upper: (days_order.index(day_upper) if day_upper in days_order else len(days_order), day_upper))

    for day_upper in sorted_days:
        # Получаем оригинальное название дня (с правильным регистром) для вывода
        original_day = next((d for d in schedule_data if d.upper() == day_upper), day_upper.capitalize())
        output_lines.append(f"<b>{original_day}:</b>") # День недели жирным
        lessons = schedule_data_upper_keys[day_upper] # lessons теперь содержат список уже отформатированных строк занятий
        if lessons:
            # Просто добавляем список отформатированных строк
            output_lines.extend(lessons)
        # Пустая строка между днями не нужна при таком формате

    final_output = "\n".join(output_lines).strip() # Объединяем все строки
    log.debug(f"Сформировано расписание для группы {target_group_upper} (длина {len(final_output)})")
    return final_output


async def escape_html(text: str) -> str:
     """Простая функция для экранирования основных HTML символов."""
     if not isinstance(text, str): return text # Если не строка, возвращаем как есть
     return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# --- Обработчики команд и сообщений ---

# Убедитесь, что инициализация bot и dp идет ДО этого блока
@dp.message_handler(commands=['start'], state='*')
async def send_welcome(message: types.Message, state: FSMContext):
    await state.finish()
    log.info(f"User {message.from_user.id} started bot.")
    await message.answer("Привет! Я чат-бот поддержки студентов. Выбери нужный раздел:", reply_markup=kb_student)

# --- Функционал для студентов ---

@dp.message_handler(text="Набор на 2026 год", state='*')
async def send_admission_info(message: types.Message):
    file_id = bot_data.get('admission_file_id')
    if file_id:
        try:
            log.info(f"Sending admission file {file_id} to {message.from_user.id}")
            await bot.send_document(message.chat.id, file_id, caption="Информация о наборе на 2026 год.")
        except Exception as e:
            log.error(f"Error sending admission file {file_id} to {message.from_user.id}: {e}")
            await message.answer("Не удалось отправить файл. Возможно, он был удален. Сообщите администратору.")
    else:
        log.warning(f"Admission file not found for user {message.from_user.id}")
        await message.answer("Файл с информацией о наборе еще не загружен.")

@dp.message_handler(text="Мероприятия", state='*')
async def send_events_info(message: types.Message):
    events_text = bot_data.get('events_text', "Информация о мероприятиях еще не добавлена.")
    log.info(f"Sending events info to {message.from_user.id}")
    # Пытаемся отправить с HTML, если не вышло - как есть
    try: await message.answer(events_text, parse_mode=ParseMode.HTML, disable_web_page_preview=True)
    except CantParseEntities:
         log.warning(f"Failed to parse HTML for events text for user {message.from_user.id}. Sending plain text.")
         await message.answer(events_text, disable_web_page_preview=True)

# --- Новая логика для кнопки "Расписание" ---

@dp.message_handler(text="Расписание", state='*')
async def show_schedule_groups(message: types.Message, state: FSMContext):
    await state.finish()
    user_id = message.from_user.id
    log.info(f"User {user_id} requested schedule. Getting group list.")
    msg_wait = await message.answer("🔍 Ищу список групп в файле...")

    group_names = get_group_names_from_schedule(schedule_path)

    if group_names is None:
        await msg_wait.edit_text("❌ Ошибка чтения файла расписания. Сообщите администратору.")
        log.error(f"Failed to get group names from schedule file for user {user_id}.")
        return
    if not group_names:
         await msg_wait.edit_text("😕 Список групп в файле расписания не найден. Администратор должен проверить файл.")
         log.warning(f"No group names found in schedule file for user {user_id}.")
         return

    inline_kb_groups = InlineKeyboardMarkup(row_width=3) # 3 кнопки в ряд
    groups_added = 0
    for group in group_names:
        callback_data = f"schedule_group:{group}"
        # Проверка длины callback_data
        if len(callback_data.encode('utf-8')) > 64:
             log.warning(f"Callback data too long for group '{group}' ({len(callback_data.encode('utf-8'))} bytes), button skipped.")
             continue
        inline_kb_groups.insert(InlineKeyboardButton(text=group, callback_data=callback_data))
        groups_added += 1

    if groups_added > 0:
         await msg_wait.edit_text("⬇️ Выберите вашу группу:", reply_markup=inline_kb_groups)
    else:
         await msg_wait.edit_text("😕 Не найдено групп (возможно, их имена слишком длинные). Сообщите администратору.")
         log.warning(f"No groups added to inline keyboard for user {user_id}.")


@dp.callback_query_handler(lambda c: c.data.startswith('schedule_group:'), state='*')
async def send_schedule_for_group(call: types.CallbackQuery, state: FSMContext):
    group_name = call.data.split(':', 1)[1]
    user_id = call.from_user.id
    log.info(f"User {user_id} selected group '{group_name}' for schedule.")
    await call.answer()

    # Показываем статус загрузки
    try:
        # Используем исходное название группы для отображения
        await call.message.edit_text(f"⏳ Загружаю расписание для группы <b>{await escape_html(group_name)}</b>...", reply_markup=None)
    except MessageNotModified: pass
    except Exception as e: log.error(f"Failed to edit 'Select group' message for {user_id}: {e}")

    await asyncio.sleep(0.2) # Короткая пауза

    try:
        schedule_text = await parse_schedule(schedule_path, group_name)

        # parse_schedule теперь возвращает сообщения об ошибках или отсутствии записей
        if schedule_text:
             log.info(f"Sending schedule for group {group_name} to user {user_id}")
             # Сначала удаляем сообщение "Загружаю..." или "Выберите группу..."
             try: await call.message.delete()
             except (MessageToDeleteNotFound, MessageCantBeDeleted): pass # Игнорируем, если не получилось удалить
             except Exception as e: log.warning(f"Could not delete previous message for {user_id}: {e}")

             # Отправляем расписание или сообщение от парсера
             if len(schedule_text) > 4096:
                 log.warning(f"Schedule text for group {group_name} too long ({len(schedule_text)} bytes), splitting.")
                 # Разбиваем на части, стараясь сохранить целостность строк
                 messages = []
                 current_part = ""
                 # Разбиваем сначала по дням, потом уже внутри дней, если нужно
                 # Ищем заголовки дней или начало расписания
                 sections = schedule_text.split('<b>') # Разделяем по началу жирного текста (дни или заголовок)

                 for section in sections:
                      if not section: continue

                      # Восстанавливаем маркер жирного шрифта, если это не первая часть (заголовок)
                      section_to_add = '<b>' + section if not sections[0].startswith(section) else section

                      # Используем запас, чтобы избежать ошибок парсинга HTML
                      # Проверяем не просто длину, а длину после объединения
                      # Учитываем, что символы HTML разметки тоже занимают место
                      if len((current_part + section_to_add).encode('utf-8')) > 3800: # С запасом на HTML теги
                          if current_part.strip():
                              messages.append(current_part.strip())
                          current_part = section_to_add
                      else:
                          current_part += section_to_add


                 if current_part.strip():
                      messages.append(current_part.strip())


                 for part in messages:
                      await bot.send_message(call.message.chat.id, part) # Parse mode уже HTML из настроек бота
                      await asyncio.sleep(0.3)

             else:
                 await bot.send_message(call.message.chat.id, schedule_text) # Parse mode HTML

        # else: # parse_schedule теперь не должен возвращать None при найденной группе
        #      log.error(f"parse_schedule returned None unexpectedly for group {group_name} ({user_id}).")
        #      await call.message.edit_text(f"❌ Неизвестная ошибка при получении расписания для группы <b>{await escape_html(group_name)}</b>.")


    except Exception as e:
        log.error(f"Critical error processing schedule request for group {group_name} ({user_id}): {e}", exc_info=True)
        error_message = f"❌ Произошла ошибка при обработке расписания для группы <b>{await escape_html(group_name)}</b>.\nИнформация об ошибке: {e}" # Добавил информацию об ошибке
        try:
             # Пробуем редактировать ИСХОДНОЕ сообщение с кнопками (call.message)
             # или сообщение о загрузке, если оно еще не удалено
             # Если editMessageText выдает ошибку "message to edit not found",
             # это нормально, т.к. мы могли его успешно удалить ранее.
             await call.message.edit_text(error_message, reply_markup=None) # Убираем клавиатуру при ошибке
        except MessageNotModified: # Игнорируем, если сообщение не изменилось
             pass
        except Exception:
             # Если редактирование не удалось, отправляем новое сообщение с ошибкой
             await bot.send_message(call.message.chat.id, error_message)


# --- Обработчики для системы "Вопросы-Ответы" ---

@dp.message_handler(text="❓ Задать вопрос", state='*')
async def ask_question_start(message: types.Message, state: FSMContext):
    await state.finish()
    await message.answer("Напишите ваш вопрос одним сообщением. Он будет передан администратору.")
    await StudentStates.waiting_for_question.set()
    log.info(f"User {message.from_user.id} initiated question. State set: waiting_for_question.")

@dp.message_handler(state=StudentStates.waiting_for_question, content_types=ContentType.ANY)
async def forward_question_to_admin(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    log.info(f"Received question from {user_id}. Forwarding to admins: {ADMIN_IDS}")
    forward_success_count = 0
    if not ADMIN_IDS:
        log.warning("ADMIN_IDS list is empty! Cannot forward question.")
        await message.answer("❌ Не удалось отправить вопрос: не настроены администраторы.")
        await state.finish(); return

    for admin_id in ADMIN_IDS:
        try:
            fwd_msg = await bot.forward_message(chat_id=admin_id, from_chat_id=message.chat.id, message_id=message.message_id)
            await bot.send_message(admin_id,
                                   f"👆 Новый вопрос от студента!\n"
                                   f"👤: {await escape_html(message.from_user.full_name)}\n"
                                   f"🆔: <code>{user_id}</code>\n\n"
                                   f"❗️Чтобы ответить, используйте функцию <b>'Ответить'</b> (Reply) на <i>пересланное сообщение</i> выше.",
                                   parse_mode=ParseMode.HTML) # Используем HTML для форматирования ID
            forward_success_count += 1
            await asyncio.sleep(0.1)
        except Exception as e: log.error(f"Failed to forward question to admin {admin_id} from {user_id}: {e}")

    await message.answer("✅ Ваш вопрос отправлен администратору." if forward_success_count > 0 else "❌ Не удалось отправить ваш вопрос ни одному администратору.")
    await state.finish(); log.info(f"Finished state waiting_for_question for user {user_id}")

@dp.message_handler(lambda msg: msg.from_user.id in ADMIN_IDS and msg.reply_to_message is not None and msg.reply_to_message.forward_from is not None, content_types=ContentType.ANY)
async def answer_to_student(message: types.Message):
    student_id = None
    try:
        # Пытаемся получить ID студента из пересланного сообщения
        if message.reply_to_message.forward_from:
             student_id = message.reply_to_message.forward_from.id
             student_info = await escape_html(message.reply_to_message.forward_from.full_name)
        elif message.reply_to_message.text:
             # Если это не пересланное сообщение, но админ ответил на сообщение бота,
             # которое содержит ID студента в формате <code>ID</code>
             # Ищем ID в тексте оригинального сообщения бота
             import re
             match = re.search(r'🆔:\s*<code>(\d+)</code>', message.reply_to_message.html_text or message.reply_to_message.text) # Ищем и в html_text и в text
             if match:
                 student_id = int(match.group(1))
                 student_info = f"ID {student_id}" # Имя может быть недоступно
             else:
                 await message.reply("❌ Не удалось определить ID студента из пересланного сообщения или по формату бота.")
                 log.warning(f"Admin {message.from_user.id} replied to message, but could not extract student ID.")
                 return
        else:
            await message.reply("❌ Не удалось определить, кому ответить. Ответьте (Reply) на пересланное сообщение бота с вопросом студента.")
            log.warning(f"Admin {message.from_user.id} replied to message, but it's not a forwarded message from a student or a bot message with ID.")
            return

        admin_id = message.from_user.id
        log.info(f"Admin {admin_id} is replying to student {student_id} ({student_info})")

        # Формируем префикс ответа
        reply_prefix = "<b>Ответ от администратора:</b>\n"

        # Отправляем ответ студенту, копируя исходное сообщение админа
        # copy_message лучше сохраняет форматирование и медиа
        await bot.copy_message(
            chat_id=student_id, from_chat_id=message.chat.id, message_id=message.message_id,
            caption=(f"{reply_prefix}{await escape_html(message.caption)}" if message.caption else reply_prefix.strip()), # Добавляем префикс к caption, strip() если caption пустой
            reply_markup=message.reply_markup,
            parse_mode=ParseMode.HTML # Указываем parse_mode явно для copy_message
        )

        # Если copy_message не отправило текст (например, если было только текст без медиа), отправляем текст отдельно
        # Проверяем, что это не сообщение, на которое отвечает бот (иначе может зациклиться)
        if message.text and not message.caption and message.content_type == ContentType.TEXT and message.message_id != message.reply_to_message.message_id:
             # Отправляем текст с префиксом
             log.debug(f"Sending text part of reply separately for student {student_id}")
             await bot.send_message(student_id, f"{reply_prefix}\n{await escape_html(message.text)}", parse_mode=ParseMode.HTML)


        await message.reply("✅ Ответ отправлен студенту.")
        log.info(f"Reply from {admin_id} successfully sent to student {student_id}")

    except Exception as e:
        log.error(f"Error sending reply to student (ID {student_id}) from admin {message.from_user.id}: {e}", exc_info=True)
        await message.reply(f"❌ Не удалось отправить ответ студенту.\nОшибка: {e}")


# --- Функционал для администраторов ---

@dp.message_handler(commands=['admin'], state='*')
async def admin_panel(message: types.Message, state: FSMContext):
    await state.finish()
    if message.from_user.id in ADMIN_IDS:
        log.info(f"Admin {message.from_user.id} accessed admin panel.")
        await message.answer("Добро пожаловать в админ-панель!", reply_markup=kb_admin)
    else:
        log.warning(f"User {message.from_user.id} tried to access admin panel.")
        await message.answer("У вас нет прав доступа.", reply_markup=kb_student)

@dp.message_handler(lambda message: message.text == "Выйти из админ-панели" and message.from_user.id in ADMIN_IDS, state='*')
async def exit_admin_panel(message: types.Message, state: FSMContext):
    await state.finish()
    log.info(f"Admin {message.from_user.id} exited admin panel.")
    await message.answer("Вы вышли из админ-панели.", reply_markup=kb_student)

# --- Обработка загрузки файлов и текста админом ---

@dp.message_handler(lambda message: message.text == "Загрузить файл 'Набор на 2026'" and message.from_user.id in ADMIN_IDS, state='*')
async def request_admission_file(message: types.Message):
    log.info(f"Admin {message.from_user.id} initiated admission file upload.")
    await message.answer("Отправьте файл (документ)...")
    await AdminStates.waiting_for_admission_file.set()

@dp.message_handler(content_types=ContentType.DOCUMENT, state=AdminStates.waiting_for_admission_file)
async def process_admission_file(message: types.Message, state: FSMContext):
    if message.document:
        bot_data['admission_file_id'] = message.document.file_id
        save_bot_data(bot_data)
        log.info(f"Admin {message.from_user.id} uploaded admission file. ID: {message.document.file_id}")
        await message.answer("✅ Файл 'Набор на 2026' сохранен!", reply_markup=kb_admin)
    else: await message.answer("❌ Ошибка: Нужен ДОКУМЕНТ.")
    await state.finish()

@dp.message_handler(lambda message: message.text == "Загрузить текст 'Мероприятия'" and message.from_user.id in ADMIN_IDS, state='*')
async def request_events_text(message: types.Message):
    log.info(f"Admin {message.from_user.id} initiated events text upload.")
    await message.answer("Введите текст для 'Мероприятия' (можно использовать HTML)...")
    await AdminStates.waiting_for_events_text.set()

@dp.message_handler(content_types=ContentType.TEXT, state=AdminStates.waiting_for_events_text)
async def process_events_text(message: types.Message, state: FSMContext):
    if message.text:
         bot_data['events_text'] = message.html_text # Сохраняем с HTML разметкой
         save_bot_data(bot_data)
         log.info(f"Admin {message.from_user.id} uploaded events text.")
         await message.answer("✅ Текст 'Мероприятия' сохранен!", reply_markup=kb_admin)
    else: await message.answer("❌ Ошибка: Нужен ТЕКСТ.")
    await state.finish()

@dp.message_handler(lambda message: message.text == "Загрузить файл 'Расписание' (.docx)" and message.from_user.id in ADMIN_IDS, state='*')
async def request_schedule_file(message: types.Message):
    log.info(f"Admin {message.from_user.id} initiated schedule file upload.")
    await message.answer("Отправьте файл расписания в формате <b>.docx</b>")
    await AdminStates.waiting_for_schedule_file.set()

@dp.message_handler(content_types=ContentType.DOCUMENT, state=AdminStates.waiting_for_schedule_file)
async def process_schedule_file(message: types.Message, state: FSMContext):
    if message.document and message.document.file_name and message.document.file_name.lower().endswith('.docx'):
        try:
            file_info = await bot.get_file(message.document.file_id)
            log.info(f"Admin {message.from_user.id} uploaded schedule file: {message.document.file_name}")
            if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR); log.info(f"Created directory {DATA_DIR}")
            await bot.download_file(file_info.file_path, schedule_path)
            log.info(f"Schedule file saved as {schedule_path}")
            await message.answer("✅ Файл расписания обновлен!", reply_markup=kb_admin)
        except Exception as e:
            log.error(f"Error saving schedule file {schedule_path}: {e}", exc_info=True)
            await message.answer(f"❌ Ошибка сохранения файла: {e}", reply_markup=kb_admin)
    else: await message.answer("❌ Ошибка: Нужен файл <b>.docx</b>")
    await state.finish()

# --- Обработчик для любых других сообщений ---
@dp.message_handler(state='*')
async def handle_other_messages(message: types.Message):
    log.debug(f"Unhandled message from {message.from_user.id}: {message.text or message.content_type}")
    # await message.answer("Неизвестная команда. Используйте кнопки.", reply_markup=kb_student)

# --- Запуск бота ---
if __name__ == '__main__':
    log.info("Запуск бота...")
    executor.start_polling(dp, skip_updates=True)
    log.info("Бот остановлен.")