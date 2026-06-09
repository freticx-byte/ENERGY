import asyncio
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Command
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.utils import executor
from openpyxl import Workbook, load_workbook
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ============ ТОКЕН БОТА ============
TOKEN = "8976307638:AAGZNiGdfhYeYTjWVvWS2g3bAFM5RiwLi1g"

# ============ СОЗДАНИЕ БОТА ============
storage = MemoryStorage()
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)

# ============ ФАЙЛЫ ============
EXCEL_ELE_FILE = "Energy_rep_ELE.xlsx"
EXCEL_RES_FILE = "Energy_rep_RES.xlsx"
MENU_ELE_FILE = "Energy_menu_ELE.xlsx"
MENU_RES_FILE = "Energy_menu_RES.xlsx"

# ============ НАСТРОЙКИ EMAIL ============
EMAIL_TO = "a.misyunas@uvelka.ru"
EMAIL_FROM = "uvenergorusursy@gmail.com"
EMAIL_PASSWORD = "bmdt pzqh qdme wgnc"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# ============ ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ДЛЯ МЕНЮ ============
ELE_MENU = {}  # Иерархическое меню для ПЛК
RES_MENU = {}  # Иерархическое меню для Ресурс
ALL_ELE_COUNTERS = []  # Все счётчики ПЛК (плоский список)
ALL_RES_COUNTERS = []  # Все счётчики Ресурс (плоский список)
ELE_COUNTER_TO_PATH = {}  # Счётчик -> путь в меню
RES_COUNTER_TO_PATH = {}  # Счётчик -> путь в меню


class EnergyForm(StatesGroup):
    choosing_object = State()
    choosing_level1 = State()
    choosing_level2 = State()
    choosing_counter = State()
    entering_value = State()


# ============ ПАРСИНГ МЕНЮ ИЗ EXCEL ============
def parse_menu_from_excel(file_path, sheet_name):
    """Парсит иерархическое меню из Excel файла"""
    if not os.path.exists(file_path):
        print(f"Файл {file_path} не найден")
        return {}
    
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    
    menu = {}
    all_counters = []
    counter_to_path = {}
    
    for row in ws.iter_rows(min_row=2, max_col=15, values_only=True):
        if not row or all(cell is None for cell in row):
            continue
        
        level1 = row[3] if len(row) > 3 and row[3] else None  # Уровень 1 (ЦРП, ТП1...)
        level2 = row[4] if len(row) > 4 and row[4] else None  # Уровень 2 (ЦРП В1, ТП1 СГП...)
        level3 = row[5] if len(row) > 5 and row[5] else None  # Уровень 3 (конкретный счётчик)
        
        if level1 and level1 not in menu:
            menu[level1] = {}
        
        if level2 and level1:
            if level2 not in menu[level1]:
                menu[level1][level2] = []
        
        if level3 and level2 and level1:
            if level3 not in menu[level1][level2]:
                menu[level1][level2].append(level3)
                all_counters.append(level3)
                counter_to_path[level3] = (level1, level2, level3)
    
    wb.close()
    return menu, all_counters, counter_to_path


def load_menus():
    """Загружает меню для обоих объектов"""
    global ELE_MENU, ALL_ELE_COUNTERS, ELE_COUNTER_TO_PATH
    global RES_MENU, ALL_RES_COUNTERS, RES_COUNTER_TO_PATH
    
    # Загружаем меню из файлов (если есть)
    if os.path.exists(MENU_ELE_FILE):
        ELE_MENU, ALL_ELE_COUNTERS, ELE_COUNTER_TO_PATH = parse_menu_from_excel(MENU_ELE_FILE, "Меню ПЛК")
    else:
        # Встроенное меню по умолчанию
        ELE_MENU = {
            "ЦРП 10кВ": {
                "ЦРП В1": ["ЦРП В1"],
                "ЦРП В2": ["ЦРП В2"]
            },
            "ТП1": {
                "ТП1 СГП В1": ["ТП1 СГП В1"],
                "ТП1 СГП В2": ["ТП1 СГП В2"],
                "ТП1 ППУ В1": ["ТП1 ППУ В1"],
                "ТП1 ППУ В2": ["ТП1 ППУ В2"]
            },
            "ТП2": {
                "ТП2 Цэх1 В1": ["ТП2 Цэх1 В1"],
                "ТП2 Цэх1 В2": ["ТП2 Цэх1 В2"],
                "ТП2 ГРЩ В1": ["ТП2 ГРЩ В1"],
                "ТП2 ГРЩ В2": ["ТП2 ГРЩ В2"]
            },
            "ТП3": {
                "ТП3 Элеватор В1": ["ТП3 Элеватор В1"],
                "ТП3 Элеватор В2": ["ТП3 Элеватор В2"],
                "ТП3 Лузговая В1": ["ТП3 Лузговая В1"],
                "ТП3 Лузговая В2": ["ТП3 Лузговая В2"],
                "ТП3 ЛОС В1": ["ТП3 ЛОС В1"],
                "ТП3 ЛОС В2": ["ТП3 ЛОС В2"],
                "ТП3 КНС В1": ["ТП3 КНС В1"],
                "ТП3 КНС В2": ["ТП3 КНС В2"],
                "ТП3 Насосная В1": ["ТП3 Насосная В1"],
                "ТП3 Насосная В2": ["ТП3 Насосная В2"],
                "ТП3 Газовая котельная В1": ["ТП3 Газовая котельная В1"],
                "ТП3 Газовая котельная В2": ["ТП3 Газовая котельная В2"]
            },
            "ТП4": {
                "ТП4 Элеватор В1": ["ТП4 Элеватор В1"],
                "ТП4 Элеватор В2": ["ТП4 Элеватор В2"],
                "ТП4 ЛОС В1": ["ТП4 ЛОС В1"],
                "ТП4 ЛОС В2": ["ТП4 ЛОС В2"]
            },
            "ТП5": {
                "ТП5 ССиТ В1": ["ТП5 ССиТ В1"],
                "ТП5 ССиТ В2": ["ТП5 ССиТ В2"]
            },
            "ТП6": {
                "ТП6 ГЦ В1": ["ТП6 ГЦ В1"],
                "ТП6 ГЦ В2": ["ТП6 ГЦ В2"]
            },
            "ТП7": {
                "ТП7 СТЗ": ["ТП7 СТЗ"]
            },
            "ГПУ КТПН": {
                "КТПН ГПУ →": ["КТПН ГПУ →"],
                "КТПН ГПУ ←": ["КТПН ГПУ ←"],
                "КТПН ГПУ1 →": ["КТПН ГПУ1 →"],
                "КТПН ГПУ1 ←": ["КТПН ГПУ1 ←"],
                "КТПН ГПУ2 →": ["КТПН ГПУ2 →"],
                "КТПН ГПУ2 ←": ["КТПН ГПУ2 ←"],
                "КТПН ГПУ3 →": ["КТПН ГПУ3 →"],
                "КТПН ГПУ3 ←": ["КТПН ГПУ3 ←"],
                "КТПН ГПУ4 →": ["КТПН ГПУ4 →"],
                "КТПН ГПУ4 ←": ["КТПН ГПУ4 ←"]
            },
            "Радуга Склад": {
                "ШВ-6": ["ШВ-6"],
                "ШС-6": ["ШС-6"],
                "ЩС Зар В1": ["ЩС Зар В1"],
                "ЩС Зар В2": ["ЩС Зар В2"]
            }
        }
        ALL_ELE_COUNTERS = []
        for level1 in ELE_MENU:
            for level2 in ELE_MENU[level1]:
                for counter in ELE_MENU[level1][level2]:
                    ALL_ELE_COUNTERS.append(counter)
                    ELE_COUNTER_TO_PATH[counter] = (level1, level2, counter)
    
    # Загружаем меню для Ресурс
    if os.path.exists(MENU_RES_FILE):
        RES_MENU, ALL_RES_COUNTERS, RES_COUNTER_TO_PATH = parse_menu_from_excel(MENU_RES_FILE, "Меню Ресурс")
    else:
        # Встроенное меню по умолчанию
        RES_MENU = {
            "Цех1": {
                "РП3": ["РП3"]
            },
            "Цех3": {
                "ВРУ-1 (ТП-304)": ["ВРУ-1 (ТП-304)"],
                "ВРУ-1 (ТП-324)": ["ВРУ-1 (ТП-324)"],
                "ВРУ-2 (ТП-304)": ["ВРУ-2 (ТП-304)"],
                "ВРУ-2 (ТП-324)": ["ВРУ-2 (ТП-324)"]
            },
            "Цех4": {
                "ТЛ1 ШУ-3-1 ШУ6": ["ТЛ1 ШУ-3-1 ШУ6"],
                "ТЛ2 ШУ-3-2 ШУ6": ["ТЛ2 ШУ-3-2 ШУ6"],
                "ТЛ3 ШУ-3-1 ШУ5": ["ТЛ3 ШУ-3-1 ШУ5"],
                "Нории14-23 ШУ5": ["Нории14-23 ШУ5"]
            },
            "РМЦ": {
                "РМЦ": ["РМЦ"],
                "Карный участок": ["Карный участок"]
            },
            "Компрессорная": {
                "ПК3 ТП304 В1": ["ПК3 ТП304 В1"],
                "ПК3 ТП324 В2": ["ПК3 ТП324 В2"],
                "ПК2 Компр.4": ["ПК2 Компр.4"],
                "ПК2 Компр.5": ["ПК2 Компр.5"],
                "ПК2 Компр.6": ["ПК2 Компр.6"],
                "ПК1 Компр.7": ["ПК1 Компр.7"],
                "ПК1 Компр.8": ["ПК1 Компр.8"]
            },
            "РХУ": {
                "Скл. бестарного хранения": ["Скл. бестарного хранения"],
                "Хлопья Шулле": ["Хлопья Шулле"]
            },
            "Фасовка": {
                "Фасовка": ["Фасовка"],
                "ЩВ фасовка": ["ЩВ фасовка"]
            },
            "ССиТ": {
                "ВРУ4 ТП304 В1": ["ВРУ4 ТП304 В1"],
                "ВРУ4 ТП324 В2": ["ВРУ4 ТП324 В2"]
            },
            "Офис": {
                "Офис (Админ)": ["Офис (Админ)"]
            },
            "Столовая": {
                "ШР столовая": ["ШР столовая"]
            },
            "КГУ": {
                "Выработка КГУ": ["Выработка КГУ"]
            }
        }
        ALL_RES_COUNTERS = []
        for level1 in RES_MENU:
            for level2 in RES_MENU[level1]:
                for counter in RES_MENU[level1][level2]:
                    ALL_RES_COUNTERS.append(counter)
                    RES_COUNTER_TO_PATH[counter] = (level1, level2, counter)


# ============ ИНИЦИАЛИЗАЦИЯ EXCEL ============
def init_excel(file_path, counters):
    if not os.path.exists(file_path):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время записи"] + counters)
        wb.save(file_path)
        print(f"✅ Создан файл {file_path} с {len(counters)} счётчиками")
    else:
        wb = load_workbook(file_path)
        ws = wb.active
        headers = [str(cell.value) if cell.value else "" for cell in ws[1]]
        
        added = 0
        for counter in counters:
            if counter not in headers:
                ws.cell(1, ws.max_column + 1, value=counter)
                added += 1
        
        if added > 0:
            for row in range(2, ws.max_row + 1):
                for col in range(len(headers) + 1, ws.max_column + 1):
                    ws.cell(row, col, value=0)
            wb.save(file_path)
            print(f"✅ В файл {file_path} добавлено {added} новых счётчиков")
        wb.close()


def get_today_str():
    return datetime.now().strftime("%Y-%m-%d")


def ensure_today_exists(file_path, counters):
    today = get_today_str()
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        date_exists = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == today:
                date_exists = True
                break
        if not date_exists:
            ws.append([today, ""] + [0] * len(counters))
            wb.save(file_path)
        wb.close()
    except Exception as e:
        print(f"Ошибка: {e}")


def update_reading(file_path, counter_name, value, record_time, counters):
    try:
        today = get_today_str()
        ensure_today_exists(file_path, counters)
        
        wb = load_workbook(file_path)
        ws = wb.active
        headers = [str(cell.value) if cell.value else "" for cell in ws[1]]
        
        if counter_name not in headers:
            wb.close()
            return False
        
        col_idx = headers.index(counter_name) + 1
        
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == today:
                ws.cell(row, col_idx).value = float(value)
                ws.cell(row, 2).value = record_time
                wb.save(file_path)
                wb.close()
                return True
        wb.close()
        return False
    except Exception as e:
        print(f"Ошибка: {e}")
        return False


def get_all_data(file_path):
    try:
        wb = load_workbook(file_path)
        ws = wb.active
        headers = [str(cell.value) if cell.value else "" for cell in ws[1]][2:]
        result = []
        for row in range(2, ws.max_row + 1):
            date = ws.cell(row, 1).value
            record_time = ws.cell(row, 2).value
            if not date:
                continue
            row_data = {"date": date, "time": record_time, "counters": {}, "total": 0}
            for idx, counter in enumerate(headers, start=3):
                val = ws.cell(row, idx).value
                if val and isinstance(val, (int, float)) and val != 0:
                    row_data["counters"][counter] = val
                    row_data["total"] += val
            if row_data["counters"]:
                result.append(row_data)
        wb.close()
        return result
    except Exception as e:
        print(f"Ошибка: {e}")
        return []


def get_counters_with_data(file_path):
    data = get_all_data(file_path)
    counters_with_data = set()
    for day in data:
        for counter in day['counters'].keys():
            counters_with_data.add(counter)
    return list(counters_with_data)


# ============ EMAIL ============
async def send_email_report(file_path, object_name):
    if not os.path.exists(file_path):
        print(f"Файл {file_path} не найден")
        return False
    try:
        msg = EmailMessage()
        msg['Subject'] = f"Отчёт по {object_name} за {datetime.now().strftime('%d.%m.%Y')}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg.set_content(f"Отчёт по {object_name}\nДата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=f"{object_name}_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"Отчёт по {object_name} отправлен на {EMAIL_TO}")
        return True
    except Exception as e:
        print(f"Ошибка email: {e}")
        return False


# ============ ГЛАВНОЕ МЕНЮ ============
def get_main_menu():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(KeyboardButton("📊 ЭЭ ПЛК"))
    keyboard.add(KeyboardButton("📊 ЭЭ Ресурс"))
    keyboard.add(KeyboardButton("📁 Скачать Excel"))
    keyboard.add(KeyboardButton("❓ Помощь"))
    return keyboard


def get_object_keyboard():
    keyboard = InlineKeyboardMarkup(row_width=2)
    keyboard.add(InlineKeyboardButton("📊 ЭЭ ПЛК", callback_data="obj_ELE"))
    keyboard.add(InlineKeyboardButton("📊 ЭЭ Ресурс", callback_data="obj_RES"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


async def show_menu_level1(chat_id, object_type, menu, message_id=None):
    """Показывает первый уровень меню"""
    keyboard = InlineKeyboardMarkup(row_width=1)
    for level1 in menu.keys():
        keyboard.add(InlineKeyboardButton(level1, callback_data=f"{object_type}_l1_{level1}"))
    keyboard.add(InlineKeyboardButton("🔙 Назад", callback_data="back_to_main"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    
    text = f"📁 {object_type} - Выберите объект:"
    
    if message_id:
        await bot.edit_message_text(text, chat_id, message_id, reply_markup=keyboard)
    else:
        await bot.send_message(chat_id, text, reply_markup=keyboard)


async def show_menu_level2(chat_id, object_type, level1, menu, message_id=None):
    """Показывает второй уровень меню"""
    keyboard = InlineKeyboardMarkup(row_width=1)
    for level2 in menu[level1].keys():
        keyboard.add(InlineKeyboardButton(level2, callback_data=f"{object_type}_l2_{level1}_{level2}"))
    keyboard.add(InlineKeyboardButton("🔙 Назад", callback_data=f"{object_type}_back_l1"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    
    text = f"📁 {object_type} - {level1} → Выберите подгруппу:"
    
    if message_id:
        await bot.edit_message_text(text, chat_id, message_id, reply_markup=keyboard)
    else:
        await bot.send_message(chat_id, text, reply_markup=keyboard)


async def show_counters(chat_id, object_type, level1, level2, menu, message_id=None):
    """Показывает список счётчиков"""
    keyboard = InlineKeyboardMarkup(row_width=1)
    for counter in menu[level1][level2]:
        text = counter[:35] if len(counter) > 35 else counter
        keyboard.add(InlineKeyboardButton(text, callback_data=f"{object_type}_cnt_{counter}"))
    keyboard.add(InlineKeyboardButton("🔙 Назад", callback_data=f"{object_type}_back_l2_{level1}"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    
    text = f"📁 {object_type} - {level1} / {level2} → Выберите счётчик:"
    
    if message_id:
        await bot.edit_message_text(text, chat_id, message_id, reply_markup=keyboard)
    else:
        await bot.send_message(chat_id, text, reply_markup=keyboard)


# ============ ОБРАБОТЧИКИ КОМАНД ============
@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    load_menus()
    init_excel(EXCEL_ELE_FILE, ALL_ELE_COUNTERS)
    init_excel(EXCEL_RES_FILE, ALL_RES_COUNTERS)
    ensure_today_exists(EXCEL_ELE_FILE, ALL_ELE_COUNTERS)
    ensure_today_exists(EXCEL_RES_FILE, ALL_RES_COUNTERS)
    
    data_ele = get_all_data(EXCEL_ELE_FILE)
    data_res = get_all_data(EXCEL_RES_FILE)
    
    today = datetime.now().strftime("%d.%m.%Y")
    
    text = (
        f"🏭 ЭНЕРГОУЧЁТ\n\n"
        f"📅 Сегодня: {today}\n"
        f"📊 ЭЭ ПЛК: {len(data_ele)} записей, {len(ALL_ELE_COUNTERS)} счётчиков\n"
        f"📊 ЭЭ Ресурс: {len(data_res)} записей, {len(ALL_RES_COUNTERS)} счётчиков\n\n"
        f"💡 Нажмите '📊 ЭЭ ПЛК' или '📊 ЭЭ Ресурс'"
    )
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["help"])
@dp.message_handler(lambda message: message.text == "❓ Помощь")
async def help_command(message: types.Message):
    text = (
        "📘 ПОМОЩЬ И КОМАНДЫ\n\n"
        "ОСНОВНЫЕ КОМАНДЫ:\n"
        "/start - Главное меню\n"
        "/add - Ввести показания счётчика\n"
        "/stats - Показать статистику\n"
        "/file - Скачать Excel файл\n"
        "/help - Эта справка\n\n"
        "ДОПОЛНИТЕЛЬНЫЕ КОМАНДЫ:\n"
        "/test_email - Проверить отправку email\n"
        "/update - Обновить список счётчиков\n\n"
        "СТРУКТУРА:\n"
        "• Два объекта: ЭЭ ПЛК и ЭЭ Ресурс\n"
        "• Иерархическое меню (объект → группа → подгруппа → счётчик)\n"
        "• Все показания записываются на СЕГОДНЯ\n"
        "• Отчёт на email: каждый понедельник в 12:00\n\n"
        f"📊 ЭЭ ПЛК: {len(ALL_ELE_COUNTERS)} счётчиков\n"
        f"📊 ЭЭ Ресурс: {len(ALL_RES_COUNTERS)} счётчиков"
    )
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["test_email"])
async def test_email(message: types.Message):
    await message.answer("📧 Отправляю отчёты...")
    result1 = await send_email_report(EXCEL_ELE_FILE, "ЭЭ_ПЛК")
    result2 = await send_email_report(EXCEL_RES_FILE, "ЭЭ_Ресурс")
    if result1 or result2:
        await message.answer(f"✅ Письма отправлены! ПЛК: {result1}, Ресурс: {result2}")
    else:
        await message.answer("❌ Ошибка отправки")


@dp.message_handler(commands=["update"])
async def update_counters(message: types.Message):
    load_menus()
    init_excel(EXCEL_ELE_FILE, ALL_ELE_COUNTERS)
    init_excel(EXCEL_RES_FILE, ALL_RES_COUNTERS)
    await message.answer(f"✅ Обновлено! ПЛК: {len(ALL_ELE_COUNTERS)} счётчиков, Ресурс: {len(ALL_RES_COUNTERS)} счётчиков")


@dp.message_handler(commands=["add"])
async def add_reading(message: types.Message, state: FSMContext):
    await message.answer("📁 Выберите объект:", reply_markup=get_object_keyboard())
    await state.set_state(EnergyForm.choosing_object)


@dp.message_handler(lambda message: message.text == "📊 ЭЭ ПЛК")
async def add_ele(message: types.Message, state: FSMContext):
    await state.update_data(object_type="ELE", file_path=EXCEL_ELE_FILE, counters=ALL_ELE_COUNTERS)
    await show_menu_level1(message.chat.id, "ELE", ELE_MENU)
    await state.set_state(EnergyForm.choosing_level1)


@dp.message_handler(lambda message: message.text == "📊 ЭЭ Ресурс")
async def add_res(message: types.Message, state: FSMContext):
    await state.update_data(object_type="RES", file_path=EXCEL_RES_FILE, counters=ALL_RES_COUNTERS)
    await show_menu_level1(message.chat.id, "RES", RES_MENU)
    await state.set_state(EnergyForm.choosing_level1)


@dp.message_handler(lambda message: message.text == "📁 Скачать Excel")
async def send_excel_file(message: types.Message):
    keyboard = InlineKeyboardMarkup(row_width=2)
    keyboard.add(InlineKeyboardButton("📊 ЭЭ ПЛК", callback_data="download_ELE"))
    keyboard.add(InlineKeyboardButton("📊 ЭЭ Ресурс", callback_data="download_RES"))
    await message.answer("Выберите файл для скачивания:", reply_markup=keyboard)


# ============ CALLBACK ОБРАБОТЧИКИ ============
@dp.callback_query_handler(lambda c: c.data.startswith("download_"), state="*")
async def download_file(callback_query: types.CallbackQuery):
    object_type = callback_query.data.replace("download_", "")
    file_path = EXCEL_ELE_FILE if object_type == "ELE" else EXCEL_RES_FILE
    name = "ЭЭ ПЛК" if object_type == "ELE" else "ЭЭ Ресурс"
    
    if os.path.exists(file_path):
        doc = types.InputFile(file_path)
        await bot.send_document(callback_query.from_user.id, doc, caption=f"📊 {name}\n{datetime.now().strftime('%d.%m.%Y %H:%M')}")
    else:
        await bot.send_message(callback_query.from_user.id, "❌ Файл не создан")
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("ELE_l1_"), state=EnergyForm.choosing_level1)
async def ele_level1_selected(callback_query: types.CallbackQuery, state: FSMContext):
    level1 = callback_query.data.replace("ELE_l1_", "")
    await state.update_data(selected_level1=level1)
    await show_menu_level2(callback_query.from_user.id, "ELE", level1, ELE_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level2)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("RES_l1_"), state=EnergyForm.choosing_level1)
async def res_level1_selected(callback_query: types.CallbackQuery, state: FSMContext):
    level1 = callback_query.data.replace("RES_l1_", "")
    await state.update_data(selected_level1=level1)
    await show_menu_level2(callback_query.from_user.id, "RES", level1, RES_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level2)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("ELE_l2_"), state=EnergyForm.choosing_level2)
async def ele_level2_selected(callback_query: types.CallbackQuery, state: FSMContext):
    parts = callback_query.data.replace("ELE_l2_", "").split("_", 1)
    level1 = parts[0]
    level2 = parts[1]
    await state.update_data(selected_level2=level2)
    await show_counters(callback_query.from_user.id, "ELE", level1, level2, ELE_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_counter)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("RES_l2_"), state=EnergyForm.choosing_level2)
async def res_level2_selected(callback_query: types.CallbackQuery, state: FSMContext):
    parts = callback_query.data.replace("RES_l2_", "").split("_", 1)
    level1 = parts[0]
    level2 = parts[1]
    await state.update_data(selected_level2=level2)
    await show_counters(callback_query.from_user.id, "RES", level1, level2, RES_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_counter)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("ELE_cnt_"), state=EnergyForm.choosing_counter)
async def ele_counter_selected(callback_query: types.CallbackQuery, state: FSMContext):
    counter_name = callback_query.data.replace("ELE_cnt_", "")
    await state.update_data(selected_counter=counter_name)
    today = datetime.now().strftime("%d.%m.%Y")
    await callback_query.message.edit_text(
        f"✅ {counter_name}\n\n"
        f"📅 Дата: {today}\n"
        f"✏️ Введите показание (кВт·ч):"
    )
    await state.set_state(EnergyForm.entering_value)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("RES_cnt_"), state=EnergyForm.choosing_counter)
async def res_counter_selected(callback_query: types.CallbackQuery, state: FSMContext):
    counter_name = callback_query.data.replace("RES_cnt_", "")
    await state.update_data(selected_counter=counter_name)
    today = datetime.now().strftime("%d.%m.%Y")
    await callback_query.message.edit_text(
        f"✅ {counter_name}\n\n"
        f"📅 Дата: {today}\n"
        f"✏️ Введите показание (кВт·ч):"
    )
    await state.set_state(EnergyForm.entering_value)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == "ELE_back_l1", state=EnergyForm.choosing_level2)
async def ele_back_to_level1(callback_query: types.CallbackQuery, state: FSMContext):
    await show_menu_level1(callback_query.from_user.id, "ELE", ELE_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level1)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == "RES_back_l1", state=EnergyForm.choosing_level2)
async def res_back_to_level1(callback_query: types.CallbackQuery, state: FSMContext):
    await show_menu_level1(callback_query.from_user.id, "RES", RES_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level1)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("ELE_back_l2_"), state=EnergyForm.choosing_counter)
async def ele_back_to_level2(callback_query: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    level1 = data.get('selected_level1')
    await show_menu_level2(callback_query.from_user.id, "ELE", level1, ELE_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level2)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("RES_back_l2_"), state=EnergyForm.choosing_counter)
async def res_back_to_level2(callback_query: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    level1 = data.get('selected_level1')
    await show_menu_level2(callback_query.from_user.id, "RES", level1, RES_MENU, callback_query.message.message_id)
    await state.set_state(EnergyForm.choosing_level2)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == "back_to_main", state="*")
async def back_to_main(callback_query: types.CallbackQuery, state: FSMContext):
    await state.finish()
    await callback_query.message.edit_text("Главное меню:", reply_markup=get_main_menu())
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("obj_"), state=EnergyForm.choosing_object)
async def object_selected(callback_query: types.CallbackQuery, state: FSMContext):
    object_type = callback_query.data.replace("obj_", "")
    if object_type == "ELE":
        await state.update_data(object_type="ELE", file_path=EXCEL_ELE_FILE, counters=ALL_ELE_COUNTERS)
        await show_menu_level1(callback_query.from_user.id, "ELE", ELE_MENU, callback_query.message.message_id)
        await state.set_state(EnergyForm.choosing_level1)
    else:
        await state.update_data(object_type="RES", file_path=EXCEL_RES_FILE, counters=ALL_RES_COUNTERS)
        await show_menu_level1(callback_query.from_user.id, "RES", RES_MENU, callback_query.message.message_id)
        await state.set_state(EnergyForm.choosing_level1)
    await callback_query.answer()


@dp.message_handler(state=EnergyForm.entering_value)
async def value_entered(message: types.Message, state: FSMContext):
    try:
        value = float(message.text.replace(",", "."))
        if value < 0:
            await message.answer("❌ Введите положительное число")
            return
        
        data = await state.get_data()
        counter = data['selected_counter']
        file_path = data['file_path']
        counters = data['counters']
        object_type = data.get('object_type', 'ELE')
        record_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if update_reading(file_path, counter, value, record_time, counters):
            await message.answer(
                f"✅ Сохранено!\n\n"
                f"📊 {object_type}\n"
                f"🏭 {counter}\n"
                f"⚡ {value:,.2f} кВт·ч\n"
                f"⏰ {datetime.now().strftime('%H:%M:%S')}",
                reply_markup=get_main_menu()
            )
        else:
            await message.answer("❌ Ошибка при сохранении!", reply_markup=get_main_menu())
        
        await state.finish()
        
    except ValueError:
        await message.answer("❌ Введите число (например: 125.5)")


@dp.callback_query_handler(lambda c: c.data == "cancel", state="*")
async def cancel_action(callback_query: types.CallbackQuery, state: FSMContext):
    await state.finish()
    await callback_query.message.edit_text("❌ Отменено")
    await callback_query.message.answer("Главное меню:", reply_markup=get_main_menu())
    await callback_query.answer()


# ============ ЗАПУСК ============
if __name__ == "__main__":
    print("🚀 Бот Энергоучёт запущен!")
    print(f"📁 Файлы: {EXCEL_ELE_FILE}, {EXCEL_RES_FILE}")
    
    load_menus()
    init_excel(EXCEL_ELE_FILE, ALL_ELE_COUNTERS)
    init_excel(EXCEL_RES_FILE, ALL_RES_COUNTERS)
    ensure_today_exists(EXCEL_ELE_FILE, ALL_ELE_COUNTERS)
    ensure_today_exists(EXCEL_RES_FILE, ALL_RES_COUNTERS)
    
    print(f"🏭 ЭЭ ПЛК: {len(ALL_ELE_COUNTERS)} счётчиков")
    print(f"🏭 ЭЭ Ресурс: {len(ALL_RES_COUNTERS)} счётчиков")
    print("=" * 40)
    
    scheduler = AsyncIOScheduler(timezone="Europe/Moscow")
    scheduler.add_job(lambda: send_email_report(EXCEL_ELE_FILE, "ЭЭ_ПЛК"), 'cron', day_of_week='mon', hour=12, minute=0)
    scheduler.add_job(lambda: send_email_report(EXCEL_RES_FILE, "ЭЭ_Ресурс"), 'cron', day_of_week='mon', hour=12, minute=0)
    scheduler.start()
    print("⏰ Отчёты: каждый понедельник в 12:00 МСК")
    print("=" * 40)
    
    executor.start_polling(dp, skip_updates=True)
