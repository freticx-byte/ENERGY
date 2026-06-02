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

# ============ ТОКЕН БОТА =============
TOKEN = "8976307638:AAEyUMxOzc5Wy7JSHThXxPV_v1bbazZRSYQ"

# ============ СОЗДАНИЕ БОТА ============
storage = MemoryStorage()
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)

EXCEL_FILE = "Energy_rep_ELE.xlsx"

# ============ НАСТРОЙКИ EMAIL ============
EMAIL_TO = "a.misyunas@uvelka.ru"
EMAIL_FROM = "uvenergorusursy@gmail.com"
EMAIL_PASSWORD = "bmdt pzqh qdme wgnc"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# ============ ВСЕ СЧЁТЧИКИ (53 штуки) ============
ALL_COUNTERS = [
    "ЦРП В1", "ЦРП В2", "ЦРП ТП1, СШ1", "ЦРП ТП1, СШ2",
    "ЦРП ТП2, СШ1", "ЦРП ТП2, СШ2", "ЦРП ТП3, СШ1", "ЦРП ТП3, СШ2",
    'ЦРП КСО "Радуга"',
    "ТП1 ГРЩ. ППУ В1", "ТП1 ГРЩ. ППУ В2",
    "ТП1 РУ0,4 СГП В1", "ТП1 РУ0,4 СГП В2",
    "ТП2 ГРЩ В1", "ТП2 ГРЩ В2",
    "ТП2 РУ0,4 В1", "ТП2 РУ0,4 В2",
    "ТП3 РУ0,4 Элеватор В1", "ТП3 РУ0,4 Элеватор В2",
    "ТП3 Лузговая В1", "ТП3 Лузговая В2",
    "ТП3 Элеваторный В1", "ТП3 Элеваторный В2",
    "ТП4 Элеватор В1", "ТП4 Элеватор В2",
    "ТП5 ССиТ В1", "ТП5 ССиТ В2",
    "ТП6 ГЦ В1", "ТП6 ГЦ В2",
    "ТП7 СТЗ В1",
    "КТПН ГПУ Вход", "КТПН ГПУ Выход",
    "ТП Луговская", "Насосная В1", "Насосная В2",
    "ЛОС В1", "ЛОС В2", "КНС В1", "КНС В2",
    "Склад газации", "Теплосети ИТП",
    "Газовая котельная №1", "Газовая котельная №2", "Временно ТП-3"
]

# ============ ГРУППЫ СЧЁТЧИКОВ ============
COUNTER_GROUPS = {
    "ЦРП (Вводы и ТП)": ALL_COUNTERS[0:9],
    "ТП1 (ГРЩ и СГП)": ALL_COUNTERS[9:15],
    "ТП2 (ГРЩ и цеха)": ALL_COUNTERS[15:21],
    "ТП3 (Элеватор, котельная)": ALL_COUNTERS[21:29],
    "ТП4, ТП5, ТП6, ТП7": ALL_COUNTERS[29:38],
    "КТПН и прочие": ALL_COUNTERS[38:53],
}


class EnergyForm(StatesGroup):
    choosing_group = State()
    choosing_counter = State()
    entering_value = State()


# ============ ГЛАВНОЕ МЕНЮ ============
def get_main_menu():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(KeyboardButton("📝 Ввести показания"))
    keyboard.add(KeyboardButton("📊 Статистика"))
    keyboard.add(KeyboardButton("📋 Все счётчики"))
    keyboard.add(KeyboardButton("✅ Счётчики с данными"))
    keyboard.add(KeyboardButton("📁 Скачать Excel"))
    keyboard.add(KeyboardButton("❓ Помощь"))
    return keyboard


# ============ ФУНКЦИИ EMAIL =============
async def send_email_report():
    if not os.path.exists(EXCEL_FILE):
        print("Файл не найден")
        return False
    try:
        msg = EmailMessage()
        msg['Subject'] = f"Отчёт по энергоучёту за {datetime.now().strftime('%d.%m.%Y')}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg.set_content(f"Отчёт по потреблению электроэнергии.\nДата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        with open(EXCEL_FILE, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=f"Energy_rep_ELE_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)
        print(f"Отчёт отправлен на {EMAIL_TO}")
        return True
    except Exception as e:
        print(f"Ошибка email: {e}")
        return False
        
# ============ РАБОТА С EXCEL ============
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время записи"] + ALL_COUNTERS)
        wb.save(EXCEL_FILE)
        print(f"✅ Создан файл с {len(ALL_COUNTERS)} счётчиками")


def get_today_str():
    return datetime.now().strftime("%Y-%m-%d")


def ensure_today_exists():
    today = get_today_str()
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        date_exists = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == today:
                date_exists = True
                break
        if not date_exists:
            ws.append([today, ""] + [0] * len(ALL_COUNTERS))
            wb.save(EXCEL_FILE)
        wb.close()
    except Exception as e:
        print(f"Ошибка: {e}")


def update_reading(counter_name, value, record_time):
    try:
        today = get_today_str()
        ensure_today_exists()
        wb = load_workbook(EXCEL_FILE)
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
                wb.save(EXCEL_FILE)
                wb.close()
                return True
        wb.close()
        return False
    except Exception as e:
        print(f"Ошибка: {e}")
        return False


def get_all_data():
    try:
        wb = load_workbook(EXCEL_FILE)
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


def get_counters_with_data():
    data = get_all_data()
    counters_with_data = set()
    for day in data:
        for counter in day['counters'].keys():
            counters_with_data.add(counter)
    return list(counters_with_data)


# ============ КЛАВИАТУРЫ ============
async def groups_keyboard():
    keyboard = InlineKeyboardMarkup(row_width=1)
    for group_name in COUNTER_GROUPS.keys():
        keyboard.add(InlineKeyboardButton(group_name, callback_data=f"group_{group_name}"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


async def counters_keyboard(group_name):
    keyboard = InlineKeyboardMarkup(row_width=1)
    counters = COUNTER_GROUPS.get(group_name, [])
    for counter in counters:
        text = counter[:35] if len(counter) > 35 else counter
        keyboard.add(InlineKeyboardButton(text, callback_data=f"cnt_{counter}"))
    keyboard.add(InlineKeyboardButton("🔙 Назад к группам", callback_data="back_to_groups"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


# ============ ОБРАБОТЧИКИ КОМАНД ============
@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    init_excel()
    ensure_today_exists()
    data = get_all_data()
    counters_with_data = get_counters_with_data()
    today = datetime.now().strftime("%d.%m.%Y")
    
    text = (
        f"🏭 ЭНЕРГОУЧЁТ\n\n"
        f"📅 Сегодня: {today}\n"
        f"📊 Записей: {len(data)}\n"
        f"🏭 Счётчиков: {len(ALL_COUNTERS)}\n"
        f"✅ Счётчиков с данными: {len(counters_with_data)}\n\n"
        f"💡 Нажмите '📝 Ввести показания'"
    )
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["test_email"])
async def test_email(message: types.Message):
    await message.answer("📧 Отправляю...")
    result = await send_email_report()
    if result:
        await message.answer("✅ Письмо отправлено!")
    else:
        await message.answer("❌ Ошибка")


@dp.message_handler(commands=["update"])
async def update_counters(message: types.Message):
    """Обновляет список счётчиков в Excel файле"""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        current_headers = [str(cell.value) if cell.value else "" for cell in ws[1]]
        
        new_counters = []
        for counter in ALL_COUNTERS:
            if counter not in current_headers:
                new_counters.append(counter)
        
        if new_counters:
            for counter in new_counters:
                ws.cell(1, ws.max_column + 1, value=counter)
            for row in range(2, ws.max_row + 1):
                for col in range(len(current_headers) + 1, ws.max_column + 1):
                    ws.cell(row, col, value=0)
            wb.save(EXCEL_FILE)
            await message.answer(f"✅ Добавлено {len(new_counters)} новых счётчиков")
        else:
            await message.answer("✅ Все счётчики уже есть")
        wb.close()
    except Exception as e:
        await message.answer(f"❌ Ошибка: {e}")


@dp.message_handler(lambda message: message.text == "📝 Ввести показания")
@dp.message_handler(commands=["add"])
async def add_button(message: types.Message, state: FSMContext):
    await message.answer("📁 Выберите группу:", reply_markup=await groups_keyboard())
    await state.set_state(EnergyForm.choosing_group)


@dp.callback_query_handler(lambda c: c.data.startswith("group_"), state=EnergyForm.choosing_group)
async def group_selected(callback_query: types.CallbackQuery, state: FSMContext):
    group_name = callback_query.data.replace("group_", "")
    await state.update_data(selected_group=group_name)
    await callback_query.message.edit_text(
        f"📁 {group_name}\n\n👇 Выберите счётчик:",
        reply_markup=await counters_keyboard(group_name)
    )
    await state.set_state(EnergyForm.choosing_counter)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == "back_to_groups", state=EnergyForm.choosing_counter)
async def back_to_groups(callback_query: types.CallbackQuery, state: FSMContext):
    await callback_query.message.edit_text("📁 Выберите группу:", reply_markup=await groups_keyboard())
    await state.set_state(EnergyForm.choosing_group)
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("cnt_"), state=EnergyForm.choosing_counter)
async def counter_selected(callback_query: types.CallbackQuery, state: FSMContext):
    counter_name = callback_query.data.replace("cnt_", "")
    await state.update_data(selected_counter=counter_name)
    today = datetime.now().strftime("%d.%m.%Y")
    await callback_query.message.edit_text(
        f"✅ {counter_name}\n\n"
        f"📅 Дата: {today}\n"
        f"✏️ Введите показание (кВт·ч):"
    )
    await state.set_state(EnergyForm.entering_value)
    await callback_query.answer()


@dp.message_handler(state=EnergyForm.entering_value)
async def value_entered(message: types.Message, state: FSMContext):
    try:
        value = float(message.text.replace(",", "."))
        if value < 0:
            await message.answer("❌ Введите положительное число")
            return
        data = await state.get_data()
        record_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if update_reading(data['selected_counter'], value, record_time):
            await message.answer(
                f"✅ Сохранено!\n\n"
                f"🏭 {data['selected_counter']}\n"
                f"⚡ {value:,.2f} кВт·ч\n"
                f"⏰ {datetime.now().strftime('%H:%M:%S')}",
                reply_markup=get_main_menu()
            )
        else:
            await message.answer("❌ Ошибка!", reply_markup=get_main_menu())
        await state.finish()
    except ValueError:
        await message.answer("❌ Введите число")


@dp.message_handler(lambda message: message.text == "📊 Статистика")
@dp.message_handler(commands=["stats"])
async def show_stats(message: types.Message):
    data = get_all_data()
    if not data:
        await message.answer("📊 Нет данных", reply_markup=get_main_menu())
        return
    text = "📊 СТАТИСТИКА\n\n"
    total = 0
    for day in data[-10:]:
        text += f"📅 {day['date']}: {day['total']:,.2f} кВт·ч\n"
        total += day['total']
    text += f"\n💰 ИТОГО: {total:,.2f} кВт·ч"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(lambda message: message.text == "📋 Все счётчики")
@dp.message_handler(commands=["counters"])
async def show_all_counters(message: types.Message):
    text = "📋 ВСЕ СЧЁТЧИКИ:\n\n"
    for group, counters in COUNTER_GROUPS.items():
        text += f"📁 {group}:\n"
        for i, c in enumerate(counters, 1):
            text += f"   {i}. {c}\n"
        text += "\n"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(lambda message: message.text == "✅ Счётчики с данными")
@dp.message_handler(commands=["data_counters"])
async def show_counters_with_data(message: types.Message):
    counters = get_counters_with_data()
    if not counters:
        await message.answer("✅ Нет данных", reply_markup=get_main_menu())
        return
    text = "✅ СЧЁТЧИКИ С ДАННЫМИ:\n\n"
    for i, c in enumerate(counters, 1):
        text += f"{i}. {c}\n"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(lambda message: message.text == "📁 Скачать Excel")
@dp.message_handler(commands=["file"])
async def send_excel_file(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        doc = types.InputFile(EXCEL_FILE)
        await message.answer_document(doc, caption=f"📊 Energy_rep_ELE\n{datetime.now().strftime('%d.%m.%Y %H:%M')}")
    else:
        await message.answer("❌ Файл не создан")


@dp.message_handler(lambda message: message.text == "❓ Помощь")
@dp.message_handler(commands=["help"])
async def help_command(message: types.Message):
    text = (
        "📘 *ПОМОЩЬ И КОМАНДЫ*\n\n"
        "📌 *ОСНОВНЫЕ КОМАНДЫ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "🏠 `/start` - Главное меню\n"
        "📝 `/add` - Ввести показания счётчика\n"
        "📊 `/stats` - Показать статистику\n"
        "📋 `/counters` - Список всех счётчиков\n"
        "✅ `/data_counters` - Счётчики с показаниями\n"
        "📁 `/file` - Скачать Excel файл\n"
        "❓ `/help` - Эта справка\n\n"
        "📌 *ДОПОЛНИТЕЛЬНЫЕ КОМАНДЫ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "📧 `/test_email` - Проверить отправку email\n"
        "🔄 `/update` - Обновить список счётчиков\n\n"
        "📌 *КАК ВВЕСТИ ПОКАЗАНИЯ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "1️⃣ Нажмите '📝 Ввести показания' или /add\n"
        "2️⃣ Выберите группу счётчиков\n"
        "3️⃣ Выберите конкретный счётчик\n"
        "4️⃣ Введите показание в кВт·ч\n\n"
        "📌 *СТАТИСТИКА:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "• Все показания записываются на СЕГОДНЯ\n"
        "• Время записи фиксируется автоматически\n"
        "• Отчёт на email: каждый понедельник в 12:00\n\n"
        "📌 *ВСЕГО СЧЁТЧИКОВ:* 53\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "📁 Скачать файл: /file (Energy_rep_ELE.xlsx)\n"
        "📧 Проверить email: /test_email"
    )
    await message.answer(text, parse_mode="Markdown", reply_markup=get_main_menu())


@dp.callback_query_handler(lambda c: c.data == "cancel", state="*")
async def cancel_action(callback_query: types.CallbackQuery, state: FSMContext):
    await state.finish()
    await callback_query.message.edit_text("❌ Отменено")
    await callback_query.message.answer("Главное меню:", reply_markup=get_main_menu())
    await callback_query.answer()


# ============ ЗАПУСК ============
if __name__ == "__main__":
    print("🚀 Бот Энергоучёт запущен!")
    print(f"📁 Файл: {EXCEL_FILE}")
    print(f"🏭 Счётчиков: {len(ALL_COUNTERS)}")
    print("=" * 40)
    
    init_excel()
    ensure_today_exists()
    
    scheduler = AsyncIOScheduler(timezone="Europe/Moscow")
    scheduler.add_job(send_email_report, 'cron', day_of_week='mon', hour=12, minute=0)
    scheduler.start()
    print("⏰ Отчёт: каждый понедельник в 12:00 МСК")
    print("=" * 40)
    
    executor.start_polling(dp, skip_updates=True)
