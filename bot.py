import asyncio
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Command
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.utils import executor
from openpyxl import Workbook, load_workbook
from apscheduler.schedulers.asyncio import AsyncIOScheduler

TOKEN = "8976307638:AAGZNiGdfhYeYTjWVvWS2g3bAFM5RiwLi1g"

storage = MemoryStorage()
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)

EXCEL_FILE = "Energy_rep_ELE.xlsx"

EMAIL_TO = "a.misyunas@uvelka.ru"
EMAIL_FROM = "uvenergorusursy@gmail.com"
EMAIL_PASSWORD = "bmdt pzqh qdme wgnc"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# ============ МЕНЮ (2 УРОВНЯ: ГРУППА → СЧЁТЧИК) ============
MENU = {
    "ЦРП 10кВ": ["ЦРП В1", "ЦРП В2", "ЦРП ТП1, СШ1", "ЦРП ТП1, СШ2",
                 "ЦРП ТП2, СШ1", "ЦРП ТП2, СШ2", "ЦРП ТП3, СШ1", "ЦРП ТП3, СШ2",
                 'ЦРП КСО "Радуга"'],
    "ТП1": ["ТП1 ГРЩ. ППУ В1", "ТП1 ГРЩ. ППУ В2", "ТП1 СГП В1", "ТП1 СГП В2"],
    "ТП2": ["ТП2 ГРЩ В1", "ТП2 ГРЩ В2", "ТП2 Цэх1 В1", "ТП2 Цэх1 В2"],
    "ТП3": [
        "ТП3 Элеватор В1", "ТП3 Элеватор В2",
        "ТП3 Лузговая В1", "ТП3 Лузговая В2",
        "ТП3 ЛОС В1", "ТП3 ЛОС В2",
        "ТП3 КНС В1", "ТП3 КНС В2",
        "ТП3 Насосная В1", "ТП3 Насосная В2",
        "ТП3 Газовая котельная В1", "ТП3 Газовая котельная В2"
    ],
    "ТП4": ["ТП4 Элеватор В1", "ТП4 Элеватор В2", "ТП4 ЛОС В1", "ТП4 ЛОС В2"],
    "ТП5": ["ТП5 ССиТ В1", "ТП5 ССиТ В2"],
    "ТП6": ["ТП6 ГЦ В1", "ТП6 ГЦ В2"],
    "ТП7": ["ТП7 СТЗ"],
    "ГПУ КТПН": [
        "КТПН ГПУ →", "КТПН ГПУ ←",
        "КТПН ГПУ1 →", "КТПН ГПУ1 ←",
        "КТПН ГПУ2 →", "КТПН ГПУ2 ←",
        "КТПН ГПУ3 →", "КТПН ГПУ3 ←",
        "КТПН ГПУ4 →", "КТПН ГПУ4 ←"
    ],
    "Радуга Склад": ["ШВ-6", "ШС-6", "ЩС Зар В1", "ЩС Зар В2"],
    "Теплосети": ["Теплосети ИТП"],
    "Прочее": ["ТП Луговская", "Склад газации", "Временно ТП-3"]
}

ALL_COUNTERS = []
for counters in MENU.values():
    ALL_COUNTERS.extend(counters)


class EnergyForm(StatesGroup):
    choosing_group = State()
    entering_value = State()


# ============ ФУНКЦИИ EXCEL ============
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время записи"] + ALL_COUNTERS)
        wb.save(EXCEL_FILE)


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
    except:
        pass


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
    except:
        return False


def get_all_data():
    try:
        if not os.path.exists(EXCEL_FILE):
            return []
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        result = []
        for row in range(2, ws.max_row + 1):
            date = ws.cell(row, 1).value
            if not date:
                continue
            total = 0
            for col in range(3, ws.max_column + 1):
                val = ws.cell(row, col).value
                if val and isinstance(val, (int, float)) and val != 0:
                    total += val
            if total > 0:
                result.append({"date": date, "total": total})
        wb.close()
        return result
    except:
        return []


def get_counters_with_data():
    try:
        if not os.path.exists(EXCEL_FILE):
            return []
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        headers = [str(cell.value) if cell.value else "" for cell in ws[1]][2:]
        counters_with_data = set()
        for row in range(2, ws.max_row + 1):
            for idx, counter in enumerate(headers, start=3):
                val = ws.cell(row, idx).value
                if val and isinstance(val, (int, float)) and val != 0:
                    counters_with_data.add(counter)
        wb.close()
        return list(counters_with_data)
    except:
        return []


# ============ EMAIL ============
async def send_email_report():
    if not os.path.exists(EXCEL_FILE):
        return False
    try:
        msg = EmailMessage()
        msg['Subject'] = f"Отчёт за {datetime.now().strftime('%d.%m.%Y')}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg.set_content(f"Отчёт по потреблению\nДата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        with open(EXCEL_FILE, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=f"Energy_rep_ELE_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)
        return True
    except:
        return False


# ============ КЛАВИАТУРЫ ============
def get_main_menu():
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.add(KeyboardButton("📝 Ввести показания"))
    keyboard.add(KeyboardButton("📊 Статистика"))
    keyboard.add(KeyboardButton("📋 Все счётчики"))
    keyboard.add(KeyboardButton("✅ Счётчики с данными"))
    keyboard.add(KeyboardButton("📁 Скачать Excel"))
    keyboard.add(KeyboardButton("❓ Помощь"))
    return keyboard


def get_group_keyboard():
    keyboard = InlineKeyboardMarkup(row_width=1)
    for group_name in MENU.keys():
        keyboard.add(InlineKeyboardButton(group_name, callback_data=f"group_{group_name}"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


def get_counters_keyboard(group_name):
    keyboard = InlineKeyboardMarkup(row_width=1)
    for counter in MENU[group_name]:
        text = counter[:40] if len(counter) > 40 else counter
        keyboard.add(InlineKeyboardButton(text, callback_data=f"cnt_{counter}"))
    keyboard.add(InlineKeyboardButton("🔙 Назад к группам", callback_data="back_to_groups"))
    keyboard.add(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    return keyboard


# ============ ОБРАБОТЧИКИ ============
@dp.message_handler(commands=["start"])
async def start(message: types.Message):
    init_excel()
    ensure_today_exists()
    data = get_all_data()
    counters_with_data = get_counters_with_data()
    
    text = (
        f"🏭 ЭНЕРГОУЧЁТ\n\n"
        f"📅 Сегодня: {datetime.now().strftime('%d.%m.%Y')}\n"
        f"📊 Записей: {len(data)}\n"
        f"🏭 Счётчиков: {len(ALL_COUNTERS)}\n"
        f"✅ Счётчиков с данными: {len(counters_with_data)}\n\n"
        f"💡 Нажмите '📝 Ввести показания'"
    )
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["help"])
@dp.message_handler(lambda message: message.text == "❓ Помощь")
async def help_command(message: types.Message):
    text = (
        "📘 ПОМОЩЬ\n\n"
        "/start - Главное меню\n"
        "📝 Ввести показания - Выбрать счётчик\n"
        "📊 Статистика - Общая статистика\n"
        "📋 Все счётчики - Список всех\n"
        "✅ Счётчики с данными - Только с показаниями\n"
        "📁 Скачать Excel - Скачать файл\n\n"
        f"📊 Всего счётчиков: {len(ALL_COUNTERS)}"
    )
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(lambda message: message.text == "📝 Ввести показания")
@dp.message_handler(commands=["add"])
async def add_reading(message: types.Message, state: FSMContext):
    await message.answer("📁 Выберите группу счётчиков:", reply_markup=get_group_keyboard())
    await state.set_state(EnergyForm.choosing_group)


@dp.callback_query_handler(lambda c: c.data.startswith("group_"), state=EnergyForm.choosing_group)
async def group_selected(callback_query: types.CallbackQuery, state: FSMContext):
    group_name = callback_query.data.replace("group_", "")
    await state.update_data(selected_group=group_name)
    await callback_query.message.edit_text(
        f"📁 {group_name}\n\n👇 Выберите счётчик:",
        reply_markup=get_counters_keyboard(group_name)
    )
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data == "back_to_groups", state=EnergyForm.choosing_group)
async def back_to_groups(callback_query: types.CallbackQuery, state: FSMContext):
    await callback_query.message.edit_text(
        "📁 Выберите группу счётчиков:",
        reply_markup=get_group_keyboard()
    )
    await callback_query.answer()


@dp.callback_query_handler(lambda c: c.data.startswith("cnt_"), state=EnergyForm.choosing_group)
async def counter_selected(callback_query: types.CallbackQuery, state: FSMContext):
    counter_name = callback_query.data.replace("cnt_", "")
    await state.update_data(selected_counter=counter_name)
    today = datetime.now().strftime("%d.%m.%Y")
    await callback_query.message.edit_text(
        f"✅ Выбран счётчик: {counter_name}\n\n"
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
        counter = data['selected_counter']
        group = data['selected_group']
        record_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if update_reading(counter, value, record_time):
            # Отправляем подтверждение и сразу возвращаемся к списку счётчиков
            await message.answer(
                f"✅ Сохранено!\n\n"
                f"🏭 {counter}\n"
                f"⚡ {value:,.2f} кВт·ч"
            )
            # Возвращаемся к выбору счётчика в той же группе (НОВЫМ СООБЩЕНИЕМ)
            await message.answer(
                f"📁 {group}\n\n👇 Выберите следующий счётчик:",
                reply_markup=get_counters_keyboard(group)
            )
            await state.set_state(EnergyForm.choosing_group)
        else:
            await message.answer("❌ Ошибка при сохранении!", reply_markup=get_main_menu())
            await state.finish()
        
    except ValueError:
        await message.answer("❌ Введите число (например: 125.5)")


@dp.message_handler(commands=["stats"])
@dp.message_handler(lambda message: message.text == "📊 Статистика")
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


@dp.message_handler(commands=["counters"])
@dp.message_handler(lambda message: message.text == "📋 Все счётчики")
async def show_all_counters(message: types.Message):
    text = "📋 ВСЕ СЧЁТЧИКИ:\n\n"
    for group, counters in MENU.items():
        text += f"📁 {group}:\n"
        for i, c in enumerate(counters, 1):
            text += f"   {i}. {c}\n"
        text += "\n"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["data_counters"])
@dp.message_handler(lambda message: message.text == "✅ Счётчики с данными")
async def show_counters_with_data(message: types.Message):
    counters = get_counters_with_data()
    if not counters:
        await message.answer("✅ Нет счётчиков с данными", reply_markup=get_main_menu())
        return
    
    text = "✅ СЧЁТЧИКИ С ДАННЫМИ:\n\n"
    for i, c in enumerate(counters, 1):
        text += f"{i}. {c}\n"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message_handler(commands=["file"])
@dp.message_handler(lambda message: message.text == "📁 Скачать Excel")
async def send_excel_file(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        doc = types.InputFile(EXCEL_FILE)
        await message.answer_document(doc, caption=f"📊 Файл учёта\n{datetime.now().strftime('%d.%m.%Y %H:%M')}")
    else:
        await message.answer("❌ Файл не создан")


@dp.message_handler(commands=["test_email"])
async def test_email(message: types.Message):
    await message.answer("📧 Отправляю...")
    result = await send_email_report()
    await message.answer(f"✅ Результат: {result}")


@dp.callback_query_handler(lambda c: c.data == "cancel", state="*")
async def cancel_action(callback_query: types.CallbackQuery, state: FSMContext):
    await state.finish()
    await callback_query.message.delete()
    await callback_query.message.answer("❌ Отменено", reply_markup=get_main_menu())
    await callback_query.answer()


if __name__ == "__main__":
    print("🚀 Бот запущен!")
    print(f"📁 Файл: {EXCEL_FILE}")
    print(f"🏭 Счётчиков: {len(ALL_COUNTERS)}")
    
    init_excel()
    ensure_today_exists()
    
    scheduler = AsyncIOScheduler(timezone="Europe/Moscow")
    scheduler.add_job(send_email_report, 'cron', day_of_week='mon', hour=12, minute=0)
    scheduler.start()
    
    executor.start_polling(dp, skip_updates=True)
