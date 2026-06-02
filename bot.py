import asyncio
import os
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, F, types
from aiogram.dispatcher.filters import Command
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram import executor

from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from aiogram.types import FSInputFile, KeyboardButton
from openpyxl import Workbook, load_workbook
from apscheduler.schedulers.asyncio import AsyncIOScheduler

# ============ ТОКЕН БОТА ============
TOKEN = "8976307638:AAEyUMxOzc5Wy7JSHThXxPV_v1bbazZRSYQ"
bot = Bot(token=TOKEN)
dp = Dispatcher()

EXCEL_FILE = "Ж учета энергоресурсов.xlsx"

# ============ НАСТРОЙКИ EMAIL ============
EMAIL_TO = "a.misyunas@uvelka.ru"  # ← ИСПРАВЛЕНО
EMAIL_FROM = "uvenergorusursy@gmail.com"
EMAIL_PASSWORD = "bmdt pzqh qdme wgnc"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# ============ ВСЕ СЧЁТЧИКИ (45 штук) ============
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
    "ЦРП (Вводы и ТП)": [
        "ЦРП В1", "ЦРП В2", "ЦРП ТП1, СШ1", "ЦРП ТП1, СШ2",
        "ЦРП ТП2, СШ1", "ЦРП ТП2, СШ2", "ЦРП ТП3, СШ1", "ЦРП ТП3, СШ2",
        'ЦРП КСО "Радуга"'
    ],
    "ТП1 (ГРЩ и СГП)": [
        "ТП1 ГРЩ. ППУ В1", "ТП1 ГРЩ. ППУ В2",
        "ТП1 РУ0,4 СГП В1", "ТП1 РУ0,4 СГП В2"
    ],
    "ТП2 (ГРЩ и цеха)": [
        "ТП2 ГРЩ В1", "ТП2 ГРЩ В2",
        "ТП2 РУ0,4 В1", "ТП2 РУ0,4 В2"
    ],
    "ТП3 (Элеватор, котельная)": [
        "ТП3 РУ0,4 Элеватор В1", "ТП3 РУ0,4 Элеватор В2",
        "ТП3 Лузговая В1", "ТП3 Лузговая В2",
        "ТП3 Элеваторный В1", "ТП3 Элеваторный В2"
    ],
    "ТП4, ТП5, ТП6, ТП7": [
        "ТП4 Элеватор В1", "ТП4 Элеватор В2",
        "ТП5 ССиТ В1", "ТП5 ССиТ В2",
        "ТП6 ГЦ В1", "ТП6 ГЦ В2",
        "ТП7 СТЗ В1"
    ],
    "КТПН и прочие": [
        "КТПН ГПУ Вход", "КТПН ГПУ Выход",
        "ТП Луговская", "Насосная В1", "Насосная В2",
        "ЛОС В1", "ЛОС В2", "КНС В1", "КНС В2",
        "Склад газации", "Теплосети ИТП",
        "Газовая котельная №1", "Газовая котельная №2", "Временно ТП-3"
    ]
}


class EnergyForm(StatesGroup):
    choosing_group = State()
    choosing_counter = State()
    entering_value = State()


# ============ ГЛАВНОЕ МЕНЮ ============
def get_main_menu():
    builder = ReplyKeyboardBuilder()
    builder.add(KeyboardButton(text="📝 Ввести показания"))
    builder.add(KeyboardButton(text="📊 Статистика"))
    builder.add(KeyboardButton(text="📋 Все счётчики"))
    builder.add(KeyboardButton(text="✅ Счётчики с данными"))
    builder.add(KeyboardButton(text="📁 Скачать Excel"))
    builder.add(KeyboardButton(text="❓ Помощь"))
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)


# ============ ФУНКЦИИ EMAIL ============
async def send_email_report():
    """Отправляет Excel файл на email"""
    if not os.path.exists(EXCEL_FILE):
        print("Файл не найден для отправки")
        return False

    try:
        msg = EmailMessage()
        msg['Subject'] = f"Отчёт по энергоучёту за {datetime.now().strftime('%d.%m.%Y')}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        msg.set_content(
            f"Еженедельный отчёт по потреблению электроэнергии.\nДата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}")

        with open(EXCEL_FILE, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application',
                               subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               filename=f"energy_report_{datetime.now().strftime('%Y%m%d')}.xlsx")

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)

        print(f"Отчёт отправлен на {EMAIL_TO}")
        return True
    except Exception as e:
        print(f"Ошибка отправки email: {e}")
        return False


# ============ РАБОТА С EXCEL ============
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время записи"] + ALL_COUNTERS)
        wb.save(EXCEL_FILE)
        print(f"✅ Создан новый файл с {len(ALL_COUNTERS)} счётчиками")
    else:
        # Проверяем и добавляем недостающие колонки
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        headers = [str(cell.value) if cell.value else "" for cell in ws[1]]
        
        added = 0
        for counter in ALL_COUNTERS:
            if counter not in headers:
                ws.cell(1, ws.max_column + 1, value=counter)
                added += 1
        
        if added > 0:
            # Заполняем нулями существующие строки
            for row in range(2, ws.max_row + 1):
                for col in range(len(headers) + 1, ws.max_column + 1):
                    ws.cell(row, col, value=0)
            wb.save(EXCEL_FILE)
            print(f"✅ Добавлено {added} новых счётчиков")
        wb.close()


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
    builder = InlineKeyboardBuilder()
    for group_name in COUNTER_GROUPS.keys():
        builder.button(text=group_name, callback_data=f"group_{group_name}")
    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(1)
    return builder.as_markup()


async def counters_keyboard(group_name):
    builder = InlineKeyboardBuilder()
    counters = COUNTER_GROUPS.get(group_name, [])
    for counter in counters:
        text = counter[:35] if len(counter) > 35 else counter
        builder.button(text=text, callback_data=f"cnt_{counter}")
    builder.button(text="🔙 Назад к группам", callback_data="back_to_groups")
    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(1)
    return builder.as_markup()


# ============ КОМАНДЫ И ОБРАБОТЧИКИ ============
@dp.message(Command("start"))
async def start(message: types.Message):
    init_excel()
    ensure_today_exists()
    data = get_all_data()
    counters_with_data = get_counters_with_data()

    today = datetime.now().strftime("%d.%m.%Y")

    text = (
        f"🏭 ЭНЕРГОУЧЁТ\n\n"
        f"📅 Сегодня: {today}\n"
        f"📊 Записей всего: {len(data)}\n"
        f"🏭 Счётчиков: {len(ALL_COUNTERS)}\n"
        f"✅ Счётчиков с данными: {len(counters_with_data)}\n\n"
        f"💡 Нажмите кнопку '📝 Ввести показания'"
    )

    await message.answer(text, reply_markup=get_main_menu())


@dp.message(Command("help"))
async def help_command(message: types.Message):
    await message.answer(
        "📘 *ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ*\n\n"
        "📌 *ОСНОВНЫЕ КОМАНДЫ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "📝 `/add` - Ввести показания счётчика\n"
        "📊 `/stats` - Показать статистику потребления\n"
        "📋 `/counters` - Список всех счётчиков\n"
        "✅ `/data_counters` - Счётчики с показаниями\n"
        "📁 `/file` - Скачать Excel файл\n"
        "🔄 `/update` - Обновить список счётчиков\n"
        "📧 `/test_email` - Проверить отправку email\n"
        "❓ `/help` - Эта инструкция\n\n"
        "📌 *КАК ВВЕСТИ ПОКАЗАНИЯ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "1️⃣ Нажмите кнопку '📝 Ввести показания'\n"
        "2️⃣ Выберите группу счётчиков\n"
        "3️⃣ Выберите конкретный счётчик\n"
        "4️⃣ Введите показание в кВт·ч\n\n"
        "📌 *ГРУППЫ СЧЁТЧИКОВ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "• ЦРП (Вводы и ТП) - 9 счётчиков\n"
        "• ТП1 (ГРЩ и СГП) - 6 счётчиков\n"
        "• ТП2 (ГРЩ и цеха) - 6 счётчиков\n"
        "• ТП3 (Элеватор, котельная) - 8 счётчиков\n"
        "• ТП4, ТП5, ТП6, ТП7 - 9 счётчиков\n"
        "• КТПН и прочие - 15 счётчиков\n\n"
        "📌 *СТАТИСТИКА:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "• Всего счётчиков: 53\n"
        "• Все показания записываются на СЕГОДНЯ\n"
        "• Время записи фиксируется автоматически\n"
        "• Отчёт на email: каждый понедельник в 12:00\n\n"
        "📌 *ПОЛЕЗНЫЕ ССЫЛКИ:*\n"
        "━━━━━━━━━━━━━━━━━━━━━\n"
        "• Скачать Excel файл: /file\n"
        "• Проверить email: /test_email\n"
        "• Обновить счётчики: /update",
        parse_mode="Markdown",
        reply_markup=get_main_menu()
    )


@dp.message(Command("update"))
async def update_counters(message: types.Message):
    """Добавляет новые колонки в Excel файл"""
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
            await message.answer(f"✅ Добавлено {len(new_counters)} новых счётчиков:\n" + "\n".join(new_counters[:10]))
        else:
            await message.answer("✅ Все счётчики уже есть в таблице")
        
        wb.close()
    except Exception as e:
        await message.answer(f"❌ Ошибка: {e}")


@dp.message(Command("test_email"))
async def test_email(message: types.Message):
    await message.answer("📧 Отправляю тестовое письмо...")
    result = await send_email_report()
    if result:
        await message.answer("✅ Письмо отправлено! Проверь почту.")
    else:
        await message.answer("❌ Ошибка. Проверь настройки:\n"
                           f"Email отправителя: {EMAIL_FROM}\n"
                           f"Email получателя: {EMAIL_TO}")


@dp.message(F.text == "📝 Ввести показания")
@dp.message(Command("add"))
async def add_button(message: types.Message, state: FSMContext):
    await message.answer("📁 Выберите группу счётчиков:", reply_markup=await groups_keyboard())
    await state.set_state(EnergyForm.choosing_group)


@dp.callback_query(F.data.startswith("group_"))
async def group_selected(callback: types.CallbackQuery, state: FSMContext):
    group_name = callback.data.replace("group_", "")
    await state.update_data(selected_group=group_name)

    await callback.message.edit_text(
        f"📁 Выбрана группа: {group_name}\n\n👇 Выберите счётчик:",
        reply_markup=await counters_keyboard(group_name)
    )
    await state.set_state(EnergyForm.choosing_counter)
    await callback.answer()


@dp.callback_query(F.data == "back_to_groups")
async def back_to_groups(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.edit_text("📁 Выберите группу счётчиков:")
    await callback.message.answer("Группы:", reply_markup=await groups_keyboard())
    await state.set_state(EnergyForm.choosing_group)
    await callback.answer()


@dp.callback_query(F.data.startswith("cnt_"))
async def counter_selected(callback: types.CallbackQuery, state: FSMContext):
    counter_name = callback.data.replace("cnt_", "")
    await state.update_data(selected_counter=counter_name)

    today = datetime.now().strftime("%d.%m.%Y")

    await callback.message.edit_text(
        f"✅ Выбран счётчик: {counter_name}\n\n"
        f"📅 Дата записи: {today} (сегодня)\n"
        f"✏️ Введите показание в кВт·ч:"
    )
    await state.set_state(EnergyForm.entering_value)
    await callback.answer()


@dp.message(EnergyForm.entering_value)
async def value_entered(message: types.Message, state: FSMContext):
    try:
        value = float(message.text.replace(",", "."))
        if value < 0:
            await message.answer("❌ Введите положительное число")
            return

        data = await state.get_data()
        counter = data['selected_counter']
        record_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today = datetime.now().strftime("%d.%m.%Y")

        if update_reading(counter, value, record_time):
            await message.answer(
                f"✅ Показания сохранены!\n\n"
                f"🏭 Счётчик: {counter}\n"
                f"📅 Дата: {today}\n"
                f"⚡ Значение: {value:,.2f} кВт·ч\n"
                f"⏰ Время: {datetime.now().strftime('%H:%M:%S')}",
                reply_markup=get_main_menu()
            )
        else:
            await message.answer("❌ Ошибка при сохранении!", reply_markup=get_main_menu())

        await state.clear()

    except ValueError:
        await message.answer("❌ Введите число (например: 125.5)")


@dp.message(F.text == "📊 Статистика")
@dp.message(Command("stats"))
async def show_stats(message: types.Message):
    data = get_all_data()
    if not data:
        await message.answer("📊 Нет данных. Введите показания через '📝 Ввести показания'", reply_markup=get_main_menu())
        return

    text = "📊 СТАТИСТИКА ПОТРЕБЛЕНИЯ\n\n"
    total_all = 0
    for day in data[-10:]:
        text += f"📅 {day['date']}\n"
        if day.get('time'):
            text += f"   🕐 Время: {day['time'][11:16] if len(day['time']) > 11 else day['time']}\n"
        text += f"   ⚡ Сумма: {day['total']:,.2f} кВт·ч\n\n"
        total_all += day['total']

    text += f"💰 ИТОГО ЗА ВСЁ ВРЕМЯ: {total_all:,.2f} кВт·ч"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message(F.text == "📋 Все счётчики")
@dp.message(Command("counters"))
async def show_all_counters(message: types.Message):
    text = "📋 ВСЕ СЧЁТЧИКИ ПО ГРУППАМ:\n\n"
    for group, counters in COUNTER_GROUPS.items():
        text += f"📁 {group}:\n"
        for i, c in enumerate(counters, 1):
            text += f"   {i}. {c}\n"
        text += "\n"
        if len(text) > 3500:
            await message.answer(text)
            text = ""
    if text:
        await message.answer(text, reply_markup=get_main_menu())


@dp.message(F.text == "✅ Счётчики с данными")
@dp.message(Command("data_counters"))
async def show_counters_with_data(message: types.Message):
    counters_with_data = get_counters_with_data()
    if not counters_with_data:
        await message.answer("✅ Нет счётчиков с данными", reply_markup=get_main_menu())
        return

    text = "✅ СЧЁТЧИКИ С ПОКАЗАНИЯМИ:\n\n"
    for i, c in enumerate(counters_with_data, 1):
        text += f"{i}. {c}\n"
    await message.answer(text, reply_markup=get_main_menu())


@dp.message(F.text == "📁 Скачать Excel")
@dp.message(Command("file"))
async def send_excel_file(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        try:
            document = FSInputFile(EXCEL_FILE)
            await message.answer_document(
                document=document,
                caption=f"📊 Файл учёта энергоресурсов\n📅 {datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )
        except Exception as e:
            await message.answer(f"❌ Ошибка: {e}")
    else:
        await message.answer("❌ Файл ещё не создан")


@dp.message(F.text == "❓ Помощь")
async def help_button(message: types.Message):
    await help_command(message)


@dp.callback_query(F.data == "cancel")
async def cancel_action(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("❌ Отменено")
    await callback.message.answer("Главное меню:", reply_markup=get_main_menu())
    await callback.answer()


# ============ ЗАПУСК ============
async def main():
    print("🚀 Бот Энергоучёт запущен!")
    print(f"📁 Файл: {EXCEL_FILE}")
    print(f"🏭 Счётчиков: {len(ALL_COUNTERS)}")
    print("========================================")
    
    init_excel()
    ensure_today_exists()

    scheduler = AsyncIOScheduler(timezone="Europe/Moscow")
    scheduler.add_job(send_email_report, 'cron', day_of_week='mon', hour=12, minute=0)
    scheduler.start()
    print("⏰ Планировщик запущен: отчёт каждый понедельник в 12:00 МСК")
    print("========================================")

    await dp.start_polling(bot)


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)
