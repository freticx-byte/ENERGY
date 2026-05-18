import asyncio
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters.command import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder, ReplyKeyboardBuilder
from aiogram.types import FSInputFile, ReplyKeyboardMarkup, KeyboardButton
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
import os

bot = Bot(token="8976307638:AAEyUMxOzc5Wy7JSHThXxPV_v1bbazZRSYQ")
dp = Dispatcher()

EXCEL_FILE = "Ж учета энергоресурсов.xlsx"


class EnergyForm(StatesGroup):
    choosing_counter = State()
    entering_value = State()


# ГЛАВНОЕ МЕНЮ
def get_main_menu():
    builder = ReplyKeyboardBuilder()
    builder.add(KeyboardButton(text="📊 Статистика"))
    builder.add(KeyboardButton(text="📝 Ввести показания"))
    builder.add(KeyboardButton(text="📋 Все счётчики"))
    builder.add(KeyboardButton(text="✅ Счётчики с данными"))
    builder.add(KeyboardButton(text="📁 Скачать Excel"))
    builder.add(KeyboardButton(text="❓ Помощь"))
    builder.adjust(2)
    return builder.as_markup(resize_keyboard=True)


# СПИСОК СЧЁТЧИКОВ
COUNTERS = [
    "ЦРП В1", "ЦРП В2", "ЦРП ТП1, СШ1", "ЦРП ТП1, СШ2",
    "ЦРП ТП2, СШ1", "ЦРП ТП2, СШ2", "ЦРП ТП3, СШ1", "ЦРП ТП3, СШ2",
    'ЦРП КСО "Радуга"', "ТП1 ГРЩ. ППУ В1", "ТП1 ГРЩ. ППУ В2",
    "ТП2 ГРЩ В1", "ТП2 ГРЩ В2", "ТП1 РУ0,4 СГП В1", "ТП1 РУ0,4 СГП В2",
    "ТП2 РУ0,4 В1", "ТП2 РУ0,4 В2", "ТП3 РУ0,4 Элеватор В1",
    "ТП3 РУ0,4 Элеватор В2", "ТП3 Лузговая В1", "ТП3 Лузговая В2",
    "ТП3 Элеваторный В1", "ТП3 Элеваторный В2", "ТП Луговская",
    "ТП-4 СТЗ", "Насосная В1", "Насосная В2", "ЛОС В1", "ЛОС В2",
    "КНС В1", "КНС В2", "Склад газации", "Теплосети ИТП",
    "Газовая котельная №1", "Газовая котельная №2", "Временно ТП-3"
]


def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время записи"] + COUNTERS)
        wb.save(EXCEL_FILE)


def get_today_str():
    return datetime.now().strftime("%Y-%m-%d")


def ensure_today_exists():
    """Проверяет, есть ли сегодняшняя дата в файле"""
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
            ws.append([today, ""] + [0] * len(COUNTERS))
            wb.save(EXCEL_FILE)

        wb.close()
    except Exception as e:
        print(f"Ошибка: {e}")


def update_reading(counter_name, value, record_time):
    """Обновляет показание счётчика на сегодняшнюю дату"""
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


async def counters_keyboard():
    builder = InlineKeyboardBuilder()

    for i, counter in enumerate(COUNTERS[:20]):
        text = counter[:35] if len(counter) > 35 else counter
        builder.button(text=text, callback_data=f"cnt_{i}")

    if len(COUNTERS) > 20:
        builder.button(text="📋 Ещё", callback_data="more_cnt")

    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(1)
    return builder.as_markup()


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
        f"🏭 Счётчиков: {len(COUNTERS)}\n"
        f"✅ Счётчиков с данными: {len(counters_with_data)}\n\n"
        f"💡 Нажмите кнопку '📝 Ввести показания'\n"
        f"   или используйте команды:"
    )

    await message.answer(text, reply_markup=get_main_menu())


@dp.message(F.text == "📊 Статистика")
async def stats_button(message: types.Message):
    await show_stats(message)


@dp.message(F.text == "📝 Ввести показания")
async def add_button(message: types.Message, state: FSMContext):
    await message.answer("Выберите счётчик:", reply_markup=await counters_keyboard())
    await state.set_state(EnergyForm.choosing_counter)


@dp.message(F.text == "📋 Все счётчики")
async def all_counters_button(message: types.Message):
    await show_all_counters(message)


@dp.message(F.text == "✅ Счётчики с данными")
async def counters_with_data_button(message: types.Message):
    await show_counters_with_data(message)


@dp.message(F.text == "📁 Скачать Excel")
async def file_button(message: types.Message):
    await send_excel_file(message)


@dp.message(F.text == "❓ Помощь")
async def help_button(message: types.Message):
    await help_command(message)


@dp.message(Command("help"))
async def help_command(message: types.Message):
    await message.answer(
        "ИНСТРУКЦИЯ\n\n"
        "📝 Ввести показания:\n"
        "   1. Нажмите кнопку '📝 Ввести показания'\n"
        "   2. Выберите счётчик из списка\n"
        "   3. Введите показание (число в кВт·ч)\n\n"
        "📊 Статистика - показывает общее потребление\n"
        "📁 Скачать Excel - получить файл с данными\n\n"
        f"📅 Все показания записываются на сегодняшнюю дату: {datetime.now().strftime('%d.%m.%Y')}",
        reply_markup=get_main_menu()
    )


@dp.message(Command("counters"))
async def show_all_counters(message: types.Message):
    text = "ВСЕ СЧЁТЧИКИ:\n\n"
    for i, c in enumerate(COUNTERS, 1):
        text += f"{i}. {c}\n"
        if len(text) > 3500:
            await message.answer(text)
            text = ""

    if text:
        await message.answer(text, reply_markup=get_main_menu())


@dp.message(Command("data_counters"))
async def show_counters_with_data(message: types.Message):
    counters_with_data = get_counters_with_data()

    if not counters_with_data:
        await message.answer("Нет счётчиков с данными", reply_markup=get_main_menu())
        return

    text = "СЧЁТЧИКИ С ПОКАЗАНИЯМИ:\n\n"
    for i, c in enumerate(counters_with_data, 1):
        text += f"{i}. {c}\n"

    await message.answer(text, reply_markup=get_main_menu())


@dp.message(Command("file"))
async def send_excel_file(message: types.Message):
    if os.path.exists(EXCEL_FILE):
        try:
            document = FSInputFile(EXCEL_FILE)
            await message.answer_document(
                document=document,
                caption=f"Файл учёта энергоресурсов\n{datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )
        except Exception as e:
            await message.answer(f"Ошибка: {e}")
    else:
        await message.answer("Файл ещё не создан")


@dp.message(Command("add"))
async def add_reading(message: types.Message, state: FSMContext):
    await message.answer("Выберите счётчик:", reply_markup=await counters_keyboard())
    await state.set_state(EnergyForm.choosing_counter)


@dp.message(Command("stats"))
async def show_stats(message: types.Message):
    data = get_all_data()

    if not data:
        await message.answer("Нет данных. Введите показания через кнопку '📝 Ввести показания'",
                             reply_markup=get_main_menu())
        return

    text = "СТАТИСТИКА ПОТРЕБЛЕНИЯ\n\n"
    total_all = 0

    # Показываем последние 10 дней (или все, если меньше)
    for day in data[-10:]:
        text += f"📅 {day['date']}\n"
        if day.get('time'):
            text += f"   🕐 Время: {day['time'][11:16] if len(day['time']) > 11 else day['time']}\n"
        text += f"   ⚡ Сумма: {day['total']:,.2f} кВт·ч\n"
        text += f"   📊 Счётчиков: {len(day['counters'])}\n\n"
        total_all += day['total']

    text += f"💰 ИТОГО за всё время: {total_all:,.2f} кВт·ч"
    await message.answer(text, reply_markup=get_main_menu())


@dp.callback_query(F.data == "more_cnt")
async def more_counters(callback: types.CallbackQuery):
    builder = InlineKeyboardBuilder()

    for i, counter in enumerate(COUNTERS[20:40], 20):
        text = counter[:35] if len(counter) > 35 else counter
        builder.button(text=text, callback_data=f"cnt_{i}")

    builder.button(text="🔙 Назад", callback_data="back_main")
    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(1)

    await callback.message.edit_text("Выберите счётчик:", reply_markup=builder.as_markup())
    await callback.answer()


@dp.callback_query(F.data == "back_main")
async def back_main(callback: types.CallbackQuery):
    await callback.message.edit_text("Выберите счётчик:", reply_markup=await counters_keyboard())
    await callback.answer()


@dp.callback_query(F.data.startswith("cnt_"))
async def counter_selected(callback: types.CallbackQuery, state: FSMContext):
    idx = int(callback.data.split("_")[1])

    if idx >= len(COUNTERS):
        await callback.answer("Ошибка")
        return

    selected = COUNTERS[idx]
    await state.update_data(selected_counter=selected)

    today = datetime.now().strftime("%d.%m.%Y")

    await callback.message.edit_text(
        f"✅ Выбран счётчик: {selected}\n\n"
        f"📅 Дата записи: {today} (сегодня)\n"
        f"⏰ Время будет записано автоматически\n\n"
        f"✏️ Введите показание в кВт·ч:"
    )
    await state.set_state(EnergyForm.entering_value)
    await callback.answer()


@dp.message(EnergyForm.entering_value)
async def value_entered(message: types.Message, state: FSMContext):
    try:
        value = float(message.text.replace(",", "."))
        if value < 0:
            await message.answer("Введите положительное число")
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
                f"⏰ Время: {datetime.now().strftime('%H:%M:%S')}\n\n"
                f"📊 /stats - статистика\n"
                f"📁 /file - скачать Excel",
                reply_markup=get_main_menu()
            )
        else:
            await message.answer("❌ Ошибка при сохранении!", reply_markup=get_main_menu())

        await state.clear()

    except ValueError:
        await message.answer("❌ Введите число (например: 125.5)")


@dp.callback_query(F.data == "cancel")
async def cancel_action(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("❌ Отменено", reply_markup=get_main_menu())
    await callback.answer()


async def main():
    print("Бот Энергоучёт запущен!")
    print(f"Файл: {EXCEL_FILE}")
    print(f"Счётчиков: {len(COUNTERS)}")
    print("Дата для записи: ТОЛЬКО ТЕКУЩАЯ")

    init_excel()
    ensure_today_exists()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())