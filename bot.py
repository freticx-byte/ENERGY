import asyncio
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters.command import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.types import FSInputFile
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
import os

bot = Bot(token="8976307638:AAEyUMxOzc5Wy7JSHThXxPV_v1bbazZRSYQ")
dp = Dispatcher()

EXCEL_FILE = "Ж учета энергоресурсов.xlsx"


class EnergyForm(StatesGroup):
    choosing_counter = State()
    entering_value = State()
    choosing_date = State()


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


def ensure_date_exists(date_str):
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        date_exists = False
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == date_str:
                date_exists = True
                break

        if not date_exists:
            ws.append([date_str, ""] + [0] * len(COUNTERS))
            wb.save(EXCEL_FILE)

        wb.close()
    except Exception as e:
        print(f"Ошибка: {e}")


def update_reading(counter_name, date_str, value, record_time):
    try:
        ensure_date_exists(date_str)

        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        headers = [str(cell.value) if cell.value else "" for cell in ws[1]]

        if counter_name not in headers:
            wb.close()
            return False

        col_idx = headers.index(counter_name) + 1

        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == date_str:
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


def get_future_dates(days=10):
    today = datetime.now()
    return [(today + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(days)]


async def counters_keyboard():
    builder = InlineKeyboardBuilder()

    for i, counter in enumerate(COUNTERS[:20]):
        text = counter[:35] if len(counter) > 35 else counter
        builder.button(text=text, callback_data=f"cnt_{i}")

    if len(COUNTERS) > 20:
        builder.button(text="📋 Ещё", callback_data="more_cnt")

    builder.button(text="📊 Статистика", callback_data="total_stats")
    builder.button(text="📋 Счётчики с данными", callback_data="counters_with_data")
    builder.button(text="📁 Скачать Excel", callback_data="download_file")
    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(1)
    return builder.as_markup()


def date_keyboard(dates):
    builder = InlineKeyboardBuilder()
    for date in dates:
        date_obj = datetime.strptime(date, "%Y-%m-%d")
        display_date = date_obj.strftime("%d.%m.%Y")
        builder.button(text=display_date, callback_data=f"dt_{date}")
    builder.button(text="❌ Отмена", callback_data="cancel")
    builder.adjust(2)
    return builder.as_markup()


@dp.message(Command("start"))
async def start(message: types.Message):
    init_excel()
    data = get_all_data()
    counters_with_data = get_counters_with_data()

    text = (
        "🏭 ЭНЕРГОУЧЁТ\n\n"
        f"📊 Статистика:\n"
        f"   Записей: {len(data)}\n"
        f"   Счётчиков всего: {len(COUNTERS)}\n"
        f"   Счётчиков с данными: {len(counters_with_data)}\n\n"
        "📌 Команды:\n"
        "   /add - Ввести показания\n"
        "   /stats - Статистика\n"
        "   /counters - Все счётчики\n"
        "   /data_counters - Счётчики с данными\n"
        "   /file - Скачать Excel\n"
        "   /help - Помощь\n\n"
        f"📅 Доступны даты: следующие 10 дней\n"
        f"⏰ Время записи: фиксируется автоматически"
    )

    await message.answer(text)


@dp.message(Command("help"))
async def help_command(message: types.Message):
    await message.answer(
        "ИНСТРУКЦИЯ\n\n"
        "/add - выбрать счётчик → дату → ввести показание\n"
        "/stats - общая статистика\n"
        "/counters - список всех счётчиков\n"
        "/data_counters - счётчики с показаниями\n"
        "/file - скачать Excel файл\n\n"
        "Даты: сегодня + следующие 9 дней\n"
        "Время записи сохраняется автоматически"
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
        await message.answer(text)


@dp.message(Command("data_counters"))
async def show_counters_with_data(message: types.Message):
    counters_with_data = get_counters_with_data()

    if not counters_with_data:
        await message.answer("Нет счётчиков с данными. Добавьте показания через /add")
        return

    text = "СЧЁТЧИКИ С ПОКАЗАНИЯМИ:\n\n"
    for i, c in enumerate(counters_with_data, 1):
        text += f"{i}. {c}\n"
        if len(text) > 3500:
            await message.answer(text)
            text = ""

    if text:
        await message.answer(text)


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
            await message.answer(f"Ошибка при отправке файла: {e}")
    else:
        await message.answer("Файл ещё не создан. Добавьте первые показания через /add")


@dp.message(Command("add"))
async def add_reading(message: types.Message, state: FSMContext):
    await message.answer("Выберите счётчик:", reply_markup=await counters_keyboard())
    await state.set_state(EnergyForm.choosing_counter)


@dp.message(Command("stats"))
async def show_stats(message: types.Message):
    data = get_all_data()

    if not data:
        await message.answer("Нет данных. Добавьте через /add")
        return

    text = "СТАТИСТИКА ПОТРЕБЛЕНИЯ\n\n"
    total_all = 0

    for day in data[-10:]:
        text += f"📅 {day['date']}\n"
        if day.get('time'):
            text += f"   Время: {day['time']}\n"
        text += f"   Сумма: {day['total']:,.2f} кВт·ч\n"
        text += f"   Счётчиков: {len(day['counters'])}\n\n"
        total_all += day['total']

    text += f"ИТОГО: {total_all:,.2f} кВт·ч"
    await message.answer(text)


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


@dp.callback_query(F.data == "counters_with_data")
async def counters_with_data_callback(callback: types.CallbackQuery):
    counters_with_data = get_counters_with_data()

    if not counters_with_data:
        await callback.message.edit_text("Нет счётчиков с данными")
        await callback.answer()
        return

    text = "СЧЁТЧИКИ С ПОКАЗАНИЯМИ:\n\n"
    for i, c in enumerate(counters_with_data, 1):
        text += f"{i}. {c}\n"

    await callback.message.edit_text(text)
    await callback.answer()


@dp.callback_query(F.data == "download_file")
async def download_file_callback(callback: types.CallbackQuery):
    if os.path.exists(EXCEL_FILE):
        try:
            document = FSInputFile(EXCEL_FILE)
            await callback.message.answer_document(
                document=document,
                caption=f"Файл учёта энергоресурсов\n{datetime.now().strftime('%d.%m.%Y %H:%M')}"
            )
        except Exception as e:
            await callback.message.answer(f"Ошибка: {e}")
    else:
        await callback.message.answer("Файл ещё не создан")
    await callback.answer()


@dp.callback_query(F.data.startswith("cnt_"))
async def counter_selected(callback: types.CallbackQuery, state: FSMContext):
    idx = int(callback.data.split("_")[1])

    if idx >= len(COUNTERS):
        await callback.answer("Ошибка")
        return

    selected = COUNTERS[idx]
    await state.update_data(selected_counter=selected)

    future_dates = get_future_dates(10)

    await callback.message.edit_text(
        f"Выбран счётчик: {selected}\n\nВыберите дату (доступны следующие 10 дней):",
        reply_markup=date_keyboard(future_dates)
    )
    await state.set_state(EnergyForm.choosing_date)
    await callback.answer()


@dp.callback_query(F.data.startswith("dt_"))
async def date_selected(callback: types.CallbackQuery, state: FSMContext):
    date_str = callback.data.replace("dt_", "")
    await state.update_data(selected_date=date_str)

    data = await state.get_data()
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    display_date = date_obj.strftime("%d.%m.%Y")

    await callback.message.edit_text(
        f"Введите показание\n\n"
        f"Счётчик: {data['selected_counter']}\n"
        f"Дата: {display_date}\n"
        f"Время записи: {datetime.now().strftime('%H:%M:%S')}\n\n"
        f"Введите значение в кВт·ч:"
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
        date = data['selected_date']
        record_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if update_reading(counter, date, value, record_time):
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            display_date = date_obj.strftime("%d.%m.%Y")

            await message.answer(
                f"✅ Показания сохранены!\n\n"
                f"Счётчик: {counter}\n"
                f"Дата: {display_date}\n"
                f"Значение: {value:,.2f} кВт·ч\n"
                f"Время записи: {record_time}\n\n"
                f"/stats - статистика\n"
                f"/data_counters - счётчики с данными\n"
                f"/file - скачать Excel"
            )
        else:
            await message.answer("Ошибка при сохранении!")

        await state.clear()

    except ValueError:
        await message.answer("Введите число (например: 125.5)")


@dp.callback_query(F.data == "total_stats")
async def total_stats_callback(callback: types.CallbackQuery):
    data = get_all_data()

    if not data:
        await callback.message.edit_text("Нет данных. /add - добавить")
        await callback.answer()
        return

    text = "ОБЩАЯ СТАТИСТИКА\n\n"
    total = sum(d['total'] for d in data)

    for day in data[:15]:
        text += f"{day['date']}: {day['total']:,.2f} кВт·ч\n"

    text += f"\nВсего: {total:,.2f} кВт·ч"
    await callback.message.edit_text(text)
    await callback.answer()


@dp.callback_query(F.data == "cancel")
async def cancel_action(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("Отменено")
    await callback.answer()


async def main():
    print("Бот Энергоучёт запущен!")
    print(f"Файл: {EXCEL_FILE}")
    print(f"Счётчиков: {len(COUNTERS)}")

    init_excel()
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())