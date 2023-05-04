import asyncio
from aiogram import Bot, Dispatcher, types, executor
from aiogram.types.message import ContentTypes
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import StatesGroup, State
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, InputFile
import my_settings
import os
import bank_requester
from database import Database
import excel


class States(StatesGroup):
    timer: State = State()
    start_work: State = State()
    end_work: State = State()


storage = MemoryStorage()
bot = Bot(token=my_settings.bot_token)
dp = Dispatcher(bot, storage=storage)


@dp.message_handler(content_types=ContentTypes.DOCUMENT)
async def get_document(message: types.Message):
    file_id = message.document.file_id
    success, errors, number = await download_file(file_id, message.from_user.id)
    if (success, errors, number) == (False, False, False):
        await bot.send_message(message.from_user.id, 'Неверный формат документа, требуется файл .xlsx')
        return
    filename = excel.create_workbook(success, errors, number)
    caption = f'Всего запросов: {len(errors)+len(success)-2}\n' \
              f'Успешно: {len(success)-1}\n' \
              f'С ошибками: {len(errors)-1}\n'
    await bot.send_document(chat_id=message.from_user.id, document=InputFile(filename), caption=caption)
    await asyncio.sleep(10)
    os.remove(filename)


@dp.message_handler(lambda message: message.text == "Настройки" and message.from_user.id in my_settings.admins)
async def settings(message: types.Message):
    b1 = InlineKeyboardButton('Частота запросов', callback_data='timer')
    b2 = InlineKeyboardButton('Начало дня', callback_data='start')
    b3 = InlineKeyboardButton('Конец дня', callback_data='end')
    buttons = InlineKeyboardMarkup().add(b1).add(b2).add(b3)
    await message.answer("Выберите, что хотите изменить:", reply_markup=buttons)


@dp.message_handler(lambda message: message.text == "Статистика" and message.from_user.id in my_settings.admins)
async def settings(message: types.Message):
    data = await Database.get_statistics()
    for i in data:
        await bot.send_document(chat_id=message.from_user.id, document=InputFile(i.filename))
        os.remove(i.filename)
    string = f'Запросы к API банка отправляются раз в {my_settings.timer} секунд.\n' \
             f'Рабочий день длится с {my_settings.start_working_time} до {my_settings.end_working_time}.\n' \
             f'За сегодня запросов:\n{data[0]}' \
             f'За неделю:\n{data[1]}' \
             f'За месяц:\n{data[2]}'
    await message.answer(string)


@dp.message_handler(lambda message: message.from_user.id in my_settings.admins)
async def init_admin(message: types.Message):
    buttons = [[types.KeyboardButton(text="Настройки")],
              [types.KeyboardButton(text="Статистика")],]
    keyboard = types.ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)
    await message.answer("Вы являетесь администратором.", reply_markup=keyboard)


@dp.callback_query_handler(lambda c: c.data == 'timer' and c.from_user.id in my_settings.admins)
async def click_timer(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer(f"Запросы к API Альфа-банка отправляются раз в {my_settings.timer} секунд. Введите новый промежуток времени в секундах:")
    await state.set_state(States.timer)


@dp.callback_query_handler(lambda c: c.data == 'start' and c.from_user.id in my_settings.admins)
async def click_start_time(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer(f"Бот начинает отправку запросов в {my_settings.start_working_time}. Введите новое время начала работы:")
    await state.set_state(States.start_work)


@dp.callback_query_handler(lambda c: c.data == 'end' and c.from_user.id in my_settings.admins)
async def click_end_time(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer(f"Бот заканчивает отправку запросов в {my_settings.end_working_time}. Введите новое время окончания работы:")
    await state.set_state(States.end_work)


@dp.message_handler(state=States.timer)
async def change_timer(message: types.Message, state: FSMContext):
    await state.reset_state()
    text = message.text.strip()
    if not text.isdigit():
        await message.answer("Неправильно введено время.")
        return
    my_settings.timer = int(text)
    await message.answer("Задержка между отпрвкой запросов успешно изменена.")


@dp.message_handler(state=States.start_work)
async def change_start_time(message: types.Message, state: FSMContext):
    await state.reset_state()
    text = message.text.strip()
    time = text.split(':')
    is_digit = [i.isdigit() for i in time]
    if len(time) != 3 or not all(is_digit) or int(time[0]) >= 25 or int(time[1]) >= 60 or int(time[2]) >= 60:
        await message.answer("Неправильно введено время.")
        return
    my_settings.start_working_time = text
    await message.answer("Время начала работы успешно изменено.")


@dp.message_handler(state=States.end_work)
async def change_start_time(message: types.Message, state: FSMContext):
    await state.reset_state()
    text = message.text.strip()
    time = text.split(':')
    is_digit = [i.isdigit() for i in time]
    if len(time) != 3 or not all(is_digit) or int(time[0]) >= 25 or int(time[1]) >= 60 or int(time[2]) >= 60:
        await message.answer("Неправильно введено время.")
        return
    my_settings.end_working_time = text
    await message.answer("Время окончания работы успешно изменено.")


async def download_file(file_id, user_id):
    global count_files
    now_count = count_files
    count_files += 1
    
    file = await bot.get_file(file_id)
    file_path = file.file_path
    if not '.xlsx' in file_path:
        return False, False, False
    dest_path = os.path.join(excel.files_path, f'file{now_count}.xlsx')
    await bot.download_file(file_path, dest_path)
    await bot.send_message(user_id, 'Файл принят в работу.')
    success, errors = await bank_requester.parse_file(dest_path, now_count)
    await asyncio.sleep(10)
    os.remove(dest_path)
    return success, errors, now_count


async def get_count():
    global count_files
    count_files = await Database.init_db()


if __name__ == "__main__":
    count_files = 1
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.create_task(get_count())
    executor.start_polling(dp, skip_updates=True)

