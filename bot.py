# gitADD https://github.com/stepannhl/mini_bot_telegram.git
import logging
# Импорт библиотек для бота
from random import randint

from openpyxl import load_workbook

from aiogram import Bot, types
from aiogram.utils import executor
from aiogram.dispatcher import Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware

import messages
from config import TOKEN

# Подключение логов
logging.basicConfig(format=u'%(filename)+13s [ LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
                    level=logging.DEBUG)
# инициализация бота через токен
bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())
# инициализация бд - эксель табличка
wb = load_workbook(filename='Films.xlsx', read_only=True)
sheet_ranges = wb['lst']

# комменты убрать при добавлении сериалов
buttons = ["Случайный фильм🎬",
           "Случайный фильм по категории🍕"
           # "Случайный Сериал🎞",
           # "Случайный сериал по категории🍪"
           ]
# задача кнопок ко всему боту
menu_buttons_f = ["Фантастика👩‍🚀",
                  "Драма💧",
                  "Комедия💃",
                  "Ужастик👻",
                  "Триллер🍿",
                  "Мультфильм🥺"]

menu_buttons_s = ["◀ Назад",
                  "Случайный сериал по категории🍪"]


# dry насколько это возможно
def msg_caption(i):
    string = ('✅ Название: ' + str(sheet_ranges[f'A{i}'].value)
              + '\n\n' +
              '🎫 ' + str(sheet_ranges[f'B{i}'].value))
    return string


# Задаем метку старта бота
@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    # инициализация keyboard и дальнейщее присваивание кнопок
    board = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    board.add(*buttons)
    # дергаем имя от аккаунта пишущего боту
    name = message.from_user.first_name
    await message.answer(messages.start(name), reply_markup=board)


# начало костяка бота
@dp.message_handler(content_types=['text'])
async def send_random_value(message: types.Message):
    # если бот увидел /back, он отображает экран главного меню по кнопкам
    if message.text == '/back':
        board = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
        board.add(*buttons)
        name = message.from_user.first_name
        await message.answer(messages.menu(name), reply_markup=board)
    # оформление работы одноименной кнопки
    elif message.text == 'Случайный фильм🎬':
        # обычный рандомайзер на размер всех фильмов в бд
        r = str(randint(1, 31))
        await bot.send_photo(chat_id=message.from_user.id,
                             # дергаем фотки из локального адреса
                             photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                             caption=msg_caption(r))
    # тк бот асинхронный, для простоты любое вхождение по меню мы обрабатываем соответствюще
    elif message.text in ('Случайный фильм по категории🍕',
                          'Фантастика👩‍🚀',
                          'Драма💧',
                          'Комедия💃',
                          'Ужастик👻',
                          "Триллер🍿",
                          "Мультфильм🥺"
                          ):
        # задаем и показываем кнопки 2 = сколько в ряду кнопок отображается
        board = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
        board.add(*menu_buttons_f)
        # при тапе на случайный фильм, мы выводим текстовку онбординга пользователя
        # в случае чего возврат по /back
        if message.text == 'Случайный фильм по категории🍕':
            await message.answer(messages.text_menu_f(), reply_markup=board)
        # аналогично точке /start
        if message.text == 'Фантастика👩‍🚀':
            r = str(randint(1, 5))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)
        elif message.text == 'Драма💧':
            r = str(randint(6, 11))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)
        elif message.text == 'Комедия💃':
            r = str(randint(12, 16))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)
        elif message.text == 'Ужастик👻':
            r = str(randint(17, 21))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)
        elif message.text == 'Триллер🍿':
            r = str(randint(22, 26))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)
        elif message.text == 'Мультфильм🥺':
            r = str(randint(27, 31))
            await bot.send_photo(chat_id=message.from_user.id,
                                 photo=open(f'C:/Users/stepa/Desktop/bot/bot start/pict/{r}.jpg', 'rb'),
                                 caption=msg_caption(r), reply_markup=board)

        # каждый раз по тапу на любую категорию, показываем текстовку из начала, для возможности к навигации
        if not message.text == 'Случайный фильм по категории🍕':
            await message.answer(messages.text_menu_f(), reply_markup=board)
    else:
        # обработка некорректных комманд
        await message.answer(messages.no_command())


# отключение
async def shutdown(dispatcher: Dispatcher):
    await dispatcher.storage.close()
    await dispatcher.storage.wait_closed()


# первичная инициализация
if __name__ == '__main__':
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    button_1 = types.KeyboardButton(
        text="/start")
    keyboard.add(
        button_1)
    executor.start_polling(dp, on_shutdown=shutdown)
