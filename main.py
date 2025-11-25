import telebot
import datetime
import openpyxl
import os
from datetime import datetime, timedelta
from telebot import types
from dotenv import load_dotenv


load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
SCHEDULE_FILE = os.getenv("SCHEDULE_FILE", "schedule.xlsx")

bot = telebot.TeleBot(BOT_TOKEN)

# Переменная для отслеживания текущего раздела
user_states = {}


# Функция для чтения расписания из Excel
def get_schedule(day, parity):
    try:
        # Загружаем файл Excel
        wb = openpyxl.load_workbook('schedule.xlsx')
        sheet = wb.active

        schedule = []

        # Проходим по всем строкам (начиная со 2-й, так как 1-я - заголовки)
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # row[0] - день недели, row[1] - четность, row[2] - номер пары
            if row[0] == day and row[1] == parity and row[2] is not None and row[2] != "":
                # Преобразуем время в минуты от начала дня
                time_str = row[5]  # Время из столбца F
                minutes_to_first_pair = convert_time_to_minutes(time_str)


                subject = {
                    'pair_number': row[2],
                    'name': row[3],
                    'type': row[4],
                    'time': row[5],
                    'location': row[6],
                    'minutes_to_first_pair': minutes_to_first_pair,
                }

                schedule.append(subject)

        # Сортируем по номеру пары
        schedule.sort(key=lambda x: x['pair_number'])
        return schedule

    except Exception as e:
        print(f"Ошибка при чтении Excel: {e}")
        return []

# --- Фиксированное время подъёма и выхода ---
TIME_TABLE = {
    'МП': {
        1: {'wake': '06:20', 'leave': '08:00'},
        2: {'wake': '08:00', 'leave': '09:40'},
        3: {'wake': '10:00', 'leave': '11:40'},
        4: {'wake': '11:00', 'leave': '13:20'},
        5: {'wake': '12:00', 'leave': '15:20'},
    },
    'СТ': {
        1: {'wake': '05:50', 'leave': '07:30'},
        2: {'wake': '07:30', 'leave': '09:00'},
        3: {'wake': '09:20', 'leave': '11:00'},
        4: {'wake': '', 'leave': ''},
        5: {'wake': '', 'leave': ''},
    },
    'ПВ': {
        1: {'wake': '06:30', 'leave': '08:10'},
        2: {'wake': '08:20', 'leave': '10:00'},
        3: {'wake': '10:20', 'leave': '12:00'},
        4: {'wake': '12:00', 'leave': '13:30'},
        5: {'wake': '', 'leave': ''},
    }
}

def get_wake_leave(pair_number, location):
    """Возвращает время подъема и выхода по номеру пары и корпусу"""
    if 'МП' in location:
        campus = 'МП'
    elif 'С-20' in location:
        campus = 'СТ'
    elif 'В-78' in location:
        campus = 'ПВ'
    else:
        campus = None

    if campus and pair_number in TIME_TABLE[campus]:
        return TIME_TABLE[campus][pair_number]['wake'], TIME_TABLE[campus][pair_number]['leave']
    else:
        return None, None


# Функция для преобразования времени в минуты от начала дня
def convert_time_to_minutes(time_str):
    if not time_str:
        return 540  # 9:00 по умолчанию

    try:
        # Разбираем время вида "14:20 -15:50" или "09:00 - 10:30"
        start_time_str = time_str.split('-')[0].strip()
        hours, minutes = map(int, start_time_str.split(':'))
        return hours * 60 + minutes
    except:
        return 540  # 9:00 по умолчанию


# Функция для расчета времени подъема
def calculate_wake_up_time(schedule, day_type):
    if not schedule:
        return "На выбранный день пар нет! Можно поспать подольше 😴"

    first_pair = min(schedule, key=lambda x: x['pair_number'])
    pair = first_pair['pair_number']
    location = first_pair['location']
    wake, leave = get_wake_leave(pair, location)

    # Если данных нет в словаре — подстрахуемся
    if not wake:
        return f"Не найдено время подъёма для пары {pair} ({location})"

    text = (
        f"📅 {day_type.capitalize()}\n\n"
        f"⏰ Встать: {wake}\n"
        f"🚪 Выйти: {leave}\n\n"
        f"📚 Первая пара: {first_pair['time']} — {first_pair['name']} ({first_pair['type']})\n"
        f"🏫 Место: {location}\n"
    )

    return text



# Функция для форматирования расписания в красивый текст
def format_schedule(schedule, day_name, parity_name):
    if not schedule:
        return f"📅 {day_name} ({parity_name} неделя)\n\nПар нет! 🎉"

    result = f"📅 {day_name} ({parity_name} неделя)\n\n"

    for subject in schedule:
        result += f"🕒 {subject['time']}\n"
        result += f" {subject['name']}\n"
        result += f" {subject['location']}"
        result += f" {subject['type']}\n"
        result += "─" * 30 + "\n"

    return result


# Функция для получения русского названия дня недели
def get_russian_day_name(weekday):
    days = {
        0: "Понедельник",
        1: "Вторник",
        2: "Среда",
        3: "Четверг",
        4: "Пятница",
        5: "Суббота",
        6: "Воскресенье"
    }
    return days.get(weekday, "Неизвестный день")


# Функция для получения названия четности
def get_parity_name(parity):
    return "чётная" if parity == 0 else "нечётная"


@bot.message_handler(commands=['start'])
def main(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)

    btn1 = types.KeyboardButton('Встать')
    markup.row(btn1)

    btn2 = types.KeyboardButton('Расписание')
    btn3 = types.KeyboardButton('coming soon...')
    markup.row(btn2, btn3)
    file = open('./tosya.jpg', 'rb')
    bot.send_photo(message.chat.id, file, reply_markup=markup, caption="<b>Привет</b>, я тут чтобы помочь тебе",
                   parse_mode='HTML')


@bot.message_handler(content_types=['text'])
def text_button(message):
    if message.text == 'Встать':
        # Создаем клавиатуру для Встать
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('Сегодня')
        btn2 = types.KeyboardButton('Завтра')
        markup.row(btn1, btn2)
        btn_back = types.KeyboardButton('Назад')
        markup.row(btn_back)

        # Сохраняем состояние пользователя
        user_states[message.chat.id] = 'встать'
        bot.send_message(message.chat.id, "Сегодня или завтра?", reply_markup=markup)

    elif message.text == 'Расписание':
        # Создаем клавиатуру для расписания
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('Сегодня')
        btn2 = types.KeyboardButton('Завтра')
        markup.row(btn1, btn2)
        btn_back = types.KeyboardButton('Назад')
        markup.row(btn_back)

        # Сохраняем состояние пользователя
        user_states[message.chat.id] = 'расписание'
        bot.send_message(message.chat.id, "Выбери день:", reply_markup=markup)


    elif message.text == 'Сегодня':
        # Проверяем из какого раздела пришло нажатие
        current_state = user_states.get(message.chat.id)
        # Получаем данные для сегодня
        today_date = datetime.now()
        today_weekday = today_date.weekday()  # 0 = понедельник
        parity = 1 - (today_date.isocalendar()[1] % 2)  # чётность недели

        if current_state == 'встать':
            # Получаем расписание для расчета времени подъема
            schedule = get_schedule(today_weekday, parity)
            wake_up_info = calculate_wake_up_time(schedule, "сегодня")
            bot.send_message(message.chat.id, wake_up_info)


        elif current_state == 'расписание':
            # Получаем расписание из Excel
            schedule = get_schedule(today_weekday, parity)
            day_name = get_russian_day_name(today_weekday)
            parity_name = get_parity_name(parity)
            formatted_schedule = format_schedule(schedule, day_name, parity_name)
            bot.send_message(message.chat.id, formatted_schedule)


    elif message.text == 'Завтра':
        # Проверяем из какого раздела пришло нажатие
        current_state = user_states.get(message.chat.id)
        # Получаем данные для завтра
        tomorrow_date = datetime.now() + timedelta(days=1)
        tomorrow_weekday = tomorrow_date.weekday()  # 0 = понедельник
        parity = 1 - (tomorrow_date.isocalendar()[1] % 2)  # чётность недели

        if current_state == 'встать':
            # Получаем расписание для расчета времени подъема
            schedule = get_schedule(tomorrow_weekday, parity)
            wake_up_info = calculate_wake_up_time(schedule, "завтра")
            bot.send_message(message.chat.id, wake_up_info)


        elif current_state == 'расписание':
            # Получаем расписание из Excel
            schedule = get_schedule(tomorrow_weekday, parity)
            day_name = get_russian_day_name(tomorrow_weekday)
            parity_name = get_parity_name(parity)
            formatted_schedule = format_schedule(schedule, day_name, parity_name)
            bot.send_message(message.chat.id, formatted_schedule)

    elif message.text == 'Назад':
        # Возвращаем главное меню
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        btn1 = types.KeyboardButton('Встать')
        markup.row(btn1)
        btn2 = types.KeyboardButton('Расписание')
        btn3 = types.KeyboardButton('coming soon...')
        markup.row(btn2, btn3)

        # Очищаем состояние пользователя
        user_states[message.chat.id] = 'главное'
        bot.send_message(message.chat.id, "Главное меню:", reply_markup=markup)

    elif message.text == 'coming soon...':
        bot.send_message(message.chat.id, "В разработке")

    else:
        bot.send_message(message.chat.id, 'не понял')


bot.polling(none_stop=True)